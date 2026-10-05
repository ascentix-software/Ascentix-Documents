using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace Ascentix.Documents.SdkTests;

/// <summary>
/// Wraps MemoryService with direct Update/Delete so plug-in step and runtime-table writes can be
/// observed. Runtime row CAS updates still flow through MemoryService.Execute.
/// </summary>
internal sealed class StepService : IOrganizationService
{
    internal readonly DurableWorkerTests.MemoryService Memory = new();
    internal readonly List<string> Writes = new();
    internal bool IgnoreStepWrites { get; set; }
    internal HashSet<string> MissingTables { get; set; } = new();

    /// <summary>Users that real Dataverse would fault on (Guid.Empty always does).</summary>
    internal HashSet<Guid> UnknownUsers { get; set; } = new();

    /// <summary>When true the worker holds no prvReadAsyncOperation.</summary>
    internal bool WorkerLacksSystemJobs { get; set; }

    /// <summary>Worker Read depth per table; a table not listed is held at Global, null = not held.</summary>
    internal Dictionary<string, PrivilegeDepth?> WorkerReadDepth { get; set; } = new();

    internal static Guid ReadPrivilegeId(string table) =>
        new Guid(
            System
                .Security.Cryptography.MD5.Create()
                .ComputeHash(System.Text.Encoding.UTF8.GetBytes("read:" + table))
        );

    internal static readonly Guid SystemJobsPrivilegeId = Guid.Parse(
        "00000000-0000-0000-0000-00000000a5c0"
    );

    public Guid Create(Entity entity)
    {
        Writes.Add("Create " + entity.LogicalName);
        return Memory.Create(entity);
    }

    public void Update(Entity entity)
    {
        Writes.Add("Update " + entity.LogicalName);
        if (IgnoreStepWrites && entity.LogicalName == "sdkmessageprocessingstep")
            return;
        foreach (var pair in entity.Attributes)
            Memory.Rows[entity.Id][pair.Key] = pair.Value;
    }

    public void Delete(string name, Guid id)
    {
        Writes.Add("Delete " + name);
        Memory.Delete(name, id);
    }

    public OrganizationResponse Execute(OrganizationRequest request)
    {
        if (request is RetrieveEntityRequest entity)
        {
            if (MissingTables.Contains(entity.LogicalName))
                throw new InvalidOperationException(
                    "Entity " + entity.LogicalName + " does not exist."
                );
            var metadata = new EntityMetadata { LogicalName = entity.LogicalName };
            typeof(EntityMetadata)
                .GetProperty("Privileges")!
                .SetValue(metadata, new[] { ReadPrivilege(entity.LogicalName) }, null);
            var response = new RetrieveEntityResponse();
            response.Results["EntityMetadata"] = metadata;
            return response;
        }
        if (request is RetrieveUserPrivilegesRequest user)
        {
            if (user.UserId == Guid.Empty || UnknownUsers.Contains(user.UserId))
                throw new InvalidOperationException("SystemUser Does Not Exist");
            var held = new List<RolePrivilege>();
            if (!WorkerLacksSystemJobs)
                held.Add(
                    new RolePrivilege((int)PrivilegeDepth.Basic, SystemJobsPrivilegeId, Guid.Empty)
                    {
                        PrivilegeName = "prvReadAsyncOperation",
                    }
                );
            foreach (var table in Tables)
            {
                var depth = WorkerReadDepth.TryGetValue(table, out var d)
                    ? d
                    : PrivilegeDepth.Global;
                if (depth.HasValue)
                    held.Add(
                        new RolePrivilege((int)depth.Value, ReadPrivilegeId(table), Guid.Empty)
                    );
            }
            var privileges = new RetrieveUserPrivilegesResponse();
            privileges.Results["RolePrivileges"] = held.ToArray();
            return privileges;
        }
        return Memory.Execute(request);
    }

    public Entity Retrieve(string name, Guid id, ColumnSet columns) =>
        Memory.Retrieve(name, id, columns);

    public EntityCollection RetrieveMultiple(QueryBase query)
    {
        // Dataverse faults when an EntityName condition names a table that no longer exists in metadata.
        if (
            query is QueryExpression q
            && q.EntityName == "sdkmessagefilter"
            && q.Criteria.Conditions.Any(c =>
                c.AttributeName == "primaryobjecttypecode"
                && c.Values.Any(v => v is string s && MissingTables.Contains(s))
            )
        )
            throw new InvalidOperationException("simulated EntityName conversion fault");
        return Memory.RetrieveMultiple(query);
    }

    public void Associate(string n, Guid id, Relationship r, EntityReferenceCollection e) =>
        throw new NotSupportedException();

    public void Disassociate(string n, Guid id, Relationship r, EntityReferenceCollection e) =>
        throw new NotSupportedException();

    private static SecurityPrivilegeMetadata ReadPrivilege(string table)
    {
        var privilege = (SecurityPrivilegeMetadata)
            System.Runtime.Serialization.FormatterServices.GetUninitializedObject(
                typeof(SecurityPrivilegeMetadata)
            );
        foreach (
            var pair in new (string, object)[]
            {
                ("PrivilegeId", ReadPrivilegeId(table)),
                ("PrivilegeType", PrivilegeType.Read),
            }
        )
            typeof(SecurityPrivilegeMetadata)
                .GetProperty(pair.Item1)!
                .GetSetMethod(true)!
                .Invoke(privilege, new[] { pair.Item2 });
        return privilege;
    }

    private IEnumerable<string> Tables =>
        Memory
            .Rows.Values.Where(r => r.LogicalName == "sdkmessagefilter")
            .Select(r => r.GetAttributeValue<string>("primaryobjecttypecode"))
            .Distinct();

    internal int StepWrites => Writes.FindAll(w => w.EndsWith("sdkmessageprocessingstep")).Count;
}
