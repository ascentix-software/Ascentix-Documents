using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
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

    public OrganizationResponse Execute(OrganizationRequest request) => Memory.Execute(request);

    public Entity Retrieve(string name, Guid id, ColumnSet columns) =>
        Memory.Retrieve(name, id, columns);

    public EntityCollection RetrieveMultiple(QueryBase query) => Memory.RetrieveMultiple(query);

    public void Associate(string n, Guid id, Relationship r, EntityReferenceCollection e) =>
        throw new NotSupportedException();

    public void Disassociate(string n, Guid id, Relationship r, EntityReferenceCollection e) =>
        throw new NotSupportedException();

    internal int StepWrites => Writes.FindAll(w => w.EndsWith("sdkmessageprocessingstep")).Count;
}
