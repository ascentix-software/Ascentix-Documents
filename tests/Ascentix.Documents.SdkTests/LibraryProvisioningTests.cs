using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class LibraryProvisioningTests
{
    [Fact]
    public void CreatesLibraryBoundaryAndOwnerGrantThenQueuesInitialTeamSync()
    {
        var f = new Fixture();
        var queued = f.Queue();
        var result = f.Run(queued.Key);
        Assert.Equal("AccessPending", result.Status);
        Assert.Equal(3, f.Posts.Count);
        Assert.Contains(f.Posts, p => p == "_api/web/lists");
        Assert.Contains(
            f.Posts,
            p =>
                p.EndsWith(
                    "breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)",
                    StringComparison.Ordinal
                )
        );
        Assert.Contains(
            f.Posts,
            p =>
                p.EndsWith("addroleassignment(principalid=7,roledefid=5)", StringComparison.Ordinal)
        );
        var catalog = Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
        Assert.True(catalog.GetAttributeValue<bool>("asx_approved"));
        Assert.False(catalog.GetAttributeValue<bool>("asx_policyapplied"));
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "sharepointdocumentlocation");
        var registration = f.Store.Require<TeamRegistration>(
            "asx_teamregistration",
            "team:" + f.Team.ToString("N")
        );
        Assert.True(registration.Value.Enabled);
        var policy = f.Store.Require<PolicyDocument>(
            "asx_policy",
            "policy:" + catalog.Id.ToString("N")
        );
        Assert.Equal(f.Team, Assert.Single(policy.Value.Queued).TeamId);
        Assert.NotNull(policy.Value.OperationKey);
        var operation = f.Store.Require<SecurityOperation>(
            "asx_operation",
            policy.Value.OperationKey!
        );
        operation.Value.Status = "Applied";
        f.Store.Save(operation);
        Assert.Equal("Ready", f.Worker.Inspect(queued.Key).Status);
        Assert.Equal(queued.Key, f.Queue().Key);
        Assert.Equal(3, f.Posts.Count);
    }

    [Fact]
    public void ExistingSameNameLibraryIsNeverChangedByCreate()
    {
        var f = new Fixture { Exists = true };
        var result = f.Run(f.Queue().Key);
        Assert.Equal("Blocked", result.Status);
        Assert.Empty(f.Posts);
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
    }

    [Fact]
    public void UnknownCreationIsQuarantinedWithoutDeletingOrAdoptingContent()
    {
        var f = new Fixture { UnknownCreate = true };
        var queued = f.Queue();
        Assert.Equal("Quarantined", f.Run(queued.Key).Status);
        Assert.Single(f.Posts);
        Assert.Equal("ExternalUnknown", f.Worker.Inspect(queued.Key).Status);
        Assert.Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
            f.Worker.Execute(new WorkerRequest { Command = "Retry", Key = queued.Key }, true)
        );
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
    }

    [Fact]
    public void UnknownLibraryRecoveryRequiresOriginalResponseThenResumesWithoutRecreating()
    {
        var f = new Fixture { UnknownCreate = true };
        var queued = f.Queue();
        f.Run(queued.Key);
        var claim = f.Store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(f.Service, queued.Key)
        );
        claim.Value.LeaseUntilUtc = DateTime.UtcNow.AddMinutes(-1);
        f.Store.Save(claim);
        var request = new WorkerRequest
        {
            Key = queued.Key,
            RunId = claim.Value.RunId!,
            Token = claim.Value.Token,
            Evidence =
                "Original run terminated; successful create response recovered from run history.",
        };
        var recovery = new WorkerCoordinator(f.Service);
        Assert.Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
            f.Service.Transaction(() => recovery.PermitRecovery(request, true))
        );
        request.ResponseBody = JsonWire.Write(
            new ODataEnvelope<CreatedLibrary>
            {
                Data = new CreatedLibrary { Id = f.List, Title = "Different" },
            }
        );
        Assert.Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
            f.Service.Transaction(() => recovery.PermitRecovery(request, true))
        );
        request.ResponseBody = JsonWire.Write(
            new ODataEnvelope<CreatedLibrary>
            {
                Data = new CreatedLibrary { Id = f.List, Title = "Documents" },
            }
        );
        Assert.Equal(
            "RecoveryPermitted",
            f.Service.Transaction(() => recovery.PermitRecovery(request, true)).Status
        );
        Assert.Equal("AccessPending", f.Run(queued.Key).Status);
        Assert.Single(f.Posts, p => p == "_api/web/lists");
        Assert.DoesNotContain(f.Posts, p => p.Contains("delete"));
    }

    private sealed class Fixture
    {
        public readonly DurableWorkerTests.MemoryService Service =
            new DurableWorkerTests.MemoryService();
        public DocumentStore Store => new DocumentStore(Service);
        public LibraryProvisioning Worker => new LibraryProvisioning(Service);
        public Guid Site = Guid.NewGuid(),
            Native = Guid.NewGuid(),
            Web = Guid.NewGuid(),
            List = Guid.NewGuid(),
            Root = Guid.NewGuid(),
            Team = Guid.NewGuid();
        public bool Exists,
            Unique,
            OwnerAccess,
            UnknownCreate;
        public List<string> Posts = new List<string>();

        public Fixture()
        {
            Service.Seed(
                new Entity("asx_site", Site)
                {
                    ["asx_nativeid"] = new EntityReference("sharepointsite", Native),
                    ["asx_webid"] = Web.ToString("D"),
                    ["asx_url"] = "https://example.sharepoint.com/sites/test",
                    ["asx_approved"] = true,
                }
            );
            Service.Seed(
                new Entity("team", Team)
                {
                    ["teamtype"] = new OptionSetValue(0),
                    ["isdefault"] = false,
                }
            );
        }

        public CatalogResult Queue() =>
            Service.Transaction(() =>
                Worker.Queue(
                    new CatalogRequest
                    {
                        SiteId = Site,
                        Name = "Documents",
                        RequestId = Guid.NewGuid(),
                        Entries = new[]
                        {
                            new PolicyEntry { TeamId = Team, Access = "Read" },
                        },
                    }
                )
            );

        private static string Body<T>(T data) =>
            JsonWire.Write(new ODataEnvelope<T> { Data = data });

        public WorkerResult Run(string key)
        {
            var work = Call("Claim", new WorkerResult { Key = key });
            for (int i = 0; i < 30; i++)
            {
                if (work.Status == "ReadyToCreate")
                {
                    work = Call("PrepareCreate", work);
                    continue;
                }
                if (work.Status == "Create")
                {
                    Posts.Add(work.Http!.RelativeUri);
                    int status = 200;
                    string body = "{}";
                    if (work.Http.RelativeUri == "_api/web/lists")
                    {
                        Exists = true;
                        status = UnknownCreate ? 0 : 201;
                        body = Body(new CreatedLibrary { Id = List, Title = "Documents" });
                    }
                    else if (work.Http.RelativeUri.Contains("breakroleinheritance"))
                        Unique = true;
                    else if (work.Http.RelativeUri.Contains("addroleassignment"))
                        OwnerAccess = true;
                    else
                        throw new Exception("Unexpected mutation");
                    work = Call("CreateResponse", work, body, status);
                    continue;
                }
                if (work.Status != "Read")
                    return work;
                string response;
                int code = 200;
                switch (work.ProbeKind)
                {
                    case "Owner":
                        response = Body(new SiteGroup { Id = 7, Type = 8 });
                        break;
                    case "Roles":
                        response = Body(
                            new ODataRows<SecurityRoleObservation>
                            {
                                Rows = new[]
                                {
                                    new SecurityRoleObservation
                                    {
                                        Id = 2,
                                        Type = 2,
                                        Permissions = new PermissionMask
                                        {
                                            High = "176",
                                            Low = "138612833",
                                        },
                                    },
                                    new SecurityRoleObservation
                                    {
                                        Id = 3,
                                        Type = 3,
                                        Permissions = new PermissionMask
                                        {
                                            High = "432",
                                            Low = "1011028719",
                                        },
                                    },
                                    new SecurityRoleObservation
                                    {
                                        Id = 5,
                                        Type = 5,
                                        Permissions = new PermissionMask
                                        {
                                            High = "2147483647",
                                            Low = "4294967295",
                                        },
                                    },
                                },
                            }
                        );
                        break;
                    case "Library":
                        code = Exists ? 200 : 404;
                        response = Body(
                            new CatalogLibraryObservation
                            {
                                Id = List,
                                Unique = Unique,
                                Root = new FolderObservation
                                {
                                    Id = Root,
                                    Path = "/sites/test/Documents",
                                },
                            }
                        );
                        break;
                    case "Acl":
                        response = Body(
                            new ODataRows<AclAssignment>
                            {
                                Rows = OwnerAccess
                                    ? new[]
                                    {
                                        new AclAssignment
                                        {
                                            Member = new AclMember { Id = 7, Type = 8 },
                                            Roles = new ODataRows<AclRole>
                                            {
                                                Rows = new[]
                                                {
                                                    new AclRole
                                                    {
                                                        Id = 5,
                                                        Permissions = new PermissionMask
                                                        {
                                                            High = "2147483647",
                                                            Low = "4294967295",
                                                        },
                                                    },
                                                },
                                            },
                                        },
                                    }
                                    : Array.Empty<AclAssignment>(),
                            }
                        );
                        break;
                    default:
                        throw new Exception(work.ProbeKind);
                }
                work = Call("Observe", work, response, code);
            }
            throw new Exception("Library setup did not terminate");
        }

        private WorkerResult Call(
            string command,
            WorkerResult work,
            string? body = null,
            int status = 0
        ) =>
            Service.Transaction(() =>
                Worker.Execute(
                    new WorkerRequest
                    {
                        Command = command,
                        Key = work.Key,
                        RunId = "library-test",
                        Token = work.Token,
                        ProbeId = work.ProbeId,
                        ProbeKind = work.ProbeKind,
                        ResponseBody = body,
                        HttpStatus = status,
                    },
                    true
                )
            );
    }
}
