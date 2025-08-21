using System;
using System.Linq;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
        MessageNameEnum.Associate,
        "none",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperation_MirrorWorkspaceUsersToTask_Associate",
        1,
        IsolationModeEnum.Sandbox,
        Description = "Mirrors workspace-user associations onto task-user relationship.")]
    public class PostOperation_MirrorWorkspaceUsersToTask_Associate : IPlugin
    {
        private const string WorkspaceEntity = "ts_workorderservicetaskworkspace";
        private const string TaskEntity = "msdyn_workorderservicetask";
        private const string UserEntity = "systemuser";

        // Lookup field on workspace that points to the msdyn_workorderservicetask
        private const string Workspace_TaskLookup = "ts_workorderservicetask";

        // Relationship schema names (NOT table names)
        // Workspace <-> User M:N relationship schema name
        private const string Workspace_User_Relationship = "ts_WorkOrderServiceTaskWorkspace_SystemUser_SystemUser";
        // Task <-> User M:N relationship schema name to mirror onto
        private const string Task_User_Relationship = "ts_msdyn_workorderservicetask_systemuser";

        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = factory.CreateOrganizationService(context.UserId);
            var trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                trace.Trace("Mirror Associate start. CorrelationId={0}, Depth={1}", context.CorrelationId, context.Depth);

                // Guard: avoid recursion
                if (context.Depth > 1)
                {
                    trace.Trace("Depth > 1 detected. Exiting to prevent recursion.");
                    return;
                }

                if (!(context.InputParameters.Contains("Relationship") &&
                      context.InputParameters["Relationship"] is Relationship relationship))
                {
                    trace.Trace("Missing Relationship parameter. Exiting.");
                    return;
                }

                if (!string.Equals(relationship.SchemaName, Workspace_User_Relationship, StringComparison.OrdinalIgnoreCase))
                {
                    trace.Trace("Relationship {0} does not match expected {1}. Exiting.",
                        relationship.SchemaName, Workspace_User_Relationship);
                    return;
                }

                if (!(context.InputParameters.Contains("Target") &&
                      context.InputParameters["Target"] is EntityReference target))
                {
                    trace.Trace("Missing Target parameter. Exiting.");
                    return;
                }

                if (!(context.InputParameters.Contains("RelatedEntities") &&
                      context.InputParameters["RelatedEntities"] is EntityReferenceCollection relatedEntities) ||
                    relatedEntities.Count == 0)
                {
                    trace.Trace("Missing RelatedEntities parameter. Exiting.");
                    return;
                }

                // Determine which side is the workspace and which side is/are the user(s)
                EntityReference workspaceRef = null;
                var userRefs = new EntityReferenceCollection();

                if (string.Equals(target.LogicalName, WorkspaceEntity, StringComparison.OrdinalIgnoreCase))
                {
                    workspaceRef = target;
                    foreach (var re in relatedEntities.Where(e => e.LogicalName == UserEntity))
                        userRefs.Add(re);
                }
                else if (string.Equals(target.LogicalName, UserEntity, StringComparison.OrdinalIgnoreCase))
                {
                    var reWorkspace = relatedEntities.FirstOrDefault(e => e.LogicalName == WorkspaceEntity);
                    if (reWorkspace != null)
                    {
                        workspaceRef = reWorkspace;
                        userRefs.Add(target);
                    }
                }

                if (workspaceRef == null || userRefs.Count == 0)
                {
                    trace.Trace("Could not resolve workspace and user(s) from the association. Exiting.");
                    return;
                }

                // Retrieve the linked msdyn_workorderservicetask from the workspace
                var workspace = service.Retrieve(WorkspaceEntity, workspaceRef.Id, new ColumnSet(Workspace_TaskLookup));
                var taskRef = workspace.GetAttributeValue<EntityReference>(Workspace_TaskLookup);
                if (taskRef == null)
                {
                    trace.Trace("Workspace {0} does not have {1} populated. Exiting.", workspaceRef.Id, Workspace_TaskLookup);
                    return;
                }

                // Mirror the association(s) on the Task <-> User relationship
                var relationshipToMirror = new Relationship(Task_User_Relationship);

                foreach (var userRef in userRefs)
                {
                    try
                    {
                        var associate = new AssociateRequest
                        {
                            Target = new EntityReference(TaskEntity, taskRef.Id),
                            Relationship = relationshipToMirror,
                            RelatedEntities = new EntityReferenceCollection { new EntityReference(UserEntity, userRef.Id) }
                        };

                        service.Execute(associate);
                        trace.Trace("Mirrored association Task({0}) <-> User({1}) via relationship {2}.",
                            taskRef.Id, userRef.Id, Task_User_Relationship);
                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        var msg = ex.Detail?.Message ?? string.Empty;
                        if (msg.IndexOf("already associated", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            msg.IndexOf("duplicate", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            trace.Trace("Association already exists for Task({0}) and User({1}); ignoring.", taskRef.Id, userRef.Id);
                        }
                        else
                        {
                            trace.Trace("Associate fault for Task({0}) and User({1}): {2}", taskRef.Id, userRef.Id, ex.Message);
                            throw;
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
        }
    }

    [CrmPluginRegistration(
        MessageNameEnum.Disassociate,
        "none", // "none" for Associate/Disassociate
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "", // no filtered attributes; filter by relationship in code
        "TSIS2.Plugins.PostOperation_MirrorWorkspaceUsersToTask_Disassociate",
        1,
        IsolationModeEnum.Sandbox,
        Description = "Mirrors workspace-user disassociations by removing task-user links.")]
    public class PostOperation_MirrorWorkspaceUsersToTask_Disassociate : IPlugin
    {
        private const string WorkspaceEntity = "ts_workorderservicetaskworkspace";
        private const string TaskEntity = "msdyn_workorderservicetask";
        private const string UserEntity = "systemuser";

        // Lookup field on workspace that points to the msdyn_workorderservicetask
        private const string Workspace_TaskLookup = "ts_workorderservicetask";

        // Relationship schema names (NOT table names)
        private const string Workspace_User_Relationship = "ts_WorkOrderServiceTaskWorkspace_SystemUser_SystemUser";
        private const string Task_User_Relationship = "ts_msdyn_workorderservicetask_systemuser";

        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = factory.CreateOrganizationService(context.UserId);
            var trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                trace.Trace("Mirror Disassociate start. CorrelationId={0}, Depth={1}", context.CorrelationId, context.Depth);

                // Guard: avoid recursion
                if (context.Depth > 1)
                {
                    trace.Trace("Depth > 1 detected. Exiting to prevent recursion.");
                    return;
                }

                if (!(context.InputParameters.Contains("Relationship") &&
                      context.InputParameters["Relationship"] is Relationship relationship))
                {
                    trace.Trace("Missing Relationship parameter. Exiting.");
                    return;
                }

                if (!string.Equals(relationship.SchemaName, Workspace_User_Relationship, StringComparison.OrdinalIgnoreCase))
                {
                    trace.Trace("Relationship {0} does not match expected {1}. Exiting.",
                        relationship.SchemaName, Workspace_User_Relationship);
                    return;
                }

                if (!(context.InputParameters.Contains("Target") &&
                      context.InputParameters["Target"] is EntityReference target))
                {
                    trace.Trace("Missing Target parameter. Exiting.");
                    return;
                }

                if (!(context.InputParameters.Contains("RelatedEntities") &&
                      context.InputParameters["RelatedEntities"] is EntityReferenceCollection relatedEntities) ||
                    relatedEntities.Count == 0)
                {
                    trace.Trace("Missing RelatedEntities parameter. Exiting.");
                    return;
                }

                // Determine which side is the workspace and which side is/are the user(s)
                EntityReference workspaceRef = null;
                var userRefs = new EntityReferenceCollection();

                if (string.Equals(target.LogicalName, WorkspaceEntity, StringComparison.OrdinalIgnoreCase))
                {
                    workspaceRef = target;
                    foreach (var re in relatedEntities.Where(e => e.LogicalName == UserEntity))
                        userRefs.Add(re);
                }
                else if (string.Equals(target.LogicalName, UserEntity, StringComparison.OrdinalIgnoreCase))
                {
                    var reWorkspace = relatedEntities.FirstOrDefault(e => e.LogicalName == WorkspaceEntity);
                    if (reWorkspace != null)
                    {
                        workspaceRef = reWorkspace;
                        userRefs.Add(target);
                    }
                }

                if (workspaceRef == null || userRefs.Count == 0)
                {
                    trace.Trace("Could not resolve workspace and user(s) from the disassociation. Exiting.");
                    return;
                }

                // Retrieve the linked msdyn_workorderservicetask from the workspace
                var workspace = service.Retrieve(WorkspaceEntity, workspaceRef.Id, new ColumnSet(Workspace_TaskLookup));
                var taskRef = workspace.GetAttributeValue<EntityReference>(Workspace_TaskLookup);
                if (taskRef == null)
                {
                    trace.Trace("Workspace {0} does not have {1} populated. Exiting.", workspaceRef.Id, Workspace_TaskLookup);
                    return;
                }

                // Mirror the disassociation(s) on the Task <-> User relationship
                var relationshipToMirror = new Relationship(Task_User_Relationship);

                foreach (var userRef in userRefs)
                {
                    try
                    {
                        var disassociate = new DisassociateRequest
                        {
                            Target = new EntityReference(TaskEntity, taskRef.Id),
                            Relationship = relationshipToMirror,
                            RelatedEntities = new EntityReferenceCollection { new EntityReference(UserEntity, userRef.Id) }
                        };

                        service.Execute(disassociate);
                        trace.Trace("Mirrored disassociation Task({0}) !- User({1}) via relationship {2}.",
                            taskRef.Id, userRef.Id, Task_User_Relationship);
                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        var msg = ex.Detail?.Message ?? string.Empty;
                        if (msg.IndexOf("does not exist", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            msg.IndexOf("not associated", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            trace.Trace("No existing association Task({0}) !- User({1}); ignoring.", taskRef.Id, userRef.Id);
                        }
                        else
                        {
                            trace.Trace("Disassociate fault for Task({0}) and User({1}): {2}", taskRef.Id, userRef.Id, ex.Message);
                            throw;
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
        }
    }
}