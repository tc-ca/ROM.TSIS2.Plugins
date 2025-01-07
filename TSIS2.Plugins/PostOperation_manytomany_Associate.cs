using System;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
        MessageNameEnum.Associate,
        "none",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperation_manytomany_Associate Plugin",
        1,
        IsolationModeEnum.Sandbox,
        Description = "This plugin fires after every time a many-to-many relationship record is created for any entity")]
    public class PostOperation_manytomany_Associate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service reference which you will need for
            // web service calls.
            IOrganizationServiceFactory serviceFactory =
                (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            string relationshipName = string.Empty;

            try
            {
                tracingService.Trace("Get the Relationship Key from context");
                if (context.InputParameters.Contains("Relationship"))
                    relationshipName = context.InputParameters["Relationship"].ToString();

                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference &&
                    context.InputParameters.Contains("RelatedEntities") && context.InputParameters["RelatedEntities"] is EntityReferenceCollection)
                {
                    tracingService.Trace("Check the Relationship Name is part of ts_entityrisk");
                    {
                        //// for the Entity Risk we populate the ts_name with the ID of the related record in the many-to-many
                        //// this helps with integrity so we only have one entity risk for that year
                        //// Entity Risk - Account (Stakeholder)
                        //if (relationshipName.ToLower() == "ts_entityrisk_account_account.referencing")
                        //{           
                        //    tracingService.Trace("Entity Risk - Account (Stakeholder)");
                            
                        //    EntityReference targetEntity = (EntityReference)context.InputParameters["Target"];

                        //    EntityReferenceCollection targetEntityReferences = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];

                        //    using (var serviceContext = new Xrm(service))
                        //    {
                        //        tracingService.Trace("Get the Stakeholder");
                        //        var myStakeholder = serviceContext.AccountSet.Where(a => a.Id == targetEntity.Id).FirstOrDefault();

                        //        tracingService.Trace("Get the Entity Risk");
                        //        ts_EntityRisk myEntityRisk = null;
                        //        EntityReferenceCollection myTargetEntityReferences = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];
                                
                        //        foreach (var myEntityReference in myTargetEntityReferences)
                        //        {
                        //            myEntityRisk = serviceContext.ts_EntityRiskSet.Where(f => f.Id == myEntityReference.Id).FirstOrDefault();
                        //        }

                        //        tracingService.Trace("check if anything is null");
                        //        if (myEntityRisk != null && myStakeholder != null) {
                        //            tracingService.Trace($"We have Entity Risk: {myEntityRisk.Id}");
                        //            tracingService.Trace($"We have Stakeholder: {myStakeholder.Id}");

                        //            tracingService.Trace("Update the Entity Risk Record's Name, EntityID, and EntityName");
                        //            myEntityRisk.ts_Name = myStakeholder.Name;
                        //            myEntityRisk.ts_EntityID = myStakeholder.Id.ToString();
                        //            myEntityRisk.ts_EntityName = ts_EntityRisk_ts_EntityName.Stakeholder;
                        //            serviceContext.UpdateObject(myEntityRisk);
                        //            serviceContext.SaveChanges();
                        //        }
                        //    }
                        //}

                        //// Entity Risk - Functional Location (Site) 
                        //if (relationshipName.ToLower() == "ts_EntityRisk_msdyn_FunctionalLocation_msdyn_FunctionalLocation.referencing")
                        //{
                        //    tracingService.Trace("Entity Risk - Functional Location (Site)");

                        //    EntityReference targetEntity = (EntityReference)context.InputParameters["Target"];

                        //    EntityReferenceCollection targetEntityReferences = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];

                        //    using (var serviceContext = new Xrm(service))
                        //    {
                        //        tracingService.Trace("Get the Sites");
                        //        var mySite = serviceContext.msdyn_FunctionalLocationSet.Where(a => a.Id == targetEntity.Id).FirstOrDefault();

                        //        tracingService.Trace("Get the Entity Risk");
                        //        ts_EntityRisk myEntityRisk = null;
                        //        EntityReferenceCollection myTargetEntityReferences = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];

                        //        foreach (var myEntityReference in myTargetEntityReferences)
                        //        {
                        //            myEntityRisk = serviceContext.ts_EntityRiskSet.Where(f => f.Id == myEntityReference.Id).FirstOrDefault();
                        //        }

                        //        tracingService.Trace("check if anything is null");
                        //        if (myEntityRisk != null && mySite != null)
                        //        {
                        //            tracingService.Trace($"We have Entity Risk: {myEntityRisk.Id}");
                        //            tracingService.Trace($"We have Site: {mySite.Id}");

                        //            tracingService.Trace("Update the Entity Risk Record's Name, EntityID, and EntityName");
                        //            myEntityRisk.ts_Name = mySite.msdyn_Name;
                        //            myEntityRisk.ts_EntityID = mySite.Id.ToString();
                        //            myEntityRisk.ts_EntityName = ts_EntityRisk_ts_EntityName.Site;
                        //            serviceContext.UpdateObject(myEntityRisk);
                        //            serviceContext.SaveChanges();
                        //        }
                        //    }
                        //}

                        //// Entity Risk - Incident Type  (Activity Type)
                        //if (relationshipName.ToLower() == "ts_EntityRisk_msdyn_incidenttype_msdyn_incidenttype.referencing")
                        //{
                        //    tracingService.Trace("Entity Risk - Incident Type  (Activity Type)");

                        //    EntityReference targetEntity = (EntityReference)context.InputParameters["Target"];

                        //    EntityReferenceCollection targetEntityReferences = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];

                        //    using (var serviceContext = new Xrm(service))
                        //    {
                        //        tracingService.Trace("Get the Activity Type");
                        //        var myIncidentType = serviceContext.msdyn_incidenttypeSet.Where(a => a.Id == targetEntity.Id).FirstOrDefault();

                        //        tracingService.Trace("Get the Entity Risk");
                        //        ts_EntityRisk myEntityRisk = null;
                        //        EntityReferenceCollection myTargetEntityReferences = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];

                        //        foreach (var myEntityReference in myTargetEntityReferences)
                        //        {
                        //            myEntityRisk = serviceContext.ts_EntityRiskSet.Where(f => f.Id == myEntityReference.Id).FirstOrDefault();
                        //        }

                        //        tracingService.Trace("check if anything is null");
                        //        if (myEntityRisk != null && myIncidentType != null)
                        //        {
                        //            tracingService.Trace($"We have Entity Risk: {myEntityRisk.Id}");
                        //            tracingService.Trace($"We have ActivityType: {myIncidentType.Id}");

                        //            tracingService.Trace("Update the Entity Risk Record's Name, EntityID, and EntityName");
                        //            myEntityRisk.ts_Name = myIncidentType.ovs_IncidentTypeNameEnglish;
                        //            myEntityRisk.ts_EntityID = myIncidentType.Id.ToString();
                        //            myEntityRisk.ts_EntityName = ts_EntityRisk_ts_EntityName.ActivityType;
                        //            serviceContext.UpdateObject(myEntityRisk);
                        //            serviceContext.SaveChanges();
                        //        }
                        //    }
                        //}

                        //// Entity Risk - Operation
                        //if (relationshipName.ToLower() == "ts_EntityRisk_ovs_operation_ovs_operation.referencing")
                        //{
                        //    tracingService.Trace("Entity Risk - Operation");

                        //    EntityReference targetEntity = (EntityReference)context.InputParameters["Target"];

                        //    EntityReferenceCollection targetEntityReferences = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];

                        //    using (var serviceContext = new Xrm(service))
                        //    {
                        //        tracingService.Trace("Get the Operation");
                        //        var myOperation = serviceContext.ovs_operationSet.Where(a => a.Id == targetEntity.Id).FirstOrDefault();

                        //        tracingService.Trace("Get the Entity Risk");
                        //        ts_EntityRisk myEntityRisk = null;
                        //        EntityReferenceCollection myTargetEntityReferences = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];

                        //        foreach (var myEntityReference in myTargetEntityReferences)
                        //        {
                        //            myEntityRisk = serviceContext.ts_EntityRiskSet.Where(f => f.Id == myEntityReference.Id).FirstOrDefault();
                        //        }

                        //        tracingService.Trace("check if anything is null");
                        //        if (myEntityRisk != null && myOperation != null)
                        //        {
                        //            tracingService.Trace($"We have Entity Risk: {myEntityRisk.Id}");
                        //            tracingService.Trace($"We have Operation: {myOperation.Id}");

                        //            tracingService.Trace("Update the Entity Risk Record's Name, EntityID, and EntityName");
                        //            myEntityRisk.ts_Name = myOperation.ovs_name;
                        //            myEntityRisk.ts_EntityID = myOperation.Id.ToString();
                        //            myEntityRisk.ts_EntityName = ts_EntityRisk_ts_EntityName.Operation;
                        //            serviceContext.UpdateObject(myEntityRisk);
                        //            serviceContext.SaveChanges();
                        //        }
                        //    }
                        //}

                        //// Entity Risk - Operation Type
                        //if (relationshipName.ToLower() == "ts_EntityRisk_ovs_operationtype_ovs_operationtype.referencing")
                        //{
                        //    tracingService.Trace("Entity Risk - Operation Type");

                        //    EntityReference targetEntity = (EntityReference)context.InputParameters["Target"];

                        //    EntityReferenceCollection targetEntityReferences = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];

                        //    using (var serviceContext = new Xrm(service))
                        //    {
                        //        tracingService.Trace("Get the Operation Type");
                        //        var myOperationType = serviceContext.ovs_operationtypeSet.Where(a => a.Id == targetEntity.Id).FirstOrDefault();

                        //        tracingService.Trace("Get the Entity Risk");
                        //        ts_EntityRisk myEntityRisk = null;
                        //        EntityReferenceCollection myTargetEntityReferences = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];

                        //        foreach (var myEntityReference in myTargetEntityReferences)
                        //        {
                        //            myEntityRisk = serviceContext.ts_EntityRiskSet.Where(f => f.Id == myEntityReference.Id).FirstOrDefault();
                        //        }

                        //        tracingService.Trace("check if anything is null");
                        //        if (myEntityRisk != null && myOperationType != null)
                        //        {
                        //            tracingService.Trace($"We have Entity Risk: {myEntityRisk.Id}");
                        //            tracingService.Trace($"We have Operation Type: {myOperationType.Id}");

                        //            tracingService.Trace("Update the Entity Risk Record's Name, EntityID, and EntityName");
                        //            myEntityRisk.ts_Name = myOperationType.ovs_name;
                        //            myEntityRisk.ts_EntityID = myOperationType.Id.ToString();
                        //            myEntityRisk.ts_EntityName = ts_EntityRisk_ts_EntityName.OperationType;
                        //            serviceContext.UpdateObject(myEntityRisk);
                        //            serviceContext.SaveChanges();
                        //        }
                        //    }
                        //}

                        //// Entity Risk - Program Area
                        //if (relationshipName.ToLower() == "ts_EntityRisk_ts_programarea_ts_programarea.referencing")
                        //{
                        //    tracingService.Trace("Entity Risk - Program Area");

                        //    EntityReference targetEntity = (EntityReference)context.InputParameters["Target"];

                        //    EntityReferenceCollection targetEntityReferences = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];

                        //    using (var serviceContext = new Xrm(service))
                        //    {
                        //        tracingService.Trace("Get the Program Area");
                        //        var myProgramArea = serviceContext.ts_programareaSet.Where(a => a.Id == targetEntity.Id).FirstOrDefault();

                        //        tracingService.Trace("Get the Entity Risk");
                        //        ts_EntityRisk myEntityRisk = null;
                        //        EntityReferenceCollection myTargetEntityReferences = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];

                        //        foreach (var myEntityReference in myTargetEntityReferences)
                        //        {
                        //            myEntityRisk = serviceContext.ts_EntityRiskSet.Where(f => f.Id == myEntityReference.Id).FirstOrDefault();
                        //        }

                        //        tracingService.Trace("check if anything is null");
                        //        if (myEntityRisk != null && myProgramArea != null)
                        //        {
                        //            tracingService.Trace($"We have Entity Risk: {myEntityRisk.Id}");
                        //            tracingService.Trace($"We have Program Area: {myProgramArea.Id}");

                        //            tracingService.Trace("Update the Entity Risk Record's Name, EntityID, and EntityName");
                        //            myEntityRisk.ts_Name = myProgramArea.ts_ProgramAreaEN;
                        //            myEntityRisk.ts_EntityID = myProgramArea.Id.ToString();
                        //            myEntityRisk.ts_EntityName = ts_EntityRisk_ts_EntityName.ProgramArea;
                        //            serviceContext.UpdateObject(myEntityRisk);
                        //            serviceContext.SaveChanges();
                        //        }
                        //    }
                        //}
                    }
                }
            }
            catch (Exception e)
            {
                tracingService.Trace("Exception occurred in PostOperation_manytomany_Associate: {0}", e.ToString());
                tracingService.Trace($"MESSAGE: {e.Message}");
                tracingService.Trace($"STACK TRACE: {e.InnerException?.Message}");
                tracingService.Trace($"SOURCE: {e.Source}");
            }
        }
    }
}