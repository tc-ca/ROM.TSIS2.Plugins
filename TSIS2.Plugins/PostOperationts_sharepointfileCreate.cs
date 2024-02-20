using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
    MessageNameEnum.Create,
    "ts_sharepointfile",
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PostOperationts_sharepointfileCreate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Description = "After a ts_sharepointfile record is created, handle the logic to setup the ts_sharepointgroup")]
    public class PostOperationts_sharepointfileCreate : IPlugin
    {
        // Static Variables
        public static string CASE = "Case";
        public static string CASE_FR = "Cas";
        private static string EXEMPTION = "Exemption";
        //private static string EXEMPTION_FR = "Exemption";
        private static string OPERATION = "Operation";
        //private static string OPERATION_FR = "Opération";
        private static string SECURITY_INCIDENT = "Security Incident";
        //private static string SECURITY_INCIDENT_FR = "Incidents de sûreté";
        private static string SITE = "Site";
        //private static string SITE_FR = "Site";
        private static string STAKEHOLDER = "Stakeholder";
        //private static string STAKEHOLDER_FR = "Partie prenante";
        public static string WORK_ORDER = "Work Order";
        public static string WORK_ORDER_FR = "Ordre de travail";
        public static string WORK_ORDER_SERVICE_TASK = "Work Order Service Task";
        public static string WORK_ORDER_SERVICE_TASK_FR = "Tâche de service de l'ordre de travail";

        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity target = (Entity)context.InputParameters["Target"];

                // Obtain the preimage entity
                Entity preImageEntity = (context.PreEntityImages != null && context.PreEntityImages.Contains("PreImage")) ? context.PreEntityImages["PreImage"] : null;

                // Obtain the organization service reference which you will need for
                // web service calls.
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    if (target.LogicalName.Equals(ts_SharePointFile.EntityLogicalName))
                    {
                        //ts_SharePointFile mySharePointFile = target.ToEntity<ts_SharePointFile>();
                        //var mySharePointFileGroup = mySharePointFile.ts_SharePointFileGroup;


                        //// check if we are working with a Case, Work Order, or Work Order Service Task
                        //if (mySharePointFile.ts_TableName == CASE || 
                        //    mySharePointFile.ts_TableName == WORK_ORDER ||
                        //    mySharePointFile.ts_TableName == WORK_ORDER_SERVICE_TASK)
                        //{
                        //    var sharePointFileTableName = mySharePointFile.ts_TableName;

                        //    // Case
                        //    if (sharePointFileTableName == CASE)
                        //    {
                        //        // check if the case has a SharePoint File Group
                        //        if (mySharePointFileGroup == null)
                        //        {
                        //            // create the SharePoint File Group
                        //            Guid myCaseSharePointFileGroupID = CreateSharePointFileGroup(mySharePointFile,service);

                        //            Guid myCaseID = new Guid(mySharePointFile.ts_TableRecordID);

                        //            UpdateRelatedWorkOrders(service, myCaseID, myCaseSharePointFileGroupID,mySharePointFile.ts_TableRecordOwner);
                        //        }
                        //        else
                        //        {
                        //            // do nothing
                        //        }
                        //    }

                        //    // Work Order
                        //    if (sharePointFileTableName == WORK_ORDER)
                        //    {
                        //        // find out if the Work Order has a Case assigned to it
                        //        using(var serviceContext = new Xrm(service))
                        //        {
                        //            Guid myWorkOrderID = new Guid(mySharePointFile.ts_TableRecordID);

                        //            msdyn_workorder myWorkOrder = serviceContext.msdyn_workorderSet.Where(wo => wo.Id == myWorkOrderID).FirstOrDefault();

                        //            if (myWorkOrder.msdyn_ServiceRequest != null)
                        //            {
                        //                // if we have a Case - find out if it has a SharePoint File
                        //                string myCaseIdString = myWorkOrder.msdyn_ServiceRequest.Id.ToString().ToUpper();

                        //                ts_SharePointFile myCaseSharePointFile = null;

                        //                myCaseSharePointFile = CheckSharePointFile(serviceContext, myCaseIdString, CASE);

                        //                Guid myCaseSharePointFileGroupID = new Guid();

                        //                if (myCaseSharePointFile == null)
                        //                {
                        //                    // if it doesn't have a SharePoint File, create it
                        //                    Guid newCaseSharePointFileID = CreateSharePointFile(myWorkOrder.msdyn_ServiceRequest.Name,CASE,CASE_FR,myCaseIdString,myWorkOrder.msdyn_ServiceRequest.Name,mySharePointFile.ts_TableRecordOwner,service);

                        //                    myCaseSharePointFile = serviceContext.ts_SharePointFileSet.Where(sf => sf.Id == newCaseSharePointFileID).FirstOrDefault();

                        //                    // Create the SharePoint File Group
                        //                    myCaseSharePointFileGroupID = CreateSharePointFileGroup(myCaseSharePointFile, service);

                        //                    // Assign the Work Order SharePoint File Group to the SharePoint File Group of the Case
                        //                    service.Update(new ts_SharePointFile
                        //                    {
                        //                        Id = mySharePointFile.Id,
                        //                        ts_SharePointFileGroup = new EntityReference(ts_sharepointfilegroup.EntityLogicalName, myCaseSharePointFileGroupID)
                        //                    });

                        //                    // do this here because we have created a new SharePoint File for Case with a SharePoint File Group
                        //                    UpdateRelatedWorkOrders(service, myWorkOrder.msdyn_ServiceRequest.Id, myCaseSharePointFileGroupID, mySharePointFile.ts_TableRecordOwner);
                        //                }
                        //                else
                        //                {
                        //                    // if it does, then get the SharePoint Group file and assign it to the Work Order
                        //                    service.Update(new ts_SharePointFile
                        //                    {
                        //                        Id=mySharePointFile.Id,
                        //                        ts_SharePointFileGroup = myCaseSharePointFile.ts_SharePointFileGroup
                        //                    });

                        //                    //if a Case already has a SharePoint File, all other Work Orders are already updated with the proper SharePoint Files
                        //                }
                        //            }
                        //            else
                        //            {
                        //                // if we don't have a Case - create a SharePoint File Group for the Work Order and assign it
                        //                var mySharePointFileGroupId = CreateSharePointFileGroup(mySharePointFile, service);

                        //                // update all related Work Order Service Tasks
                        //                UpdateRelatedWorkOrderServiceTasks(service, myWorkOrderID, mySharePointFileGroupId, mySharePointFile.ts_TableRecordOwner);
                        //            }
                        //        }
                        //    }


                        //    // Work Order Service Task
                        //    if (sharePointFileTableName == WORK_ORDER_SERVICE_TASK)
                        //    {
                        //        // find out if the Work Order Service Task has a Work Order assigned to it
                        //        using (var serviceContext = new Xrm(service))
                        //        {
                        //            Guid myWorkOrderServiceTaskID = new Guid(mySharePointFile.ts_TableRecordID);

                        //            msdyn_workorderservicetask myWorkOrderServiceTask = serviceContext.msdyn_workorderservicetaskSet.Where(wost => wost.Id == myWorkOrderServiceTaskID).FirstOrDefault();

                        //            if (myWorkOrderServiceTask.msdyn_WorkOrder != null)
                        //            {
                        //                // Find out if the Work Order has a Case
                        //                var myWorkOrderCase = (CaseID: "", CaseName: "");

                        //                string workOrderCaseFetchXML = $@"
                        //                    <fetch>
                        //                        <entity name='msdyn_workorder'>
                        //                        <link-entity name='incident' to='msdyn_servicerequest' from='incidentid' alias='incident' link-type='inner'>
                        //                            <attribute name='incidentid' alias='CASE_ID' />
                        //                            <attribute name='title' alias='CASE_NAME' />
                        //                        </link-entity>
                        //                        <filter>
                        //                            <condition attribute='msdyn_workorderid' operator='eq' value='{myWorkOrderServiceTask.msdyn_WorkOrder.Id.ToString()}' />
                        //                        </filter>
                        //                        </entity>
                        //                    </fetch>
                        //                ";

                        //                EntityCollection myWorkOrderCaseEntityCollection = service.RetrieveMultiple(new FetchExpression(workOrderCaseFetchXML));

                        //                foreach (var myCaseItem in myWorkOrderCaseEntityCollection.Entities)
                        //                {
                        //                    string caseID = "";
                        //                    string caseName = "";

                        //                    if (myCaseItem.Attributes["CASE_ID"] is AliasedValue aliasedCaseId)
                        //                    {
                        //                        caseID = aliasedCaseId.Value.ToString();
                        //                    }

                        //                    if (myCaseItem.Attributes["CASE_NAME"] is AliasedValue aliasedCaseName)
                        //                    {
                        //                        caseName = aliasedCaseName.Value.ToString();
                        //                    }

                        //                    myWorkOrderCase = (
                        //                        CaseID: caseID,
                        //                        CaseName: caseName
                        //                        );
                        //                }

                        //                if (myWorkOrderCase.CaseID != "")
                        //                {
                        //                    // if we have a Case - find out if it has a SharePoint File
                        //                    ts_SharePointFile myCaseSharePointFile = null;

                        //                    myCaseSharePointFile = CheckSharePointFile(serviceContext, myWorkOrderCase.CaseID, CASE);

                        //                    if (myCaseSharePointFile == null)
                        //                    {
                        //                        // if it doesn't have a SharePoint File, create it
                        //                        Guid newCaseSharePointFileID = CreateSharePointFile(myWorkOrderCase.CaseName, CASE, CASE_FR, myWorkOrderCase.CaseID, myWorkOrderCase.CaseName, mySharePointFile.ts_TableRecordOwner, service);

                        //                        myCaseSharePointFile = serviceContext.ts_SharePointFileSet.Where(sf => sf.Id == newCaseSharePointFileID).FirstOrDefault();

                        //                        // Create the SharePoint File Group
                        //                        Guid newCaseSharePointFileGroupID = CreateSharePointFileGroup(myCaseSharePointFile, service);

                        //                        // Update all Work Orders, and Work Order Service Tasks related to the Case with the SharePoint Group File
                        //                        UpdateRelatedWorkOrders(service, new Guid(myWorkOrderCase.CaseID), newCaseSharePointFileGroupID, mySharePointFile.ts_TableRecordOwner);
                        //                    }
                        //                    else
                        //                    {
                        //                        // only update the Work Order Service Task with the SharePoint Group File for the Case
                        //                        service.Update(new ts_SharePointFile
                        //                        {
                        //                            Id = mySharePointFile.Id,
                        //                            ts_SharePointFileGroup = myCaseSharePointFile.ts_SharePointFileGroup
                        //                        });
                        //                    }
                        //                }
                        //                else
                        //                {
                        //                    // If the Work Order doesn't have a Case

                        //                    // Find out if the Work Order has a SharePoint File
                        //                    ts_SharePointFile myWorkOrderSharePointFile = CheckSharePointFile(serviceContext, myWorkOrderServiceTask.msdyn_WorkOrder.Id.ToString().ToUpper(), WORK_ORDER);
                        //                    ts_sharepointfilegroup myWorkOrderSharePointFileGroup = null;

                        //                    // Create the SharePoint File if it doesn't exists, along with the SharePoint File Group
                        //                    if (myWorkOrderSharePointFile == null)
                        //                    {
                        //                        Guid newWorkOrderSharePointFileID = CreateSharePointFile(myWorkOrderServiceTask.msdyn_WorkOrder.Name, WORK_ORDER, WORK_ORDER_FR, myWorkOrderServiceTask.msdyn_WorkOrder.Id.ToString().ToUpper(), myWorkOrderServiceTask.msdyn_WorkOrder.Name, mySharePointFile.ts_TableRecordOwner, service);
                        //                        myWorkOrderSharePointFile = CheckSharePointFile(serviceContext, myWorkOrderServiceTask.msdyn_WorkOrder.Id.ToString().ToUpper(), WORK_ORDER);
                        //                        Guid newWorkOrderSharePointFileGroupID = CreateSharePointFileGroup(myWorkOrderSharePointFile, service);
                        //                        myWorkOrderSharePointFileGroup = serviceContext.ts_sharepointfilegroupSet.Where(sg => sg.Id == newWorkOrderSharePointFileGroupID).FirstOrDefault();

                        //                        // Update all other Work Order Service Tasks related to the Work Order with the SharePoint File Group
                        //                        // We go over this because we want the Work Order Service Task to show all related attachments, even if it doesn't have one.

                        //                        UpdateRelatedWorkOrderServiceTasks(service, myWorkOrderServiceTask.msdyn_WorkOrder.Id, myWorkOrderSharePointFileGroup.Id, mySharePointFile.ts_TableRecordOwner);
                        //                    }
                        //                    else
                        //                    {
                        //                        // Check if the SharePoint File has a SharePoint File Group
                        //                        if (myWorkOrderSharePointFile.ts_SharePointFileGroup == null)
                        //                        {
                        //                            // if it doesn't create it
                        //                            Guid newWorkOrderSharePointFileGroupID = CreateSharePointFileGroup(myWorkOrderSharePointFile, service);
                        //                            myWorkOrderSharePointFileGroup = serviceContext.ts_sharepointfilegroupSet.Where(sg => sg.Id == newWorkOrderSharePointFileGroupID).FirstOrDefault();
                        //                        }
                        //                        else
                        //                        {
                        //                            myWorkOrderSharePointFileGroup = serviceContext.ts_sharepointfilegroupSet.Where(sg => sg.Id == myWorkOrderSharePointFile.ts_SharePointFileGroup.Id).FirstOrDefault();
                        //                        }

                        //                        // Update the Work Order Service Task with the Work Order SharePoint File Group 
                        //                        service.Update(new ts_SharePointFile
                        //                        {
                        //                            Id = mySharePointFile.Id,
                        //                            ts_SharePointFileGroup = myWorkOrderSharePointFileGroup.ToEntityReference()
                        //                        });
                        //                    }
                        //                }
                        //            }
                        //            else
                        //            {
                        //                // do nothing since the Work Order Service Task doesn't have a Work Order
                        //            }
                        //        }
                        //    }
                        //}
                        //else if (mySharePointFile.ts_TableName == STAKEHOLDER ||
                        //         mySharePointFile.ts_TableName == SITE ||
                        //         mySharePointFile.ts_TableName == OPERATION ||
                        //         mySharePointFile.ts_TableName == SECURITY_INCIDENT ||
                        //         mySharePointFile.ts_TableName == EXEMPTION)
                        //{
                        //    // check if the SharePoint File has a SharePoint File Group - might include this in the future if they want to group the above table names
                        //    //if (mySharePointFileGroup == null)
                        //    //{
                        //    //    {
                        //    //        // create the SharePoint File Group
                        //    //        CreateSharePointFileGroup(mySharePointFile, service);
                        //    //    }
                        //    //}
                        //}
                        //else
                        //{
                        //    // for everything else do nothing - this is where new SharePoint Files go when they are initially created as placeholders in this plugin
                        //}
                    }
                }
                catch (Exception ex)
                {
                    tracingService.Trace("PostOperationts_fileCreate Plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    
        public static Guid CreateSharePointFile(string ts_name, string ts_tablename, string ts_tablenamefrench, string ts_tablerecordid, string ts_tablerecordname, string ts_tablerecordowner, IOrganizationService service)
        {
            ts_SharePointFile newSharePointFile = new ts_SharePointFile();
            newSharePointFile.ts_Name = ts_name;

            // NOTE: When we create the SharePoint File this plugin will run again, however nothing will happen because it has no ts_TableName
            Guid newSharePointFileID = service.Create(newSharePointFile);

            service.Update(new ts_SharePointFile
            {
                Id = newSharePointFileID,
                ts_TableName = ts_tablename,
                ts_TableNameFrench = ts_tablenamefrench,
                ts_TableRecordID = ts_tablerecordid,
                ts_TableRecordName = ts_tablerecordname,
                ts_TableRecordOwner = ts_tablerecordowner
            });

            return newSharePointFileID;
        }

        public static Guid CreateSharePointFileGroup(ts_SharePointFile mySharePointFile, IOrganizationService service)
        {
            // create the SharePoint File Group
            ts_sharepointfilegroup newSharePointFileGroup = new ts_sharepointfilegroup();
            newSharePointFileGroup.ts_Name = mySharePointFile.ts_TableRecordName;
            Guid newSharePointFileGroupId = service.Create(newSharePointFileGroup);

            // assign the SharePoint File the new SharePoint File Group
            service.Update(new ts_SharePointFile
            {
                Id = mySharePointFile.Id,
                ts_SharePointFileGroup = new EntityReference(ts_sharepointfilegroup.EntityLogicalName, newSharePointFileGroupId)
            });

            return newSharePointFileGroupId;
        }

        public static ts_SharePointFile CheckSharePointFile(Xrm serviceContext, string myTableRecordID, string myTableName)
        {
            ts_SharePointFile mySharePointFile = null;

            try
                {
                mySharePointFile = serviceContext.ts_SharePointFileSet
                    .FirstOrDefault(file => file.ts_TableRecordID == myTableRecordID && file.ts_TableName == myTableName);
                }
            catch (Exception ex)
            {
                // Log the exception
                System.Diagnostics.Trace.TraceError("An error occurred while running PostOperationSharePointFileCreate.CheckSharePointFile(): " + ex.ToString());
            }

            return mySharePointFile;
        }
        
        public static void UpdateRelatedWorkOrders(IOrganizationService service, Guid myCaseID, Guid myCaseSharePointFileGroupID, string recordOwner)
        {
            using (var serviceContext = new Xrm(service))
            {
                var myWorkOrders = serviceContext.msdyn_workorderSet.Where(wo => wo.msdyn_ServiceRequest.Id == myCaseID);

                foreach (var myWorkOrder in myWorkOrders)
                {
                    // check if there are any SharePointFiles
                    string myWorkOrderIdString = myWorkOrder.Id.ToString().ToUpper();

                    ts_SharePointFile myWorkOrderSharePointFile = null;

                    myWorkOrderSharePointFile = CheckSharePointFile(serviceContext, myWorkOrderIdString, WORK_ORDER);

                    if (myWorkOrderSharePointFile == null)
                    {
                        // if it doesn't have a SharePoint File, create it
                        Guid newWorkOrderSharePointFileID = CreateSharePointFile(myWorkOrder.msdyn_name, WORK_ORDER, WORK_ORDER_FR, myWorkOrderIdString, myWorkOrder.msdyn_name, recordOwner, service);

                        myWorkOrderSharePointFile = serviceContext.ts_SharePointFileSet.Where(sf => sf.Id == newWorkOrderSharePointFileID).FirstOrDefault();
                    }

                    // Assign the Work Order SharePoint File Group to the SharePoint File Group of the Case
                    service.Update(new ts_SharePointFile
                    {
                        Id = myWorkOrderSharePointFile.Id,
                        ts_SharePointFileGroup = new EntityReference(ts_sharepointfilegroup.EntityLogicalName, myCaseSharePointFileGroupID)
                    });

                    // Now go through each Work Order Service Tasks for the Work Order
                    var myWorkOrderServiceTasks = serviceContext.msdyn_workorderservicetaskSet.Where(wost => wost.msdyn_WorkOrder.Id == myWorkOrder.Id);

                    foreach (var myWorkOrderServiceTask in myWorkOrderServiceTasks)
                    {
                        // check if there are any SharePointFiles
                        string myWorkOrderServiceTaskIdString = myWorkOrderServiceTask.Id.ToString().ToUpper();

                        ts_SharePointFile myWorkOrderServiceTaskSharePointFile = null;

                        myWorkOrderServiceTaskSharePointFile = CheckSharePointFile(serviceContext, myWorkOrderServiceTaskIdString, WORK_ORDER_SERVICE_TASK);

                        if (myWorkOrderServiceTaskSharePointFile == null)
                        {
                            // if it doesn't have a SharePoint File, create it
                            Guid newWorkOrderServiceTaskSharePointFileID = CreateSharePointFile(myWorkOrderServiceTask.msdyn_name, WORK_ORDER_SERVICE_TASK, WORK_ORDER_SERVICE_TASK_FR, myWorkOrderServiceTaskIdString,myWorkOrderServiceTask.msdyn_name, recordOwner, service);

                            myWorkOrderServiceTaskSharePointFile = serviceContext.ts_SharePointFileSet.Where(sf => sf.Id == newWorkOrderServiceTaskSharePointFileID).FirstOrDefault();
                        }

                        // Assign the Work Order Service Task SharePoint File Group to the SharePoint File Group of the Case
                        service.Update(new ts_SharePointFile
                        {
                            Id = myWorkOrderServiceTaskSharePointFile.Id,
                            ts_SharePointFileGroup = new EntityReference(ts_sharepointfilegroup.EntityLogicalName, myCaseSharePointFileGroupID)
                        });
                    }
                }
            }
        }

        public static void UpdateRelatedWorkOrderServiceTasks(IOrganizationService service, Guid myWorkOrderID, Guid myWorkOrderSharePointFileGroupID, string recordOwner)
        {
            using (var serviceContext = new Xrm(service))
            {
                var myWorkOrders = serviceContext.msdyn_workorderSet.Where(wo => wo.Id == myWorkOrderID);

                foreach (var myWorkOrder in myWorkOrders)
                {
                    // check if there are any SharePointFiles
                    string myWorkOrderIdString = myWorkOrder.Id.ToString().ToUpper();

                    ts_SharePointFile myWorkOrderSharePointFile = null;

                    myWorkOrderSharePointFile = CheckSharePointFile(serviceContext, myWorkOrderIdString, WORK_ORDER);

                    if (myWorkOrderSharePointFile == null)
                    {
                        // if it doesn't have a SharePoint File, create it
                        Guid newWorkOrderSharePointFileID = CreateSharePointFile(myWorkOrder.msdyn_name, WORK_ORDER, WORK_ORDER_FR, myWorkOrderIdString, myWorkOrder.msdyn_name, recordOwner, service);

                        myWorkOrderSharePointFile = serviceContext.ts_SharePointFileSet.Where(sf => sf.Id == newWorkOrderSharePointFileID).FirstOrDefault();
                    }

                    // Assign the Work Order SharePoint File Group to the SharePoint File Group of the Case
                    service.Update(new ts_SharePointFile
                    {
                        Id = myWorkOrderSharePointFile.Id,
                        ts_SharePointFileGroup = new EntityReference(ts_sharepointfilegroup.EntityLogicalName, myWorkOrderSharePointFileGroupID)
                    });

                    // Now go through each Work Order Service Tasks for the Work Order
                    var myWorkOrderServiceTasks = serviceContext.msdyn_workorderservicetaskSet.Where(wost => wost.msdyn_WorkOrder.Id == myWorkOrder.Id);

                    foreach (var myWorkOrderServiceTask in myWorkOrderServiceTasks)
                    {
                        // check if there are any SharePointFiles
                        string myWorkOrderServiceTaskIdString = myWorkOrderServiceTask.Id.ToString().ToUpper();

                        ts_SharePointFile myWorkOrderServiceTaskSharePointFile = null;

                        myWorkOrderServiceTaskSharePointFile = CheckSharePointFile(serviceContext, myWorkOrderServiceTaskIdString, WORK_ORDER_SERVICE_TASK);

                        if (myWorkOrderServiceTaskSharePointFile == null)
                        {
                            // if it doesn't have a SharePoint File, create it
                            Guid newWorkOrderServiceTaskSharePointFileID = CreateSharePointFile(myWorkOrderServiceTask.msdyn_name, WORK_ORDER_SERVICE_TASK, WORK_ORDER_SERVICE_TASK_FR, myWorkOrderServiceTaskIdString, myWorkOrderServiceTask.msdyn_name, recordOwner, service);

                            myWorkOrderServiceTaskSharePointFile = serviceContext.ts_SharePointFileSet.Where(sf => sf.Id == newWorkOrderServiceTaskSharePointFileID).FirstOrDefault();
                        }

                        // Assign the Work Order Service Task SharePoint File Group to the SharePoint File Group of the Case
                        service.Update(new ts_SharePointFile
                        {
                            Id = myWorkOrderServiceTaskSharePointFile.Id,
                            ts_SharePointFileGroup = new EntityReference(ts_sharepointfilegroup.EntityLogicalName, myWorkOrderSharePointFileGroupID)
                        });
                    }
                }
            }
        }
    
        public static string GetWorkOrderOwner(IOrganizationService serviceContext, Guid workOrderId)
        {
            string myOwner = "";

            // get the owner
            string ownerFetchXML = $@"
                                        <fetch>
                                            <entity name='msdyn_workorder'>
                                            <link-entity name='ovs_operationtype' to='ovs_operationtypeid' from='ovs_operationtypeid' alias='ovs_operationtype' link-type='inner'>
                                                <link-entity name='team' to='owningteam' from='teamid' alias='team' link-type='inner'>
                                                <attribute name='name' alias='OWNER_NAME' />
                                                </link-entity>
                                            </link-entity>
                                            <filter>
                                                <condition attribute='msdyn_workorderid' operator='eq' value='{workOrderId.ToString()}' />
                                            </filter>
                                            </entity>
                                        </fetch>                                                
            ";

            var myWorkOrderEntityCollection = serviceContext.RetrieveMultiple(new FetchExpression(ownerFetchXML));

            foreach (var item in myWorkOrderEntityCollection.Entities)
            {
                if (item.Attributes["OWNER_NAME"] is AliasedValue aliasedOwner)
                {
                    myOwner = aliasedOwner.Value.ToString();
                }
            }

            // make this adjustment for the DEV Environment
            if (myOwner == "Intermodal Surface Security Oversight (ISSO) (dev)")
            {
                myOwner = "Intermodal Surface Security Oversight (ISSO)";
            }

            return myOwner;
        }
    }
}
