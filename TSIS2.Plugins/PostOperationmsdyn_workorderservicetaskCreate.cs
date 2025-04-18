﻿// <copyright file="PostOperationmsdyn_workorderservicetaskCreate.cs" company="">
// Copyright (c) 2018 All Rights Reserved
// </copyright>
// <author>Hong Liu</author>
// <date>9/20/2018 10:21:27 AM</date>
// <summary>Implements the PostOperationmsdyn_workorderservicetaskCreate Plugin.</summary>
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.1
// </auto-generated>

using System;
using System.Linq;
using System.Web.Services.Description;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
    MessageNameEnum.Create,
    "msdyn_workorderservicetask",
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PostOperationmsdyn_workorderservicetaskCreate Plugin",
    1,
    IsolationModeEnum.Sandbox, 
    Description = "Copies the JSON Questionnaire Definition from a Service Task to a Work Order Service Task.")]
    /// <summary>
    /// PostOperationmsdyn_workorderservicetaskCreate Plugin.
    /// </summary>    
    public class PostOperationmsdyn_workorderservicetaskCreate : PluginBase
    {
        //private readonly string postImageAlias = "PostImage";

        /// <summary>
        /// Initializes a new instance of the <see cref="PostOperationmsdyn_workorderservicetaskCreate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PostOperationmsdyn_workorderservicetaskCreate(string unsecure, string secure)
            : base(typeof(PostOperationmsdyn_workorderservicetaskCreate))
        {
            //if (secure != null &&!secure.Equals(string.Empty))
            //{

            //}
        }

        /// <summary>
        /// Main entry point for he business logic that the plug-in is to execute.
        /// </summary>
        /// <param name="localContext">The <see cref="LocalPluginContext"/> which contains the
        /// <see cref="IPluginExecutionContext"/>,
        /// <see cref="IOrganizationService"/>
        /// and <see cref="ITracingService"/>
        /// </param>
        /// <remarks>
        /// For improved performance, Microsoft Dynamics 365 caches plug-in instances.
        /// The plug-in's Execute method should be written to be stateless as the constructor
        /// is not called for every invocation of the plug-in. Also, multiple system threads
        /// could execute the plug-in at the same time. All per invocation state information
        /// is stored in the context. This means that you should not use global variables in plug-ins.
        /// </remarks>
        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException("localContext");
            }

            IPluginExecutionContext context = localContext.PluginExecutionContext;
            ITracingService tracingService = localContext.TracingService;
            Entity target = (Entity)context.InputParameters["Target"];

            //Entity postImageEntity = (context.PostEntityImages != null && context.PostEntityImages.Contains(this.postImageAlias)) ? context.PostEntityImages[this.postImageAlias] : null;
            tracingService.Trace("Entering ExecuteCrmPlugin method.");
            try
            {
                // Log the system username at the start
                var systemUser = localContext.OrganizationService.Retrieve("systemuser", context.InitiatingUserId, new ColumnSet("fullname"));
                tracingService.Trace("Plugin executed by user: {0}", systemUser.GetAttributeValue<string>("fullname"));

                tracingService.Trace("PostOperationmsdyn_workorderservicetaskCreate: Begin.");
                if (target.LogicalName.Equals(msdyn_workorderservicetask.EntityLogicalName))
                {
                    if (target.Attributes.Contains("msdyn_tasktype") && target.Attributes["msdyn_tasktype"] != null)
                    {
                        EntityReference tasktype = (EntityReference)target.Attributes["msdyn_tasktype"];
                        using (var servicecontext = new Xrm(localContext.OrganizationService))
                        {
                            tracingService.Trace("Task type found: {0}", tasktype.Id);
                            var rclegislations = (from tt in servicecontext.ovs_msdyn_servicetasktype_qm_rclegislationSet
                                                  join le in servicecontext.qm_rclegislationSet
                                                  on tt.qm_rclegislationid.Value equals le.qm_rclegislationId.Value
                                                   where tt.msdyn_servicetasktypeid == tasktype.Id
                                                   select new
                                                   {
                                                       tt.qm_rclegislationid,
                                                       le.qm_name
                                                   }).ToList();
                            foreach (var le in rclegislations)
                            {
                                if (le.qm_rclegislationid != null)
                                {
                                    ovs_workorderservicetaskprovision wp = new ovs_workorderservicetaskprovision();
                                    wp.ovs_ProvisionId = new EntityReference(qm_rclegislation.EntityLogicalName, le.qm_rclegislationid.Value);
                                    wp.ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, target.Id);
                                    wp.ovs_name = le.qm_name;
                                    localContext.OrganizationService.Create(wp);
                                }
                            }
                        }
                    }

                    tracingService.Trace("Find out if the Work Order has a recored in ts_sharepointfile.");
                    {
                        using (var servicecontext = new Xrm(localContext.OrganizationService))
                        {
                            msdyn_workorderservicetask workOrderServiceTask = target.ToEntity<msdyn_workorderservicetask>();

                            tracingService.Trace("Get the selected Work Order.");
                            var workOrder = servicecontext.msdyn_workorderSet.Where(wo => wo.Id == workOrderServiceTask.msdyn_WorkOrder.Id).FirstOrDefault();

                            // Log the Work Order Service Task Name
                            tracingService.Trace("Work Order Service Task Name: {0}", workOrderServiceTask.msdyn_name);

                            tracingService.Trace("Check if the Work Order has a SharePoint File.");
                            var myWorkOrderSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(servicecontext, workOrder.Id.ToString().ToUpper().Trim(), PostOperationts_sharepointfileCreate.WORK_ORDER);

                            if (myWorkOrderSharePointFile != null)
                            {
                                tracingService.Trace("Retrieve the name.");
                                string myWorkOrderServiceTaskName = servicecontext.msdyn_workorderservicetaskSet.Where(wost => wost.Id == workOrderServiceTask.Id).FirstOrDefault().msdyn_name;

                                tracingService.Trace("Retrieve the owner.");
                                string owner = PostOperationts_sharepointfileCreate.GetWorkOrderOwner(localContext.OrganizationService, workOrder.Id);

                                tracingService.Trace("Create the SharePointFile for the Work Order Service Task.");
                                Guid myWorkOrderServiceTaskSharePointFileId = PostOperationts_sharepointfileCreate.CreateSharePointFile(
                                    myWorkOrderServiceTaskName,
                                    PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK,
                                    PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK_FR,
                                    workOrderServiceTask.Id.ToString().ToUpper(),
                                    myWorkOrderServiceTaskName,
                                    owner,
                                    localContext.OrganizationService);

                                tracingService.Trace("Update the Work Order Service Tasks with the SharePointFile Group for the Work Order.");
                                localContext.OrganizationService.Update(new ts_SharePointFile
                                {
                                    Id = myWorkOrderServiceTaskSharePointFileId,
                                    ts_SharePointFileGroup = myWorkOrderSharePointFile.ts_SharePointFileGroup
                                });
                            }
                            else
                            {
                                tracingService.Trace("No SharePoint File for the Work Order");
                            }
                        };
                    }
                }
            }
            catch (Exception e)
            {
                tracingService.Trace("Exception occurred: {0}", e.ToString());
                throw new InvalidPluginExecutionException(e.Message);
            }
        }
    }
}