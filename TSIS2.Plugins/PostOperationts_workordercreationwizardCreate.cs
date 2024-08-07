﻿// <copyright file="PostOperationts_workordercreationwizardCreate.cs" company="">
// Copyright (c) 2018 All Rights Reserved
// </copyright>
// <author>Hong Liu</author>
// <date>9/20/2018 10:21:27 AM</date>
// <summary>Implements the PostOperationts_workordercreationwizardCreate Plugin.</summary>
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.1
// </auto-generated>

using System;
using System.Linq;
using Microsoft.Xrm.Sdk;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
    MessageNameEnum.Create,
    "ts_workordercreationwizard",
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PostOperationts_workordercreationwizardCreate Plugin",
    1,
    IsolationModeEnum.Sandbox, 
    Description = "Create a list of Activity Types for the new Work Order Wizard.")]
    /// <summary>
    /// PostOperationts_workordercreationwizardCreate Plugin.
    /// </summary>    
    public class PostOperationts_workordercreationwizardCreate : PluginBase
    {
        //private readonly string postImageAlias = "PostImage";

        /// <summary>
        /// Initializes a new instance of the <see cref="PostOperationts_workordercreationwizardCreate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PostOperationts_workordercreationwizardCreate(string unsecure, string secure)
            : base(typeof(PostOperationts_workordercreationwizardCreate))
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
            Entity target = (Entity)context.InputParameters["Target"];
            //Entity postImageEntity = (context.PostEntityImages != null && context.PostEntityImages.Contains(this.postImageAlias)) ? context.PostEntityImages[this.postImageAlias] : null;

            try
            {
                if (target.LogicalName.Equals(ts_workordercreationwizard.EntityLogicalName))
                {
                    if (target.Attributes.Contains("ts_workordertypeid") && target.Attributes["ts_workordertypeid"] != null && target.Attributes.Contains("ts_operationtypeid") && target.Attributes["ts_operationtypeid"] != null)
                    {
                        EntityReference workordertype = (EntityReference)target.Attributes["ts_workordertypeid"];
                        EntityReference operationtype = (EntityReference)target.Attributes["ts_operationtypeid"];
                        using (var servicecontext = new Xrm(localContext.OrganizationService))
                        {
                            var incidenttypes = (from tt in servicecontext.msdyn_incidenttypeSet
                                                 where tt.msdyn_DefaultWorkOrderType.Id == workordertype.Id
                                                 // Commenting out for now due to remodeling of Operation Type relationship
                                                 //where tt.ts_ovs_operationtype.Id == operationtype.Id
                                                 where tt.statecode == 0
                                                 select new
                                                   {
                                                       tt.msdyn_incidenttypeId,
                                                       tt.msdyn_name
                                                   }).ToList();
                            foreach (var le in incidenttypes)
                            {
                                if (le.msdyn_incidenttypeId != null)
                                {
                                    ts_workorderactivitytype wa = new ts_workorderactivitytype();
                                    wa.ts_ActivityTypeId = new EntityReference(msdyn_incidenttype.EntityLogicalName, le.msdyn_incidenttypeId.Value);
                                    wa.ts_WorkOrderWizardId = new EntityReference(ts_workordercreationwizard.EntityLogicalName, target.Id);
                                    wa.ts_name = le.msdyn_name;
                                    localContext.OrganizationService.Create(wa);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }
    }
}