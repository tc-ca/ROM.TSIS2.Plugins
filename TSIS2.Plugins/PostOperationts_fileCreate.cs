﻿using System;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
        MessageNameEnum.Create,
        "ts_file",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperationts_fileCreate Plugin",
        1,
        IsolationModeEnum.Sandbox,
        Description = "After File created, automatically give the requested team(s) access to the uploaded file")]
    public class PostOperationts_fileCreate : IPlugin
    {
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
                    if (target.LogicalName.Equals(ts_File.EntityLogicalName))
                    {
                        ts_File myFile = target.ToEntity<ts_File>();

                        // if the uploaded file is visible to other programs
                        if (myFile.ts_VisibletoOtherPrograms == true)
                        {
                            // retrieve the list of teams that the file is being shared with
                            var teams = myFile.ts_ProgramAccessTeamNameID.Split(',').ToList();

                            // drop the last item in the array if it is empty
                            if (String.IsNullOrWhiteSpace(teams[teams.Count - 1]))
                                teams.RemoveAt(teams.Count - 1);

                            // go through each team that was selected and share the uploaded file with that team
                            foreach (var teamItem in teams)
                            {
                                // create the team entity ref
                                var teamID = new Guid(teamItem);
                                var teamRef = new EntityReference("team", teamID);

                                // create the file entity ref
                                var fileRef = new EntityReference(target.LogicalName, target.Id);

                                // create the grant access request for the file entity
                                var grantAccess = new GrantAccessRequest
                                {
                                    PrincipalAccess = new PrincipalAccess
                                    {
                                        AccessMask = AccessRights.ReadAccess,
                                        Principal = teamRef
                                    },
                                    Target = fileRef
                                };

                                // set the many-to-many with team
                                EntityReferenceCollection relatedEntities = new EntityReferenceCollection();

                                relatedEntities.Add(teamRef);

                                Relationship relationship = new Relationship("ts_File_Team_Team");

                                service.Associate(fileRef.LogicalName,fileRef.Id,relationship,relatedEntities);

                                // give the team access to the uploaded file
                                service.Execute(grantAccess);
                            }
                        }
                    }
                }
                catch (NotImplementedException ex)
                {
                    // If exceptions from mocking library, just continue. If not from there, we should still throw the error.
                    if (ex.Source == "XrmMockup365" && ex.Message == "No implementation for expression operator 'Count'")
                    {
                        // continue
                    }
                    else
                    {
                        tracingService.Trace("PostOperationts_fileCreate Plugin: {0}", ex.ToString());
                        throw;
                    }
                }
                catch (Exception ex)
                {
                    tracingService.Trace("PostOperationts_fileCreate Plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}