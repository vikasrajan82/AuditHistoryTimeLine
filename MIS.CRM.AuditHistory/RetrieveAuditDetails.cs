// <copyright file="StartRecertificationProcess.cs" company="MS">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author></author>
// <date>12/22/2015 12:43:59 PM</date>
// <summary>Implements the StartRecertificationProcess Workflow Activity.</summary>
namespace MIS.CRM.AuditHistory
{
    using System;
    using System.Activities;
    using System.Globalization;
    using System.ServiceModel;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Workflow;
    using MIS.CRM.AuditHistory.AuditHistory;


    /// <summary>
    /// Start recertification process
    /// </summary>
    public sealed class RetrieveAuditDetails : CodeActivity
    {
        /// <summary>
        /// Gets or sets whether participant is pregnant
        /// </summary>
        [Input("Entity Name")]
        [RequiredArgument]
        public InArgument<string> EntityName { get; set; }

        /// <summary>
        /// Gets or sets whether participant is pregnant
        /// </summary>
        [Input("Entity Id")]
        [RequiredArgument]
        public InArgument<string> EntityId { get; set; }

        /// <summary>
        /// Gets or sets whether participant is pregnant
        /// </summary>
        [Output("Audit Details")]
        [RequiredArgument]
        public OutArgument<string> AuditDetails { get; set; }

        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext)
        {
            using (AuditHistoryExtension ext = new AuditHistoryExtension())
            {
                ext.SetWorkflowContext(executionContext);

                ext.SetEntityDetails(this.EntityName.Get(executionContext), this.EntityId.Get(executionContext));
                
                ext.RetrieveAuditDetails();

                this.AuditDetails.Set(executionContext, ext.ConvertToJsonString());
            }
        }
    }
}