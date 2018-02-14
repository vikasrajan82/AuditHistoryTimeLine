// <copyright file="WorkflowContext.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Workflow Context</summary>
namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Activities;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Workflow;

    /// <summary>
    /// Workflow Context
    /// </summary>
    internal class WorkflowContext : IContext
    {
        /// <summary>
        /// Gets the Tracing Service
        /// </summary>
        /// <param name="context">Workflow Context</param>
        /// <returns>Tracing Service</returns>
        public ITracingService GetTracingService(object context)
        {
            if (context.GetType() == typeof(CodeActivityContext))
            {
                var localWorkflowContext = (CodeActivityContext)context;
                return localWorkflowContext.GetExtension<ITracingService>();
            }

            return null;
        }

        /// <summary>
        /// Gets the Organization Service
        /// </summary>
        /// <param name="context">Workflow Context</param>
        /// <returns>Organization Service</returns>
        public IOrganizationService GetOrganizationService(object context)
        {
            if (context.GetType() == typeof(CodeActivityContext))
            {
                // Create the context
                var executionContext = (CodeActivityContext)context;
                IWorkflowContext workflowContext = executionContext.GetExtension<IWorkflowContext>();
                return executionContext.GetExtension<IOrganizationServiceFactory>().CreateOrganizationService(workflowContext.UserId);
            }

            throw new ArgumentNullException("context");
        }

        /// <summary>
        /// Gets the Organization Service Factory
        /// </summary>
        /// <param name="context">Workflow Context</param>
        /// <returns>Organization Service Factory</returns>
        public IOrganizationServiceFactory GetOrganizationServiceFactory(object context)
        {
            if (context.GetType() == typeof(CodeActivityContext))
            {
                var executionContext = (CodeActivityContext)context;
                return executionContext.GetExtension<IOrganizationServiceFactory>();
            }

            throw new ArgumentNullException("context");
        }

        /// <summary>
        /// Gets the organization service for impersonating user.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="impersonatingUserId">The impersonating user identifier.</param>
        /// <returns>Organization Service</returns>
        public IOrganizationService GetOrganizationServiceForImpersonatingUser(object context, Guid impersonatingUserId)
        {
            if (context.GetType() == typeof(CodeActivityContext))
            {
                // Create the context
                var executionContext = (CodeActivityContext)context;
                return executionContext.GetExtension<IOrganizationServiceFactory>().CreateOrganizationService(impersonatingUserId);
            }

            throw new ArgumentNullException("context");
        }

        /// <summary>
        /// Gets the Pre Entity Image
        /// </summary>
        /// <param name="context">Workflow Context</param>
        /// <param name="imageAlias">Pre-Image Alias Name</param>
        /// <returns>Entity Object</returns>
        public object GetPreEntityImage(object context, string imageAlias)
        {
            throw new NotSupportedException("This is not supported within a workflow context");
        }

        /// <summary>
        /// Gets Input Parameter
        /// </summary>
        /// <param name="context">Workflow Context</param>
        /// <returns>Entity object</returns>
        public object GetTargetInputParameter(object context)
        {
            throw new NotSupportedException("This is not supported within a workflow context");
        }

        /// <summary>
        /// Gets the Input Parameters in the context
        /// </summary>
        /// <param name="context">Workflow Context</param>
        /// <param name="parameterName">Input Parameter Name</param>
        /// <returns>Entity object</returns>
        public object GetInputParameter(object context, string parameterName)
        {
            throw new NotSupportedException("This is not supported within a workflow context");
        }

        /// <summary>
        /// Gets the Post Entity Image
        /// </summary>
        /// <param name="context">Workflow Context</param>
        /// <param name="imageAlias">Post-Image Alias Name</param>
        /// <returns>Entity Object</returns>
        public object GetPostEntityImage(object context, string imageAlias)
        {
            throw new NotSupportedException("This is not supported within a workflow context");
        }

        /// <summary>
        /// Get Initiating User Id
        /// </summary>
        /// <param name="context">Workflow Context</param>
        /// <returns>Initiating User Id</returns>
        public object GetInitiatingUserId(object context)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get Execution Stage
        /// </summary>
        /// <param name="context">Workflow Context</param>
        /// <returns>Execution stage</returns>
        public string GetExecutionStage(object context)
        {
            throw new NotImplementedException();
        }
    }
}
