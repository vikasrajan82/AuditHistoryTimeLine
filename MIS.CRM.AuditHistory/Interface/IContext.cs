// <copyright file="IContext.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Interface for the context</summary>
namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Xrm.Sdk;
    
    /// <summary>
    /// Context interface
    /// </summary>
    internal interface IContext
    {
        /// <summary>
        /// Gets the Tracing Service
        /// </summary>
        /// <param name="context">Workflow or Plugin Context</param>
        /// <returns>Tracing Service</returns>
        ITracingService GetTracingService(object context);

        /// <summary>
        /// Gets the Organization Service
        /// </summary>
        /// <param name="context">Workflow or Plugin Context</param>
        /// <returns>Organization Service</returns>
        IOrganizationService GetOrganizationService(object context);

        /// <summary>
        /// Gets the Organization Service Factory
        /// </summary>
        /// <param name="context">Workflow or Plugin Context</param>
        /// <returns>Organization Service Factory</returns>
        IOrganizationServiceFactory GetOrganizationServiceFactory(object context);

        /// <summary>
        /// Gets the organization service for impersonating user.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="impersonatingUserId">The impersonating user identifier.</param>
        /// <returns>Organization Service</returns>
        IOrganizationService GetOrganizationServiceForImpersonatingUser(object context, Guid impersonatingUserId);

        /// <summary>
        /// Gets the Pre Entity Image
        /// </summary>
        /// <param name="context">Workflow or Plugin Context</param>
        /// <param name="imageAlias">Pre-Image Alias Name</param>
        /// <returns>Entity Object</returns>
        object GetPreEntityImage(object context, string imageAlias);

        /// <summary>
        /// Gets the Post Entity Image
        /// </summary>
        /// <param name="context">Workflow or Plugin Context</param>
        /// <param name="imageAlias">Post-Image Alias Name</param>
        /// <returns>Entity Object</returns>
        object GetPostEntityImage(object context, string imageAlias);

        /// <summary>
        /// Gets Input Parameter
        /// </summary>
        /// <param name="context">Workflow or Plugin Context</param>
        /// <returns>Entity object</returns>
        object GetTargetInputParameter(object context);

        /// <summary>
        /// Gets Input Parameter
        /// </summary>
        /// <param name="context">Workflow or Plugin Context</param>
        /// <param name="parameterName">Input Parameter Name</param>
        /// <returns>Entity object</returns>
        object GetInputParameter(object context, string parameterName);

        /// <summary>
        /// Gets the Initiating User Id
        /// </summary>
        /// <param name="context">Workflow or Plugin Context</param>
        /// <returns>Initiating User Id</returns>
        object GetInitiatingUserId(object context);

        /// <summary>
        /// Get Execution Stage
        /// </summary>
        /// <param name="context">Workflow Context</param>
        /// <returns>Execution stage</returns>
        string GetExecutionStage(object context);
    }
}
