// <copyright file="PluginContext.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Plugin Context Class</summary>
namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Xrm.Sdk;

    /// <summary>
    /// Plugin Context
    /// </summary>
    internal class PluginContext : IContext
    {
        /// <summary>
        /// Gets the Tracing Service
        /// </summary>
        /// <param name="context">Plugin Context</param>
        /// <returns>Tracing Service</returns>
        public ITracingService GetTracingService(object context)
        {
            if (context.GetType() == typeof(BasePlugin.LocalPluginContext))
            {
                var localPluginContext = (BasePlugin.LocalPluginContext)context;
                return localPluginContext.TracingService;
            }

            return null;
        }

        /// <summary>
        /// Gets the Organization Service
        /// </summary>
        /// <param name="context">Plugin Context</param>
        /// <returns>Organization Service</returns>
        public IOrganizationService GetOrganizationService(object context)
        {
            if (context.GetType() == typeof(BasePlugin.LocalPluginContext))
            {
                var localPluginContext = (BasePlugin.LocalPluginContext)context;
                return localPluginContext.OrganizationService;
            }

            throw new ArgumentNullException("context");
        }

        /// <summary>
        /// Gets the Organization Service Factory
        /// </summary>
        /// <param name="context">Plugin Context</param>
        /// <returns>Organization Service Factory</returns>
        public IOrganizationServiceFactory GetOrganizationServiceFactory(object context)
        {
            if (context.GetType() == typeof(BasePlugin.LocalPluginContext))
            {
                var localPluginContext = (BasePlugin.LocalPluginContext)context;
                return localPluginContext.OrganizationServiceFactory;
            }

            throw new ArgumentNullException("context");
        }

        /// <summary>
        /// Gets the Pre Entity Image
        /// </summary>
        /// <param name="context">Plugin Context</param>
        /// <param name="imageAlias">Pre-Image Alias Name</param>
        /// <returns>Entity Object</returns>
        public object GetPreEntityImage(object context, string imageAlias)
        {
            if (context.GetType() == typeof(BasePlugin.LocalPluginContext))
            {
                var basePluginContext = (BasePlugin.LocalPluginContext)context;

                if (basePluginContext.PluginExecutionContext != null
                        && basePluginContext.PluginExecutionContext.PreEntityImages != null
                        && basePluginContext.PluginExecutionContext.PreEntityImages.Contains(imageAlias))
                {
                    return basePluginContext.PluginExecutionContext.PreEntityImages[imageAlias];
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the Post Entity Image
        /// </summary>
        /// <param name="context">Plugin Context</param>
        /// <param name="imageAlias">Post-Image Alias Name</param>
        /// <returns>Entity Object</returns>
        public object GetPostEntityImage(object context, string imageAlias)
        {
            if (context.GetType() == typeof(BasePlugin.LocalPluginContext))
            {
                var basePluginContext = (BasePlugin.LocalPluginContext)context;

                if (basePluginContext.PluginExecutionContext != null
                        && basePluginContext.PluginExecutionContext.PostEntityImages != null
                        && basePluginContext.PluginExecutionContext.PostEntityImages.Contains(imageAlias))
                {
                    return basePluginContext.PluginExecutionContext.PostEntityImages[imageAlias];
                }
            }

            return null;
        }

        /// <summary>
        /// Gets Input Parameter
        /// </summary>
        /// <param name="context">Plugin Context</param>
        /// <returns>Entity object</returns>
        public object GetTargetInputParameter(object context)
        {
            return this.GetInputParameter(context, "Target");
        }

        /// <summary>
        /// Gets the Input Parameters in the context
        /// </summary>
        /// <param name="context">Plugin Context</param>
        /// <param name="parameterName">Input Parameter Name</param>
        /// <returns>Entity object</returns>
        public object GetInputParameter(object context, string parameterName)
        {
            // TODO: Needs to be extended for other input parameter types
            if (context.GetType() == typeof(BasePlugin.LocalPluginContext))
            {
                var basePluginContext = (BasePlugin.LocalPluginContext)context;

                if (basePluginContext.PluginExecutionContext != null
                         && basePluginContext.PluginExecutionContext.InputParameters.Contains(parameterName))
                {
                    return basePluginContext.PluginExecutionContext.InputParameters[parameterName];
                }
            }

            return null;
        }

        /// <summary>
        /// Get Initiating User Id
        /// </summary>
        /// <param name="context">Base Plugin Context</param>
        /// <returns>Initiating User Id</returns>
        public object GetInitiatingUserId(object context)
        {
            if (context.GetType() == typeof(BasePlugin.LocalPluginContext))
            {
                var basePluginContext = (BasePlugin.LocalPluginContext)context;
                return basePluginContext.PluginExecutionContext.InitiatingUserId;
            }

            return Guid.Empty;
        }

        /// <summary>
        /// Get Execution Stage
        /// </summary>
        /// <param name="context">Base Plugin Context</param>
        /// <returns>Execution Stage</returns>
        public string GetExecutionStage(object context)
        {
            if (context.GetType() == typeof(BasePlugin.LocalPluginContext))
            {
                var basePluginContext = (BasePlugin.LocalPluginContext)context;
                return basePluginContext.PluginExecutionContext.MessageName;
            }

            return string.Empty;
        }

        /// <summary>
        /// Gets the organization service for impersonating user.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="impersonatingUserId">The impersonating user identifier.</param>
        /// <returns>
        /// Organization Service
        /// </returns>
        public IOrganizationService GetOrganizationServiceForImpersonatingUser(object context, Guid impersonatingUserId)
        {
            if (context.GetType() == typeof(BasePlugin.LocalPluginContext))
            {
                var localPluginContext = (BasePlugin.LocalPluginContext)context;
                return localPluginContext.OrganizationService;
            }

            throw new ArgumentNullException("context");
        }
    }
}
