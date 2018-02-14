// <copyright file="Extension.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Base for all extensions</summary>
namespace MIS.CRM.AuditHistory.BusinessProcesses
{   
    using System;
    using System.Activities;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Workflow;

    /// <summary>
    /// Extension Type
    /// </summary>
    public enum ExtensionType
    {
        /// <summary>
        /// Plugin Type
        /// </summary>
        Plugin,

        /// <summary>
        /// Workflow Type
        /// </summary>
        Workflow
    }

    /// <summary>
    /// Base Class for all extensions
    /// </summary>
    public class Extension : IDisposable
    {
        /// <summary>
        /// Checks if the class is disposed
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Execution Context 
        /// </summary>
        private IContext executionContext;

        /// <summary>
        /// Extension Type
        /// </summary>
        private ExtensionType extensionType;

        /// <summary>
        /// Pre-Image Alias
        /// </summary>
        private string preImageAlias;

        /// <summary>
        /// Post-Image Alias
        /// </summary>
        private string postImageAlias;

        /// <summary>
        /// The unsecure configuration
        /// </summary>
        private string unsecureConfig;

        /// <summary>
        /// The secure configuration
        /// </summary>
        private string secureConfig;

        /// <summary>
        /// The current date time
        /// </summary>
        private DateTime currentDateTime;

        /// <summary>
        /// Exception Handler
        /// </summary>
        private IExceptionManager exceptionHandler;

        /// <summary>
        /// Initializes a new instance of the <see cref="Extension"/> class
        /// </summary>
        protected Extension()
        {
            this.CrmHandler = new CrmConnection();
            this.exceptionHandler = new ExceptionManager();
            this.FetchXmlHandler = new FetchXmlConnection();
        }
  
        /// <summary>
        /// Finalizes an instance of the <see cref="Extension"/> class.
        /// </summary>
        ~Extension()
        {
            this.Dispose(false);
        }

        /// <summary>
        /// Gets the Plugin Context
        /// </summary>
        public int ContextDepth
        {
            get
            {
                return this.BasePluginContext.PluginExecutionContext.Depth;
            }
        }

        /// <summary>
        /// Gets the Plugin Context
        /// </summary>
        public string ContextMessage
        {
            get
            {
                return this.BasePluginContext.PluginExecutionContext.MessageName;
            }
        }

        /// <summary>
        /// Gets the Initiating User Id
        /// </summary>
        public object InitiatingUserId
        {
            get
            {
                switch (this.extensionType)
                {
                    case ExtensionType.Plugin:
                        return this.executionContext.GetInitiatingUserId(this.BasePluginContext);
                    case ExtensionType.Workflow:
                        return this.WorkflowContext.InitiatingUserId;
                }

                return null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this is a create message
        /// </summary>
        public bool IsCreateMessage
        {
            get
            {
                return this.executionContext.GetExecutionStage(this.BasePluginContext).ToUpperInvariant() == "CREATE";
            }
        }

        /// <summary>
        /// Gets a value indicating whether this is a update message
        /// </summary>
        public bool IsUpdateMessage
        {
            get
            {
                return this.executionContext.GetExecutionStage(this.BasePluginContext).ToUpperInvariant() == "UPDATE";
            }
        }

        /// <summary>
        /// Gets a value indicating whether this is a delete message
        /// </summary>
        public bool IsDeleteMessage
        {
            get
            {
                return this.executionContext.GetExecutionStage(this.BasePluginContext).ToUpperInvariant() == "DELETE";
            }
        }

        /// <summary>
        /// Gets a value indicating whether this is a Associate message
        /// </summary>
        public bool IsAssociateMessage
        {
            get
            {
                return this.executionContext.GetExecutionStage(this.BasePluginContext).ToUpperInvariant() == "ASSOCIATE";
            }
        }

        /// <summary>
        /// Gets a value indicating whether this is a Disassociate message
        /// </summary>
        public bool IsDisassociateMessage
        {
            get
            {
                return this.executionContext.GetExecutionStage(this.BasePluginContext).ToUpperInvariant() == "DISASSOCIATE";
            }
        }

        /// <summary>
        /// Gets the Current Date Time
        /// </summary>
        protected DateTime CurrentDateTime
        {
            get
            {
                if (this.currentDateTime == DateTime.MinValue)
                {
                    this.currentDateTime = DateTime.UtcNow.LocalDateTime(this.OrgService);
                }

                return this.currentDateTime;
            }
        }

        /// <summary>
        /// Gets the unsecure configuration.
        /// </summary>
        /// <value>
        /// The unsecure configuration.
        /// </value>
        protected string UnsecureConfig
        {
            get
            {
                return this.unsecureConfig;
            }
        }

        /// <summary>
        /// Gets the secure configuration.
        /// </summary>
        /// <value>
        /// The secure configuration.
        /// </value>
        protected string SecureConfig
        {
            get
            {
                return this.secureConfig;
            }
        }

        /// <summary>
        /// Gets or sets Interface reference to CRM Handler Object
        /// </summary>
        protected ICrmConnection CrmHandler
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets Interface reference to Fetch Xml Handler Object
        /// </summary>
        protected IFetchXmlConnection FetchXmlHandler
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the Organization Service
        /// </summary>
        protected IOrganizationService OrgService
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the Organization Service Factory
        /// </summary>
        protected IOrganizationServiceFactory OrgServiceFactory
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the Tracing Service
        /// </summary>
        protected ITracingService TracingService
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the Base Plugin Context
        /// </summary>
        protected BasePlugin.LocalPluginContext BasePluginContext
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the workflow code activity context.
        /// </summary>
        /// <value>
        /// The workflow code activity context.
        /// </value>
        protected CodeActivityContext CodeActivityContext
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the Workflow Context
        /// </summary>
        protected IWorkflowContext WorkflowContext
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the exception handler reference
        /// </summary>
        protected IExceptionManager ExceptionHandler
        {
            get
            {
                return this.exceptionHandler;
            }
        }

        /// <summary>
        /// Gets the input parameter for target (Entity)
        /// </summary>
        protected Entity GetEntityInputParameter
        {
            get
            {
                return (Entity)this.executionContext.GetTargetInputParameter(this.BasePluginContext);
            }
        }

        /// <summary>
        /// Gets the input parameter for target (Entity Reference)
        /// </summary>
        protected EntityReference GetEntityReferenceInputParameter
        {
            get
            {
                return (EntityReference)this.executionContext.GetTargetInputParameter(this.BasePluginContext);
            }
        }

        /// <summary>
        /// Gets the Pre-Entity image
        /// </summary>
        protected Entity GetPreEntityImage
        {
            get
            {
                return (Entity)this.executionContext.GetPreEntityImage(this.BasePluginContext, this.preImageAlias);
            }
        }

        /// <summary>
        /// Gets the Post-Entity image
        /// </summary>
        protected Entity GetPostEntityImage
        {
            get
            {
                return (Entity)this.executionContext.GetPostEntityImage(this.BasePluginContext, this.postImageAlias);
            }
        }

        /// <summary>
        /// Public Method to check if the arguments are null
        /// </summary>
        /// <param name="value">Method argument</param>
        /// <param name="ex">Exception to be returned if the parameter is null</param>
        /// <returns>Indicates if the object is null</returns>
        public static bool IsObjectNotNull(object value, Exception ex)
        {
            if (value == null)
            {
                throw ex;
            }

            return true;
        }

        /// <summary>
        /// Get Boolean Value
        /// </summary>
        /// <param name="boolObject">Boolean Object</param>
        /// <returns>Boolean value</returns>
        public static bool GetBooleanValue(bool? boolObject)
        {
            return boolObject.HasValue && boolObject.Value;
        }

        /// <summary>
        /// Get Aliased Value
        /// </summary>
        /// <param name="entity">Entity Object</param>
        /// <param name="aliasedAttributeName">Aliased attribute name</param>
        /// <returns>Object value</returns>
        public static object GetAliasedValue(Entity entity, string aliasedAttributeName)
        {
            if (entity != null && entity.Contains(aliasedAttributeName) && entity.Attributes[aliasedAttributeName] != null)
            {
                return ((AliasedValue)entity.Attributes[aliasedAttributeName]).Value;
            }

            return null;
        }

        /// <summary>
        /// Get Aliased Value
        /// </summary>
        /// <typeparam name="T">Type parameter</typeparam>
        /// <param name="entity">Entity object</param>
        /// <param name="aliasName">Alias name</param>
        /// <param name="attributeName">Attribute name</param>
        /// <returns>An instance of value of type specified by Type parameter</returns>
        public static T GetAliasedValue<T>(Entity entity, string aliasName, string attributeName)
        {
            string aliasedAttributeName = string.Format(CultureInfo.InvariantCulture, "{0}.{1}", aliasName, attributeName);
            var aliasedValue = Extension.GetAliasedValue(entity, aliasedAttributeName);
            if (aliasedValue != null && aliasedValue is T)
            {
                return (T)aliasedValue;
            }

            return default(T);
        }

        /// <summary>
        /// Constructs the early bound entity.
        /// </summary>
        /// <typeparam name="T">The type parameter for the early bound entity</typeparam>
        /// <param name="dataEntity">The data entity</param>
        /// <param name="entityAliasName">Entity alias name</param>
        /// <returns>An early bound entity of the specified type</returns>
        public static T ConstructEarlyBoundEntity<T>(Entity dataEntity, string entityAliasName) where T : Entity, new()
        {
            if (dataEntity == null)
            {
                return null;
            }

            var attributeNames = new List<string>();
            var idAttributeName = string.Empty;
            var entityAliasNamePresent = !string.IsNullOrWhiteSpace(entityAliasName);
            foreach (var property in typeof(T).GetProperties())
            {
                var customAttributes = property.GetCustomAttributes(typeof(AttributeLogicalNameAttribute), false);
                if (customAttributes.Count() > 0)
                {
                    var attributeLogicalName = ((AttributeLogicalNameAttribute)customAttributes[0]).LogicalName;
                    if (property.Name == "Id")
                    {
                        idAttributeName = attributeLogicalName;
                    }

                    var attributeKey = attributeLogicalName;
                    if (entityAliasNamePresent)
                    {
                        attributeKey = string.Format(CultureInfo.InvariantCulture, "{0}.{1}", entityAliasName, attributeLogicalName);
                    }

                    if (attributeLogicalName == idAttributeName || dataEntity.Contains(attributeKey))
                    {
                        attributeNames.Add(attributeLogicalName);
                    }
                }
            }

            attributeNames = attributeNames.Distinct().ToList();

            var entity = new T();
            Action<string> setIdAttributeAction = (attributeName) =>
            {
                Guid? id;
                if (entityAliasNamePresent)
                {
                    id = Extension.GetAliasedValue<Guid?>(dataEntity, entityAliasName, attributeName);
                }
                else
                {
                    id = dataEntity.GetAttributeValue<Guid?>(attributeName) ?? dataEntity.Id;
                }

                if (id.HasValue)
                {
                    entity.Id = id.Value;
                }
            };

            Action<string> setAttributeAction = (attributeName) =>
            {
                if (entityAliasNamePresent)
                {
                    entity.Attributes.Add(attributeName, Extension.GetAliasedValue<object>(dataEntity, entityAliasName, attributeName));
                }
                else
                {
                    entity.Attributes.Add(attributeName, dataEntity.GetAttributeValue<object>(attributeName));
                }
            };

            foreach (var attributeName in attributeNames)
            {
                if (attributeName == idAttributeName)
                {
                    setIdAttributeAction(attributeName);
                }
                else
                {
                    setAttributeAction(attributeName);
                }
            }

            return entity;
        }

        /// <summary>
        /// Constructs the early bound entity.
        /// </summary>
        /// <typeparam name="T">The type parameter for the early bound entity</typeparam>
        /// <param name="dataEntity">The data entity</param>
        /// <returns>An early bound entity of the specified type</returns>
        public static T ConstructEarlyBoundEntity<T>(Entity dataEntity) where T : Entity, new()
        {
            return Extension.ConstructEarlyBoundEntity<T>(dataEntity, null);
        }

        /// <summary>
        /// Sets the Plugin Context
        /// </summary>
        /// <param name="pluginContext">Plugin Context Object</param>
        public void SetPluginContext(object pluginContext)
        {
            this.extensionType = ExtensionType.Plugin;

            if (pluginContext != null)
            {
                this.executionContext = new PluginContext();

                this.BasePluginContext = (BasePlugin.LocalPluginContext)pluginContext;
                this.TracingService = this.executionContext.GetTracingService(pluginContext);
                this.OrgService = this.executionContext.GetOrganizationService(pluginContext);
                this.OrgServiceFactory = this.executionContext.GetOrganizationServiceFactory(pluginContext);
            }
        }

        /// <summary>
        /// Sets the Workflow Context
        /// </summary>
        /// <param name="workflowContext">Workflow Context Object</param>
        public void SetWorkflowContext(object workflowContext)
        {
            this.extensionType = ExtensionType.Workflow;

            if (workflowContext != null)
            {
                this.executionContext = new WorkflowContext();
                this.CodeActivityContext = (CodeActivityContext)workflowContext;
                this.WorkflowContext = this.CodeActivityContext.GetExtension<IWorkflowContext>();
                this.TracingService = this.executionContext.GetTracingService(workflowContext);
                this.OrgService = this.executionContext.GetOrganizationService(workflowContext);
                this.OrgServiceFactory = this.executionContext.GetOrganizationServiceFactory(workflowContext);
            }
        }

        /// <summary>
        /// Sets the plugin image
        /// </summary>
        /// <param name="preImageName">Pre-Image for the entity</param>
        public void SetPluginPreImage(string preImageName)
        {
            this.preImageAlias = preImageName;
        }

        public void SetOrganizationService(IOrganizationService orgService)
        {
            this.OrgService = orgService;
        }

        /// <summary>
        /// Sets the plugin image
        /// </summary>
        /// <param name="postImageName">Post-Image for the entity</param>
        public void SetPluginPostImage(string postImageName)
        {
            this.postImageAlias = postImageName;
        }

        /// <summary>
        /// Sets the plugin unsecure configuration.
        /// </summary>
        /// <param name="unsecureConfiguration">The unsecure configuration.</param>
        public void SetPluginUnsecureConfig(string unsecureConfiguration)
        {
            this.unsecureConfig = unsecureConfiguration;
        }

        /// <summary>
        /// Sets the plugin secure configuration.
        /// </summary>
        /// <param name="secureConfiguration">The secure configuration.</param>
        public void SetPluginSecureConfig(string secureConfiguration)
        {
            this.secureConfig = secureConfiguration;
        }

        /// <summary>
        /// Start Logging
        /// </summary>
        public void StartLogging()
        {
            if (this.extensionType == ExtensionType.Plugin)
            {
                this.BasePluginContext.StartLogging();
            }
        }

        /// <summary>
        /// End Logging
        /// </summary>
        public void EndLogging()
        {
            if (this.extensionType == ExtensionType.Plugin)
            {
                this.BasePluginContext.EndLogging();
            }
        }

        /// <summary>
        /// Dispose the object
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Logs the Message
        /// </summary>
        /// <param name="type">Log Type</param>
        /// <param name="message">Message to be logged</param>
        /// <param name="args">Values to be substituted in the message</param>
        public void LogMessage(LogType type, string message, params object[] args)
        {
            if (this.TracingService != null)
            {
                Logger.Log(type, ExtensionBase.ConcatenatedString(message, args), this.TracingService);
            }
        }

        /// <summary>
        /// Dispose the Object
        /// </summary>
        /// <param name="disposing">Indicates if the object has disposed</param>
        protected virtual void Dispose(bool disposing)
        {
            if (this.disposed)
            {
                return;
            }

            if (disposing)
            {
                this.BasePluginContext = null;
                this.TracingService = null;
                this.OrgService = null;
                this.CrmHandler = null;
            }

            this.disposed = true;
        }

        /// <summary>
        /// Gets Input Parameter Value
        /// </summary>
        /// <param name="parameterValue">Parameter to be retrieved</param>
        /// <returns>Parameter Value</returns>
        protected object GetInputParameter(string parameterValue)
        {
            return this.executionContext.GetInputParameter(this.BasePluginContext, parameterValue);
        }
    }
}