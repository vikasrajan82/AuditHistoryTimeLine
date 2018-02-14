// <copyright file="RetrieveAttributeResponseWrapper.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author></author>
// <date>9/15/2015 12:09:38 PM</date>
// <summary>Retrieve Attribute Response Wrapper</summary>
//--------------------------------------------------------------------------------
namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.Serialization;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Metadata;

    /// <summary>  
    /// Wrapper class for the Retrieve Attribute Response class. Primarily used to support during testing.  
    /// </summary>  
    [DataContract(Namespace = "http://schemas.microsoft.com/xrm/2011/Contracts")]
    public class RetrieveAttributeResponseWrapper : OrganizationResponse
    {
        /// <summary>
        /// meta data attribute
        /// </summary>
        private AttributeMetadata metadata;

        /// <summary>
        /// Initializes a new instance of the RetrieveAttributeResponseWrapper class
        /// </summary>
        /// <param name="response">organization response</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "In case any exception, existing flow should not break")]
        public RetrieveAttributeResponseWrapper(OrganizationResponse response)
        {
            try
            {
                this.metadata = ((RetrieveAttributeResponseWrapper)response).AttributeMetadata;
            }
            catch
            {
                this.metadata = ((RetrieveAttributeResponse)response).AttributeMetadata;              
            }
        }
    
        /// <summary>
        /// Gets or sets Attribute Metadata
        /// </summary>
        public AttributeMetadata AttributeMetadata
        {
            get
            {
                return this.metadata;
            }

            set
            {
                this.metadata = value;
            }
        }
    }
}
