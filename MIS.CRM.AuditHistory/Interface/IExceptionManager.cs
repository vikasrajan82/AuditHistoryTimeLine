// <copyright file="IExceptionManager.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Crm Connection Interface</summary>
namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Xrm.Sdk;

    /// <summary>
    /// Exception Manager Interface
    /// </summary>
    public interface IExceptionManager
    {
        /// <summary>
        /// Throws an Argument Null Exception
        /// </summary>
        /// <param name="parameterName">Parameter Name</param>
        /// <param name="message">Message to be included in the exception</param>
        /// <returns>Argument Null Exception</returns>
        ArgumentNullException ArgumentNullException(string parameterName, string message);

        /// <summary>
        /// Throws an Argument Null Exception
        /// </summary>
        /// <param name="parameterName">Parameter Name</param>
        /// <param name="message">Message to be included in the exception</param>
        /// <param name="args">Values to be substituted in the message</param>
        /// <returns>Argument Null Exception</returns>
        ArgumentNullException ArgumentNullException(string parameterName, string message, params object[] args);

        /// <summary>
        /// Throws an Invalid Plugin Exception
        /// </summary>
        /// <param name="message">Message to be included in the exception</param>
        /// <returns>Invalid Plugin Exception</returns>
        InvalidPluginExecutionException InvalidPluginExecutionException(string message);

        /// <summary>
        /// Throws an Invalid Plugin Exception
        /// </summary>
        /// <param name="message">Message to be included in the exception</param>
        /// <param name="args">Values to be substituted in the message</param>
        /// <returns>Invalid Plugin Exception</returns>
        InvalidPluginExecutionException InvalidPluginExecutionException(string message, params object[] args);

        /// <summary>
        /// Throws an Argument Out of Range Exception
        /// </summary>
        /// <param name="parameterName">Parameter Name</param>
        /// <param name="message">Message to be included in the exception</param>
        /// <returns>Argument Out of Range Exception</returns>
        ArgumentOutOfRangeException ArgumentOutOfRangeException(string parameterName, string message);

        /// <summary>
        /// Throws an Argument Out of Range Exception
        /// </summary>
        /// <param name="parameterName">Parameter Name</param>
        /// <param name="message">Message to be included in the exception</param>
        /// <param name="args">Values to be substituted in the message</param>
        /// <returns>Argument Out of Range Exception</returns>
        ArgumentOutOfRangeException ArgumentOutOfRangeException(string parameterName, string message, params object[] args);
    }
}
