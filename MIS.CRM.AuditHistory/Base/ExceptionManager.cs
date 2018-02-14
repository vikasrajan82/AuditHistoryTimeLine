// <copyright file="ExceptionManager.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Manages the Exception Generation</summary>
namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Xrm.Sdk;

    /// <summary>
    /// Exception Manager Class
    /// </summary>
    public class ExceptionManager : IExceptionManager
    {
        /// <summary>
        /// Throws an Argument Null Exception
        /// </summary>
        /// <param name="parameterName">Parameter Name</param>
        /// <param name="message">Message to be included in the exception</param>
        /// <returns>Argument Null Exception</returns>
        public ArgumentNullException ArgumentNullException(string parameterName, string message)
        {
            return this.ArgumentNullException(parameterName, message, null);
        }

        /// <summary>
        /// Throws an Argument Null Exception
        /// </summary>
        /// <param name="parameterName">Parameter Name</param>
        /// <param name="message">Message to be included in the exception</param>
        /// <param name="args">Values to be substituted in the message</param>
        /// <returns>Argument Null Exception</returns>
        public ArgumentNullException ArgumentNullException(string parameterName, string message, params object[] args)
        {
            return new ArgumentNullException(parameterName, ExtensionBase.ConcatenatedString(message, args));
        }

        /// <summary>
        /// Throws an Invalid Plugin Exception
        /// </summary>
        /// <param name="message">Message to be included in the exception</param>
        /// <returns>Invalid Plugin Exception</returns>
        public InvalidPluginExecutionException InvalidPluginExecutionException(string message)
        {
            return this.InvalidPluginExecutionException(message, null);
        }

        /// <summary>
        /// Throws an Invalid Plugin Exception
        /// </summary>
        /// <param name="message">Message to be included in the exception</param>
        /// <param name="args">Values to be substituted in the message</param>
        /// <returns>Invalid Plugin Exception</returns>
        public InvalidPluginExecutionException InvalidPluginExecutionException(string message, params object[] args)
        {
            return new InvalidPluginExecutionException(ExtensionBase.ConcatenatedString(message, args));
        }

        /// <summary>
        /// Throws an Argument Out of Range Exception
        /// </summary>
        /// <param name="parameterName">Parameter Name</param>
        /// <param name="message">Message to be included in the exception</param>
        /// <returns>Argument Out of Range Exception</returns>
        public ArgumentOutOfRangeException ArgumentOutOfRangeException(string parameterName, string message)
        {
            return this.ArgumentOutOfRangeException(parameterName, message, null);
        }

        /// <summary>
        /// Throws an Argument Out of Range Exception
        /// </summary>
        /// <param name="parameterName">Parameter Name</param>
        /// <param name="message">Message to be included in the exception</param>
        /// <param name="args">Values to be substituted in the message</param>
        /// <returns>Argument Out of Range Exception</returns>
        public ArgumentOutOfRangeException ArgumentOutOfRangeException(string parameterName, string message, params object[] args)
        {
            return new ArgumentOutOfRangeException(parameterName, ExtensionBase.ConcatenatedString(message, args));
        }
    }
}
