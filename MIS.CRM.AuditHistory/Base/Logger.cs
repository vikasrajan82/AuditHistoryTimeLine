// <copyright file="Logger.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author></author>
// <date>10/8/2015 12:38:06 PM</date>
// <summary>Implements the GetFamilyPermittedIncome workflow activity.</summary>

namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Globalization;
    using System.Reflection;
    using System.Text;
    using Microsoft.Xrm.Sdk;

    /// <summary>
    /// Represents type of Log
    /// </summary>
    public enum LogType
    {
        /// <summary>
        /// Log type of type Info
        /// </summary>
        Info,

        /// <summary>
        /// Log type of type Warning
        /// </summary>
        Warning,

        /// <summary>
        /// Log type of type Error
        /// </summary>
        Error
    }

    /// <summary>
    /// Implements logging to tracing
    /// </summary>
    public static class Logger
    {
        /// <summary>
        /// Gets current timestamp
        /// </summary>
        public static string UniqueKey
        {
            get
            {
                return DateTime.Now.ToString("yyyyMMddHHmmssfff", CultureInfo.CurrentCulture);
            }
        }

        /// <summary>
        /// Logs the message passed to tracing service
        /// </summary>
        /// <param name="logType">pass type of log of type LogType</param>
        /// <param name="message">pass message of type String</param>
        /// <param name="timestamp">pass time stamp of type string.</param>
        /// <param name="tracingService">pass trace service of type ITracingService.</param>
        public static void Log(LogType logType, string message, string timestamp, ITracingService tracingService)
        {
            if (tracingService != null)
            {
                tracingService.Trace(LogMessage(logType, message, timestamp, null));
            }
        }

        /// <summary>
        /// Logs the message passed to tracing service
        /// </summary>
        /// <param name="logType">pass type of log of type LogType</param>
        /// <param name="message">pass message of type String</param>
        /// <param name="timestamp">pass time stamp of type string.</param>
        /// <param name="tracingService">pass trace service of type ITracingService.</param>
        /// <param name="ex">pass exception of type System.Exception</param>
        public static void Log(LogType logType, string message, string timestamp, ITracingService tracingService, Exception ex)
        {
            if (tracingService != null)
            {
                tracingService.Trace(LogMessage(logType, message, timestamp, ex));
            }
        }

        /// <summary>
        /// Logs the message passed to tracing service
        /// </summary>
        /// <param name="logType">pass type of log of type LogType</param>
        /// <param name="message">pass message of type String</param>
        /// <param name="tracingService">pass trace service of type ITracingService.</param>
        public static void Log(LogType logType, string message, ITracingService tracingService)
        {
            if (tracingService != null)
            {
                tracingService.Trace(LogMessage(logType, message, null, null));
            }
        }

        /// <summary>
        /// Logs the message passed to tracing service
        /// </summary>
        /// <param name="logType">pass type of log of type LogType</param>
        /// <param name="message">pass message of type String</param>
        /// <param name="tracingService">pass trace service of type ITracingService.</param>
        /// /// <param name="ex">pass exception of type System.Exception</param>
        public static void Log(LogType logType, string message, ITracingService tracingService, Exception ex)
        {
            if (tracingService != null)
            {
                tracingService.Trace(LogMessage(logType, message, null, ex));
            }
        }

        /// <summary>
        /// Logs the Exception raised from the Base Plugin
        /// </summary>
        /// <param name="tracingService">pass trace service of type ITracingService.</param>
        /// <param name="ex">pass exception of type System.Exception</param>
        public static void LogBasePluginException(ITracingService tracingService, Exception ex)
        {
            var timestamp = UniqueKey;

            if (tracingService != null && ex != null)
            {
                Log(LogType.Error, ex.Message, timestamp, tracingService, ex);
            }

            throw new InvalidPluginExecutionException(string.Format(CultureInfo.CurrentCulture, Resources.ContactAdminError, timestamp));
        }

        /// <summary>
        /// Prepares log message and returns
        /// </summary>
        /// <param name="logType">pass type of log of type LogType</param>
        /// <param name="message">pass message of type String</param>
        /// <param name="timestamp">pass the timestamp</param>
        /// <param name="ex">pass exception of type System.Exception</param>
        /// <returns>returns log message of type String.</returns>
        private static string LogMessage(LogType logType, string message, string timestamp, Exception ex)
        {
            StringBuilder sb = new StringBuilder();
            switch (logType)
            {
                case LogType.Info:
                    sb.Append("Log Type:- ");
                    sb.AppendLine("Information");
                    if (!string.IsNullOrEmpty(timestamp))
                    {
                        sb.AppendLine(string.Format(CultureInfo.CurrentCulture, "Time Stamp:- {0}", timestamp));
                    }

                    GetMessage(message, ex, sb);
                    break;
                case LogType.Warning:
                    sb.Append("Log Type:- ");
                    sb.AppendLine("Warning");
                    if (!string.IsNullOrEmpty(timestamp))
                    {
                        sb.AppendLine(string.Format(CultureInfo.CurrentCulture, "Time Stamp:- {0}", timestamp));
                    }

                    GetMessage(message, ex, sb);
                    break;
                case LogType.Error:
                    sb.Append("Log Type:- ");
                    sb.AppendLine("Error");
                    if (!string.IsNullOrEmpty(timestamp))
                    {
                        sb.AppendLine(string.Format(CultureInfo.CurrentCulture, "Time Stamp:- {0}", timestamp));
                    }

                    GetMessage(message, ex, sb);
                    break;
            }

            return sb.ToString();
        }

        /// <summary>
        /// Prepares log message and returns
        /// </summary>
        /// <param name="message">pass message of type String</param>
        /// <param name="ex">pass exception of type System.Exception</param>
        /// <param name="sb">pass string builder of type StringBuilder.</param>
        private static void GetMessage(string message, Exception ex, StringBuilder sb)
        {
            var stackTrace = new System.Diagnostics.StackTrace(3, false);
            MethodBase method = stackTrace.GetFrame(0).GetMethod();
            string namespaceName = method.DeclaringType.Namespace;
            string methodName = string.Format(CultureInfo.CurrentCulture, "{0}.{1}", namespaceName, method.Name);
            string classname = method.DeclaringType.Name;

            sb.Append("Class:- ");
            sb.AppendLine(classname);

            sb.Append("Method:- ");
            sb.AppendLine(methodName);

            if (!string.IsNullOrEmpty(message))
            {
                sb.Append("Message:- ");
                sb.AppendLine(message);
            }

            Exception nextException = ex;

            while (nextException != null)
            {
                sb.AppendLine(nextException.Message);

                sb.Append("Type:- ");
                sb.AppendLine(nextException.GetType().ToString());

                sb.Append("Source:- ");
                sb.AppendLine(nextException.Source);

                sb.AppendLine();
                sb.AppendLine(nextException.StackTrace);

                nextException = nextException.InnerException;

                if (nextException != null)
                {
                    sb.AppendLine("\r\n ---- Inner Exception ----");
                }
            }
        }
    }
}
