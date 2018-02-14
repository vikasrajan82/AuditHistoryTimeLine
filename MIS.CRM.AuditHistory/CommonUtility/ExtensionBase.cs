// <copyright file="ExtensionBase.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Base class for extensions</summary>
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
    /// Base class for extensions
    /// </summary>
    public static class ExtensionBase
    {
        /// <summary>
        /// Private Method to check if the arguments are null
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
        /// Returns the Object Value as mentioned
        /// </summary>
        /// <typeparam name="T">Expected Type</typeparam>
        /// <param name="value">Object Value</param>
        /// <returns>Value converted to the required type</returns>
        public static T GetObjectValue<T>(object value)
        {
            if (value != null)
            {
                object convertedValue = null;
                switch (value.GetType().ToString())
                {
                    case "System.DateTime":
                        convertedValue = Convert.ToDateTime(value, CultureInfo.InvariantCulture).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                        break;
                    case "Microsoft.Xrm.Sdk.EntityReference":
                        convertedValue = ((EntityReference)value).Name;
                        break;
                    case "System.Boolean":
                        convertedValue = Convert.ToBoolean(value, CultureInfo.InvariantCulture) ? "Yes" : "No";
                        break;
                    case "System.Int32":
                    case "System.Int16":
                    case "System.Int64":
                        convertedValue = value.ToString();
                        break;
                    case "System.Decimal":
                        convertedValue = Math.Round(Convert.ToDecimal(value, CultureInfo.InvariantCulture), 2).ToString(CultureInfo.InvariantCulture);
                        break;
                    case "Microsoft.Xrm.Sdk.Money":
                        convertedValue = Math.Round(((Money)value).Value, 2).ToString(CultureInfo.InvariantCulture);
                        break;
                    default:
                        convertedValue = string.Empty;

                        // throw new ArgumentOutOfRangeException("value", "Invalid Exception" + value.GetType().ToString());
                        break;
                }

                return (T)Convert.ChangeType(convertedValue, typeof(T), CultureInfo.InvariantCulture);
            }

            return default(T);
        }

        /// <summary>
        /// Substitutes the string with the arguments
        /// </summary>
        /// <param name="message">full message text</param>
        /// <param name="args">arguments to be substituted</param>
        /// <returns>replaced string</returns>
        public static string ConcatenatedString(string message, params object[] args)
        {
            if (args != null)
            {
                return string.Format(CultureInfo.InvariantCulture, message, args);
            }
            else
            {
                return message;
            }
        }
    }
}
