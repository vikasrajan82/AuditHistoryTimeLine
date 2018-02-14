// <copyright file="CRMHelper.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author></author>
// <date>10/20/2015 12:38:06 PM</date>
// <summary>Implements the Blood works related transactions.</summary>

namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Runtime.Serialization.Json;
    using System.Text;
    using System.Xml.Linq;
    //using EntityWrapper.Entities;
    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Metadata;
    using Microsoft.Xrm.Sdk.Query;
    //using MIS.CRM.EntityWrapper.Constants;

    /// <summary>
    /// Performs CRM related transactions
    /// </summary>
    public static class CRMHelper
    {
        #region Variables

        /// <summary>
        /// Gets or sets User entity
        /// </summary>
        private static SystemUser User
        {
            get;
            set;
        }

        #endregion Variables

        #region CRM Methods

        /// <summary>
        /// Prepares the attribute value.
        /// </summary>
        /// <param name="attributeName">Name of the attribute.</param>
        /// <param name="value">The value as string.</param>
        /// <param name="typeCode">The type code.</param>
        /// <returns>An instance of key value pair</returns>
        public static KeyValuePair<string, object>? PrepareAttributeValue(string attributeName, string value, int typeCode)
        {
            KeyValuePair<string, object>? attributeValue = null;

            if (string.IsNullOrWhiteSpace(value))
            {
                return new KeyValuePair<string, object>(attributeName, null);
            }

            long longAttributeValue;
            bool boolAttributeValue;
            DateTime dateTimeAttributeValue;
            decimal decimalAttributeValue;
            double doubleAttributeValue;
            int intAttributeValue;
            Guid guidAttributeValue;

            switch (typeCode)
            {
                case (int)AttributeTypeCode.BigInt:
                    if (long.TryParse(value, out longAttributeValue))
                    {
                        attributeValue = new KeyValuePair<string, object>(attributeName, longAttributeValue);
                    }

                    break;
                case (int)AttributeTypeCode.Boolean:
                    if (bool.TryParse(value, out boolAttributeValue))
                    {
                        attributeValue = new KeyValuePair<string, object>(attributeName, boolAttributeValue);
                    }

                    break;
                case (int)AttributeTypeCode.DateTime:
                    if (DateTime.TryParse(value, out dateTimeAttributeValue))
                    {
                        attributeValue = new KeyValuePair<string, object>(attributeName, dateTimeAttributeValue);
                    }

                    break;
                case (int)AttributeTypeCode.Decimal:
                case (int)AttributeTypeCode.Money:
                    if (decimal.TryParse(value, out decimalAttributeValue))
                    {
                        attributeValue = new KeyValuePair<string, object>(attributeName, decimalAttributeValue);
                    }

                    break;
                case (int)AttributeTypeCode.Double:
                    if (double.TryParse(value, out doubleAttributeValue))
                    {
                        attributeValue = new KeyValuePair<string, object>(attributeName, doubleAttributeValue);
                    }

                    break;
                case (int)AttributeTypeCode.EntityName:
                case (int)AttributeTypeCode.Memo:
                case (int)AttributeTypeCode.String:
                    attributeValue = new KeyValuePair<string, object>(attributeName, value);
                    break;
                case (int)AttributeTypeCode.Integer:
                    if (int.TryParse(value, out intAttributeValue))
                    {
                        attributeValue = new KeyValuePair<string, object>(attributeName, intAttributeValue);
                    }

                    break;
                case (int)AttributeTypeCode.Picklist:
                case (int)AttributeTypeCode.State:
                case (int)AttributeTypeCode.Status:
                    if (int.TryParse(value, out intAttributeValue))
                    {
                        attributeValue = new KeyValuePair<string, object>(attributeName, new Microsoft.Xrm.Sdk.OptionSetValue(intAttributeValue));
                    }

                    break;
                case (int)AttributeTypeCode.Uniqueidentifier:
                    if (Guid.TryParse(value, out guidAttributeValue))
                    {
                        attributeValue = new KeyValuePair<string, object>(attributeName, guidAttributeValue);
                    }

                    break;
                case (int)AttributeTypeCode.Lookup:
                case (int)AttributeTypeCode.Owner:
                case (int)AttributeTypeCode.Customer:
                    var tokens = value.Split(":".ToCharArray());
                    if (tokens.Length == 2 && !string.IsNullOrEmpty(tokens[0]) && Guid.TryParse(tokens[1], out guidAttributeValue))
                    {
                        attributeValue = new KeyValuePair<string, object>(attributeName, new EntityReference(tokens[0], guidAttributeValue));
                    }

                    break;
                default:
                    // Do nothing. Do not set value
                    break;
            }

            return attributeValue;
        }

        /// <summary>
        /// Sets the attribute of the specified entity.
        /// </summary>
        /// <param name="entity">An instance of <see cref="Microsoft.Xrm.Sdk.Entity"/></param>
        /// <param name="attributeName">Name of the attribute.</param>
        /// <param name="value">The value.</param>
        /// <param name="typeCode">The type code.</param>
        public static void SetAttribute(Entity entity, string attributeName, string value, int typeCode)
        {
            if (entity != null)
            {
                var pair = PrepareAttributeValue(attributeName, value, typeCode);
                if (pair != null)
                {
                    entity.Attributes[attributeName] = pair.Value.Value;
                }
            }
        }

        /// <summary>
        /// Sets multiple attributes of the specified entity.
        /// </summary>
        /// <param name="entity">An instance of <see cref="Microsoft.Xrm.Sdk.Entity"/></param>
        /// <param name="attributes">The attributes.</param>
        public static void SetAttributes(Entity entity, Dictionary<string, object> attributes)
        {
            if (entity != null && attributes != null)
            {
                foreach (var pair in attributes)
                {
                    entity.Attributes[pair.Key] = pair.Value;
                }
            }
        }

        /// <summary>
        /// Get a collection of Business Entities where the specified attribute(s) have the specified value(s)
        /// </summary>
        /// <param name="service">IOrganizationService to execute the request</param>
        /// <param name="entityName">Entity type to return e.g. "contact"</param>
        /// <param name="attributeNames">string array of attribute names e.g. {"field","field"}</param>
        /// <param name="attributeValues">object array of attribute values e.g. {"John", "Smith"}</param>
        /// <param name="columnsToReturn">Optional string array of columns to return. If null or empty, all columns are returned.</param>
        /// <returns>EntityCollection of entities with attribute values that match the ones specified</returns>
        /// <example>GetEntityCollectionByAttributes(service, "contact", {"field","field"},{"value","value"},{"columns"}) returns all contact entity instances where field=john and field=smith. Returns contactIDs only.</example>
        public static EntityCollection GetEntityCollectionByAttributes(IOrganizationService service, string entityName, string[] attributeNames, object[] attributeValues, string[] columnsToReturn)
        {
            if (service != null && attributeNames != null && attributeValues != null)
            {
                if (attributeNames.Length != attributeValues.Length)
                {
                    throw new ArgumentException("Number of attribute names is not the same as number of attribute values");
                }

                QueryByAttribute attributeQuery = new QueryByAttribute();
                attributeQuery.EntityName = entityName;
                attributeQuery.Attributes.AddRange(attributeNames);
                if (columnsToReturn != null && columnsToReturn.Length > 0)
                {
                    ColumnSet columnSet = new ColumnSet();
                    columnSet.AddColumns(columnsToReturn);
                    attributeQuery.ColumnSet = columnSet;
                }

                attributeQuery.Values.AddRange(attributeValues);
                RetrieveMultipleRequest request = new RetrieveMultipleRequest();
                request.Query = attributeQuery;

                RetrieveMultipleResponse response = (RetrieveMultipleResponse)service.Execute(request);
                EntityCollection returnedEntities = response.EntityCollection;

                return returnedEntities;
            }

            return null;
        }

        /// <summary>
        /// Get a collection of Business Entities where the specified attribute(s) have the specified value(s)
        /// </summary>
        /// <param name="service">IOrganizationService to execute CRM requests.</param>
        /// <param name="entityName">Entity type to return e.g. "contact"</param>
        /// <param name="attributeNames">string array of attribute names e.g. {"field","field"}</param>
        /// <param name="attributeValues">object array of attribute values e.g. {"John", "Smith"}</param>
        /// <param name="columnsToReturn">Optional string array of columns to return. If null or empty, all columns are returned.</param>
        /// <param name="sortAttributeName">string of sorting attribute name</param>
        /// <param name="sortOrder">OrderType to order returned results either Ascending or Descending</param>
        /// <returns>EntityCollection of entities with attribute values that match the ones specified</returns>
        /// <example>GetEntityCollectionByAttributes(service, "contact", {"field","field"},{"John","Smith"},{"column"}) returns all contact entity instances where field=john and field=smith. Returns contactIDs only.</example>
        public static EntityCollection GetEntityCollectionByAttributes(IOrganizationService service, string entityName, string[] attributeNames, object[] attributeValues, string[] columnsToReturn, string sortAttributeName, OrderType sortOrder)
        {
            if (service != null && attributeNames != null && attributeValues != null)
            {
                if (attributeNames.Length != attributeValues.Length)
                {
                    throw new ArgumentException("Number of attribute names is not the same as number of attribute values");
                }

                QueryByAttribute attributeQuery = new QueryByAttribute();
                attributeQuery.EntityName = entityName;
                attributeQuery.Attributes.AddRange(attributeNames);
                if (columnsToReturn != null && columnsToReturn.Length > 0)
                {
                    ColumnSet columnSet = new ColumnSet();
                    columnSet.AddColumns(columnsToReturn);
                    attributeQuery.ColumnSet = columnSet;
                }

                attributeQuery.Values.AddRange(attributeValues);
                attributeQuery.AddOrder(sortAttributeName, sortOrder);
                RetrieveMultipleRequest request = new RetrieveMultipleRequest();
                request.Query = attributeQuery;

                RetrieveMultipleResponse response = (RetrieveMultipleResponse)service.Execute(request);
                EntityCollection returnedEntities = response.EntityCollection;

                return returnedEntities;
            }

            return null;
        }

        /// <summary>
        /// Get the overlap check filter expression based on dynamic type passed
        /// </summary>
        /// <typeparam name="T">T can be any CRM type to compare data</typeparam>
        /// <param name="startAttribute">string to pass field to compare start attribute in CRM</param>
        /// <param name="endAttribute">string to pass field to compare end attribute in CRM</param>
        /// <param name="startAttributeValue">pass dynamic CRM type for start attribute</param>
        /// <param name="endAttributeValue">pass dynamic CRM type for end attribute</param>
        /// <returns>returns filter expression of overlap check.</returns>
        public static FilterExpression GetOverlapCheckExpression<T>(string startAttribute, string endAttribute, T startAttributeValue, T endAttributeValue)
        {
            return new FilterExpression()
            {
                FilterOperator = LogicalOperator.And,
                Filters =
                    {
                        new FilterExpression
                        {
                            FilterOperator = LogicalOperator.Or,
                            Filters =
                            {
                                new FilterExpression
                                {
                                    FilterOperator = LogicalOperator.And,
                                    Conditions =
                                         {
                                               new ConditionExpression(startAttribute, ConditionOperator.LessEqual, startAttributeValue),
                                              //// new ConditionExpression(endAttribute, ConditionOperator.GreaterEqual, startAttributeValue)
                                         },
                                         Filters =
                                         {
                                         new FilterExpression
                                         {
                                        FilterOperator = LogicalOperator.Or,

                                         Conditions =
                                         {
                                              new ConditionExpression(endAttribute, ConditionOperator.Null),
                                              new ConditionExpression(endAttribute, ConditionOperator.GreaterEqual, startAttributeValue)
                                         },
                                         }
                                         }
                                },
                                new FilterExpression
                                {
                                    FilterOperator = LogicalOperator.And,
                                    Conditions =
                                         {
                                               new ConditionExpression(startAttribute, ConditionOperator.LessEqual, endAttributeValue),
                                              //// new ConditionExpression(endAttribute, ConditionOperator.GreaterEqual, endAttributeValue)
                                         },
                                          Filters =
                                         {
                                         new FilterExpression
                                         {
                                        FilterOperator = LogicalOperator.Or,

                                         Conditions =
                                         {
                                              new ConditionExpression(endAttribute, ConditionOperator.Null),
                                             new ConditionExpression(endAttribute, ConditionOperator.GreaterEqual, endAttributeValue)
                                         },
                                         }
                                         }
                      },
                                new FilterExpression
                                {
                                    FilterOperator = LogicalOperator.And,
                                    Conditions =
                                         {
                                               new ConditionExpression(startAttribute, ConditionOperator.GreaterEqual, startAttributeValue),
                                           ////    new ConditionExpression(endAttribute, ConditionOperator.LessEqual, endAttributeValue)
                                         },
                                          Filters =
                                         {
                                         new FilterExpression
                                         {
                                        FilterOperator = LogicalOperator.Or,

                                         Conditions =
                                         {
                                              new ConditionExpression(endAttribute, ConditionOperator.Null),
                                              new ConditionExpression(endAttribute, ConditionOperator.LessEqual, endAttributeValue)
                                         },
                                         }
                                         }
                                }
                            }
                        }
                    }
            };
        }

        /// <summary>
        /// retrieve a CRM entity based on a simple query (generally the Id)
        /// </summary>
        /// <param name="service">IOrganizationService to execute requests</param>
        /// <param name="entityName">Name of CRM entity</param>
        /// <param name="columnNames">Column Names</param>
        /// <param name="attributeName">The "column" name that will be searched</param>
        /// <param name="keyValue">The Search Key/Id</param>
        /// <returns>returns CRM Entity record requested</returns>
        public static Entity GetCrmEntityUsingKey(IOrganizationService service, string entityName, string[] columnNames, string attributeName, string keyValue)
        {
            if (service != null)
            {
                QueryExpression query = new QueryExpression(entityName);
                query.ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet(columnNames);
                ConditionExpression con = new ConditionExpression();
                con.AttributeName = attributeName;
                con.Operator = ConditionOperator.Equal;
                con.Values.Add(keyValue);

                FilterExpression filter = new FilterExpression();
                filter.Conditions.Add(con);

                query.Criteria.AddFilter(filter);
                EntityCollection resultSet = service.RetrieveMultiple(query);
                if (resultSet.HasRecords() && resultSet.Entities.Count > 0)
                {
                    return resultSet.Entities[0];
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the CRM entity using keys.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="entityName">Name of the entity.</param>
        /// <param name="columnNames">column names.</param>
        /// <param name="attributeName">Name of the attribute.</param>
        /// <param name="keyValue">The key value.</param>
        /// <returns>Returns Auto number configuration Records</returns>
        public static EntityCollection GetCrmEntityUsingKeys(IOrganizationService service, string entityName, string[] columnNames, string attributeName, string[] keyValue)
        {
            if (service != null && keyValue != null)
            {
                QueryExpression query = new QueryExpression(entityName);
                query.ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet(columnNames);

                ConditionExpression condition = new ConditionExpression();
                FilterExpression filter = new FilterExpression();

                condition.AttributeName = attributeName;
                condition.Operator = ConditionOperator.In;
                for (int i = 0; i < keyValue.Length; i++)
                {
                    condition.Values.Add(keyValue[i]);
                    filter.Conditions.Add(condition);
                }

                query.Criteria.AddFilter(filter);
                EntityCollection resultSet = service.RetrieveMultiple(query);
                if (resultSet.Entities.Count > 0)
                {
                    return resultSet;
                }
            }

            return null;
        }

        /// <summary>
        ///  This method converts time to 24hours format time.
        /// </summary>
        /// <param name="time"> time in string </param>
        /// <returns> formatted timespan </returns>
        public static TimeSpan To24HRTime(string time)
        {
            TimeSpan timeSpan = new TimeSpan();
            if (time != null)
            {
                char[] delimiters = new char[] { ':', ' ' };
                string[] spltTime = time.Split(delimiters);

                int hour = Convert.ToInt16(spltTime[0], CultureInfo.CurrentCulture);
                int minute = Convert.ToInt16(spltTime[1], CultureInfo.CurrentCulture);
                int seconds = 0;

                string timeSuffix = spltTime[2];

                if (timeSuffix.ToUpper(CultureInfo.CurrentCulture) == Resources.PM.ToString())
                {
                    hour = (hour % 12) + 12;
                }
                else if (timeSuffix.ToUpper(CultureInfo.CurrentCulture) == Resources.AM.ToString() && hour == 12)
                {
                    hour = 0;
                }

                timeSpan = new TimeSpan(hour, minute, seconds);
            }

            return timeSpan;
        }

        /// <summary>
        /// This method gets option set label.
        /// </summary>
        /// <param name="entityName"> entity name </param>
        /// <param name="fieldName"> field name </param>
        /// <param name="optionSetValue"> option set value </param>
        /// <param name="service"> organization service </param>
        /// <returns> name formed </returns>
        public static string GetOptionSetValueLabel(string entityName, string fieldName, int optionSetValue, IOrganizationService service)
        {
            if (service != null)
            {
                var attReq = new RetrieveAttributeRequest();
                attReq.EntityLogicalName = entityName;
                attReq.LogicalName = fieldName;
                attReq.RetrieveAsIfPublished = true;

                var attResponse = (RetrieveAttributeResponse)service.Execute(attReq);
                var attMetadata = (EnumAttributeMetadata)attResponse.AttributeMetadata;
                if (attMetadata != null && attMetadata.OptionSet != null &&
                    attMetadata.OptionSet.Options != null && attMetadata.OptionSet.Options.Any())
                {
                    var option = attMetadata.OptionSet.Options.FirstOrDefault(x => x.Value == optionSetValue);
                    if (option != null)
                    {
                        return option.Label.UserLocalizedLabel.Label;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Get aliased value
        /// </summary>
        /// <param name="recordCollection">record Collection</param>
        /// <param name="aliasName">alias name</param>
        /// <returns>returns the aliased value</returns>
        public static int GetAliasedValue(EntityCollection recordCollection, string aliasName)
        {
            if (recordCollection != null && recordCollection.HasRecords())
            {
                return Convert.ToInt32(recordCollection[0].GetAttributeValue<AliasedValue>(aliasName).Value, CultureInfo.InvariantCulture);
            }

            return -1;
        }

        /// <summary>
        /// Get Option set value labels
        /// </summary>
        /// <param name="entityName">entity name</param>
        /// <param name="fieldName">field name</param>
        /// <param name="service">the service</param>
        /// <returns>returns the labels</returns>
        public static Dictionary<int, string> GetOptionSetValueLabels(string entityName, string fieldName, IOrganizationService service)
        {
            if (service != null)
            {
                Dictionary<int, string> optionSetValues = new Dictionary<int, string>();
                var attReq = new RetrieveAttributeRequest();
                attReq.EntityLogicalName = entityName;
                attReq.LogicalName = fieldName;
                attReq.RetrieveAsIfPublished = true;

                var attResponse = (RetrieveAttributeResponse)service.Execute(attReq);
                var attMetadata = (EnumAttributeMetadata)attResponse.AttributeMetadata;

                attMetadata.OptionSet.Options.ToList().ForEach(optionset =>
                    {
                        optionSetValues.Add(optionset.Value.Value, optionset.Label.UserLocalizedLabel.Label);
                    });

                return optionSetValues;
            }

            return null;
        }

        #endregion CRM Methods

        #region Non-CRM Methods

        /// <summary>
        /// method to calculate age. Returns age in months.
        /// </summary>
        /// <param name="dateOfBirth">birth date of type DateTime</param>
        /// <param name="toDate">to date of type DateTime</param>
        /// <returns>age in months of type Integer</returns>        
        public static int CalculateAge(DateTime dateOfBirth, DateTime toDate)
        {
            dateOfBirth = dateOfBirth.Date;
            toDate = toDate.Date;
            return GetAgeInMonths(dateOfBirth, toDate);
        }

        /// <summary>
        /// Overload for Calculating Age in Months in decimal.
        /// </summary>
        /// <param name="dateOfBirth"> date of birth </param>
        /// <returns> months in decimal </returns>
        public static decimal CalculateAge(DateTime dateOfBirth)
        {
            return CalculateAgeInMonthsDecimal(dateOfBirth, DateTime.Today);
        }

        /// <summary>
        /// method to calculate age in months and days. Returns age in months.
        /// </summary>
        /// <param name="dateOfBirth">birth date of type DateTime</param>
        /// <param name="toDate">to date of type DateTime</param>
        /// <returns>age in months and days of type decimal</returns>  
        public static decimal CalculateAgeInMonthsDecimal(DateTime dateOfBirth, DateTime toDate)
        {
            dateOfBirth = dateOfBirth.Date;
            toDate = toDate.Date;
            if (toDate > dateOfBirth)
            {
                int monthDiff = GetAgeInMonths(dateOfBirth, toDate);
                int daysLeft = GetRemainingDays(dateOfBirth.AddMonths(monthDiff), toDate);
                return monthDiff + ((decimal)daysLeft / 31m);
            }

            return 0m;
        }

        /// <summary>
        /// Calculating the age in months and days
        /// </summary>
        /// <param name="dateOfBirth">pass date of birth</param>
        /// <returns>age in months and days</returns>
        public static string CalculateAgeInMonthsAndDays(DateTime dateOfBirth)
        {
            dateOfBirth = dateOfBirth.Date;
            DateTime toDate = DateTime.Today;
            if (toDate > dateOfBirth)
            {
                int monthDiff = GetAgeInMonths(dateOfBirth, toDate);
                int daysLeft = GetRemainingDays(dateOfBirth.AddMonths(monthDiff), toDate);
                return monthDiff + "m" + " " + daysLeft + "d";
            }

            return 0 + "m" + " " + 0 + "d";
        }

        /// <summary>
        /// Calculating Age in Month and Year 
        /// </summary>
        /// <param name="dateOfBirth">date of birth</param>
        /// <returns>age in years and months</returns>
        public static string CalculateAgeInYearsAndMonths(DateTime dateOfBirth)
        {
            DateTime currentDate = DateTime.Now;
            TimeSpan difference = currentDate.Subtract(dateOfBirth);
            DateTime age = DateTime.MinValue + difference;
            int ageInYears = age.Year - 1;
            int ageInMonths = age.Month - 1;
            return ageInYears + "y" + " " + ageInMonths + "m";
        }

        /// <summary>
        /// Calculating Age in Months in decimal
        /// </summary>
        /// <param name="dateOfBirth">date of birth</param>
        /// <param name="toDate">to Date</param>
        /// <returns>age in months</returns>
        public static string DisplayAgeBasedOnCategory(DateTime dateOfBirth, DateTime toDate)
        {
            decimal age = CalculateAgeInMonthsDecimal(dateOfBirth, toDate) / 12;
            TimeSpan difference = toDate.Subtract(dateOfBirth);
            DateTime ageValue = DateTime.MinValue + difference;
            int ageInYears = ageValue.Year - 1;
            int ageInMonths = ageValue.Month - 1;
            int ageInDays = ageValue.Day - 1;

            if (age <= 2)
            {
                return ageInMonths + "m" + " " + ageInDays + "d";
            }

            if (age > 2 && age <= 9)
            {
                return ageInYears + "y" + " " + ageInMonths + "m";
            }

            if (age > 9)
            {
                return ageInYears + "y";
            }

            return string.Empty;
        }

        /// <summary>
        /// method to calculate age. Returns age in years.
        /// </summary>
        /// <param name="dateOfBirth">birth date of type DateTime</param>
        /// <param name="toDate">to date of type DateTime</param>
        /// <returns>age in months of type decimal</returns>
        public static decimal CalculateAgeYears(DateTime dateOfBirth, DateTime toDate)
        {
            return CalculateAgeInMonthsDecimal(dateOfBirth, toDate) / 12;
        }

        /// <summary>
        /// Gets days in the month for the date passed
        /// </summary>
        /// <param name="value">value of the DateTime</param>
        /// <returns>returns number of days available in the month</returns>
        public static int DaysInMonth(DateTime value)
        {
            return DateTime.DaysInMonth(value.Year, value.Month);
        }

        /// <summary>
        /// Gets last day of the month passed
        /// </summary>
        /// <param name="value">value of type DateTime</param>
        /// <returns>returns last day of the month of type DateTime</returns>
        public static DateTime LastDayOfMonth(DateTime value)
        {
            return new DateTime(value.Year, value.Month, DaysInMonth(value));
        }

        /// <summary>
        /// This method calculates last day of month from given date time.
        /// </summary>
        /// <param name="dateTime"> date passed </param>
        /// <returns> calculated date time </returns>
        public static DateTime LastDayOfMonthFromDateTime(DateTime dateTime)
        {
            DateTime firstDayOfTheMonth = new DateTime(dateTime.Year, dateTime.Month, 1, 23, 59, 00);
            return firstDayOfTheMonth.AddMonths(1).AddDays(-1);
        }

        /// <summary>
        /// JSON serializer.
        /// </summary>
        /// <typeparam name="T">Type of Object</typeparam>
        /// <param name="type">The t.</param>
        /// <returns>JSON String</returns>
        public static string JsonSerializer<T>(T type)
        {
            DataContractJsonSerializer jsonSerializer = new DataContractJsonSerializer(typeof(T));
            string jsonString = string.Empty;
            using (var memoryStream = new MemoryStream())
            {
                jsonSerializer.WriteObject(memoryStream, type);
                jsonString = Encoding.UTF8.GetString(memoryStream.ToArray());
            }

            return jsonString;
        }

        #endregion Non-CRM Methods

        
        /// <summary>
        /// Gets the users security roles by user identifier.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="columns">The columns.</param>
        /// <param name="userId">The user identifier.</param>
        /// <returns>Security Role Entity Collection</returns>
        public static EntityCollection GetUsersSecurityRolesByUserId(IOrganizationService service, string[] columns, Guid userId)
        {
            if (service != null && columns != null && columns.Length > 0)
            {
                QueryExpression query = new QueryExpression()
                {
                    EntityName = Role.EntityLogicalName,
                    ColumnSet = new ColumnSet(columns),
                    LinkEntities =
            {
              new LinkEntity
              {
                LinkFromEntityName = Role.EntityLogicalName,
                LinkFromAttributeName = "roleid",
                LinkToEntityName = SystemUserRoles.EntityLogicalName,
                LinkToAttributeName = "roleid",
                LinkCriteria = new FilterExpression
                {
                  FilterOperator = LogicalOperator.And,
                  Conditions =
                  {
                    new ConditionExpression
                    {
                      AttributeName = "systemuserid",
                      Operator = ConditionOperator.Equal,
                      Values = { userId }
                    }
                  }
                }
              }
            }
                };
                return service.RetrieveMultiple(query);
            }

            return null;
        }

        /// <summary>
        /// Retrieve Option set text
        /// </summary>
        /// <param name="entityName">entity name</param>
        /// <param name="attributeName">attribute name</param>
        /// <param name="optionSetValue">option set value</param>
        /// <param name="orgService">the service</param>
        /// <returns>returns the option set text</returns>
        public static string RetrieveOptionSetText(string entityName, string attributeName, int optionSetValue, IOrganizationService orgService)
        {
            if (orgService == null)
            {
                throw new ArgumentNullException("orgService");
            }

            string optionsetText = string.Empty;
            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest();
            retrieveAttributeRequest.EntityLogicalName = entityName;
            retrieveAttributeRequest.LogicalName = attributeName;
            retrieveAttributeRequest.RetrieveAsIfPublished = true;

            RetrieveAttributeResponse retrieveAttributeResponse =
              (RetrieveAttributeResponse)orgService.Execute(retrieveAttributeRequest);
            PicklistAttributeMetadata picklistAttributeMetadata =
              (PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;

            if (picklistAttributeMetadata != null)
            {
                OptionSetMetadata optionsetMetadata = picklistAttributeMetadata.OptionSet;

                if (optionsetMetadata != null)
                {
                    foreach (OptionMetadata optionMetadata in optionsetMetadata.Options)
                    {
                        if (optionMetadata.Value == optionSetValue)
                        {
                            optionsetText = optionMetadata.Label.UserLocalizedLabel.Label;
                            return optionsetText;
                        }
                    }
                }
            }

            return optionsetText;
        }

        /// <summary>
        /// This method calculates age in weeks
        /// </summary>
        /// <param name="startDate"> start date</param>
        /// <param name="endDate"> end date </param>
        /// <returns> returns number of weeks </returns>
        public static int CalculateAgeInWeeks(DateTime startDate, DateTime endDate)
        {
            return Convert.ToInt32(Math.Round((endDate - startDate).TotalDays / 7), CultureInfo.CurrentCulture);
        }

        /// <summary>
        /// Calculate weeks gestation field
        /// </summary>
        /// <param name="birthdate">birth date</param>
        /// <param name="expectedDateOfBirth">expected date of birth</param>
        /// <returns>returns the calculate gestation field</returns>
        public static int CalculateWeeksGestationField(DateTime birthdate, bool expectedDateOfBirth)
        {
            int diffdays = 0;
            DateTime currentDate = DateTime.Now;
            TimeSpan difference = currentDate.Subtract(birthdate);
            diffdays = difference.Days;

            double calculateWeeksGestation = 0;

            if (!expectedDateOfBirth)
            {
                calculateWeeksGestation = 40.0;
            }
            else
            {
                calculateWeeksGestation = Math.Floor(40 + (diffdays / 7.0));
            }

            return Convert.ToInt32(calculateWeeksGestation);
        }

        /// <summary>
        /// Gets first day of the month passed
        /// </summary>
        /// <param name="value">value of type DateTime</param>
        /// <returns>returns first day of the month of type DateTime</returns>
        public static DateTime FirstDayOfMonth(DateTime value)
        {
            return new DateTime(value.Year, value.Month, 1);
        }

        /// <summary>
        /// This function returns the option set text
        /// </summary>
        /// <param name="orgService">org Service</param>
        /// <param name="entityName">entity Name</param>
        /// <param name="attributeName">attribute Name</param>
        /// <param name="optionSetValue">option set value</param>
        /// <returns>option set text</returns>
        public static string GetOptionSetText(IOrganizationService orgService, string entityName, string attributeName, int optionSetValue)
        {
            string optionsetText = string.Empty;
            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest();
            retrieveAttributeRequest.EntityLogicalName = entityName;
            retrieveAttributeRequest.LogicalName = attributeName;
            retrieveAttributeRequest.RetrieveAsIfPublished = true;

            if (orgService != null)
            {
                RetrieveAttributeResponseWrapper retrieveAttributeResponse = new RetrieveAttributeResponseWrapper(orgService.Execute(retrieveAttributeRequest));
                OptionSetMetadata optionsetMetadata = RetrieveAttributeType(retrieveAttributeResponse);

                if (optionsetMetadata != null)
                {
                    foreach (OptionMetadata optionMetadata in optionsetMetadata.Options)
                    {
                        if (optionMetadata.Value == optionSetValue)
                        {
                            optionsetText = optionMetadata.Label.UserLocalizedLabel.Label;
                            return optionsetText;
                        }
                    }
                }
            }

            return optionsetText;
        }

        /// <summary>
        /// Serialize the data
        /// </summary>
        /// <typeparam name="T">pass generic type</typeparam>
        /// <param name="type">pass type</param>
        /// <returns>string data</returns>
        public static string SerializeJson<T>(T type)
        {
            DataContractJsonSerializer jsonSerializer = new DataContractJsonSerializer(typeof(T));
            string jsonString = string.Empty;
            using (var memoryStream = new MemoryStream())
            {
                jsonSerializer.WriteObject(memoryStream, type);
                jsonString = Encoding.UTF8.GetString(memoryStream.ToArray());
            }

            return jsonString;
        }

        /// <summary>
        /// Deserialize the data
        /// </summary>
        /// <typeparam name="T">pass generic type</typeparam>
        /// <param name="data">pass data</param>
        /// <returns>data collection</returns>
        public static T DeserializeJson<T>(string data)
        {
            var memoryStream = new MemoryStream(ASCIIEncoding.ASCII.GetBytes(data));
            try
            {
                DataContractJsonSerializer dataContractJsonSerializer = new DataContractJsonSerializer(typeof(T));
                var crmFoodCategorySubcategory = (T)dataContractJsonSerializer.ReadObject(memoryStream);
                return crmFoodCategorySubcategory;
            }
            finally
            {
                memoryStream.Dispose();
            }
        }

        /// <summary>
        /// Get aliased value
        /// </summary>
        /// <param name="entity">pass entity Record</param>
        /// <param name="entity1">pass primary entity</param>
        /// <param name="entity2">pass secondary entity</param>
        /// <param name="attributeName">pass attribute name</param>
        /// <returns>aliased value</returns>
        public static AliasedValue GetAliasedValue(Entity entity, string entity1, string entity2, string attributeName)
        {
            if (entity != null && entity1 != null && entity2 != null && attributeName != null)
            {
                return entity.GetAttributeValue<AliasedValue>(string.Format(CultureInfo.CurrentCulture, "{0}{1}.{2}", entity1, entity2, attributeName));
            }

            return null;
        }

        /// <summary>
        /// Get aliased value
        /// </summary>
        /// <param name="entity">pass entity Record</param>
        /// <param name="entity1">pass primary entity</param>
        /// <param name="attributeName">pass attribute name</param>
        /// <returns>aliased value</returns>
        public static AliasedValue GetAliasedValue(Entity entity, string entity1, string attributeName)
        {
            if (entity != null && entity1 != null && attributeName != null)
            {
                return entity.GetAttributeValue<AliasedValue>(string.Format(CultureInfo.CurrentCulture, "{0}.{1}", entity1, attributeName));
            }

            return null;
        }

        /// <summary>
        /// Get Max Of Sequence Dates
        /// </summary>
        /// <param name="availableDates">available Dates</param>
        /// <returns>date time</returns>
        public static DateTime GetMaxOfSequenceDates(DateTime[] availableDates)
        {
            if (availableDates.Any())
            {
                availableDates = availableDates.Distinct().OrderBy(d => d.Date).ToArray();
                for (int i = 0; i < availableDates.Length; i++)
                {
                    if (availableDates[i] == availableDates[availableDates.Length - 1] ||
                        availableDates[i + 1] != availableDates[i].AddMonths(1))
                    {
                        return availableDates[i];
                    }
                }
            }

            return DateTime.Today;
        }

        /// <summary>
        /// Retrieve Option set value
        /// </summary>
        /// <param name="entityName">entity name</param>
        /// <param name="attributeName">attribute name</param>
        /// <param name="optionSetName">option set name</param>
        /// <param name="orgService">the service</param>
        /// <returns>returns the option set value</returns>
        public static int RetrieveOptionSetValue(string entityName, string attributeName, string optionSetName, IOrganizationService orgService)
        {
            if (orgService == null)
            {
                throw new ArgumentNullException("orgService");
            }

            int optionsetValue = 0;
            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest();
            retrieveAttributeRequest.EntityLogicalName = entityName;
            retrieveAttributeRequest.LogicalName = attributeName;
            retrieveAttributeRequest.RetrieveAsIfPublished = true;

            RetrieveAttributeResponse retrieveAttributeResponse =
              (RetrieveAttributeResponse)orgService.Execute(retrieveAttributeRequest);
            PicklistAttributeMetadata picklistAttributeMetadata =
              (PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;

            if (picklistAttributeMetadata != null)
            {
                OptionSetMetadata optionsetMetadata = picklistAttributeMetadata.OptionSet;
                if (optionsetMetadata != null)
                {
                    foreach (OptionMetadata optionMetadata in optionsetMetadata.Options)
                    {
                        if (optionMetadata.Label.UserLocalizedLabel.Label == optionSetName)
                        {
                            optionsetValue = optionMetadata.Value.Value;
                            return optionsetValue;
                        }
                    }
                }
            }

            return optionsetValue;
        }

        /// <summary>
        /// Retrieves the Option set metadata options
        /// </summary>
        /// <param name="retrieveAttributeResponse">Attribute Response</param>
        /// <returns>Option set metadata options</returns>
        private static OptionSetMetadata RetrieveAttributeType(RetrieveAttributeResponseWrapper retrieveAttributeResponse)
        {
            if (retrieveAttributeResponse != null && retrieveAttributeResponse.AttributeMetadata != null)
            {
                switch (retrieveAttributeResponse.AttributeMetadata.AttributeTypeName.Value)
                {
                    case "StatusType":
                        StatusAttributeMetadata statusAttributeMetadata = (StatusAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
                        return statusAttributeMetadata.OptionSet;

                    case "PicklistType":
                        PicklistAttributeMetadata picklistAttributeMetadata = (PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
                        return picklistAttributeMetadata.OptionSet;
                }
            }

            return null;
        }

        /// <summary>
        /// Get Email Description
        /// </summary>
        /// <param name="emailDescription">email Description</param>
        /// <param name="service">org service</param>
        /// <param name="participantName">participant Name</param>
        /// <param name="age">age of Participant</param>
        /// <param name="user">Initiating user</param>
        /// <returns>email Description New</returns>
        private static string GetEmailDescription(string emailDescription, IOrganizationService service, string participantName, string age, SystemUser user)
        {
            emailDescription = emailDescription.Replace("[RequiredDate]", DateTime.UtcNow.LocalDateTime(service).ToString(CultureInfo.InvariantCulture));
            emailDescription = emailDescription.Replace("[RequiredChild]", participantName);
            emailDescription = emailDescription.Replace("[RequiredAge]", age.ToString());
            emailDescription = emailDescription.Replace("[RequiredStaffName]", user.FullName);
            emailDescription = emailDescription.Replace("[RequiredAgency]", user.BusinessUnitId.Name);
            emailDescription = emailDescription.Replace("[RequiredDFPS]", string.Empty);

            return emailDescription;
        }

        #region Private Methods

        /// <summary>
        /// Get Age In Months
        /// </summary>
        /// <param name="dateOfBirth">date Of Birth</param>
        /// <param name="toDate">to Date</param>
        /// <returns>returns months</returns>
        private static int GetAgeInMonths(DateTime dateOfBirth, DateTime toDate)
        {
            int monthDiff = 0;
            while (dateOfBirth.AddMonths(monthDiff + 1) <= toDate)
            {
                monthDiff++;
            }

            return monthDiff;
        }

        /// <summary>
        /// Get Remaining Days
        /// </summary>
        /// <param name="addedDob">added Dob</param>
        /// <param name="toDate">to Date</param>
        /// <returns>returns days</returns>
        private static int GetRemainingDays(DateTime addedDob, DateTime toDate)
        {
            return (int)(toDate - addedDob).TotalDays;
        }

        /// <summary>
        /// Get Data FromXml
        /// </summary>
        /// <param name="value">Pass value.</param>
        /// <param name="attributeName">Pass attributeName .</param>
        /// <returns>html tags</returns>
        private static string GetDataFromXml(string value, string attributeName)
        {
            if (!string.IsNullOrEmpty(value))
            {
                XDocument document = XDocument.Parse(value);
                XElement element = document.Descendants().FirstOrDefault(ele => ele.Attributes().Any(attr => attr.Name == attributeName));
                return element == null ? string.Empty : element.Value;
            }

            return string.Empty;
        }

        #endregion Private Methods
    }
}
