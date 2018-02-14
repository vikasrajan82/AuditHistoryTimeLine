// <copyright file="CrmObjectExtensions.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Extension Class</summary>
namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Globalization;
    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;
    //using MIS.CRM.EntityWrapper.Entities;

    /// <summary>
    /// CRM Object Extensions
    /// </summary>
    public static class CrmObjectExtensions
    {
        /// <summary>
        /// Checks if the entity collection has records
        /// </summary>
        /// <param name="entityCollection">Entity Collection</param>
        /// <returns>Indicates if the entity collection has records</returns>
        public static bool HasRecords(this EntityCollection entityCollection)
        {
            return entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0;
        }

        /// <summary>
        /// Extension method for universal date time zone it will need organization service and input
        /// </summary>
        /// <param name="coordinatedUniversalTime">Coordinated Universal Time</param>
        /// <param name="service">organization service</param>
        /// <returns>local user date time </returns>
        public static DateTime LocalDateTime(this DateTime coordinatedUniversalTime, IOrganizationService service)
        {
            int timeZoneCode;
            if (service != null)
            {
                timeZoneCode = GetUserTimeZoneCode(service);
                if (timeZoneCode != -1)
                {
                    var request = new LocalTimeFromUtcTimeRequest
                    {
                        TimeZoneCode = timeZoneCode,
                        UtcTime = coordinatedUniversalTime.ToUniversalTime()
                    };

                    var response = (LocalTimeFromUtcTimeResponse)service.Execute(request);
                    return response.LocalTime;
                }
            }

            return coordinatedUniversalTime;
        }

        /// <summary>
        /// This method gets user time zone code.
        /// </summary>
        /// <param name="service"> organization service </param>
        /// <returns> integer value </returns>
        public static int GetUserTimeZoneCode(IOrganizationService service)
        {
            int timeZoneCode = -1;
            if (service != null)
            {
                var currentUserSettings = service.RetrieveMultiple(
                new QueryExpression("usersettings")
                {
                    ColumnSet = new ColumnSet("timezonecode"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("systemuserid", ConditionOperator.EqualUserId)
                        }
                    }
                });

                if (currentUserSettings.HasRecords() && currentUserSettings.Entities[0].Attributes["timezonecode"] != null)
                {
                    timeZoneCode = (int)currentUserSettings.Entities[0].Attributes["timezonecode"];
                }
            }

            return timeZoneCode;
        }

        /// <summary>
        /// This method gets user time zone code.
        /// </summary>
        /// <param name="service"> organization service </param>
        /// <param name="userId"> user Id </param>
        /// <returns> integer value </returns>
        public static int GetUserTimeZoneCode(IOrganizationService service, Guid userId)
        {
            int timeZoneCode = -1;

            if (service != null)
            {
                var currentUserSettings = service.RetrieveMultiple(
                new QueryExpression("usersettings")
                {
                    ColumnSet = new ColumnSet("timezonecode"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("systemuserid", ConditionOperator.Equal, userId)
                        }
                    }
                });

                if (currentUserSettings.HasRecords() && currentUserSettings.Entities[0].Attributes["timezonecode"] != null)
                {
                    timeZoneCode = (int)currentUserSettings.Entities[0].Attributes["timezonecode"];
                }
            }

            return timeZoneCode;
        }

        /// <summary>
        /// This method gets user time zone code.
        /// </summary>
        /// <param name="impersonatedUser"> impersonated User</param>
        /// <param name="userId"> user Id </param>
        /// <param name="factory">service factory</param>
        /// <returns> integer value </returns>
        public static int GetImpersonatedUserTimeZoneCode(Guid impersonatedUser, Guid userId, IOrganizationServiceFactory factory)
        {
            int timeZoneCode = -1;

            if (factory != null)
            {
                IOrganizationService impersonatedService = factory.CreateOrganizationService(impersonatedUser);

                if (impersonatedService != null)
                {
                    var currentUserSettings = impersonatedService.RetrieveMultiple(
                    new QueryExpression("usersettings")
                    {
                        ColumnSet = new ColumnSet("timezonecode"),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                        {
                            new ConditionExpression("systemuserid", ConditionOperator.Equal, userId)
                        }
                        }
                    });

                    if (currentUserSettings.HasRecords() && currentUserSettings.Entities[0].Attributes["timezonecode"] != null)
                    {
                        timeZoneCode = (int)currentUserSettings.Entities[0].Attributes["timezonecode"];
                    }
                }
            }

            return timeZoneCode;
        }

        /// <summary>
        /// This method gets user local time (server time).
        /// </summary>
        /// <param name="coordinatedUniversalTime"> universal time </param>
        /// <param name="timeZoneCode"> time zone code </param>
        /// <param name="service"> organization service </param>
        /// <returns> date time </returns>
        public static DateTime GetUserLocalTime(DateTime coordinatedUniversalTime, int timeZoneCode, IOrganizationService service)
        {
            if (service != null)
            {
                var request = new LocalTimeFromUtcTimeRequest
                {
                    TimeZoneCode = timeZoneCode,
                    UtcTime = coordinatedUniversalTime
                };

                var response = (LocalTimeFromUtcTimeResponse)service.Execute(request);
                return response.LocalTime;
            }

            return coordinatedUniversalTime;
        }

        /// <summary>
        /// This method returns user time zone info.
        /// </summary>
        /// <param name="crmTimeZoneCode"> time zone code </param>
        /// <param name="service"> organization service </param>
        /// <returns> Time zone info </returns>
        public static TimeZoneInfo UserGetTimeZoneInfo(int crmTimeZoneCode, IOrganizationService service)
        {
            if (service != null)
            {
                var qe = new QueryExpression(TimeZoneDefinition.EntityLogicalName);
                qe.ColumnSet = new ColumnSet("standardname");
                qe.Criteria.AddCondition("timezonecode", ConditionOperator.Equal, crmTimeZoneCode);
                return TimeZoneInfo.FindSystemTimeZoneById(service.RetrieveMultiple(qe).Entities[0].ToEntity<TimeZoneDefinition>().StandardName);
            }

            return null;
        }

        /// <summary>
        /// Extension method for get date string in MM/DD/YYYY format
        /// </summary>
        /// <param name="universalTime">UTC Time</param>
        /// <returns>local user date time </returns>
        public static string GetWICDateFormat(this DateTime universalTime)
        {
            return universalTime.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Set Time to date
        /// </summary>
        /// <param name="dateTime"> date time </param>
        /// <param name="hours"> hours to set </param>
        /// <param name="minutes"> minutes to set </param>
        /// <param name="seconds"> Seconds to set </param>
        /// <returns> Date with time part </returns>
        public static DateTime SetTime(this DateTime dateTime, int hours, int minutes, int seconds)
        {
            return new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, hours, minutes, seconds, dateTime.Kind);
        }

        /// <summary>
        /// Checks whether this date falls in current month
        /// </summary>
        /// <param name="date"> date to check </param>
        /// <returns>Is this month </returns>
        public static bool IsThisMonth(this DateTime date)
        {
            date = date.Date;
            return date >= CRMHelper.FirstDayOfMonth(DateTime.Today) && date <= CRMHelper.LastDayOfMonth(DateTime.Today);
        }
    }
}
