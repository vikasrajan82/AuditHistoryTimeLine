// <copyright file="CrmConnector.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Crm Connection Class</summary>
namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Query;
    
    /// <summary>
    /// Request Type
    /// </summary>
    internal enum RequestType
    {
        /// <summary>
        /// Create Request
        /// </summary>
        Create,

        /// <summary>
        /// Update Request
        /// </summary>
        Update
    }

    /// <summary>
    /// CRM Connection class
    /// </summary>
    public static class CrmConnector
    {
        /// <summary>
        /// Retrieves the first or default entity
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityName">Name of the entity that has to be retrieved</param>
        /// <param name="columns">Array of column names </param>
        /// <param name="conditions">Filter condition</param>
        /// <param name="orders">Order condition</param>
        /// <returns>Retrieved entity</returns>
        public static Entity RetrieveFirstOrDefaultEntity(IOrganizationService service, string entityName, string[] columns, FilterExpression conditions, Collection<OrderExpression> orders)
        {
            ExtensionBase.IsObjectNotNull(service, new ArgumentNullException("service", "The service parameter passed to RetrieveFirstOrDefaultEntity cannot be null"));

            QueryExpression query = new QueryExpression();
            query.ColumnSet = new ColumnSet(columns);
            query.PageInfo = new PagingInfo();
            query.PageInfo.PageNumber = 1;
            query.PageInfo.Count = 1;
            query.EntityName = entityName;
            if (conditions != null)
            {
                query.Criteria = conditions;
            }

            SetOrderConditions(query, orders);

            EntityCollection entities = service.RetrieveMultiple(query);
            if (entities != null &&
                entities.Entities != null &&
                entities.Entities.Count > 0)
            {
                return entities.Entities[0];
            }

            return null;
        }

        /// <summary>
        /// Retrieves the first or default entity
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityName">Name of the entity that has to be retrieved</param>
        /// <param name="columns">Array of column names </param>
        /// <param name="conditions">Filter Expression</param>
        /// <returns>Retrieved entity</returns>
        public static Entity RetrieveFirstOrDefaultEntity(IOrganizationService service, string entityName, string[] columns, FilterExpression conditions)
        {
            return RetrieveFirstOrDefaultEntity(service, entityName, columns, conditions, null);
        }

        /// <summary>
        /// Retrieves the first or default entity
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityName">Name of the entity that has to be retrieved</param>
        /// <param name="columns">Array of column names </param>
        /// <param name="orders">Order condition</param>
        /// <returns>Retrieved entity</returns>
        public static Entity RetrieveFirstOrDefaultEntity(IOrganizationService service, string entityName, string[] columns, Collection<OrderExpression> orders)
        {
            return RetrieveFirstOrDefaultEntity(service, entityName, columns, null, orders);
        }

        /// <summary>
        /// Retrieves the first or default entity
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityName">Name of the entity that has to be retrieved</param>
        /// <param name="columns">Array of column names </param>
        /// <param name="conditions">Filter condition</param>
        /// <param name="orders">Order condition</param>
        /// <param name="count">count condition</param>
        /// <returns>Retrieved entity</returns>
        public static EntityCollection RetrieveMultiple(IOrganizationService service, string entityName, string[] columns, FilterExpression conditions, Collection<OrderExpression> orders, int count)
        {
            ExtensionBase.IsObjectNotNull(service, new ArgumentNullException("service", "The service parameter passed to RetrieveFirstOrDefaultEntity cannot be null"));

            QueryExpression query = new QueryExpression();
            query.ColumnSet = new ColumnSet(columns);
            query.EntityName = entityName;
            if (conditions != null)
            {
                query.Criteria = conditions;
            }

            SetOrderConditions(query, orders);

            if (count > 0)
            {
                query.PageInfo.PageNumber = 1;
                query.PageInfo.Count = count;
            }

            return service.RetrieveMultiple(query);
        }

        /// <summary>
        /// Retrieves all the records from RetrieveMultiple
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityName">Name of the entity that has to be retrieved</param>
        /// <param name="columns">Array of column names </param>
        /// <param name="conditions">Filter condition</param>
        /// <param name="orders">Order condition</param>
        /// <returns>Retrieved entity</returns>
        public static EntityCollection IterativeRetrieveMultiple(IOrganizationService service, string entityName, string[] columns, FilterExpression conditions, Collection<OrderExpression> orders)
        {
            ExtensionBase.IsObjectNotNull(service, new ArgumentNullException("service", "The service parameter passed to RetrieveFirstOrDefaultEntity cannot be null"));

            QueryExpression query = new QueryExpression();
            query.ColumnSet = new ColumnSet(columns);
            query.EntityName = entityName;
            if (conditions != null)
            {
                query.Criteria = conditions;
            }

            SetOrderConditions(query, orders);

            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;

            EntityCollection results = new EntityCollection();

            EntityCollection eachBatchResults = new EntityCollection();

            while (true)
            {
                eachBatchResults = service.RetrieveMultiple(query);

                if (eachBatchResults.HasRecords())
                {
                    results.Entities.AddRange(eachBatchResults.Entities.ToArray());
                }

                if (eachBatchResults.MoreRecords)
                {
                    // Increment the page number to retrieve the next page.
                    query.PageInfo.PageNumber++;

                    // Set the paging cookie to the paging cookie returned from current results.
                    query.PageInfo.PagingCookie = eachBatchResults.PagingCookie;
                }
                else
                {
                    // If no more records are in the result nodes, exit the loop.
                    break;
                }
            }

            return results;
        }

        /// <summary>
        /// Execute Multiple Update Request
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityCollection">Entity Collection</param>
        /// <param name="continueOnError">Continue on Error Indicator</param>
        /// <returns>Failed Record Count</returns>
        public static int ExecuteMultipleUpdateRequest(IOrganizationService service, EntityCollection entityCollection, bool continueOnError)
        {
            return ExecuteMultipleRequest(service, RequestType.Update, entityCollection, continueOnError);
        }

        /// <summary>
        /// Execute Multiple Create Request
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityCollection">Entity Collection</param>
        /// <param name="continueOnError">Continue on Error Indicator</param>
        /// <returns>Failed Record Count</returns>
        public static int ExecuteMultipleCreateRequest(IOrganizationService service, EntityCollection entityCollection, bool continueOnError)
        {
            return ExecuteMultipleRequest(service, RequestType.Create, entityCollection, continueOnError);
        }

        /// <summary>
        /// retrieve a CRM entity based on a simple query (generally the Id)
        /// </summary>
        /// <param name="service">IOrganizationService to execute requests</param>
        /// <param name="entityName">Name of CRM entity</param>
        /// <param name="attributeNames">attribute Names</param>
        /// <param name="filterAttributeName">The "column" name that will be searched</param>
        /// <param name="keyName">The Search Key/Id</param>
        /// <returns>returns CRM Entity record requested</returns>
        public static Entity GetCrmEntityUsingKey(IOrganizationService service, string entityName, string[] attributeNames, string filterAttributeName, string keyName)
        {
            if (service != null)
            {
                QueryExpression query = new QueryExpression(entityName);
                query.ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet(attributeNames);
                ConditionExpression con = new ConditionExpression();
                con.AttributeName = filterAttributeName;
                con.Operator = ConditionOperator.Equal;
                con.Values.Add(keyName);

                FilterExpression filter = new FilterExpression();
                filter.Conditions.Add(con);

                query.Criteria.AddFilter(filter);
                EntityCollection resultSet = service.RetrieveMultiple(query);
                if (resultSet.HasRecords())
                {
                    return resultSet.Entities[0];
                }
            }

            return null;
        }

        /// <summary>
        /// Set Order conditions for the query expression
        /// </summary>
        /// <param name="query">query expression</param>
        /// <param name="orders">Collection of orders</param>
        private static void SetOrderConditions(QueryExpression query, Collection<OrderExpression> orders)
        {
            if (orders != null)
            {
                foreach (OrderExpression order in orders)
                {
                    query.AddOrder(order.AttributeName, order.OrderType);
                }
            }
        }

        /// <summary>
        /// Execute Multiple Request
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="crmRequestType">Request Type</param>
        /// <param name="entityCollection">Entity Collection</param>
        /// <param name="continueOnError">Continue on Error Indicator</param>
        /// <returns>Failed Record Count</returns>
        private static int ExecuteMultipleRequest(IOrganizationService service, RequestType crmRequestType, EntityCollection entityCollection, bool continueOnError)
        {
            int failCount = 0;

            if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
            {
                OrganizationRequestCollection requestCollection = new OrganizationRequestCollection();
                foreach (var entity in entityCollection.Entities)
                {
                    switch (crmRequestType)
                    {
                        case RequestType.Create:
                            requestCollection.Add(new CreateRequest { Target = entity });
                            break;
                        case RequestType.Update:
                            requestCollection.Add(new UpdateRequest { Target = entity });
                            break;
                    }

                    if (requestCollection.Count == 1000)
                    {
                        failCount = failCount + ProcessExecuteMultipleAndReturnFailCount(requestCollection, service, continueOnError);
                        requestCollection.Clear();
                    }
                }

                if (requestCollection.Count > 0)
                {
                    failCount = failCount + ProcessExecuteMultipleAndReturnFailCount(requestCollection, service, continueOnError);
                }
            }

            return failCount;
        }

        /// <summary>
        /// Execute Multiple and return fail count.
        /// </summary>
        /// <param name="requestCollection">Request Collection</param>
        /// <param name="service">Organization Service</param>
        /// <param name="continueOnError">Continue on Error</param>
        /// <returns>Failed Record Count</returns>
        private static int ProcessExecuteMultipleAndReturnFailCount(OrganizationRequestCollection requestCollection, IOrganizationService service, bool continueOnError)
        {
            if (service != null)
            {
                int failCount = 0;
                var requestWithResults = new ExecuteMultipleRequest()
                {
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = continueOnError,
                        ReturnResponses = true
                    },
                    Requests = new OrganizationRequestCollection()
                };

                requestWithResults.Requests.AddRange(requestCollection);
                ExecuteMultipleResponse responseWithResults =
                    (ExecuteMultipleResponse)service.Execute(requestWithResults);

                // Display the results returned in the responses.
                foreach (var responseItem in responseWithResults.Responses)
                {
                    // An error has occurred.
                    if (responseItem.Fault != null)
                    {
                        failCount++;
                    }
                }

                return failCount;
            }

            return -1;
        }
    }
}
