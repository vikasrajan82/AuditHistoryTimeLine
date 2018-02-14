// <copyright file="CrmConnection.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Handles CRM Connection</summary>
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
    using Microsoft.Xrm.Sdk.Metadata;
    using Microsoft.Xrm.Sdk.Query;
    //using MIS.CRM.EntityWrapper.Constants;
    //using MIS.CRM.EntityWrapper.Entities;

    /// <summary>
    /// CRM Connection Manager
    /// </summary>
    public class CrmConnection : ICrmConnection
    {
        /// <summary>
        /// Query Expression
        /// </summary>
        private QueryExpression queryExpression;

        /// <summary>
        /// Current Linked Entity
        /// </summary>
        private LinkEntity currentLinkEntity;

        /// <summary>
        /// Holds the Filter Expression
        /// </summary>
        private FilterExpression expression;

        /// <summary>
        /// Filter expression first level
        /// </summary>
        private FilterExpression filterExpressionFirstLevel;

        /// <summary>
        /// link entity filter expression filter
        /// </summary>
        private FilterExpression linkEntityFilterExpressionFilter;

        /// <summary>
        /// Filter expression second level
        /// </summary>
        private FilterExpression filterExpressionSecondLevel;

        /// <summary>
        /// Collection of order expressions
        /// </summary>
        private Collection<OrderExpression> orders;

        /// <summary>
        /// page expression
        /// </summary>
        private PagingInfo page;

        /// <summary>
        /// Initializes a new instance of the <see cref="CrmConnection"/> class
        /// </summary>
        public CrmConnection()
        {
            this.expression = new FilterExpression();
            this.filterExpressionFirstLevel = new FilterExpression();
            this.filterExpressionSecondLevel = new FilterExpression();
            this.orders = new Collection<OrderExpression>();
            this.page = new PagingInfo();
        }

        /// <summary>
        /// Clears all array lists
        /// </summary>
        public void Clear()
        {
            this.expression.Conditions.Clear();
            this.filterExpressionFirstLevel.Conditions.Clear();
            this.filterExpressionFirstLevel.Filters.Clear();
            this.filterExpressionSecondLevel.Conditions.Clear();
            this.filterExpressionSecondLevel.Filters.Clear();
            this.expression.Filters.Clear();
            this.orders.Clear();
            this.page.Count = -1;
        }

        /// <summary>
        /// Adds the Filter Expression
        /// </summary>
        /// <param name="attributeName">Attribute Name</param>
        /// <param name="conditionOperator">Conditional Operator</param>
        /// <param name="values">Values to be included in the filter expression</param>
        public void AddFilterExpression(string attributeName, ConditionOperator conditionOperator, params object[] values)
        {
            this.expression.AddCondition(attributeName, conditionOperator, values);
        }

        /// <summary>
        /// Adds the Order Expression
        /// </summary>
        /// <param name="attributeName">Attribute Name</param>
        /// <param name="orderType">Order Type</param>
        public void AddOrderExpression(string attributeName, OrderType orderType)
        {
            this.orders.Add(new OrderExpression(attributeName, orderType));
        }

        /// <summary>
        /// Retrieves the First of Default Entity
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityName">Entity Name</param>
        /// <param name="columns">Columns to be retrieved</param>
        /// <returns>Returns the first Entity</returns>
        public Entity RetrieveFirstOrDefaultEntity(IOrganizationService service, string entityName, string[] columns)
        {
            return CrmConnector.RetrieveFirstOrDefaultEntity(service, entityName, columns, this.expression, this.orders);
        }

        /// <summary>
        /// Retrieves all the entities matching the criteria
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityName">Entity Name</param>
        /// <param name="columns">Columns to be retrieved</param>
        /// <returns>Returns the Entity Collection</returns>
        public EntityCollection RetrieveMultiple(IOrganizationService service, string entityName, string[] columns)
        {
            return CrmConnector.RetrieveMultiple(service, entityName, columns, this.expression, this.orders, this.page.Count);
        }

        /// <summary>
        /// Retrieves the First or Default Entity
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <returns>Returns the first entity</returns>
        public Entity RetrieveFirstOrDefaultEntityFromQueryExpression(IOrganizationService service)
        {
            if (service != null)
            {
                var entityCollection = service.RetrieveMultiple(this.queryExpression);
                if (entityCollection != null
                    && entityCollection.Entities != null
                    && entityCollection.Entities.Count > 0)
                {
                    return entityCollection.Entities[0];
                }
            }

            return null;
        }

        /// <summary>
        /// Retrieves the First or Default Entity
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <returns>Returns the first entity</returns>
        public EntityCollection RetrieveMultipleFromQueryExpression(IOrganizationService service)
        {
            if (service != null)
            {
                return service.RetrieveMultiple(this.queryExpression);
            }

            return null;
        }

        /// <summary>
        /// Build Query Expression
        /// </summary>
        /// <param name="entityName">Entity Name</param>
        /// <param name="columns">Columns to be fetched</param>
        public void BuildQueryExpression(string entityName, string[] columns)
        {
            this.queryExpression = new QueryExpression()
            {
                EntityName = entityName,
                ColumnSet = new ColumnSet(columns),
                Distinct = true
            };
        }

        /// <summary>
        /// Add New linked entity
        /// </summary>
        /// <param name="linkFromEntityName">From Entity Name</param>
        /// <param name="linkFromAttributeName">From Entity attribute Name</param>
        /// <param name="linkToEntityName">To Entity Name</param>
        /// <param name="linkToAttributeName">To Entity Attribute name</param>
        /// <param name="aliasName">Alias Name</param>
        /// <param name="joinOperator">Join Operator</param>
        public void AddNewLinkEntity(string linkFromEntityName, string linkFromAttributeName, string linkToEntityName, string linkToAttributeName, string aliasName, JoinOperator joinOperator)
        {
            this.currentLinkEntity = new LinkEntity()
                {
                    LinkFromEntityName = linkFromEntityName,
                    LinkFromAttributeName = linkFromAttributeName,
                    LinkToEntityName = linkToEntityName,
                    LinkToAttributeName = linkToAttributeName,
                    EntityAlias = aliasName,
                    JoinOperator = joinOperator
                };
        }

        /// <summary>
        /// Add New linked entity
        /// </summary>
        /// <param name="linkFromEntityName">From Entity Name</param>
        /// <param name="linkFromAttributeName">From Entity attribute Name</param>
        /// <param name="linkToEntityName">To Entity Name</param>
        /// <param name="linkToAttributeName">To Entity Attribute name</param>
        /// <param name="aliasName">Alias Name</param>
        public void AddNewLinkEntity(string linkFromEntityName, string linkFromAttributeName, string linkToEntityName, string linkToAttributeName, string aliasName)
        {
            this.AddNewLinkEntity(linkFromEntityName, linkFromAttributeName, linkToEntityName, linkToAttributeName, aliasName, JoinOperator.Inner);
        }

        /// <summary>
        /// Add New linked entity
        /// </summary>
        /// <param name="linkFromEntityName">From Entity Name</param>
        /// <param name="linkFromAttributeName">From Entity attribute Name</param>
        /// <param name="linkToEntityName">To Entity Name</param>
        /// <param name="linkToAttributeName">To Entity Attribute name</param>
        public void AddNewLinkEntity(string linkFromEntityName, string linkFromAttributeName, string linkToEntityName, string linkToAttributeName)
        {
            this.AddNewLinkEntity(linkFromEntityName, linkFromAttributeName, linkToEntityName, linkToAttributeName, null, JoinOperator.Inner);
        }

        /// <summary>
        /// Add Link Entity Filter condition
        /// </summary>
        /// <param name="attributeName">Attribute Name</param>
        /// <param name="conditionOperator">Conditional Operator</param>
        /// <param name="values">Values to be included in the filter expression</param>
        public void AddLinkEntityFilter(string attributeName, ConditionOperator conditionOperator, params object[] values)
        {
            this.currentLinkEntity.LinkCriteria.AddCondition(new ConditionExpression(attributeName, conditionOperator, values));
        }

        /// <summary>
        /// Add the columns to be fetched
        /// </summary>
        /// <param name="columns">List of columns to be fetched.</param>
        public void AddLinkEntityColumns(params string[] columns)
        {
            this.currentLinkEntity.Columns.AddColumns(columns);
        }

        /// <summary>
        /// Add Filter to Query Expression
        /// </summary>
        /// <param name="attributeName">Attribute Name</param>
        /// <param name="conditionOperator">Condition Operator</param>
        /// <param name="values">Values to be checked</param>
        public void AddQueryExpressionFilter(string attributeName, ConditionOperator conditionOperator, params object[] values)
        {
            this.queryExpression.Criteria.AddCondition(new ConditionExpression(attributeName, conditionOperator, values));
        }

        /// <summary>
        /// Add filter first level
        /// </summary>
        public void AddFilterFirstLevel()
        {
            this.queryExpression.Criteria.AddFilter(this.filterExpressionFirstLevel);
        }

        /// <summary>
        /// Add query expression filter first level logical operator
        /// </summary>
        /// <param name="logicalOperator">Logical operator</param>
        public void AddQueryExpressionFilterFirstLevelLogicalOperator(LogicalOperator logicalOperator)
        {
            this.filterExpressionFirstLevel = new FilterExpression(logicalOperator);
        }

        /// <summary>
        /// Add query expression filter first level logical operator
        /// </summary>
        /// <param name="logicalOperator">Logical operator</param>
        public void AddQueryExpressionLinkEntityFilterFirstLevelLogicalOperator(LogicalOperator logicalOperator)
        {
            this.linkEntityFilterExpressionFilter = new FilterExpression(logicalOperator);
        }

        /// <summary>
        /// Add link entity filter first level
        /// </summary>
        public void AddLinkEntityFilterFirstLevel()
        {
            this.currentLinkEntity.LinkCriteria.AddFilter(this.linkEntityFilterExpressionFilter);
        }

        /// <summary>
        /// Add query expression filter first level logical operator
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="conditionOperator">condition operator</param>
        /// <param name="values">filter expression columns</param>
        public void AddQueryExpressionLinkEntityFilterFirstLevel(string attributeName, ConditionOperator conditionOperator, params object[] values)
        {
            this.linkEntityFilterExpressionFilter.AddCondition(new ConditionExpression(attributeName, conditionOperator, values));
        }

        /// <summary>
        /// Add query expression filter first level logical operator
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="conditionOperator">condition operator</param>
        /// <param name="values">filter expression columns</param>
        public void AddQueryExpressionFilterFirstLevel(string attributeName, ConditionOperator conditionOperator, params object[] values)
        {
            this.filterExpressionFirstLevel.AddCondition(new ConditionExpression(attributeName, conditionOperator, values));
        }

        /// <summary>
        /// Add filter at second level
        /// </summary>
        public void AddFilterSecondLevel()
        {
            this.filterExpressionFirstLevel.AddFilter(this.filterExpressionSecondLevel);
        }

        /// <summary>
        /// Add query expression filter second level logical operator
        /// </summary>
        /// <param name="logicalOperator">logical operator</param>
        public void AddQueryExpressionFilterSecondLevelLogicalOperator(LogicalOperator logicalOperator)
        {
            this.filterExpressionSecondLevel = new FilterExpression(logicalOperator);
        }

        /// <summary>
        /// Add query expression filter second level
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="conditionOperator">condition operator</param>
        /// <param name="values">filter expression</param>
        public void AddQueryExpressionFilterSecondLevel(string attributeName, ConditionOperator conditionOperator, params object[] values)
        {
            this.filterExpressionSecondLevel.AddCondition(new ConditionExpression(attributeName, conditionOperator, values));
        }

        /// <summary>
        /// Add logical operator to Query Expression
        /// </summary>
        /// <param name="logicalOperator">Logical Operator</param>
        public void AddQueryExpressionLogicalOperator(LogicalOperator logicalOperator)
        {
            this.queryExpression.Criteria.FilterOperator = logicalOperator;
        }

        /// <summary>
        /// Add Order to Query Expression
        /// </summary>
        /// <param name="attributeName">Attribute Name</param>
        /// <param name="orderType">Order Type</param>
        public void AddQueryExpressionOrder(string attributeName, OrderType orderType)
        {
            this.queryExpression.AddOrder(attributeName, orderType);
        }

        /// <summary>
        /// add page expression
        /// </summary>
        /// <param name="count">the number of records</param>
        public void AddPageExpression(int count)
        {
            this.page.PageNumber = 1;
            this.page.Count = count;
        }

        /// <summary>
        /// Add Link Entity To first level
        /// </summary>
        public void AddLinkedEntityToFirstLevel()
        {
            this.queryExpression.LinkEntities.Add(this.currentLinkEntity);
        }

        /// <summary>
        /// Add link entity to second level
        /// </summary>
        public void AddLinkedEntityToSecondLevel()
        {
            this.AddLinkedEntityToSecondLevel(-1);
        }

        /// <summary>
        /// Add link entity to second level
        /// </summary>
        /// <param name="position">Position within the linked entity collection</param>
        public void AddLinkedEntityToSecondLevel(int position)
        {
            if (position == -1)
            {
                position = this.queryExpression.LinkEntities.Count - 1;
            }

            this.queryExpression.LinkEntities[position].LinkEntities.Add(this.currentLinkEntity);
        }

        /// <summary>
        /// Add link entity to third level
        /// </summary>
        public void AddLinkedEntityToThirdLevel()
        {
            this.AddLinkedEntityToThirdLevel(-1, -1);
        }

        /// <summary>
        /// Add link entity to third level
        /// </summary>
        /// <param name="firstLevelPosition">Position within the first level linked entity collection</param>
        /// <param name="secondLevelPosition">Position within the second level linked entity collection</param>
        public void AddLinkedEntityToThirdLevel(int firstLevelPosition, int secondLevelPosition)
        {
            if (firstLevelPosition == -1)
            {
                firstLevelPosition = this.queryExpression.LinkEntities.Count - 1;
            }

            if (secondLevelPosition == -1)
            {
                secondLevelPosition = this.queryExpression.LinkEntities[firstLevelPosition].LinkEntities.Count - 1;
            }

            this.queryExpression.LinkEntities[firstLevelPosition].LinkEntities[secondLevelPosition].LinkEntities.Add(this.currentLinkEntity);
        }

        /// <summary>
        /// Adds the linked entity to the fourth level
        /// </summary>
        public void AddLinkedEntityToFourthLevel()
        {
            this.AddLinkedEntityToFourthLevel(-1, -1, -1);
        }

        /// <summary>
        /// Adds the linked entity to the fourth level
        /// </summary>
        /// <param name="firstLevelPosition">First Level Position</param>
        /// <param name="secondLevelPosition">Second Level Position</param>
        /// <param name="thirdLevelPosition">Third Level Position</param>
        public void AddLinkedEntityToFourthLevel(int firstLevelPosition, int secondLevelPosition, int thirdLevelPosition)
        {
            if (firstLevelPosition == -1)
            {
                firstLevelPosition = this.queryExpression.LinkEntities.Count - 1;
            }

            if (secondLevelPosition == -1)
            {
                secondLevelPosition = this.queryExpression.LinkEntities[firstLevelPosition].LinkEntities.Count - 1;
            }

            if (thirdLevelPosition == -1)
            {
                thirdLevelPosition = this.queryExpression.LinkEntities[firstLevelPosition].LinkEntities[secondLevelPosition].LinkEntities.Count - 1;
            }

            this.queryExpression.LinkEntities[firstLevelPosition].LinkEntities[secondLevelPosition].LinkEntities[thirdLevelPosition].LinkEntities.Add(this.currentLinkEntity);
        }

        /// <summary>
        /// Execute Multiple Update Request
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityCollection">Entity Collection</param>
        /// <param name="continueOnError">Continue On Error indicator</param>
        /// <returns>Failed Record Count</returns>
        public int ExecuteMultipleUpdateRequest(IOrganizationService service, EntityCollection entityCollection, bool continueOnError)
        {
            return CrmConnector.ExecuteMultipleUpdateRequest(service, entityCollection, continueOnError);
        }

        /// <summary>
        /// Execute Multiple Create Request
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityCollection">Entity Collection</param>
        /// <param name="continueOnError">Continue on Error Indicator</param>
        /// <returns>Failed Record Count</returns>
        public int ExecuteMultipleCreateRequest(IOrganizationService service, EntityCollection entityCollection, bool continueOnError)
        {
            return CrmConnector.ExecuteMultipleCreateRequest(service, entityCollection, continueOnError);
        }

        /// <summary>
        /// Converts query expression to fetch xml
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <returns>Fetch XML</returns>
        public string ConvertQueryExpressionToFetchXml(IOrganizationService service)
        {
            if (service != null)
            {
                var conversionRequest = new QueryExpressionToFetchXmlRequest
                {
                    Query = this.queryExpression
                };

                var conversionResponse =
                    (QueryExpressionToFetchXmlResponse)service.Execute(conversionRequest);

                return conversionResponse.FetchXml;
            }

            return string.Empty;
        }


       
        /// <summary>
        /// Execute the Fetch Xml
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="fetchXml">Fetch Xml</param>
        /// <returns>Entity Collection result</returns>
        public EntityCollection ExecuteFetchXml(IOrganizationService service, string fetchXml)
        {
            if (service != null)
            {
                return service.RetrieveMultiple(new FetchExpression(fetchXml));
            }

            return null;
        }

        /// <summary>
        /// Execute the Fetch Xml in a recursive manner
        /// This is required to retrieve more than 5000 records for this fetch xml execution
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="fetchXml">Fetch Xml</param>
        /// <returns>Entity Collection result</returns>
        public Collection<Entity> ExecuteRecursiveFetchXml(IOrganizationService service, string fetchXml)
        {
            if (service != null)
            {
                FetchXmlAnalyzer xmlAnalyzer = new FetchXmlAnalyzer(fetchXml);
                return xmlAnalyzer.ExecuteFetchXml(service);
            }

            return null;
        }

        /// <summary>
        /// Convert to fetch xml and get aggregate count
        /// </summary>
        /// <param name="service">the service</param>
        /// <param name="attributeName">attribute name</param>
        /// <returns>returns the count of records</returns>
        public int ConvertToFetchXmlAndGetAggregateCount(IOrganizationService service, string attributeName)
        {
            int recordCount = 0;
            string fetchxmlResponse = this.ConvertQueryExpressionToFetchXml(service);

            if (fetchxmlResponse != null)
            {
                fetchxmlResponse = fetchxmlResponse.Replace("mapping=\"logical\"", "mapping=\"logical\" aggregate =\"true\"");
                fetchxmlResponse = fetchxmlResponse.Replace("name=\"" + attributeName + "\"", "name=\"" + attributeName + "\" alias=\"recordcount\" aggregate=\"count\"");
                EntityCollection recordCollection = this.ExecuteFetchXml(service, fetchxmlResponse);

                if (recordCollection.Entities.Count > 0)
                {
                    recordCount = Convert.ToInt32(recordCollection[0].GetAttributeValue<AliasedValue>("recordcount").Value, CultureInfo.InvariantCulture);
                }
            }

            return recordCount;
        }

        /// <summary>
        /// This function returns the option set text
        /// </summary>
        /// <param name="orgService">org Service</param>
        /// <param name="entityName">entity Name</param>
        /// <param name="attributeName">attribute Name</param>
        /// <param name="optionSetValue">option set value</param>
        /// <returns>option set text</returns>
        public string GetOptionSetText(IOrganizationService orgService, string entityName, string attributeName, int optionSetValue)
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
    }
}
