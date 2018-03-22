// <copyright file="ICrmConnection.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Crm Connection Interface</summary>
namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;
    
    /// <summary>
    /// CRM Connection Interface
    /// </summary>
    public interface ICrmConnection
    {
        /// <summary>
        /// Clear all array list
        /// </summary>
        void Clear();

        /// <summary>
        /// Adds the Filter Expression
        /// </summary>
        /// <param name="attributeName">Attribute Name</param>
        /// <param name="conditionOperator">Conditional Operator</param>
        /// <param name="values">Values to be included in the filter expression</param>
        void AddFilterExpression(string attributeName, ConditionOperator conditionOperator, params object[] values);

        /// <summary>
        /// Adds the Order Expression
        /// </summary>
        /// <param name="attributeName">Attribute Name</param>
        /// <param name="orderType">Order Type</param>
        void AddOrderExpression(string attributeName, OrderType orderType);

        /// <summary>
        /// Retrieves the First of Default Entity
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityName">Entity Name</param>
        /// <param name="columns">Columns to be retrieved</param>
        /// <returns>Returns the Entity</returns>
        Entity RetrieveFirstOrDefaultEntity(IOrganizationService service, string entityName, string[] columns);

        /// <summary>
        /// Retrieves the First or Default Entity
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <returns>Returns the first entity</returns>
        Entity RetrieveFirstOrDefaultEntityFromQueryExpression(IOrganizationService service);

        /// <summary>
        /// Retrieves all the entities matching the criteria
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityName">Entity Name</param>
        /// <param name="columns">Columns to be retrieved</param>
        /// <returns>Returns the Entity Collection</returns>
        EntityCollection RetrieveMultiple(IOrganizationService service, string entityName, string[] columns);

        /// <summary>
        /// Retrieves all the entities matching the criteria
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <returns>Returns the Entity Collection</returns>
        EntityCollection RetrieveMultipleFromQueryExpression(IOrganizationService service);

        /// <summary>
        /// Build Query Expression
        /// </summary>
        /// <param name="entityName">Entity Name</param>
        /// <param name="columns">Columns to be fetched</param>
        void BuildQueryExpression(string entityName, string[] columns);

        /// <summary>
        /// Add New linked entity
        /// </summary>
        /// <param name="linkFromEntityName">From Entity Name</param>
        /// <param name="linkFromAttributeName">From Entity attribute Name</param>
        /// <param name="linkToEntityName">To Entity Name</param>
        /// <param name="linkToAttributeName">To Entity Attribute name</param>
        void AddNewLinkEntity(string linkFromEntityName, string linkFromAttributeName, string linkToEntityName, string linkToAttributeName);

        /// <summary>
        /// Add New linked entity
        /// </summary>
        /// <param name="linkFromEntityName">From Entity Name</param>
        /// <param name="linkFromAttributeName">From Entity attribute Name</param>
        /// <param name="linkToEntityName">To Entity Name</param>
        /// <param name="linkToAttributeName">To Entity Attribute name</param>
        /// <param name="aliasName">Alias name</param>
        void AddNewLinkEntity(string linkFromEntityName, string linkFromAttributeName, string linkToEntityName, string linkToAttributeName, string aliasName);

        /// <summary>
        /// Add New linked entity
        /// </summary>
        /// <param name="linkFromEntityName">From Entity Name</param>
        /// <param name="linkFromAttributeName">From Entity attribute Name</param>
        /// <param name="linkToEntityName">To Entity Name</param>
        /// <param name="linkToAttributeName">To Entity Attribute name</param>
        /// <param name="aliasName">Alias name</param>
        /// <param name="joinOperator">Join Operator</param>
        void AddNewLinkEntity(string linkFromEntityName, string linkFromAttributeName, string linkToEntityName, string linkToAttributeName, string aliasName, JoinOperator joinOperator);

        /// <summary>
        /// Add Link Entity Filter condition
        /// </summary>
        /// <param name="attributeName">Attribute Name</param>
        /// <param name="conditionOperator">Conditional Operator</param>
        /// <param name="values">Values to be included in the filter expression</param>
        void AddLinkEntityFilter(string attributeName, ConditionOperator conditionOperator, params object[] values);

        /// <summary>
        /// Add Link Entity Columns
        /// </summary>
        /// <param name="columns">Columns to be added</param>
        void AddLinkEntityColumns(params string[] columns);

        /// <summary>
        /// Add Link Entity To first level
        /// </summary>
        void AddLinkedEntityToFirstLevel();

        /// <summary>
        /// Add link entity to second level
        /// </summary>
        void AddLinkedEntityToSecondLevel();

        /// <summary>
        /// Add link entity to second level
        /// </summary>
        /// <param name="position">Position within the linked entity collection</param>
        void AddLinkedEntityToSecondLevel(int position);

        /// <summary>
        /// Add link entity to third level
        /// </summary>
        void AddLinkedEntityToThirdLevel();

        /// <summary>
        /// Add link entity to third level
        /// </summary>
        /// <param name="firstLevelPosition">Position within the first level linked entity collection</param>
        /// <param name="secondLevelPosition">Position within the second level linked entity collection</param>
        void AddLinkedEntityToThirdLevel(int firstLevelPosition, int secondLevelPosition);

        /// <summary>
        /// Adds the linked entity to the fourth level
        /// </summary>
        void AddLinkedEntityToFourthLevel();

        /// <summary>
        /// Converts query expression to fetch xml
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <returns>Fetch XML</returns>
        string ConvertQueryExpressionToFetchXml(IOrganizationService service);

        /// <summary>
        /// Convert to fetch xml and get the aggregate count
        /// </summary>
        /// <param name="service">the service</param>
        /// <param name="attributeName">attribute name</param>
        /// <returns>count of records</returns>
        int ConvertToFetchXmlAndGetAggregateCount(IOrganizationService service, string attributeName);

        /// <summary>
        /// Add Filter to Query Expression
        /// </summary>
        /// <param name="attributeName">Attribute Name</param>
        /// <param name="conditionOperator">Condition Operator</param>
        /// <param name="values">Values to be checked</param>
        void AddQueryExpressionFilter(string attributeName, ConditionOperator conditionOperator, params object[] values);

        /// <summary>
        /// Add logical operator to Query Expression
        /// </summary>
        /// <param name="logicalOperator">Logical Operator</param>
        void AddQueryExpressionLogicalOperator(LogicalOperator logicalOperator);

        /// <summary>
        /// Add Order to Query Expression
        /// </summary>
        /// <param name="attributeName">Attribute Name</param>
        /// <param name="orderType">Order Type</param>
        void AddQueryExpressionOrder(string attributeName, OrderType orderType);

        /// <summary>
        /// Add Filter at first level
        /// </summary>
        void AddFilterFirstLevel();

        /// <summary>
        /// Add query expression filter first level logical operator
        /// </summary>
        /// <param name="logicalOperator">logical operator</param>
        void AddQueryExpressionFilterFirstLevelLogicalOperator(LogicalOperator logicalOperator);

        /// <summary>
        /// Add query expression link entity filter first level logical operator
        /// </summary>
        /// <param name="logicalOperator">logical operator</param>
        void AddQueryExpressionLinkEntityFilterFirstLevelLogicalOperator(LogicalOperator logicalOperator);

        /// <summary>
        /// Add query expression link entity filter first level
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="conditionOperator">condition operator</param>
        /// <param name="values">attribute names</param>
        void AddQueryExpressionLinkEntityFilterFirstLevel(string attributeName, ConditionOperator conditionOperator, params object[] values);

        /// <summary>
        /// Add link entity filter first level
        /// </summary>
        void AddLinkEntityFilterFirstLevel();

        /// <summary>
        /// Add query expression filter at first level
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="conditionOperator">conditional operator</param>
        /// <param name="values">column values</param>
        void AddQueryExpressionFilterFirstLevel(string attributeName, ConditionOperator conditionOperator, params object[] values);

        /// <summary>
        /// Add filter at second level
        /// </summary>
        void AddFilterSecondLevel();

        /// <summary>
        /// Add query expression filter at second level logical operator
        /// </summary>
        /// <param name="logicalOperator">logical operator</param>
        void AddQueryExpressionFilterSecondLevelLogicalOperator(LogicalOperator logicalOperator);

        /// <summary>
        /// Add query expression filter at second level with columns
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="conditionOperator">conditional operator</param>
        /// <param name="values">column values</param>
        void AddQueryExpressionFilterSecondLevel(string attributeName, ConditionOperator conditionOperator, params object[] values);

        /// <summary>
        /// add page count to query expression
        /// </summary>
        /// <param name="count">the number of records</param>
        void AddPageExpression(int count);

        /// <summary>
        /// Execute Multiple Update Request
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityCollection">Entity Collection</param>
        /// <param name="continueOnError">Continue on Error</param>
        /// <returns>Failed Record Count</returns>
        int ExecuteMultipleUpdateRequest(IOrganizationService service, EntityCollection entityCollection, bool continueOnError);

        /// <summary>
        /// Execute Multiple Create Request
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="entityCollection">Entity Collection</param>
        /// <param name="continueOnError">Continue on Error</param>
        /// <returns>Failed Record Count</returns>
        int ExecuteMultipleCreateRequest(IOrganizationService service, EntityCollection entityCollection, bool continueOnError);

        /// <summary>
        /// Execute the Fetch Xml
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="fetchXml">Fetch Xml</param>
        /// <returns>Entity Collection result</returns>
        EntityCollection ExecuteFetchXml(IOrganizationService service, string fetchXml);

        /// <summary>
        /// Execute the Fetch Xml in a recursive manner
        /// This is required to retrieve more than 5000 records for this fetch xml execution
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="fetchXml">Fetch Xml</param>
        /// <returns>Entity Collection result</returns>
        Collection<Entity> ExecuteRecursiveFetchXml(IOrganizationService service, string fetchXml);

        /// <summary>
        /// Gets the Option set value's display name
        /// </summary>
        /// <param name="orgService">Organization Service</param>
        /// <param name="entityName">Entity Name</param>
        /// <param name="attributeName">Attribute Name</param>
        /// <param name="optionSetValue">Option set value</param>
        /// <returns>Display name of the Option set</returns>
        //string GetOptionSetText(IOrganizationService orgService, string entityName, string attributeName, int optionSetValue);
    }
}
