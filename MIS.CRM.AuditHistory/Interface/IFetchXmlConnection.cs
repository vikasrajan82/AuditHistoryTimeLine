// <copyright file="IFetchXmlConnection.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Handles CRM Connection</summary>
namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;

    /// <summary>
    /// fetch xml condition
    /// </summary>
    public enum Condition
    {
        /// <summary>
        /// equal function
        /// </summary>
        EQ,

        /// <summary>
        /// not equal function
        /// </summary>
        NE
    }

    /// <summary>
    /// aggregate value
    /// </summary>
    public enum Aggregate
    {
        /// <summary>
        /// sum operation
        /// </summary>
        Sum,

        /// <summary>
        /// count operation
        /// </summary>
        Count
    }

    /// <summary>
    /// IFetch xml connection
    /// </summary>
    public interface IFetchXmlConnection
    {
        /// <summary>
        /// Clear all array list
        /// </summary>
        void Clear();

        /// <summary>
        /// Build fetch xml expression
        /// </summary>
        /// <param name="entityName">entity name</param>
        void BuildFetchXmlExpression(string entityName);

        /// <summary>
        /// Build fetch xml expression
        /// </summary>
        /// <param name="entityName">entity name</param>
        /// <param name="isDistinct">is distinct</param>
        void BuildFetchXmlExpression(string entityName, bool isDistinct);

        /// <summary>
        /// Build fetch xml expression
        /// </summary>
        /// <param name="entityName">entity name</param>
        /// <param name="isDistinct">is distinct</param>
        /// <param name="isAggregate">is aggregate</param>
        void BuildFetchXmlExpression(string entityName, bool isDistinct, bool isAggregate);

        /// <summary>
        /// Add fetch xml attribute to entity
        /// </summary>
        /// <param name="attributeNames">attribute name</param>
        void AddFetchXmlAttributeToEntity(string[] attributeNames);

        /// <summary>
        /// Add fetch xml aggregate attribute to entity 
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="arithmeticOperation">arithmetic Operation</param>
        void AddFetchXmlArithmeticAttributeToEntity(string attributeName, string arithmeticOperation);

        /// <summary>
        /// Add fetch xml condition
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="fetchXmlCondition">fetchXml Condition</param>
        /// <param name="attributeValue">attribute value</param>
        void AddFetchXmlCondition(string attributeName, Condition fetchXmlCondition, string attributeValue);

        /// <summary>
        /// Add fetch xml link entity first level
        /// </summary>
        /// <param name="entityName">entity name</param>
        /// <param name="fromAttributeName">from attribute name</param>
        /// <param name="toAttributeName">to attribute name</param>
        /// <param name="linkType">link type</param>
        /// <param name="aliasName">alias name</param>
        void AddFetchXmlLinkEntityFirstLevel(string entityName, string fromAttributeName, string toAttributeName, string linkType, string aliasName);

        /// <summary>
        /// Add fetch xml link entity second level
        /// </summary>
        /// <param name="entityName">entity name</param>
        /// <param name="fromAttributeName">from attribute name</param>
        /// <param name="toAttributeName">to attribute name</param>
        /// <param name="linkType">link type</param>
        /// <param name="aliasName">alias name</param>
        void AddFetchXmlLinkEntitySecondLevel(string entityName, string fromAttributeName, string toAttributeName, string linkType, string aliasName);

        /// <summary>
        /// Add fetch xml filter condition first level
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="attributeValue">attribute value</param>
        void AddFetchXmlFilterConditionFirstLevel(string attributeName, string attributeValue);

        /// <summary>
        /// Add fetch xml filter condition second level
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="attributeValue">attribute value</param>
        void AddFetchXmlFilterConditionSecondLevel(string attributeName, string attributeValue);

        /// <summary>
        /// Retrieve fetch xml query
        /// </summary>
        /// <returns>returns the xml</returns>
        string RetrieveFetchXmlQuery();

        /// <summary>
        /// Execute fetch xml and get aggregate count
        /// </summary>
        /// <param name="service">the service</param>
        /// <param name="query">the query</param>
        /// <param name="aggregateValue">aggregate Value</param>
        /// <returns>returns the count</returns>
        int ExecuteFetchXmlAndGetAggregateCount(IOrganizationService service, string query, string aggregateValue);

        /// <summary>
        /// Execute fetch xml and get aggregate count
        /// </summary>
        /// <param name="service">the service</param>
        /// <param name="query">the query</param>
        /// <param name="aggregateValue">aggregate Value</param>
        /// <returns>returns the count</returns>
        EntityCollection ExecuteFetchXmlAndGetAggregateSum(IOrganizationService service, string query, string aggregateValue);
    }
}
