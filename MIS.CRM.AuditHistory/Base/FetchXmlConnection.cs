// <copyright file="FetchXmlConnection.cs" company="Microsoft">
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
    using System.Xml;
    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Query;
    //using MIS.CRM.EntityWrapper.Constants;
    //using MIS.CRM.EntityWrapper.Entities;

    /// <summary>
    /// Fetch xml connection
    /// </summary>
    public class FetchXmlConnection : IFetchXmlConnection
    {
        /// <summary>
        /// fetch xml
        /// </summary>
        private XmlDocument fetchXml;

        /// <summary>
        /// fetch xml node
        /// </summary>
        private XmlNode fetchXmlNode;

        /// <summary>
        /// fetch xml attribute
        /// </summary>
        private XmlAttribute fetchXmlAttribute;

        /// <summary>
        /// entity root element
        /// </summary>
        private XmlNode entityRootElement;

        /// <summary>
        /// entity filter
        /// </summary>
        private XmlNode entityFilter;

        /// <summary>
        /// link entity root second level
        /// </summary>
        private XmlNode linkEntityRootSecondLevel;

        /// <summary>
        /// link entity root first level
        /// </summary>
        private XmlNode linkEntityRootFirstLevel;

        /// <summary>
        /// condition link entity 1
        /// </summary>
        private XmlNode conditionLinkEnity1;

        /// <summary>
        /// condition link entity 2
        /// </summary>
        private XmlNode conditionLinkEnity2;

        /// <summary>
        /// Filter root link entity 2
        /// </summary>
        private XmlNode filterRootLinkEnity2;

        /// <summary>
        /// Filter root link entity 1
        /// </summary>
        private XmlNode filterRootLinkEnity1;

        /// <summary>
        /// is entity filter added
        /// </summary>
        private bool isEntityFilterAdded = false;

        /// <summary>
        /// record count
        /// </summary>
        private int recordCount;

         /// <summary>
        /// Clears all array lists
        /// </summary>
        public void Clear()
        {
            if (this.fetchXml != null)
            {
                this.fetchXml.RemoveAll();
            }

            if (this.fetchXmlAttribute != null)
            {
                this.fetchXmlAttribute.RemoveAll();
            }

            if (this.entityRootElement != null)
            {
                this.entityRootElement.RemoveAll();
            }

            if (this.entityFilter != null)
            {
                this.entityFilter.RemoveAll();
            }

            if (this.linkEntityRootSecondLevel != null)
            {
                this.linkEntityRootSecondLevel.RemoveAll();
            }

            if (this.linkEntityRootFirstLevel != null)
            {
                this.linkEntityRootFirstLevel.RemoveAll();
            }

            if (this.conditionLinkEnity1 != null)
            {
                this.conditionLinkEnity1.RemoveAll();
            }

            if (this.conditionLinkEnity2 != null)
            {
                this.conditionLinkEnity2.RemoveAll();
            }

            this.isEntityFilterAdded = false;
        }

        /// <summary>
        /// Build fetch xml expression
        /// </summary>
        /// <param name="entityName">entity name</param>
        public void BuildFetchXmlExpression(string entityName)
        {
            this.CreateFetchXml();

            this.CreateEntityInFetchXml(entityName);
        }

        /// <summary>
        /// Build Fetch xml expression
        /// </summary>
        /// <param name="entityName">entity name</param>
        /// <param name="isDistinct">is distinct</param>
        public void BuildFetchXmlExpression(string entityName, bool isDistinct)
        {
            this.CreateFetchXml();

            this.CreateEntityInFetchXml(entityName);

            XmlAttribute distinctAttribute = this.fetchXml.CreateAttribute(FetchXmlKeyword.Distinct);
            distinctAttribute.Value = Convert.ToBoolean(isDistinct, CultureInfo.InvariantCulture) ? "true" : "false";
            this.fetchXmlNode.Attributes.Append(distinctAttribute);

            XmlAttribute aggregateAttribute = this.fetchXml.CreateAttribute(FetchXmlKeyword.Aggregate);
            aggregateAttribute.Value = "false";
            this.fetchXmlNode.Attributes.Append(aggregateAttribute);
        }

       /// <summary>
       /// Build fetch xml expression
       /// </summary>
       /// <param name="entityName">entity name</param>
       /// <param name="isDistinct">is distinct</param>
       /// <param name="isAggregate">is aggregate</param>
        public void BuildFetchXmlExpression(string entityName, bool isDistinct, bool isAggregate)
        {
            this.CreateFetchXml();

            this.CreateEntityInFetchXml(entityName);

            XmlAttribute distinctAttribute = this.fetchXml.CreateAttribute(FetchXmlKeyword.Distinct);
            distinctAttribute.Value = Convert.ToBoolean(isDistinct, CultureInfo.InvariantCulture) ? "true" : "false"; 
            this.fetchXmlNode.Attributes.Append(distinctAttribute);

            XmlAttribute aggregateAttribute = this.fetchXml.CreateAttribute(FetchXmlKeyword.Aggregate);
            aggregateAttribute.Value = Convert.ToBoolean(isAggregate, CultureInfo.InvariantCulture) ? "true" : "false"; 
            this.fetchXmlNode.Attributes.Append(aggregateAttribute);
        }

       /// <summary>
       /// Add fetch xml attribute to entity
       /// </summary>
       /// <param name="attributeNames">attribute names</param>
        public void AddFetchXmlAttributeToEntity(string[] attributeNames)
        {
            if (attributeNames != null)
            {
                foreach (string attributeName in attributeNames)
                {
                    this.AddColumnsToEntity(attributeName);
                }
            }
        }

        /// <summary>
        /// Add FetchXml Arithmetic Attribute To Entity
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="arithmeticOperation">arithmetic operation</param>
        public void AddFetchXmlArithmeticAttributeToEntity(string attributeName, string arithmeticOperation)
        {
            this.AddColumnsToEntity(attributeName);
            this.AddFetchXmlAttributeAlias(arithmeticOperation);
            this.AddFetchXmlAggregate(arithmeticOperation);
        }

        /// <summary>
        /// Add fetch xml condition
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="fetchXmlCondition">fetchXml Condition</param>
        /// <param name="attributeValue">attribute value</param>
        public void AddFetchXmlCondition(string attributeName, Condition fetchXmlCondition, string attributeValue)
        {
            if (!this.isEntityFilterAdded)
            {
                this.entityFilter = this.fetchXml.CreateElement(FetchXmlKeyword.Filter);
                XmlAttribute filterType = this.fetchXml.CreateAttribute(FetchXmlKeyword.Type);
                this.entityFilter.Attributes.Append(filterType);
                filterType.Value = FetchXmlKeyword.And;

                this.entityRootElement.AppendChild(this.entityFilter);

                this.isEntityFilterAdded = true;
            }

            XmlNode condition = this.fetchXml.CreateElement(FetchXmlKeyword.Condition);
            XmlAttribute conditionAttributeName = this.fetchXml.CreateAttribute(FetchXmlKeyword.Attribute);
            condition.Attributes.Append(conditionAttributeName);
            conditionAttributeName.Value = attributeName;

            XmlAttribute conditionOperator = this.fetchXml.CreateAttribute(FetchXmlKeyword.FetchxmlOperator);
            conditionOperator.Value = fetchXmlCondition.ToString().ToLowerInvariant();
            condition.Attributes.Append(conditionOperator);

            XmlAttribute conditionValue = this.fetchXml.CreateAttribute(FetchXmlKeyword.Value);
            conditionValue.Value = attributeValue;
            condition.Attributes.Append(conditionValue);

            this.entityFilter.AppendChild(condition);
        }

        /// <summary>
        /// Add fetch xml link entity first level
        /// </summary>
        /// <param name="entityName">entity name</param>
        /// <param name="fromAttributeName">from attribute name</param>
        /// <param name="toAttributeName">to attribute name</param>
        /// <param name="linkType">link type</param>
        /// <param name="aliasName">alias name</param>
        public void AddFetchXmlLinkEntityFirstLevel(string entityName, string fromAttributeName, string toAttributeName, string linkType, string aliasName)
        {
            this.linkEntityRootFirstLevel = this.fetchXml.CreateElement(FetchXmlLinkEntity.LinkEntity);

            XmlAttribute linkEntityRootFirstLevelFirstAttribute = this.fetchXml.CreateAttribute(FetchXmlLinkEntity.LinkEntityName);
            this.linkEntityRootFirstLevel.Attributes.Append(linkEntityRootFirstLevelFirstAttribute);
            linkEntityRootFirstLevelFirstAttribute.Value = entityName;

            XmlAttribute linkEntityRootFirstLevelSecondAttribute = this.fetchXml.CreateAttribute(FetchXmlLinkEntity.LinkEntityFromAttribute);
            this.linkEntityRootFirstLevel.Attributes.Append(linkEntityRootFirstLevelSecondAttribute);
            linkEntityRootFirstLevelSecondAttribute.Value = fromAttributeName;

            XmlAttribute linkEntityRootFirstLevelThirdAttribute = this.fetchXml.CreateAttribute(FetchXmlLinkEntity.LinkEntityToAttribute);
            this.linkEntityRootFirstLevel.Attributes.Append(linkEntityRootFirstLevelThirdAttribute);
            linkEntityRootFirstLevelThirdAttribute.Value = toAttributeName;

            if (linkType != null)
            {
                XmlAttribute linkEntityRootFirstLevelFifthAttribute = this.fetchXml.CreateAttribute(FetchXmlLinkEntity.LinkEntityLinkType);
                this.linkEntityRootFirstLevel.Attributes.Append(linkEntityRootFirstLevelFifthAttribute);
                linkEntityRootFirstLevelFifthAttribute.Value = linkType;
            }

            XmlAttribute linkEntityRootFirstLevelSixthAttribute = this.fetchXml.CreateAttribute(FetchXmlLinkEntity.LinkEntityAlias);
            this.linkEntityRootFirstLevel.Attributes.Append(linkEntityRootFirstLevelSixthAttribute);
            linkEntityRootFirstLevelSixthAttribute.Value = aliasName;

            this.entityRootElement.AppendChild(this.linkEntityRootFirstLevel);
        }

        /// <summary>
        /// Add fetch xml link entity second level
        /// </summary>
        /// <param name="entityName">entity name</param>
        /// <param name="fromAttributeName">from attribute name</param>
        /// <param name="toAttributeName">to attribute name</param>
        /// <param name="linkType">link type</param>
        /// <param name="aliasName">alias name</param>
        public void AddFetchXmlLinkEntitySecondLevel(string entityName, string fromAttributeName, string toAttributeName, string linkType, string aliasName)
        {
            this.linkEntityRootSecondLevel = this.fetchXml.CreateElement(FetchXmlLinkEntity.LinkEntity);

            XmlAttribute linkEntityRootFirstLevelFirstAttribute = this.fetchXml.CreateAttribute(FetchXmlLinkEntity.LinkEntityName);
            this.linkEntityRootSecondLevel.Attributes.Append(linkEntityRootFirstLevelFirstAttribute);
            linkEntityRootFirstLevelFirstAttribute.Value = entityName;

            XmlAttribute linkEntityRootFirstLevelSecondAttribute = this.fetchXml.CreateAttribute(FetchXmlLinkEntity.LinkEntityFromAttribute);
            this.linkEntityRootSecondLevel.Attributes.Append(linkEntityRootFirstLevelSecondAttribute);
            linkEntityRootFirstLevelSecondAttribute.Value = fromAttributeName;

            XmlAttribute linkEntityRootFirstLevelThirdAttribute = this.fetchXml.CreateAttribute(FetchXmlLinkEntity.LinkEntityToAttribute);
            this.linkEntityRootSecondLevel.Attributes.Append(linkEntityRootFirstLevelThirdAttribute);
            linkEntityRootFirstLevelThirdAttribute.Value = toAttributeName;

            if (linkType != null)
            {
                XmlAttribute linkEntityRootFirstLevelFifthAttribute = this.fetchXml.CreateAttribute(FetchXmlLinkEntity.LinkEntityLinkType);
                this.linkEntityRootSecondLevel.Attributes.Append(linkEntityRootFirstLevelFifthAttribute);
                linkEntityRootFirstLevelFifthAttribute.Value = linkType;
            }

            XmlAttribute linkEntityRootFirstLevelSixthAttribute = this.fetchXml.CreateAttribute(FetchXmlLinkEntity.LinkEntityAlias);
            this.linkEntityRootSecondLevel.Attributes.Append(linkEntityRootFirstLevelSixthAttribute);
            linkEntityRootFirstLevelSixthAttribute.Value = aliasName;

            this.linkEntityRootFirstLevel.AppendChild(this.linkEntityRootSecondLevel);
        }

        /// <summary>
        /// Add fetch xml filter condition first level
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="attributeValue">attribute value</param>
        public void AddFetchXmlFilterConditionFirstLevel(string attributeName, string attributeValue)
        {
            this.CreateFilterConditionFirstLinkEntity();

            this.conditionLinkEnity1 = this.fetchXml.CreateElement(FetchXmlKeyword.Condition);

            XmlAttribute conditionAttributeNameLinkEnity1 = this.fetchXml.CreateAttribute(FetchXmlKeyword.Attribute);
            this.conditionLinkEnity1.Attributes.Append(conditionAttributeNameLinkEnity1);
            conditionAttributeNameLinkEnity1.Value = attributeName;

            XmlAttribute conditionOperatorLinkEnity1 = this.fetchXml.CreateAttribute(FetchXmlKeyword.FetchxmlOperator);
            conditionOperatorLinkEnity1.Value = "eq";
            this.conditionLinkEnity1.Attributes.Append(conditionOperatorLinkEnity1);

            XmlAttribute conditionValueLinkEnity1 = this.fetchXml.CreateAttribute(FetchXmlKeyword.Value);
            conditionValueLinkEnity1.Value = attributeValue;
            this.conditionLinkEnity1.Attributes.Append(conditionValueLinkEnity1);

            this.filterRootLinkEnity1.AppendChild(this.conditionLinkEnity1);
        }

        /// <summary>
        /// Add fetch xml filter condition second level
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        /// <param name="attributeValue">attribute value</param>
        public void AddFetchXmlFilterConditionSecondLevel(string attributeName, string attributeValue)
        {
            this.CreateFilterConditionSecondLinkEntity();

            this.conditionLinkEnity2 = this.fetchXml.CreateElement(FetchXmlKeyword.Condition);

            XmlAttribute conditionAttributeNameLinkEnity2 = this.fetchXml.CreateAttribute(FetchXmlKeyword.Attribute);
            this.conditionLinkEnity2.Attributes.Append(conditionAttributeNameLinkEnity2);
            conditionAttributeNameLinkEnity2.Value = attributeName;

            XmlAttribute conditionOperatorLinkEnity2 = this.fetchXml.CreateAttribute(FetchXmlKeyword.FetchxmlOperator);
            conditionOperatorLinkEnity2.Value = "eq";
            this.conditionLinkEnity2.Attributes.Append(conditionOperatorLinkEnity2);

            XmlAttribute conditionValueLinkEnity2 = this.fetchXml.CreateAttribute(FetchXmlKeyword.Value);
            conditionValueLinkEnity2.Value = attributeValue; 
            this.conditionLinkEnity2.Attributes.Append(conditionValueLinkEnity2);

            this.filterRootLinkEnity2.AppendChild(this.conditionLinkEnity2);
        }

        /// <summary>
        /// Retrieve fetch xml query
        /// </summary>
        /// <returns>returns the xml</returns>
        public string RetrieveFetchXmlQuery()
        {
            return this.fetchXml.InnerXml;
        }

        /// <summary>
        /// Convert to fetch xml and get aggregate count
        /// </summary>
        /// <param name="service">the service</param>
        /// <param name="query">the query</param>
        /// <param name="aggregateValue">aggregate Value</param>
        /// <returns>returns the count of records</returns>
        public int ExecuteFetchXmlAndGetAggregateCount(IOrganizationService service, string query, string aggregateValue)
        {
            if (service != null)
            {
                EntityCollection recordCollection = service.RetrieveMultiple(new FetchExpression(query));

                if (recordCollection.Entities.Count > 0)
                {
                    this.recordCount = Convert.ToInt32(recordCollection[0].GetAttributeValue<AliasedValue>(aggregateValue).Value, CultureInfo.InvariantCulture);
                }
            }

            return this.recordCount;
        }

        /// <summary>
        /// Convert to fetch xml and get aggregate count
        /// </summary>
        /// <param name="service">the service</param>
        /// <param name="query">the query</param>
        /// <param name="aggregateValue">aggregate Value</param>
        /// <returns>returns the count of records</returns>
        public EntityCollection ExecuteFetchXmlAndGetAggregateSum(IOrganizationService service, string query, string aggregateValue)
        {
            if (service != null)
            {
                return service.RetrieveMultiple(new FetchExpression(query));
            }

            return null;
        }

        /// <summary>
        /// Create fetch xml
        /// </summary>
        private void CreateFetchXml()
        {
            this.fetchXml = new XmlDocument();
            this.fetchXmlNode = this.fetchXml.CreateElement(FetchXmlKeyword.Fetch);
            XmlAttribute versionAttribute = this.fetchXml.CreateAttribute(FetchXmlKeyword.Version);
            this.fetchXmlNode.Attributes.Append(versionAttribute);
            versionAttribute.Value = "1.0";

            XmlAttribute outputPlatformAttribute = this.fetchXml.CreateAttribute(FetchXmlKeyword.OutputFormat);
            outputPlatformAttribute.Value = "xml-platform";
            this.fetchXmlNode.Attributes.Append(outputPlatformAttribute);
        }

        /// <summary>
        /// Create entity in fetch xml
        /// </summary>
        /// <param name="entityName">entity name</param>
        private void CreateEntityInFetchXml(string entityName)
        {
            this.entityRootElement = this.fetchXml.CreateElement(FetchXmlKeyword.Entity);

            XmlAttribute entityAttribute = this.fetchXml.CreateAttribute(FetchXmlKeyword.FetchxmlAttributeName);
            entityAttribute.Value = entityName;
            this.entityRootElement.Attributes.Append(entityAttribute);

            this.fetchXml.AppendChild(this.fetchXmlNode);
            this.fetchXmlNode.AppendChild(this.entityRootElement);
        }

        /// <summary>
        /// Create filter condition first link entity
        /// </summary>
        private void CreateFilterConditionFirstLinkEntity()
        {
            this.filterRootLinkEnity1 = this.fetchXml.CreateElement(FetchXmlKeyword.Filter);
            XmlAttribute filterTypeLinkEntity1 = this.fetchXml.CreateAttribute(FetchXmlKeyword.Type);
            this.filterRootLinkEnity1.Attributes.Append(filterTypeLinkEntity1);
            filterTypeLinkEntity1.Value = FetchXmlKeyword.And;

            this.linkEntityRootFirstLevel.AppendChild(this.filterRootLinkEnity1);
        }

        /// <summary>
        /// Create filter condition second link entity
        /// </summary>
        private void CreateFilterConditionSecondLinkEntity()
        {
            this.filterRootLinkEnity2 = this.fetchXml.CreateElement(FetchXmlKeyword.Filter);
            XmlAttribute filterTypeLinkEntity2 = this.fetchXml.CreateAttribute(FetchXmlKeyword.Type);
            this.filterRootLinkEnity2.Attributes.Append(filterTypeLinkEntity2);
            filterTypeLinkEntity2.Value = FetchXmlKeyword.And;

            this.linkEntityRootSecondLevel.AppendChild(this.filterRootLinkEnity2);
        }

        /// <summary>
        /// Add columns to entity
        /// </summary>
        /// <param name="attributeName">attribute name</param>
        private void AddColumnsToEntity(string attributeName)
        {
            this.fetchXmlNode = this.fetchXml.CreateElement(FetchXmlKeyword.Attribute);

            this.fetchXmlAttribute = this.fetchXml.CreateAttribute(FetchXmlKeyword.AttributeColumnName);
            this.fetchXmlAttribute.Value = attributeName;
            this.fetchXmlNode.Attributes.Append(this.fetchXmlAttribute);

            this.entityRootElement.AppendChild(this.fetchXmlNode);
        }

      /// <summary>
      /// Add fetch xml attribute alias
      /// </summary>
      /// <param name="aliasName">alias name</param>
        private void AddFetchXmlAttributeAlias(string aliasName)
        {
            XmlAttribute alias = this.fetchXml.CreateAttribute(FetchXmlKeyword.Alias);
            alias.Value = aliasName;
            this.fetchXmlNode.Attributes.Append(alias);
        }

        /// <summary>
        /// Add fetch xml aggregate
        /// </summary>
        /// <param name="arithmeticOperation">arithmetic operation</param>
        private void AddFetchXmlAggregate(string arithmeticOperation)
        {
            XmlAttribute aggregate = this.fetchXml.CreateAttribute(FetchXmlKeyword.Aggregate);
            aggregate.Value = arithmeticOperation;
            this.fetchXmlNode.Attributes.Append(aggregate);

            this.entityRootElement.AppendChild(this.fetchXmlNode);
        }

        /// <summary>
        /// Fetch xml link entity
        /// </summary>
        private static class FetchXmlLinkEntity
        {
            /// <summary>
            /// name element
            /// </summary>
            public const string LinkEntityName = "name";

            /// <summary>
            /// from element
            /// </summary>
            public const string LinkEntityFromAttribute = "from";

            /// <summary>
            /// to element
            /// </summary>
            public const string LinkEntityToAttribute = "to";

            /// <summary>
            /// link type
            /// </summary>
            public const string LinkEntityLinkType = "link-type";

            /// <summary>
            /// alias key word
            /// </summary>
            public const string LinkEntityAlias = "alias";

            /// <summary>
            /// link entity
            /// </summary>
            public const string LinkEntity = "link-entity";
        }

        /// <summary>
        /// Fetch xml key word
        /// </summary>
        private static class FetchXmlKeyword
        {
            /// <summary>
            /// fetch element
            /// </summary>
            public const string Fetch = "fetch";

            /// <summary>
            /// version element
            /// </summary>
            public const string Version = "version";

            /// <summary>
            /// output format
            /// </summary>
            public const string OutputFormat = "output-format";

            /// <summary>
            /// aggregate key
            /// </summary>
            public const string Aggregate = "aggregate";

            /// <summary>
            /// distinct key
            /// </summary>
            public const string Distinct = "distinct";

            /// <summary>
            /// condition element
            /// </summary>
            public const string Condition = "condition";

            /// <summary>
            /// attribute key
            /// </summary>
            public const string Attribute = "attribute";

            /// <summary>
            /// operator element
            /// </summary>
            public const string FetchxmlOperator = "operator";

            /// <summary>
            /// value element
            /// </summary>
            public const string Value = "value";

            /// <summary>
            /// filter key
            /// </summary>
            public const string Filter = "filter";

            /// <summary>
            /// type key
            /// </summary>
            public const string Type = "type";

            /// <summary>
            /// and key
            /// </summary>
            public const string And = "and";

            /// <summary>
            /// name key
            /// </summary>
            public const string FetchxmlAttributeName = "name";

            /// <summary>
            /// entity key
            /// </summary>
            public const string Entity = "entity";

            /// <summary>
            /// name key
            /// </summary>
            public const string AttributeColumnName = "name";

            /// <summary>
            /// alias key
            /// </summary>
            public const string Alias = "alias";
        }
    }
}
