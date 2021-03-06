﻿using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MIS.CRM.AuditHistory.BusinessProcesses;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk.Query;

namespace MIS.CRM.AuditHistory.AuditHistory
{
    public class AuditHistoryExtension : AuditHistoryExtensionBase
    {
        private const string NO_VALUE = "(no value)";

        public class ModifiedAttributes
        {
            public string Attribute;
            public string AttributeName;
            public string OldValue;
            public string NewValue;
            public DateTime ChangedDate;
        }

        [DataContract]
        public class AuditRecords
        {
            [DataMember]
            public DateTime ChangedDate;

            [DataMember]
            public string ChangedBy;

            [DataMember]
            public String Action;

            [DataMember]
            public Collection<ModifiedAttributes> ModifiedAttributes;
        }

        private List<AuditRecords> auditRecordsCollection = new List<AuditRecords>();

        private StringBuilder attributesToBeRetrieved = new StringBuilder();

        private string CreatedBy = string.Empty;

        private DateTime CreatedOn = DateTime.MinValue;

        //private Collection<ModifiedAttributes> modifiedAttributes = new Collection<ModifiedAttributes>();

        public void SetEntityDetails(string entityName, string entityId)
        {
            this.EntityName = entityName;

            Guid id;

            if (!Guid.TryParse(entityId, out id))
            {
                this.ExceptionHandler.InvalidPluginExecutionException("Entity Id {0} should be of type GUID", entityId);
            }

            this.EntityId = id;
        }

        private void UpdateMissingAttributeValues(List<AuditRecords> sortedAuditRecordCollection)
        {
            Int16 indexCount = 0;
            attributesToBeRetrieved = new StringBuilder("createdby,createdon,");

            foreach (var currentAuditRecord in sortedAuditRecordCollection)
            {
                foreach (ModifiedAttributes modifiedAttribute in currentAuditRecord.ModifiedAttributes)
                {
                    if (modifiedAttribute.OldValue == NO_VALUE)
                    {
                        modifiedAttribute.OldValue = getMissingAttributeValue(modifiedAttribute.Attribute, false, indexCount, sortedAuditRecordCollection.ToList<AuditRecords>());
                    }
                    if (modifiedAttribute.NewValue == NO_VALUE)
                    {
                        modifiedAttribute.NewValue = getMissingAttributeValue(modifiedAttribute.Attribute, true, indexCount, sortedAuditRecordCollection.ToList<AuditRecords>());

                        if (modifiedAttribute.NewValue == NO_VALUE)
                        {
                            attributesToBeRetrieved.Append(modifiedAttribute.Attribute + ",");
                        }
                    }
                }
                indexCount++;
            }

            if (attributesToBeRetrieved.Length > 0)
            {
                attributesToBeRetrieved.Remove(attributesToBeRetrieved.Length - 1, 1);
            }

            this.auditRecordsCollection = sortedAuditRecordCollection;
        }

        private void RetrieveMissingAttributeValues()
        {
            if (this.attributesToBeRetrieved.Length > 0)
            {
                ColumnSet attributes = new ColumnSet(this.attributesToBeRetrieved.ToString().Split(','));

                Entity missingAttributes = this.OrgService.Retrieve(this.EntityName, this.EntityId, attributes);
                if (missingAttributes != null)
                {
                    foreach (var currentAuditRecord in this.auditRecordsCollection)
                    {
                        foreach (ModifiedAttributes modifiedAttribute in currentAuditRecord.ModifiedAttributes)
                        {
                            if (modifiedAttribute.NewValue == NO_VALUE && attributes.Columns.Contains(modifiedAttribute.Attribute))
                            {
                                modifiedAttribute.NewValue = this.GetAttributeValueBasedOnType(missingAttributes, modifiedAttribute.Attribute);
                            }
                        }
                    }

                    this.CreatedBy = this.GetAttributeValueBasedOnType(missingAttributes, "createdby");

                    this.CreatedOn = missingAttributes.Contains("createdon") && missingAttributes.Attributes["createdon"] != null ? (DateTime)missingAttributes.Attributes["createdon"] : DateTime.MinValue;
                }
            }
        }

        private string GetAttributeValueBasedOnType(Entity entityRecord, string keyName)
        {
            string value = "";

            if (entityRecord.Contains(keyName))
            {
                switch (entityRecord[keyName].GetType().FullName)
                {
                    case "Microsoft.Xrm.Sdk.EntityReference":
                        value = ((EntityReference)(entityRecord[keyName])).Name;
                        break;
                    case "Microsoft.Xrm.Sdk.OptionSetValue":
                    case "Microsoft.Xrm.Sdk.Money":
                        value = entityRecord.FormattedValues[keyName] == null ? string.Empty : entityRecord.FormattedValues[keyName].ToString();
                        break;
                    case "System.DateTime":
                        value = ((DateTime)entityRecord[keyName]).ToString("dd-MMM-yyyy hh:mm tt");
                        break;
                    default:
                        value = entityRecord[keyName].ToString();
                        break;
                }
            }

            return value;
        }

        public void RetrieveAuditDetails()
        {
            if (this.AuditCollection != null && this.AuditCollection.Count > 0)
            {
                IList<AuditRecords> auditRecCollection = new List<AuditRecords>();

                foreach (AuditDetail detail in this.AuditCollection.AuditDetails)
                {
                    // Display some of the detail information in each audit record. 
                    this.DisplayAuditDetails(detail, auditRecCollection);
                }

                if (auditRecCollection.Count == 0 || !(from record in auditRecCollection where record.Action == "Create" select record).Any<AuditRecords>())
                {
                    if (this.EntityName == "contact")
                    {
                        auditRecCollection.Add(new AuditRecords()
                        {
                            ChangedBy = "SYSTEM",
                            Action = "Create",
                            ModifiedAttributes = new Collection<ModifiedAttributes>(){
                                new ModifiedAttributes(){Attribute = "mis_participantcategory", AttributeName = "mis_participantcategory", NewValue = NO_VALUE, OldValue = string.Empty},
                                new ModifiedAttributes(){Attribute = "birthdate", AttributeName = "birthdate", NewValue = NO_VALUE, OldValue = string.Empty},
                                new ModifiedAttributes(){Attribute = "createdon", AttributeName = "createdon", NewValue = NO_VALUE, OldValue = string.Empty},
                                new ModifiedAttributes(){Attribute = "createdby", AttributeName = "createdby", NewValue = NO_VALUE, OldValue = string.Empty}
                            }
                        });
                    }
                }

                if (auditRecCollection.Count > 0)
                {
                    var sortedAuditRecordCollection = (from record in auditRecCollection
                                                       orderby record.ChangedDate ascending
                                                       select record).ToList<AuditRecords>();

                    this.UpdateMissingAttributeValues(sortedAuditRecordCollection);
                }

                this.RetrieveMissingAttributeValues();

                this.RectifyCreateEvent();
            }
        }

        private void RectifyCreateEvent()
        {
            var createEvent = (from record in this.auditRecordsCollection where record.Action == "Create" select record).FirstOrDefault<AuditRecords>();

            if (createEvent != null)
            {

                createEvent.ChangedDate = this.CreatedOn;
                createEvent.ChangedBy = this.CreatedBy;

                if (!(from record in createEvent.ModifiedAttributes where record.Attribute == "createdby" select record).Any<ModifiedAttributes>())
                {
                    createEvent.ModifiedAttributes.Add(new ModifiedAttributes()
                    {
                        Attribute = "createdby",
                        AttributeName = "createdby",
                        NewValue = createEvent.ChangedBy,
                        OldValue = string.Empty
                    });
                }

            }
        }

        public string ConvertToJsonString()
        {
            if (this.auditRecordsCollection != null && this.auditRecordsCollection.Count > 0)
            {
                return CRMHelper.JsonSerializer<List<AuditRecords>>(this.auditRecordsCollection);
            }

            return string.Empty;
        }

        private string getMissingAttributeValue(string keyName, bool upper, Int16 index, List<AuditRecords> recordCollection)
        {
            Int32 currentCount = upper ? index + 1 : index - 1;

            while (true)
            {
                if (currentCount > -1 && currentCount < recordCollection.Count)
                {
                    for (int j = 0; j < recordCollection[currentCount].ModifiedAttributes.Count; j++)
                    {
                        if (recordCollection[currentCount].ModifiedAttributes[j].Attribute == keyName)
                        {
                            if (upper)
                            {
                                if (recordCollection[currentCount].ModifiedAttributes[j].OldValue != NO_VALUE)
                                {
                                    return recordCollection[currentCount].ModifiedAttributes[j].OldValue;
                                }
                            }
                            else
                            {
                                if (recordCollection[currentCount].ModifiedAttributes[j].NewValue != NO_VALUE)
                                {
                                    return recordCollection[currentCount].ModifiedAttributes[j].NewValue;
                                }
                            }
                        }
                    }

                    currentCount = upper ? currentCount + 1 : currentCount - 1;
                }
                else
                {
                    break;
                }
            }
            return NO_VALUE;
        }

        private string getAttributeValue(Entity entityRecord, string keyName)
        {
            return this.getAttributeValue(entityRecord, keyName, null);
        }

        private string getAttributeValue(Entity entityRecord, string keyName, DataCollection<string> invalidNewValue)
        {
            string value = string.Empty;

            if (invalidNewValue != null && invalidNewValue.Contains(keyName))
            {
                return NO_VALUE;
            }

            if (entityRecord != null && entityRecord.Contains(keyName) && entityRecord[keyName] != null)
            {
                value = this.GetAttributeValueBasedOnType(entityRecord, keyName);
            }

            return value == null ? string.Empty : value;
        }

        private void DisplayAuditDetails(AuditDetail detail, IList<AuditRecords> auditRecCollection)
        {
            // Write out some of the change history information in the audit record. 
            Audit record = (Audit)detail.AuditRecord;

            // Show additional details for certain AuditDetail sub-types.
            var detailType = detail.GetType();

            if (detailType == typeof(AttributeAuditDetail) && record.Action.Value != 3)
            {
                Collection<ModifiedAttributes> modifiedAttributes = new Collection<ModifiedAttributes>();

                var attributeDetail = (AttributeAuditDetail)detail;

                if (attributeDetail.NewValue != null)
                {
                    // Display the old and new attribute values.
                    foreach (KeyValuePair<String, object> attribute in attributeDetail.NewValue.Attributes)
                    {
                        //TODO Display the lookup values of those attributes that do not contain strings.
                        modifiedAttributes.Add(new ModifiedAttributes()
                        {
                            ChangedDate = record.CreatedOn.Value.ToLocalTime(),
                            Attribute = attribute.Key,
                            AttributeName = attribute.Key,
                            OldValue = this.getAttributeValue(attributeDetail.OldValue, attribute.Key),
                            NewValue = this.getAttributeValue(attributeDetail.NewValue, attribute.Key)
                        });
                    }
                }

                if (attributeDetail.OldValue != null)
                {
                    foreach (KeyValuePair<String, object> attribute in attributeDetail.OldValue.Attributes)
                    {
                        if (!(attributeDetail.NewValue != null && attributeDetail.NewValue.Contains(attribute.Key)))
                        {
                            modifiedAttributes.Add(new ModifiedAttributes()
                            {
                                ChangedDate = record.CreatedOn.Value.ToLocalTime(),
                                Attribute = attribute.Key,
                                AttributeName = attribute.Key,
                                OldValue = this.getAttributeValue(attributeDetail.OldValue, attribute.Key),
                                NewValue = this.getAttributeValue(attributeDetail.NewValue, attribute.Key, attributeDetail.InvalidNewValueAttributes)
                            });
                        }
                    }
                }

                if (attributeDetail.InvalidNewValueAttributes != null)
                {
                    foreach (string invalidAttributeName in attributeDetail.InvalidNewValueAttributes)
                    {
                        if (!(attributeDetail.NewValue != null && attributeDetail.NewValue.Contains(invalidAttributeName)) &&
                            !(attributeDetail.OldValue != null && attributeDetail.OldValue.Contains(invalidAttributeName)))
                        {
                            modifiedAttributes.Add(new ModifiedAttributes()
                            {
                                ChangedDate = record.CreatedOn.Value.ToLocalTime(),
                                Attribute = invalidAttributeName,
                                AttributeName = invalidAttributeName,
                                OldValue = this.getAttributeValue(attributeDetail.OldValue, invalidAttributeName),
                                NewValue = this.getAttributeValue(attributeDetail.NewValue, invalidAttributeName, attributeDetail.InvalidNewValueAttributes)
                            });
                        }
                    }
                }

                if (modifiedAttributes.Count > 0)
                {
                    auditRecCollection.Add(new AuditRecords()
                    {
                        ChangedDate = record.CreatedOn.Value.ToLocalTime(),
                        ChangedBy = record.UserId.Name,
                        Action = record.Action.Value == 1 ? "Create" : "Update",
                        ModifiedAttributes = modifiedAttributes
                    });
                }
            }
        }
    }
}
