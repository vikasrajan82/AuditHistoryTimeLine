using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MIS.CRM.AuditHistory.BusinessProcesses;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;

namespace MIS.CRM.AuditHistory.AuditHistory
{
    public class AuditHistoryExtensionBase: Extension
    {
        protected string EntityName
        {
            get;
            set;
        }

        protected Guid EntityId
        {
            get;
            set;
        }

        private AuditDetailCollection auditCollection;

        protected AuditDetailCollection AuditCollection
        {
            get
            {
                if (this.auditCollection == null)
                {
                    this.RetrieveRecordChangeHistory();
                }

                return this.auditCollection;
            }
        }

        private void RetrieveRecordChangeHistory()
        {
            this.LogMessage(LogType.Info, "Retrieveing for {0}", this.EntityId);

            // Retrieve the audit history for the account and display it.
            RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
            changeRequest.Target = new EntityReference(this.EntityName, this.EntityId);

            RetrieveRecordChangeHistoryResponse changeResponse =
                (RetrieveRecordChangeHistoryResponse)this.OrgService.Execute(changeRequest);

            this.auditCollection = changeResponse.AuditDetailCollection;

        }
    }
}
