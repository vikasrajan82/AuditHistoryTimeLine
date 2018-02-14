// <copyright file="FetchXmlAnalyzer.cs" company="Microsoft">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <summary>Analyzes the Fetch Xml</summary>
namespace MIS.CRM.AuditHistory.BusinessProcesses
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Collections.Specialized;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;
    using System.Xml;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;

    /// <summary>
    /// Fetch Xml Analyzer
    /// </summary>
    public class FetchXmlAnalyzer
    {
        /// <summary>
        /// Maximum Fetch Count
        /// </summary>
        private const int MaxFetchCount = 5000;

        /// <summary>
        /// Count Attribute
        /// </summary>
        private const string CountAttribute = "count";

        /// <summary>
        /// Paging Cookie Attribute
        /// </summary>
        private const string PagingCookieAttribute = "paging-cookie";

        /// <summary>
        /// Page Attribute
        /// </summary>
        private const string PageAttribute = "page";

        /// <summary>
        /// pattern for fetch xml
        /// </summary>
        private string pattern = "@(?:[Tt]oday+([-+]\\d|))";

        /// <summary>
        /// Xml Document
        /// </summary>
        private XmlDocument doc = new XmlDocument();

        /// <summary>
        /// Indicates whether a count attribute exists within the fetch xml
        /// </summary>
        private bool containsCount = false;

        /// <summary>
        /// Provides a count of the value specified in the fetch xml
        /// </summary>
        private int countValue = 0;

        /// <summary>
        /// Count to be retrieved
        /// </summary>
        private int retrieveCount;

        /// <summary>
        /// Paging Cookie
        /// </summary>
        private string pagingCookie;

        /// <summary>
        /// Page Number
        /// </summary>
        private string pageNumber;

        /// <summary>
        /// Fetch Xml
        /// </summary>
        private string fetchXml;

        /// <summary>
        /// Initializes a new instance of the FetchXmlAnalyzer class.
        /// </summary>
        /// <param name="fetchXmlQuery">Fetch Xml</param>
        public FetchXmlAnalyzer(string fetchXmlQuery)
        {
            fetchXmlQuery = this.ChangeFetchXmlDatePart(fetchXmlQuery);

            using (StringReader stringReader = new StringReader(fetchXmlQuery))
            {
                XmlTextReader reader = new XmlTextReader(stringReader);
                this.doc.Load(reader);
            }

            if (this.doc != null && this.doc.DocumentElement != null)
            {
                this.containsCount = this.doc.DocumentElement.HasAttribute(FetchXmlAnalyzer.CountAttribute);

                if (this.containsCount)
                {
                    this.countValue = Convert.ToInt32(this.doc.DocumentElement.GetAttribute(FetchXmlAnalyzer.CountAttribute), CultureInfo.InvariantCulture);
                }
            }
        }

        /// <summary>
        /// Gets the Fetch Xml with paging cookie
        /// </summary>
        protected string FetchXml
        {
            get
            {
                if (string.IsNullOrEmpty(this.fetchXml))
                {
                    this.RemoveAttribute(FetchXmlAnalyzer.CountAttribute);
                    this.RemoveAttribute(FetchXmlAnalyzer.PageAttribute);
                    this.RemoveAttribute(FetchXmlAnalyzer.PagingCookieAttribute);

                    StringBuilder fetchXmlBuilder = new StringBuilder();
                    fetchXmlBuilder.Append("<fetch ");
                    foreach (XmlAttribute attribute in this.doc.DocumentElement.Attributes)
                    {
                        fetchXmlBuilder.Append(attribute.Name + "='" + attribute.Value + "' ");
                    }

                    fetchXmlBuilder.Append("{0} {1} {2}>");
                    fetchXmlBuilder.Append(this.doc.DocumentElement.InnerXml);
                    fetchXmlBuilder.Append("</fetch>");

                    this.fetchXml = fetchXmlBuilder.ToString();
                }

                return string.Format(
                    CultureInfo.InvariantCulture,
                    this.fetchXml,
                    string.IsNullOrEmpty(this.pagingCookie) ? string.Empty : FetchXmlAnalyzer.PagingCookieAttribute + "='" + XmlEncode(this.pagingCookie) + "'",
                    string.IsNullOrEmpty(this.pageNumber) ? string.Empty : FetchXmlAnalyzer.PageAttribute + "='" + this.pageNumber + "'",
                    this.retrieveCount == 0 ? string.Empty : FetchXmlAnalyzer.CountAttribute + "='" + this.retrieveCount + "'");
            }
        }

        /// <summary>
        /// Execute the Fetch Xml in a recursive manner
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <returns>Collection of Entity</returns>
        public Collection<Entity> ExecuteFetchXml(IOrganizationService service)
        {
            List<Entity> entityCollection = new List<Entity>();

            EntityCollection batchCollection = new EntityCollection();

            if (service != null)
            {
                while (true)
                {
                    if (this.containsCount)
                    {
                        if (this.countValue <= 0)
                        {
                            break;
                        }

                        this.retrieveCount = this.countValue > FetchXmlAnalyzer.MaxFetchCount ? FetchXmlAnalyzer.MaxFetchCount : this.countValue;

                        this.countValue -= this.retrieveCount;
                    }

                    batchCollection = service.RetrieveMultiple(new FetchExpression(this.FetchXml));

                    if (batchCollection.HasRecords())
                    {
                        entityCollection.AddRange(batchCollection.Entities);

                        if (batchCollection.MoreRecords)
                        {
                            this.pageNumber = string.IsNullOrEmpty(this.pageNumber) ? "2" : (Convert.ToInt32(this.pageNumber, CultureInfo.InvariantCulture) + 1).ToString(CultureInfo.InvariantCulture);

                            this.pagingCookie = batchCollection.PagingCookie;

                            continue;
                        }
                    }

                    break;
                }
            }

            return new Collection<Entity>(entityCollection);
        }

        /// <summary>
        /// Change fetch xml date part
        /// </summary>
        /// <param name="fetchXmlQuery">fetch xml</param>
        /// <returns>returns the fetch xml string</returns>
        protected string ChangeFetchXmlDatePart(string fetchXmlQuery)
        {
            MatchCollection matchCollection = Regex.Matches(fetchXmlQuery, this.pattern);

            if (matchCollection.Count > 0)
            {
                NameValueCollection collection = new NameValueCollection();

                int days;

                DateTime todaysDate = DateTime.Now;

                foreach (Match matches in matchCollection)
                {
                    string keyvalue = matches.Groups[1].Value;

                    string keyName = matches.Groups[0].Value;

                    days = 0;

                    if (!string.IsNullOrEmpty(keyvalue))
                    {
                        days = int.Parse(keyvalue, CultureInfo.InvariantCulture);
                    }

                    if (collection[keyName] == null)
                    {
                        collection.Add(keyName, todaysDate.AddDays(days).Date.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture));
                    }
                }

                if (collection.Count > 0)
                {
                    fetchXmlQuery = Regex.Replace(fetchXmlQuery, this.pattern, m => collection[m.Value]);
                }
            }

            return fetchXmlQuery;
        }
        
        /// <summary>
        /// Xml Encode
        /// </summary>
        /// <param name="value">Value to be encoded</param>
        /// <returns>Encoded String</returns>
        private static string XmlEncode(string value)
        {
            return value
              .Replace("&", "&amp;")
              .Replace("<", "&lt;")
              .Replace(">", "&gt;")
              .Replace("\"", "&quot;")
              .Replace("'", "&apos;");
        }

        /// <summary>
        /// Removes the Attribute from Fetch Xml
        /// </summary>
        /// <param name="attributeName">Attribute to be removed</param>
        private void RemoveAttribute(string attributeName)
        {
            if (this.doc.DocumentElement.HasAttribute(attributeName))
            {
                this.doc.DocumentElement.RemoveAttribute(attributeName);
            }
        }
    }
}
