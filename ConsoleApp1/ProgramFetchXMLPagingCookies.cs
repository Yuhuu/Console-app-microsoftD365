using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace LeadProcess
{
    public class ProgramFetchXMLPagingCookies
    {
        //This is contract ID
        //public static String id = "4D1C4BAA-E98B-E711-80D9-005056A67C5E"; Spain test


        //norway test with small amont of wo
       // public static String id = "437262B2-41B6-E211-9007-E41F13BE0AF6";

        public static int fetchCount = 2;

        //no test  7 record
        //public static string id = "19636F86-2DEB-E311-BE56-E41F13BE0AF4";

        //6666 items
        // public static String id = "CF56E92C-C0C7-4B4A-A97D-09B050398B92";


        // 1107 items
        //public static String id = "1C2D4751-D17E-4AF5-AF3C-F3511E633DAE";

        // 4 items /nor test
        // public static String id = "884C892D-BA4F-46EE-819E-B53271ABD3CB";


        // 11 items., no lead activities/nor test
        public static String id = "60AA56EB-5BBB-E511-91CC-E41F13BE0AF4";
        //norge test Abbas Marofi



        //no test with duplicated lead
        //public static string id = "AD22C2D1-CC34-E711-80D4-005056A6241E";

        //no test without duplicated lead, 7Items
        //public static String id = "46a1e267-734a-487c-abf8-2a75ddabede4";

        /*
          * Retrieve all accounts owned by the user with read access rights to the accounts 
          * and the name contains customer
          * then it is printed out in Debug output.
          **/
        public static void Run(OrganizationServiceProxy service)
        {


            String openFilter = "<condition attribute=\"statecode\" operator=\"neq\" value=\"1\" />";
            openFilter += "<condition attribute=\"statecode\" operator=\"neq\" value=\"2\" />";
            openFilter = "";


            int pageNumber = 1;


            // Initialize the number of records.
            int recordCount = 0;
            string pagingCookie = null;


            var activitiesXML =
                // "<fetch distinct=\"true\" paging-cookie=\"&lt;cookie page=&quot;1&quot;&gt;&lt;/cookie&gt;\" page=\"2\" count=\"6\" returntotalrecordcount=\"true\">" +
                 "<fetch distinct=\"false\" >" +
                //"<fetch distinct=\"true\" mapping=\"logical\" count=\"50\" paging-cookie=\"&lt;cookie page=&quot;2&quot;&gt;&lt;/cookie&gt;\" >" +
                "  <entity name=\"account\" >" +
                "    <filter>" +
                "      <condition attribute=\"accountid\" operator=\"eq\" value=\"" + id + "\" />" +
                "    </filter>" +
                "    <link-entity name=\"activityparty\" from=\"partyid\" to=\"accountid\" link-type=\"outer\" >" +
                "      <link-entity name=\"activitypointer\" from=\"activityid\" to=\"activityid\" link-type=\"inner\" alias=\"activity\" >" +
                "        <attribute name=\"createdon\" />" +
                "        <attribute name=\"activityid\" />" +
                "        <attribute name=\"modifiedon\" />" +
                "        <attribute name=\"statecode\" />" +
                "        <attribute name=\"subject\" />" +
                "        <attribute name=\"activityadditionalparams\" />" +
                "        <attribute name=\"activitytypecode\" />" +
                "        <attribute name=\"description\" />" +
                "        <attribute name=\"createdby\" />" +
                "        <attribute name=\"modifiedby\" />" +
                "        <attribute name=\"actualend\" />" +
                "        <attribute name=\"scheduledend\" />" +
                "        <attribute name=\"regardingobjectid\" />" +
                "        <filter  type=\"and\" >" +
                "          <condition attribute=\"activitytypecode\" operator=\"neq\" value=\"4401\" />" + //4401 - campaignresponse--
                openFilter +
                "        </filter>" +
                "        <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"createdby\" alias=\"createdby\" >" +
                "          <attribute name=\"fullname\" />" +
                "        </link-entity>" +
                "        <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"modifiedby\" alias=\"modifiedby\" >" +
                "          <attribute name=\"fullname\" />" +
                "        </link-entity>" +
                "      </link-entity>" +
                "    </link-entity>" +
                "  </entity>" +
                "</fetch>";

            var activitiesCase = "<fetch distinct=\"true\" >" +
                   "  <entity name=\"account\" >" +
                   "    <filter>" +
                   "      <condition attribute=\"accountid\" operator=\"eq\" value=\"" + id + "\" />" +
                   "    </filter>" +
                   "    <link-entity name=\"incident\" from=\"accountid\" to=\"accountid\" link-type=\"outer\" >" +
                   "      <link-entity name=\"activityparty\" from=\"partyid\" to=\"incidentid\" link-type=\"outer\" >" +
                   "        <link-entity name=\"activitypointer\" from=\"activityid\" to=\"activityid\" link-type=\"inner\" alias=\"activity\" >" +
                   "          <attribute name=\"createdon\" />" +
                   "          <attribute name=\"activityid\" />" +
                   "          <attribute name=\"modifiedon\" />" +
                   "          <attribute name=\"statecode\" />" +
                   "          <attribute name=\"subject\" />" +
                   "          <attribute name=\"activityadditionalparams\" />" +
                   "          <attribute name=\"activitytypecode\" />" +
                   "          <attribute name=\"description\" />" +
                   "          <attribute name=\"createdby\" />" +
                   "          <attribute name=\"modifiedby\" />" +
                   "          <attribute name=\"actualend\" />" +
                   "          <attribute name=\"scheduledend\" />" +
                   "          <attribute name=\"regardingobjectid\" />" +
                   "          <filter>" +
                   openFilter +
                   "          </filter>" +
                   "          <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"createdby\" alias=\"createdby\" >" +
                   "            <attribute name=\"fullname\" />" +
                   "          </link-entity>" +
                   "          <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"modifiedby\" alias=\"modifiedby\" >" +
                   "            <attribute name=\"fullname\" />" +
                   "          </link-entity>" +
                   "        </link-entity>" +
                   "      </link-entity>" +
                   "    </link-entity>" +
                   "  </entity>" +
                   "</fetch>";

            //count for task and lead
            var fetchPagingTask = "<fetch distinct=\"true\" mapping=\"logical\" aggregate=\"true\">" +
                    "  <entity name=\"account\" >" +
                    "    <filter>" +
                    "      <condition attribute=\"accountid\" operator=\"eq\" value=\"" + id + "\" />" +
                    "    </filter>" +
                    "    <link-entity name=\"task\" from=\"log_account\" to=\"accountid\" link-type=\"outer\" >" +
                    "      <link-entity name=\"activitypointer\" from=\"activityid\" to=\"activityid\" link-type=\"inner\" alias=\"activity\" >" +
                    "        <attribute name =\"activityid\" aggregate=\"countcolumn\" alias=\"OKmodifiedon\" distinct=\"true\" />" +
                    "        <filter>" +
                    openFilter +
                    "          <condition attribute=\"regardingobjectid\" operator=\"neq\" value=\"03F0C8D1-34BC-4D93-9CBC-0DE87AE32449\" />" +
                    "        </filter>" +
                    "      </link-entity>" +
                    "    </link-entity>" +
                  "    <link-entity name=\"lead\" from=\"log_newaccount\" to=\"accountid\" link-type=\"outer\" >" +
                  "      <link-entity name=\"activityparty\" from=\"partyid\" to=\"leadid\" link-type=\"outer\" >" +
                  "        <link-entity name=\"activitypointer\" from=\"activityid\" to=\"activityid\" link-type=\"inner\" alias=\"activity\" >" +
                  "           <attribute name =\"activityid\" aggregate=\"countcolumn\" alias=\"OKmodifiedon\" distinct=\"true\" />" +
                  "          <filter>" +
                       openFilter +
                  "          </filter>" +
                  "        </link-entity>" +
                  "      </link-entity>" +
                  "    </link-entity>" +
                    "  </entity>" +
                    "</fetch>";

            //count all activities which realted to account
            var fetchcountXMLAll = "<fetch distinct=\"true\" mapping=\"logical\" aggregate=\"true\">" +
                   "  <entity name=\"activitypointer\" >" +
                   "  <attribute name =\"activityid\" aggregate=\"countcolumn\" alias=\"CountActivities\" distinct=\"true\" />" +
                   "  <filter type=\"and\" >" +
                   "      <condition attribute=\"activitytypecode\" operator=\"neq\" value=\"4401\" />" + //4401 - campaignresponse--
                   openFilter +
                   "    </filter>" +
                   "<filter type=\"or\" >" +
                   "      <condition attribute = \"accountid\" entityname = \"au\" operator = \"eq\" uitype = \"account\" value = \"{" + id + "}\" />" +
                   "      <condition attribute=\"log_account\" entityname=\"bd\" operator=\"eq\" uitype=\"account\" value=\"{" + id + "}\" />" +
                   "</filter>" +
                   " <link-entity name=\"activityparty\" from=\"activityid\" to=\"activityid\" alias=\"at\" distinct=\"true\" >" +
                   "     <link-entity name=\"account\" from=\"accountid\" to=\"partyid\" alias=\"au\" link-type=\"outer\" />" +
                   "     <link-entity name=\"task\" from=\"activityid\" to=\"activityid\" alias=\"bd\" link-type=\"outer\" />" +
                   "</link-entity>" +
                   "  </entity>" +
                   "</fetch>";

            //count foor Lead
            var fetchPaging4 = "<fetch distinct=\"true\" mapping =\"logical\" aggregate=\"true\" >" +
                  "  <entity name=\"account\" >" +
                  "    <attribute name =\"accountid\" aggregate=\"countcolumn\" alias=\"OKmodifiedon\" distinct=\"false\" />" +
                  "    <filter>" +
                  "      <condition attribute=\"accountid\" operator=\"eq\" value=\"" + id + "\" />" +
                  "    </filter>" +
                  "    <link-entity name=\"lead\" from=\"log_newaccount\" to=\"accountid\" link-type=\"outer\" >" +
                  "      <link-entity name=\"activityparty\" from=\"partyid\" to=\"leadid\" link-type=\"outer\" >" +
                  "        <link-entity name=\"activitypointer\" from=\"activityid\" to=\"activityid\" link-type=\"inner\" alias=\"activity\" >" +
                  "          <filter>" +
                       openFilter +
                  "          </filter>" +
                  "        </link-entity>" +
                  "      </link-entity>" +
                  "    </link-entity>" +
                  "  </entity>" +
                  "</fetch>";

            //TODO: this is after changing to activitypointer as main entity
            var fetchPaging3 =
                   "<fetch distinct=\"true\" mapping =\"logical\" aggregate=\"true\" >" +
                   "  <entity name=\"activitypointer\" >" +
                   "        <attribute name=\"activityid\" aggregate=\"countcolumn\" alias=\"OK\" distinct=\"true\" />" +
                   "        <filter  type =\"and\" >" +
                   "          <condition attribute=\"activitytypecode\" operator=\"neq\" value=\"4401\" />" + //4401 - campaignresponse--
                               openFilter +
                   "        </filter>" +
                   "        <link-entity name =\"systemuser\" from=\"systemuserid\" to=\"createdby\" alias=\"createdby\" >" +
                   "         </link-entity>" +
                   "         <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"modifiedby\" alias=\"modifiedby\" >" +
                   "         </link-entity>" +
                   "         <link-entity name =\"activityparty\" from=\"activityid\" to=\"activityid\" alias=\"activity\" >" +
                   "               <filter>" +
                          "      <condition attribute=\"partyid\" operator=\"eq\" value=\"" + id + "\" />" +
                         "      </filter>" +
                   "          </link-entity> " +
                 "  </entity>" +
                 "</fetch>";

            //correct answer with count column
            var fetchPaging2 = "<fetch distinct=\"false\" mapping=\"logical\" aggregate=\"true\">" +
                   "  <entity name=\"account\" >" +
                   "    <filter>" +
                   "      <condition attribute=\"accountid\" operator=\"eq\" value=\"" + id + "\" />" +
                   "    </filter>" +
                   "    <link-entity name=\"activityparty\" from=\"partyid\" to=\"accountid\">" +
                   "      <link-entity name=\"activitypointer\" from=\"activityid\" to=\"activityid\" link-type=\"inner\" alias=\"activity\" >" +
                   "        <attribute name =\"activityid\" aggregate=\"countcolumn\" alias=\"OKmodifiedon\" distinct=\"true\" />" +
                   "        <attribute name =\"statecode\" aggregate=\"countcolumn\" alias=\"OKcreateon\" distinct=\"true\" />" +
                    "        <filter  type=\"and\" >" +
                   "          <condition attribute=\"activitytypecode\" operator=\"neq\" value=\"4401\" />" + //4401 - campaignresponse--
                   openFilter +
                   "        </filter>" +
                   "      </link-entity>" +
                   "    </link-entity>" +
                   "  </entity>" +
                   "</fetch>";

            //Lead activities
            var fetchPagingLead = "<fetch distinct=\"true\"> " +
                   "  <entity name=\"account\" >" +
                   //"    <filter>" +
                   //"      <condition attribute=\"accountid\" operator=\"eq\" value=\"" + id + "\" />" +
                   //"    </filter>" +
                   "    <link-entity name=\"lead\" from=\"log_newaccount\" to=\"accountid\" link-type=\"outer\" >" +
                   "      <attribute name=\"leadid\" />" +
                   "      <attribute name=\"subject\" />" +
                   "      <link-entity name=\"activityparty\" from=\"partyid\" to=\"leadid\" link-type=\"outer\" >" +
                   "        <link-entity name=\"activitypointer\" from=\"activityid\" to=\"activityid\" link-type=\"inner\" alias=\"activity\" >" +
                   "          <attribute name=\"createdon\" />" +
                   "          <attribute name=\"activityid\" />" +
                   "          <attribute name=\"modifiedon\" />" +
                   "          <attribute name=\"statecode\" />" +
                   "          <attribute name=\"subject\" />" +
                   "          <attribute name=\"activityadditionalparams\" />" +
                   "          <attribute name=\"activitytypecode\" />" +
                   "          <attribute name=\"description\" />" +
                   "          <attribute name=\"createdby\" />" +
                   "          <attribute name=\"modifiedby\" />" +
                   "          <attribute name=\"actualend\" />" +
                   "          <attribute name=\"scheduledend\" />" +
                   "          <attribute name=\"regardingobjectid\" />" +
                   "          <filter>" +
                        openFilter +
                   "          </filter>" +
                   "          <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"createdby\" alias=\"createdby\" >" +
                   "          </link-entity>" +
                   "          <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"modifiedby\" alias=\"modifiedby\" >" +
                   "          </link-entity>" +
                   "        </link-entity>" +
                   "      </link-entity>" +
                   "    </link-entity>" +
                   "  </entity>" +
                   "</fetch>";

            //fetchPagingWithDuplicates activities
            var fetchPagingWithDuplicates = "<fetch distinct=\"true\" mapping=\"logical\">" +
                    "  <entity name=\"activitypointer\" >" +
                    "  <filter type=\"and\" >" +
                    "   <condition attribute=\"activitytypecode\" operator=\"neq\" value=\"4401\" />" + //4401 - campaignresponse--
                        openFilter +
                    "  </filter>" +
                    "        <order attribute=\"createdon\" descending=\"true\" />" +
                    "        <attribute name=\"createdon\" />" +
                    "        <attribute name=\"activityid\" />" +
                    "        <attribute name=\"modifiedon\" />" +
                    "        <attribute name=\"statecode\" />" +
                    "        <attribute name=\"subject\" />" +
                    "        <attribute name=\"activityadditionalparams\" />" +
                    "        <attribute name=\"activitytypecode\" />" +
                    "        <attribute name=\"description\" />" +
                    "        <attribute name=\"createdby\" />" +
                    "        <attribute name=\"modifiedby\" />" +
                    "        <attribute name=\"actualend\" />" +
                    "        <attribute name=\"scheduledend\" />" +
                    "        <attribute name=\"regardingobjectid\" />" +
                    "<filter type=\"or\" >" +
                   "      <condition attribute = \"accountid\" entityname = \"au\" operator = \"eq\" uitype = \"account\" value = \"{" + id + "}\" />" +
                   "      <condition attribute=\"log_account\" entityname=\"bd\" operator=\"eq\" uitype=\"account\" value=\"{" + id + "}\" />" +
                    "</filter>" +
                    " <link-entity name=\"activityparty\" from=\"activityid\" to=\"activityid\" alias=\"at\" distinct=\"true\" >" +
                    "     <link-entity name=\"account\" from=\"accountid\" to=\"partyid\" alias=\"au\" link-type=\"outer\" />" +
                    "     <link-entity name=\"task\" from=\"activityid\" to=\"activityid\" alias=\"bd\" link-type=\"outer\" />" +
                    "</link-entity>" +
                    "       <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"createdby\" alias=\"createdby\" >" +
                    "          <attribute name=\"fullname\" />" +
                    "        </link-entity>" +
                    "        <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"modifiedby\" alias=\"modifiedby\" >" +
                    "          <attribute name=\"fullname\" />" +
                    "        </link-entity>" +
                    "  </entity>" +
                    "</fetch>";
            var sw = Stopwatch.StartNew();
            while (true)
            {
                string leadActivitiesXML = CreateXml(fetchcountXMLAll, pagingCookie, pageNumber, fetchCount);

                EntityCollection result = service.RetrieveMultiple(new FetchExpression(leadActivitiesXML));
                Debug.WriteLine(result);
                foreach (var c in result.Entities)
                {

                    var jsonString = JsonConvert.SerializeObject(c);
                    Debug.WriteLine(jsonString.ToString());
                    recordCount++;
                }

                if (result.MoreRecords)
                {
                    Debug.WriteLine("\n****************\nPage number {0}\n****************", pageNumber);

                    // Increment the page number to retrieve the next page.
                    pageNumber++;

                    // Set the paging cookie to the paging cookie returned from current results.                            
                    pagingCookie = result.PagingCookie;
                    Debug.WriteLine(pagingCookie);
                }
                else
                {
                    Debug.WriteLine("\n****************\nPage number {0}\n****************", pageNumber);
                    Debug.WriteLine("fetchcountXMLAll Activities Total record {0}", recordCount);
                    sw.Stop();
                    TimeSpan time = sw.Elapsed;
                    Debug.WriteLine(time.ToString());
                    // If no more records in the result nodes, exit the loop.
                    break;

                }
            }
            sw = Stopwatch.StartNew();
            while (true)
            {
                string leadActivitiesXML = CreateXml(fetchPaging2, pagingCookie, pageNumber, fetchCount);

                EntityCollection result = service.RetrieveMultiple(new FetchExpression(leadActivitiesXML));
                Debug.WriteLine(result);
                foreach (var c in result.Entities)
                {

                    var jsonString = JsonConvert.SerializeObject(c);
                    Debug.WriteLine(jsonString.ToString());
                    recordCount++;
                }

                if (result.MoreRecords)
                {
                    Debug.WriteLine("\n****************\nPage number {0}\n****************", pageNumber);

                    // Increment the page number to retrieve the next page.
                    pageNumber++;

                    // Set the paging cookie to the paging cookie returned from current results.                            
                    pagingCookie = result.PagingCookie;
                    Debug.WriteLine(pagingCookie);
                }
                else
                {
                    Debug.WriteLine("fetchPaging2 Activities Total record {0}", recordCount);

                    // If no more records in the result nodes, exit the loop.
                    TimeSpan time = sw.Elapsed;
                    Debug.WriteLine(time.ToString());
                    break;

                }
            }


            //foreach (var c in result2.Entities)
            //{

            //    var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(c);
            //    Debug.WriteLine(jsonString.ToString());
            //    Debug.WriteLine("Lead Total  : {0}...\n", result.Entities.Count());
            //}

            var taskActivitiesXML = "<fetch distinct=\"true\" paging-cookie=\"&lt;cookie page=&quot;1&quot;&gt;&lt;/cookie&gt;\" count=\"6\" returntotalrecordcount=\"true\">" +
                    // "<fetch distinct=\"true\" >" +
                    "  <entity name=\"account\" >" +
                    "    <filter>" +
                    "      <condition attribute=\"accountid\" operator=\"eq\" value=\"" + id + "\" />" +
                    "    </filter>" +
                    "    <link-entity name=\"task\" from=\"log_account\" to=\"accountid\" link-type=\"outer\" >" +
                    "      <attribute name=\"activityid\" />" +
                    "      <attribute name=\"subject\" />" +
                    "      <link-entity name=\"activitypointer\" from=\"activityid\" to=\"activityid\" link-type=\"inner\" alias=\"activity\" >" +
                    "        <attribute name=\"createdon\" />" +
                    "        <attribute name=\"activityid\" />" +
                    "        <attribute name=\"modifiedon\" />" +
                    "        <attribute name=\"statecode\" />" +
                    "        <attribute name=\"subject\" />" +
                    "        <attribute name=\"activityadditionalparams\" />" +
                    "        <attribute name=\"activitytypecode\" />" +
                    "        <attribute name=\"description\" />" +
                    "        <attribute name=\"createdby\" />" +
                    "        <attribute name=\"modifiedby\" />" +
                    "        <attribute name=\"actualend\" />" +
                    "        <attribute name=\"scheduledend\" />" +
                    "        <attribute name=\"regardingobjectid\" />" +
                    "        <filter>" +
                    openFilter +
                    "          <condition attribute=\"regardingobjectid\" operator=\"neq\" value=\"{ 03F0C8D1-34BC-4D93-9CBC-0DE87AE32449\" }/>" +
                    "        </filter>" +
                    "        <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"createdby\" alias=\"createdby\" >" +
                    "          <attribute name=\"fullname\" />" +
                    "        </link-entity>" +
                    "        <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"modifiedby\" alias=\"modifiedby\" >" +
                    "          <attribute name=\"fullname\" />" +
                    "        </link-entity>" +
                    "      </link-entity>" +
                    "    </link-entity>" +
                    "  </entity>" +
                    "</fetch>";

        //    EntityCollection result3 = service.RetrieveMultiple(new FetchExpression(fetchPaging2));

        //    foreach (var c in result3.Entities)
        //    {

        //        var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(c);
        //        Debug.WriteLine(jsonString.ToString());
        //        Debug.WriteLine("fetch count  : {0}...\n", result3.Entities.Count());
        //    }
        }

        public static string CreateXml(string xml, string cookie, int page, int count)
        {
            StringReader stringReader = new StringReader(xml);
            var reader = new XmlTextReader(stringReader);

            // Load document
            XmlDocument doc = new XmlDocument();
            doc.Load(reader);

            return CreateXml(doc, cookie, page, count);
        }

        public static string CreateXml(XmlDocument doc, string cookie, int page, int count)
        {
            XmlAttributeCollection attrs = doc.DocumentElement.Attributes;

            if (cookie != null)
            {
                XmlAttribute pagingAttr = doc.CreateAttribute("paging-cookie");
                pagingAttr.Value = cookie;
                attrs.Append(pagingAttr);
            }

            XmlAttribute pageAttr = doc.CreateAttribute("page");
            pageAttr.Value = System.Convert.ToString(page);
            attrs.Append(pageAttr);

            XmlAttribute countAttr = doc.CreateAttribute("count");
            countAttr.Value = System.Convert.ToString(count);
            attrs.Append(countAttr);

            StringBuilder sb = new StringBuilder(1024);
            StringWriter stringWriter = new StringWriter(sb);

            XmlTextWriter writer = new XmlTextWriter(stringWriter);
            doc.WriteTo(writer);
            writer.Close();

            return sb.ToString();
        }

        public static void RunQueryExpression(OrganizationServiceProxy service)
        {

            // Query using the paging cookie.
            // Define the paging attributes.
            // The number of records per page to retrieve.
            int queryCount = 3;

            // Initialize the page number.
            int pageNumber = 1;

            // Initialize the number of records.
            int recordCount = 0;


            var contracts = new QueryExpression("log_contract");
            contracts.ColumnSet = new ColumnSet(true);
            contracts.Criteria = new FilterExpression();

            contracts.Criteria.AddCondition("log_activateddate", ConditionOperator.NotNull);
            contracts.Criteria.AddCondition("statecode", ConditionOperator.Equal, new OptionSetValue(0).Value);
            contracts.AddLink("log_installation", "log_installationid", "log_installationid");
            contracts.LinkEntities[0].LinkCriteria.AddCondition("log_installationdate", ConditionOperator.NotNull);
            contracts.Orders.Add(new Microsoft.Xrm.Sdk.Query.OrderExpression("log_name", OrderType.Ascending));
            contracts.ColumnSet.AddColumns("log_activateddate", "log_contractid");

            // Assign the pageinfo properties to the query expression.
            contracts.PageInfo = new PagingInfo();
            contracts.PageInfo.Count = queryCount;
            contracts.PageInfo.PageNumber = pageNumber;

            // The current paging cookie. When retrieving the first page, 
            // pagingCookie should be null.
            contracts.PageInfo.PagingCookie = null;

            while (true)
            {
                // Retrieve the page.
                EntityCollection results = service.RetrieveMultiple(contracts);
                if (results.Entities != null)
                {
                    // Retrieve all records from the result set.
                    foreach (var acct in results.Entities)
                    {

                        var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(acct);
                        Debug.WriteLine(jsonString.ToString());
                        Debug.WriteLine("{0}", ++recordCount);
                    }
                }

                // Check for more records, if it returns true.
                if (results.MoreRecords)
                {

                    // Increment the page number to retrieve the next page.
                    contracts.PageInfo.PageNumber++;

                    // Set the paging cookie to the paging cookie returned from current results.
                    contracts.PageInfo.PagingCookie = results.PagingCookie;
                    Debug.WriteLine(contracts.PageInfo.PagingCookie);
                }
                else
                {
                    // If no more records are in the result nodes, exit the loop.
                    break;
                }
            }
            Console.WriteLine("Retrieving sample  records in pages...\n");
        }

        public static void RunQueryExpressionXML2(OrganizationServiceProxy service)
        {
            String openFilter = "<condition attribute=\"statecode\" operator=\"neq\" value=\"1\" />";
            openFilter += "<condition attribute=\"statecode\" operator=\"neq\" value=\"2\" />";

            int pageNumber = 1;


            // Initialize the number of records.
            int recordCount = 0;
            string pagingCookie = null;
            //Lead activities
            var fetchPagingLead = "<fetch distinct=\"true\"> " +
                   "  <entity name=\"account\" >" +
                   //"    <filter>" +
                   //"      <condition attribute=\"accountid\" operator=\"eq\" value=\"" + id + "\" />" +
                   //"    </filter>" +
                   "    <link-entity name=\"lead\" from=\"log_newaccount\" to=\"accountid\" link-type=\"outer\" >" +
                   "      <attribute name=\"leadid\" />" +
                   "      <attribute name=\"subject\" />" +
                   "      <link-entity name=\"activityparty\" from=\"partyid\" to=\"leadid\" link-type=\"outer\" >" +
                   "        <link-entity name=\"activitypointer\" from=\"activityid\" to=\"activityid\" link-type=\"inner\" alias=\"activity\" >" +
                   "          <attribute name=\"createdon\" />" +
                   "          <attribute name=\"activityid\" />" +
                   "          <attribute name=\"modifiedon\" />" +
                   "          <attribute name=\"statecode\" />" +
                   "          <attribute name=\"subject\" />" +
                   "          <attribute name=\"activityadditionalparams\" />" +
                   "          <attribute name=\"activitytypecode\" />" +
                   "          <attribute name=\"description\" />" +
                   "          <attribute name=\"createdby\" />" +
                   "          <attribute name=\"modifiedby\" />" +
                   "          <attribute name=\"actualend\" />" +
                   "          <attribute name=\"scheduledend\" />" +
                   "          <attribute name=\"regardingobjectid\" />" +
                   "          <filter>" +
                        openFilter +
                   "          </filter>" +
                   "          <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"createdby\" alias=\"createdby\" >" +
                   "          </link-entity>" +
                   "          <link-entity name=\"systemuser\" from=\"systemuserid\" to=\"modifiedby\" alias=\"modifiedby\" >" +
                   "          </link-entity>" +
                   "        </link-entity>" +
                   "      </link-entity>" +
                   "    </link-entity>" +
                   "  </entity>" +
                   "</fetch>";

            while (true)
            {
                string leadActivitiesXML = CreateXml(fetchPagingLead, pagingCookie, pageNumber, fetchCount);

                EntityCollection result = service.RetrieveMultiple(new FetchExpression(leadActivitiesXML));
                Debug.WriteLine(result);
                foreach (var c in result.Entities)
                {

                    var jsonString = JsonConvert.SerializeObject(c);
                    Debug.WriteLine(jsonString.ToString());
                    recordCount++;
                }

                if (result.MoreRecords)
                {
                    Debug.WriteLine("\n****************\nPage number {0}\n****************", pageNumber);

                    // Increment the page number to retrieve the next page.
                    pageNumber++;

                    // Set the paging cookie to the paging cookie returned from current results.                            
                    pagingCookie = result.PagingCookie;
                    Debug.WriteLine(pagingCookie);
                }
                else
                {
                    Debug.WriteLine("fetchPaging2 Activities Total record {0}", recordCount);

                    // If no more records in the result nodes, exit the loop.
                    break;

                }
            }
        }
            public static void RunQueryExpressionXML(OrganizationServiceProxy service)
        {

            // Query using the paging cookie.
            // Define the paging attributes.
            // The number of records per page to retrieve.
            //int queryCount = 1000;

            // Initialize the page number.
            int pageNumber = 1;

            // Initialize the number of records.
            int recordCount = 0;

            //REWRITE query
            var queryPointer = new QueryExpression("activitypointer");
            // queryAccount.ColumnSet = new ColumnSet(false);
            queryPointer.Distinct = true;
            queryPointer.ColumnSet = new ColumnSet("createdon", "activityid", "modifiedon", "statecode", "subject", "activityadditionalparams",
                "activitytypecode", "description", "createdby", "modifiedby", "actualend", "scheduledend", "regardingobjectid");
            //
            FilterExpression filter = new FilterExpression();
            filter.Conditions.Add(new ConditionExpression("activitytypecode", ConditionOperator.NotEqual, 4401));
            filter.Conditions.Add(new ConditionExpression("statecode", ConditionOperator.NotEqual, 1));
            filter.Conditions.Add(new ConditionExpression("statecode", ConditionOperator.NotEqual, 2));
            filter.FilterOperator = LogicalOperator.And;
            queryPointer.Criteria.AddFilter(filter);

            queryPointer.AddLink("systemuser", "createdby", "systemuserid");
            queryPointer.LinkEntities[0].EntityAlias = "createdby";
            queryPointer.LinkEntities[0].Columns = new ColumnSet("fullname");
            queryPointer.AddLink("systemuser", "modifiedby", "systemuserid");
            queryPointer.LinkEntities[1].EntityAlias = "modifiedby";
            queryPointer.LinkEntities[1].Columns = new ColumnSet("fullname");

            queryPointer.AddLink("activityparty", "activityid", "activityid", JoinOperator.Natural);
            queryPointer.LinkEntities[2].LinkCriteria.AddCondition("partyid", ConditionOperator.Equal, id);


            queryPointer.PageInfo = new PagingInfo();
            queryPointer.PageInfo.Count = fetchCount;
            queryPointer.PageInfo.PageNumber = pageNumber;

            // The current paging cookie. When retrieving the first page, 
            // pagingCookie should be null.
            queryPointer.PageInfo.PagingCookie = null;
            EntityCollection results = null;
            var pageResult = new List<Guid>();
            do
            {
                // Retrieve the page.
                results = service.RetrieveMultiple(queryPointer);
                if (results.Entities != null)
                {
                    // Retrieve all records from the result set.
                    foreach (var acct in results.Entities)
                    {

                        var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(acct);
                    //    Debug.WriteLine(jsonString.ToString());
                        recordCount++;
                    }
                }

                // Check for more records, if it returns true.
                if (results.MoreRecords)
                {
                    Debug.WriteLine("The above result is for page: {0}", queryPointer.PageInfo.PageNumber);
                    //if (pageResult.Count%queryCount==0)
                    queryPointer.PageInfo.PageNumber++;
             
                    // Set the paging cookie to the paging cookie returned from current results.
                    queryPointer.PageInfo.PagingCookie = results.PagingCookie;
                    Debug.WriteLine(queryPointer.PageInfo.PagingCookie);
                }
                else
                {
                    Debug.WriteLine("End with total items:" + recordCount);
                    // If no more records are in the result nodes, exit the loop.
                    break;

                }
            }
            while (results.MoreRecords);
        }
    }
}

