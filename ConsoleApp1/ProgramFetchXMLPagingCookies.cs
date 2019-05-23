using System;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;


namespace LeadProcess
{
    public class ProgramFetchXMLPagingCookies
    {


        /*
          * Retrieve all accounts owned by the user with read access rights to the accounts 
          * and the name contains customer
          * then it is printed out in Debug output.
          **/
        public static void Run(OrganizationServiceProxy service)
        {

            string openFilter = String.Empty;


            var fetch =
                //"<fetch distinct=\"true\" paging-cookie=\"&lt;cookie page=&quot;1&quot;&gt;&lt;/cookie&gt;\" count=\"6\" returntotalrecordcount=\"true\">" +
                "<fetch distinct=\"true\" >" +
                     "  <entity name=\"account\" >" +
                    "    <filter>" +
                    "      <condition attribute=\"accountid\" operator=\"eq\" value=\"437262B2-41B6-E211-9007-E41F13BE0AF6\" />" +
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


            EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetch));
  
            foreach (var c in result.Entities)
            {
        
                var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(c);
                Debug.WriteLine(jsonString.ToString());
            }
        }
        public static void RunQueryExpression(OrganizationServiceProxy service)
        {

            var contracts = new QueryExpression("log_contract");
            contracts.ColumnSet = new ColumnSet(true);
            contracts.Criteria = new FilterExpression();

            contracts.Criteria.AddCondition("log_activateddate", ConditionOperator.NotNull);
            contracts.Criteria.AddCondition("statecode", ConditionOperator.Equal, new OptionSetValue(0).Value);
            contracts.Criteria.AddCondition("log_contractexpirationdate", ConditionOperator.Today);
            contracts.AddLink("log_installation", "log_installationid", "log_installationid");
            contracts.LinkEntities[0].LinkCriteria.AddCondition("log_installationdate", ConditionOperator.NotNull);
            contracts.Orders.Add(new Microsoft.Xrm.Sdk.Query.OrderExpression("log_name", OrderType.Ascending));

            EntityCollection result = service.RetrieveMultiple(contracts);
            foreach (var c in result.Entities)
            {
                System.Console.WriteLine(c.Attributes["log_name"] + " " + c.GetAttributeValue<EntityReference>("log_account").Name);
            }
        }
    }
}
