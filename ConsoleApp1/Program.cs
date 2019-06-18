using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json.Linq;

namespace LeadProcess
{
    /*
    * Lead is a protential buyer. Lead Process is a class to show the customer acquisition process which
    * ldentifies potential buyers. And when the leads are considered qualified, get passed from marketing to create 
    * more entities. The purpose of it is to effective lead management is ultimately contribute to more sales* 
    */
    class LeadProcess
    {

        private const string currency = "NOK";
        private const string leadSourceName = "Test";
        public const string companyName = "Ahello";
        // account which got the sms
        // ONLY IN NO TEST
        //public const string id = "CF56E92C-C0C7-4B4A-A97D-09B050398B92";

        public const string id = "AFBA2387-B633-E711-80CB-005056A6C323";
        //public const string companyName = "5th week AS"; 

        public static void Main(string[] args)
        {
            var credentials = CredentialCache.DefaultNetworkCredentials;
            var clientCredentials = new ClientCredentials();
            clientCredentials.UserName.UserName = credentials.Domain + "\\" + credentials.UserName;
            clientCredentials.UserName.Password = credentials.Password;
            //playground
           //  var service = new OrganizationServiceProxy(new Uri("https://intmscrmtst.sectoralarm.net/SectorAlarmfrtstPLAYGROUND/XRMServices/2011/Organization.svc"), null, clientCredentials, null);
            //DEV
            //var organizationList = new List<IOrganizationService>();

            //var countryList = new List<string>() { "NO", "SE", "FI", "ES", "IE", "FR" };
            //foreach (var country in countryList)
            //{
            //    var addService = new OrganizationServiceProxy(new Uri("https://" + $"intmscrmtst.sectoralarm.net/SectorAlarm{country}tst/XRMServices/2011/Organization.svc"), null, clientCredentials, null);
            //    organizationList.Add(addService);
            //}

            //Here is service FOR PROD
            var service = new OrganizationServiceProxy(new Uri("https://intmscrm.sectoralarm.net/SectorAlarmNO/XRMServices/2011/Organization.svc"), null, clientCredentials, null);
            //var qAccount = new QueryExpression("account");
            //qAccount.Criteria.AddCondition("adsasdf", ConditionOperator.Equal, "something");
            //var getAccount = service.RetrieveMultiple(qAccount).Entities;
            //var thisAccount = getAccount.First();

            var alocalZone = TimeZoneInfo.Local;
            var throwALl = new StringBuilder();
            //throwALl.AppendLine($"Local:{alocalZone}");
            var throwThis = TimeZoneInfo.GetSystemTimeZones();
            //throwALl.AppendLine($"'My' system: {TimeZoneInfo.GetSystemTimeZones();}");

            //foreach (var org in organizationList)
            //{
            //    var alocalZone1 = TimeZoneInfo.Local;
            //    var throwALl1 = new StringBuilder();
            //    throwALl.AppendLine($"Local:{alocalZone1}");
            //}

            //throw new Exception(throwALl.ToString());

            //Guid id = CreateLeadWithName(service, companyName);
            //CreateSMSWithName(service, companyName);
            //QualifyLead(service, id);
            //Guid id = CreateLeadWithSimpleinfor(service, companyName);

            //TODO : for setting in more

            int loopnumber = 120;

            for (int i = 12001; i < loopnumber; i++)
            {
                string descrip2 = "short description " + i;
                Debug.WriteLine(descrip2);
                CreateSMS(service, descrip2);
            }


            for (int i = 1700; i < loopnumber; i++)
            {
                string descrip = "short description " + i;
                Debug.WriteLine(descrip);
                CreateACTholder(service, descrip);
            }

            //QualifyLeadWithMovingOut(service, id);
            //UpdateLead(service, id);
            //QualifyLead(service, id);
            //ProgramFetchQuery.Run(service);
            //ProgramFetchXMLPagingCookies.Run(service);
            //ProgramFetchXMLPagingCookies.RunQueryExpression(service);
            //ProgramFetchXMLPagingCookies.RunQueryExpressionXML(service);
            //ProgramFetchXMLPagingCookies.RunQueryExpressionXML(service);
            ProgramFetchXMLPagingCookies.RunQueryExpressionXML2(service);


        }

        private static Entity GetAccount(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("account")
            {
                ColumnSet = new ColumnSet(true)
            };
            query.Criteria.AddCondition("accountid", ConditionOperator.Equal, id);
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var account = resultLise.Entities.FirstOrDefault();
            return account;

        }

        private static Entity GetLeadSourceName(OrganizationServiceProxy service, String value)
        {
            var query = new QueryExpression("log_leadsource");
            query.Criteria.AddCondition("log_name", ConditionOperator.Equal, value);
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new ArgumentException(message: $"LeadSource Name Not Found by: {value}");
            return resultLise.Entities.FirstOrDefault();

        }
        private static Entity GetUser(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("systemuser");
            query.Criteria.AddCondition("domainname", ConditionOperator.Equal, @"SECTORALARM\sayahuan");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var user = resultLise.Entities.FirstOrDefault();
            return user;

        }
        //Get contract term which name is "Domestic"
        private static Entity GetContractTerms(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("log_contractterm");
            // Norsk test miljø
            query.Criteria.AddCondition("log_name", ConditionOperator.Equal, "SANO-306-10-15");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            return resultLise.Entities.FirstOrDefault();
        }


        //Get sales person which number is "GS-80064"
        private static Entity GetSalesPerson(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("log_employee");
     
            //query.Criteria.AddCondition("log_employeenumber", ConditionOperator.Equal, "GS-80064");
            query.Criteria.AddCondition("log_employeenumber", ConditionOperator.Equal, "NO-61026");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var user = resultLise.Entities.FirstOrDefault();
            return user;

        }
        //Get post code  which number is "1178"
        //Service degree is full in that area
        private static Entity GetPostCode(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("log_postcode");
            query.Criteria.AddCondition("log_name", ConditionOperator.Equal, "1178");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var post = resultLise.Entities.FirstOrDefault();
            return post;

        }
        private static Entity GetTakdOverCASE(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("incident");
            //query.Criteria.AddCondition("ticketnumber", ConditionOperator.Equal, "CAS-04306-K3D0K4");
            //norsk test
            query.Criteria.AddCondition("ticketnumber", ConditionOperator.Equal, "CAS-05418-Q8B6K5");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var post = resultLise.Entities.FirstOrDefault();
            return post;

        }

        //Get post code  which number is "1178"
        //Service degree is full in that area
        private static Entity GetIncident(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("incident");
            query.Criteria.AddCondition("ticketnumber", ConditionOperator.Equal, "CAS-05418-Q8B6K5");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var post = resultLise.Entities.FirstOrDefault();
            return post;

        }

        //Get post code  which number is "1178"
        //Service degree is full in that area
        private static Entity GetInstallation(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("log_installation");
            query.Criteria.AddCondition("log_alarmsystemid", ConditionOperator.Equal, "installationid");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var post = resultLise.Entities.FirstOrDefault();
            return post;

        }

        private static Entity GetSMSTemplate(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("log_smstemplate");
            query.Criteria.AddCondition("log_name", ConditionOperator.Equal, "Booking Alartec");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var entity = resultLise.Entities.FirstOrDefault();
            return entity;

        }

        /*
        *  Lead create Process allow end user to choose which name they want to use 
        *  The purpose of we want the name as a parameter is that it can be easily track and create.
        */
        private static Guid CreateLeadWithName(OrganizationServiceProxy service, String name)
        {

            var leadSource = GetLeadSourceName(service, leadSourceName);
            var newLead = new Entity("lead");

            var user = GetUser(service);
            newLead["companyname"] = name;
            newLead["log_leadsourceid"] = leadSource.ToEntityReference();
            //newLead["ownerid"] = user.ToEntityReference();
            newLead["log_direction"] = new OptionSetValue(182400001); //moved out
            newLead["log_contracttermsid"] = GetContractTerms(service).ToEntityReference();
            newLead["telephone2"] = "45512131";
            newLead["log_dateofbirth"] = DateTime.Now.AddYears(-19);
            //Date cannot be in the future
            newLead["log_solddate"] = DateTime.Now;
            newLead["log_movingdate"] = DateTime.Now;

            newLead["log_salespersonid"] = GetSalesPerson(service).ToEntityReference();
            //street 1
            //newLead["address2_line1"] = "address2 vitaminveien 1, oslo";
            newLead["log_address2_postalcode"] = GetPostCode(service).ToEntityReference();
            newLead["log_postalcode"] = GetPostCode(service).ToEntityReference();
            newLead["address1_line1"] = "address1 vitaminveien 1, oslo";
            newLead["log_canoverwritecreditcheck"] = true;
            newLead["log_takeovercase"] = GetTakdOverCASE(service).ToEntityReference();
            newLead["log_typeoflead"] = new OptionSetValue(182400002);
            newLead["log_movefrominstallation"] = GetInstallation(service).ToEntityReference();
            newLead["log_movetoinstallation"] = GetInstallation(service).ToEntityReference();
            return service.Create(newLead);
        }


        /*
      *  create activities holder 
        */
        private static Guid CreateACTholder(OrganizationServiceProxy service, String subject)
        {

            var phonecall = new Entity("sa_activityholder");

            var user = GetUser(service);
            phonecall["sa_subject"] = subject;
            phonecall["sa_description"] = subject;
            phonecall["sa_regarding"] = GetAccount(service).GetAttributeValue <string>("name");
            phonecall["sa_accountid"] = GetAccount(service).ToEntityReference().ToString();
            phonecall["sa_activityid"] = GetAccount(service).ToEntityReference().ToString();
            phonecall["ownerid"] = user.ToEntityReference();
            phonecall["createdby"] = user.ToEntityReference();
            phonecall["createdon"] = DateTime.Now.AddYears(-2);
            phonecall["modifiedby"] = user.ToEntityReference();
            phonecall["modifiedon"] = DateTime.Now.AddYears(-1);
            phonecall["sa_actualend"] = DateTime.UtcNow;
            phonecall["statecode"] = new OptionSetValue(2); //active
            return service.Create(phonecall);
        }

        /*

 */
        private static Guid CreateSMS(OrganizationServiceProxy service, String subject)
        {

            var phonecall = new Entity("phonecall");

            var user = GetUser(service);
            phonecall["subject"] = subject;
            phonecall["regardingobjectid"] = GetAccount(service).ToEntityReference();
            phonecall["from"] = GetAccount(service).ToEntityReference();
            phonecall["to"] = user.ToEntityReference();
            phonecall["statecode"] = new OptionSetValue(2); //active
            return service.Create(phonecall);
        }
        /*
        *  Lead create Process allow end user to choose which name they want to use 
        *  The purpose of we want the name as a parameter is that it can be easily track and create.
        */
        private static Guid CreateLeadWithSimpleinfor(OrganizationServiceProxy service, String name)
        {

            var leadSource = GetLeadSourceName(service, leadSourceName);
            var newLead = new Entity("lead");

            var user = GetUser(service);
            newLead["companyname"] = name;
            newLead["log_leadsourceid"] = leadSource.ToEntityReference();
            //newLead["ownerid"] = user.ToEntityReference();
            newLead["log_direction"] = new OptionSetValue(182400001); //moved out
            newLead["log_contracttermsid"] = GetContractTerms(service).ToEntityReference();

            newLead["log_dateofbirth"] = DateTime.Now.AddYears(-19);
            //Date cannot be in the future
            newLead["log_solddate"] = DateTime.Now;
            newLead["log_movingdate"] = DateTime.Now;

            newLead["log_salespersonid"] = GetSalesPerson(service).ToEntityReference();
            //street 1
            newLead["log_takeovercase"] = GetTakdOverCASE(service).ToEntityReference();
            newLead["log_typeoflead"] = new OptionSetValue(182400002);
            return service.Create(newLead);
        }

        /*
   *  Lead create Process allow end user to choose which name they want to use 
   *  The purpose of we want the name as a parameter is that it can be easily track and create.
   */
        private static Guid CreateSMSWithName(OrganizationServiceProxy service, String name)
        {

            var leadSource = GetLeadSourceName(service, leadSourceName);
            var newSMS = new Entity("log_sms")
            {
                LogicalName = "log_sms"
            };

            newSMS["subject"] = name;
            newSMS["log_smstemplateid"] = GetSMSTemplate(service).ToEntityReference();
            newSMS["regardingobjectid"] = GetAccount(service).ToEntityReference();
            newSMS["log_dispatchstatus"] = new OptionSetValue(182400002);
            var user = GetUser(service);
            newSMS["ownerid"] = user.ToEntityReference();
            return service.Create(newSMS);
        }



        /*
        *As all  know, we need to track, update, and track some more.So while you update your lead. be sure that you have chosen the 
        * right one to update.
        **/
        private static void UpdateLead(OrganizationServiceProxy service, Guid id)
        {
            Entity entity = new Entity("lead");
            entity["leadid"] = id;
            entity["telephone2"] = "33338888";
            service.Update(entity);
        }

        /* Before the qualifying process.-> Add all the other necessary fields  ->Add some new work order product
        * As all  know, we need to track, update, and track some more.
        * So while the leads has been trasfer to account. Dont delete it but to update the state of it to closed.
        * The key is to make sure the lead keeps moving through the sales cycle without being lost
        */
        private static void QualifyLead(OrganizationServiceProxy service, Guid id)
        {
            var createEntity = new Entity("log_workorderproduct");
            createEntity["log_leadid"] = new EntityReference("lead", id);
            service.Create(createEntity);

            var entity = new Entity("lead");
            entity["leadid"] = id;
            //entity["log_contracttermsid"] = GetContractTerms(service).ToEntityReference();

            //entity["log_dateofbirth"] = DateTime.Now.AddYears(-19);
            ////Date cannot be in the future
            //entity["log_solddate"] = DateTime.Now;
            entity["log_movingdate"] = DateTime.Now;
            service.Update(entity);

            entity["log_salespersonid"] = GetSalesPerson(service).ToEntityReference();
            //street 1
            entity["address2_line1"] = "address2 vitaminveien 1, oslo";
            entity["log_address2_postalcode"] = GetPostCode(service).ToEntityReference();
            entity["log_postalcode"] = GetPostCode(service).ToEntityReference();
            entity["address1_line1"] = "address1 vitaminveien 1, oslo";
            entity["log_canoverwritecreditcheck"] = true;
            entity["log_takeovercase"] = GetTakdOverCASE(service).ToEntityReference();
            //entity["log_typeofcoverage"] = new OptionSetValue(284390001);
            //entity["log_typeoflead"] = new OptionSetValue(182400002);
            service.Update(entity);

            //var createEntity = new Entity("log_workorderproduct");
            //createEntity["log_leadid"] = new EntityReference("lead", id);
            //service.Create(createEntity);

            entity["log_convertleadflag"] = 1;

            service.Update(entity);


        }

        private static void QualifyLeadWithMovingOut(OrganizationServiceProxy service, Guid id)
        {
            var entity = new Entity("lead");
            entity = service.Retrieve("lead", id, new ColumnSet(true));
            entity["leadid"] = id;
            entity["log_movetoinstallation"] = GetInstallation(service).ToEntityReference();
            /*entity["log_takeovercase"] = GetIncident(service).ToEntityReference();*/
            entity["log_contracttermsid"] = GetContractTerms(service).ToEntityReference();
            entity["log_dateofbirth"] = DateTime.Now.AddYears(-19);
            //Date cannot be in the future
            entity["log_solddate"] = DateTime.Now;
            // this set to null when it is 
            entity["log_movingdate"] = DateTime.Now.AddYears(-1);
            entity["log_salespersonid"] = GetSalesPerson(service).ToEntityReference();
            entity["log_address2_postalcode"] = GetPostCode(service).ToEntityReference();
            entity["log_postalcode"] = GetPostCode(service).ToEntityReference();
            entity["address1_line1"] = "address1 vitaminveien 1, oslo";

            //entity["address2_line1"] = "address2 vitaminveien 1, oslo";
            entity["log_canoverwritecreditcheck"] = true;

            var createEntity = new Entity("log_workorderproduct");
            createEntity["log_leadid"] = new EntityReference("lead", id);
            service.Create(createEntity);

            entity["log_direction"] = new OptionSetValue(182400001);
            entity["log_takeovercase"] = GetTakdOverCASE(service).ToEntityReference();
            entity["log_convertleadflag"] = 1;
            entity["log_typeoflead"] = new OptionSetValue(182400002);
            entity["log_movefrominstallation"] = GetInstallation(service).ToEntityReference();
            //moved out
            //service.Update(entity);

            entity["log_convertleadflag"] = 1;

            //entity["statecode"] = new OptionSetValue(1);
            //entity["statuscode"] = new OptionSetValue(3);

            service.Update(entity);
        }

        public static void RunQueryAndGetDictionary(OrganizationServiceProxy service)
        {
            var qWorkorder = new QueryExpression("log_workorders");// {  ColumnSet = new ColumnSet("workordertype") };
            qWorkorder.Criteria.AddCondition("log_installationid", ConditionOperator.Equal, id);

            qWorkorder.ColumnSet = new ColumnSet("log_typeofworkorder", "statuscode");
            var getTargetWithType = service.RetrieveMultiple(qWorkorder).Entities;
            var dictioinary = getTargetWithType.GroupBy(entry => entry.GetAttributeValue<OptionSetValue>("statuscode").Value.ToString(), entry => entry.GetAttributeValue<OptionSetValue>("log_typeofworkorder").Value.ToString())
                .ToDictionary(entry => entry.Key, entry => entry.ToList());

            string inputStringFromWF = @"[
                    { Statusreason : 182400002, Type: 182400010},
                    { Statusreason : 1, Type: 182400010},
                    { Statusreason : 182400002, Type: 182400001},
                    { Statusreason : 1, Type: 182400001},
                    { Statusreason : 182400002, Type: 182400004},
                    { Statusreason : 182400002, Type: 182400009},
                    { Statusreason : 182400002, Type: 182400006},
                    ]";
            var list = JsonToDictionary(inputStringFromWF);

            var i = 1;
            foreach (var inputRestult in list)
            {
                bool exist = false;
                if (dictioinary.Keys.Contains(inputRestult["Statusreason"]))
                {
                    if (dictioinary[inputRestult["Statusreason"]].Contains(inputRestult["Type"].ToString()))
                    {
                        exist = true;
                    }

                }
                Debug.WriteLine("No Existed statusreason: {0}   {1} for item nr:{2}  {3}", (inputRestult["Statusreason"]), (inputRestult["Type"]), i, exist);
                i++;
            }
            string sampleJson = "{\"results\":[" +
         "{\"Statusreason\":\"182400002\",\"Type\":\"182400001\"}," +
         "{\"Statusreason\":\"182400000\",\"Type\":\"182400001\"}," +
         "{\"Statusreason\":\"182400001\",\"Type\":\"supervisor1\"}" +
         "]}";

            //string inputStringFromWF = @"[
            //        { Statusreason : 182400002, Type: 182400010},
            //        { Statusreason : 1, Type: 182400010},
            //        { Statusreason : 182400002, Type: 182400001},
            //        { Statusreason : 1, Type: 182400001},
            //        { Statusreason : 182400002, Type: 182400004},
            //        { Statusreason : 182400002, Type: 182400009},
            //        { Statusreason : 182400002, Type: 182400006},
            //        ]";

            JObject resultsJson = JObject.Parse(sampleJson);
            Debug.WriteLine("Play with spit string");

            Func<string, string> checkStatusReason = delegate (string caseSwitch)
            {
                switch (caseSwitch)
                {

                    case "1":
                        return "Open";
                    case "2":
                        return "Approved";
                    case "182400000":
                        return "Cancelled";
                    case "182400001":
                        return "NoShow";
                    case "182400002":
                        return "Scheduled";
                    default:
                        return "Error";
                };
            };

            Func<string, string> checkWOType = delegate (string caseSwitch)
            {
                switch (caseSwitch)
                {
                    case "182400000":
                        return "Installation";
                    case "182400001":
                        return "Service";
                    case "182400002":
                        return "Flexi";
                    case "182400004":
                        return "TakeoverService";
                    case "182400009":
                        return "TakeoverUpgrade";
                    case "182400006":
                        return "Upgrade";
                    case "182400010":
                        return "RecurringMaintenance";
                    default:
                        return "Error";
                };
            };

            foreach (var result in resultsJson["results"])
            {
                // this can be a string or null
                string statusreason = (string)result["Statusreason"];

                //  this can be a string 
                string type = (string)result["Type"];

                Debug.WriteLine("Exists{0}{1}", checkStatusReason(statusreason), checkWOType(type));
            }
            Debug.WriteLine("Play with spit string end");

            Debug.WriteLine("Play with JsonToDictionary start");

            foreach (var result in JsonToDictionary(inputStringFromWF))
            {

                string statusreason = (string)result["Statusreason"];
                string type = (string)result["Type"];
                var outputParametersString = "Exists" + checkStatusReason(statusreason) + checkWOType(type);
                Debug.WriteLine(outputParametersString);

            }
            Debug.WriteLine("Play with JsonToDictionary end");


        }

        private static List<Dictionary<string, string>> JsonToDictionary(string json)
        {
            json = Regex.Replace(json, @" ", "");
            json = json.Replace("\r\n", "")
            .Replace("[", "").Replace("]", "")
            .Replace("{", "").Replace("}", "").Trim();
            var list = new List<Dictionary<string, string>>();
            var attributes = json.Split(new string[] { ":", "," }, StringSplitOptions.None);
            for (var i = 0; i < attributes.Length; i += 4)
            {
                var dictionary = new Dictionary<string, string>();
                var statusreasonKey = (attributes[i]);
                if (string.IsNullOrEmpty(statusreasonKey))
                    continue;
                var StatusreasonKeyvalue = (attributes[i + 1]);
                dictionary.Add(statusreasonKey, StatusreasonKeyvalue);
                var typeKey = (attributes[i + 2]);
                if (string.IsNullOrEmpty(typeKey))
                    continue;
                var typeValue = (attributes[i + 3]);
                dictionary.Add(typeKey, typeValue);
                list.Add(dictionary);
            }
            return list;
        }
    }
}
