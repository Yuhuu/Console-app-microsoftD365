using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Text.RegularExpressions;
using ConsoleAppCsharp.ConsoleApp1;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
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
        private const string leadSourceName = "test";
     
        // account which got the sms
        // ONLY IN NO TEST
        //public const string id = "CF56E92C-C0C7-4B4A-A97D-09B050398B92";

        //public const string id = "AFBA2387-B633-E711-80CB-005056A6C323";
        //public const string companyName = uiname="A AH ALSAHOO, FAHADs" 
        //accountID i Es test
        public const string accountId = "406f4f52-e4de-e511-bd38-e41f13be0af4";

        public static void Main(string[] args)
        {
            var credentials = CredentialCache.DefaultNetworkCredentials;
            var clientCredentials = new ClientCredentials();
            clientCredentials.UserName.UserName = credentials.Domain + "\\" + credentials.UserName;
            clientCredentials.UserName.Password = credentials.Password;
            //playground
            //  var service = new OrganizationServiceProxy(new Uri("https://intmscrmtst.sectoralarm.net/SectorAlarmfrtstPLAYGROUND/XRMServices/2011/Organization.svc"), null, clientCredentials, null);
            //DEV
            var organizationList = new List<IOrganizationService>();

            var countryList = new List<string>() { "NO", "SE", "FI", "ES", "IE", "FR" };
            foreach (var country in countryList)
            {
                var addService = new OrganizationServiceProxy(new Uri("https://" + $"intmscrmtst.sectoralarm.net/SectorAlarm{country}tst/XRMServices/2011/Organization.svc"), null, clientCredentials, null);
                organizationList.Add(addService);
            }

            //Test enviroment
            //var country = "NO";
            //var service = new OrganizationServiceProxy(new Uri("https://" + $"intmscrmtst.sectoralarm.net/SectorAlarm{country}tst/XRMServices/2011/Organization.svc"), null, clientCredentials, null);

            //Here is service FOR PROD Norway
            var service = new OrganizationServiceProxy(new Uri("https://intmscrm.sectoralarm.net/SectorAlarmNO/XRMServices/2011/Organization.svc"), null, clientCredentials, null);
            //Here is service FOR PROD Sweden
          //  var service = new OrganizationServiceProxy(new Uri("https://intmscrm.sectoralarm.net/SectorAlarmSE/XRMServices/2011/Organization.svc"), null, clientCredentials, null);
            //var qAccount = new QueryExpression("account");
            //qAccount.Criteria.AddCondition("adsasdf", ConditionOperator.Equal, "something");
            //var getAccount = service.RetrieveMultiple(qAccount).Entities;
            //var thisAccount = getAccount.First();

            var alocalZone = TimeZoneInfo.Local;
            var throwALl = new StringBuilder();
           // throwALl.AppendLine($"Local:{alocalZone}");
            var throwThis = TimeZoneInfo.GetSystemTimeZones();
            //throwALl.AppendLine($"'My' system: {TimeZoneInfo.GetSystemTimeZones();}");
            string ExpectedDateTimePattern = "yyyy'-'MM'-'dd'T'HH':'mm':'ss''zzz";
            foreach (var org in organizationList)
            {
                var throwALl1 = new StringBuilder();
                //get timezone fra config for timezone
                Entity conf = new Entity("log_configuration");
                var qe = new QueryExpression("log_configuration") { ColumnSet = new ColumnSet() };
                qe.ColumnSet.Columns.Add("sa_timezone");
                var configuration = org.RetrieveMultiple(qe).Entities.First();
                string timezone = configuration.GetAttributeValue<string>("sa_timezone");
                if (timezone == null || timezone.Length.Equals(0))
                {
                    throw new InvalidPluginExecutionException("Failed to get property timezone.");
                }
                TimeZoneInfo crmZone = TimeZoneInfo.FindSystemTimeZoneById(timezone);
                var timeUtc = DateTime.UtcNow;
                DateTime realtime = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, crmZone);
                //Here we removed tolocaltime cause we dont care about the localtime now
                var countryTime = realtime.ToString(ExpectedDateTimePattern);
                var alocalZone2 = crmZone.DaylightName;
                //throwall.appendline($"timeutc:{timeutc.toshorttimestring()}");
                //throwALl.AppendLine($"Time Local:{TimeZoneInfo.ConvertTimeFromUtc(timeUtc, TimeZoneInfo.Local).ToShortTimeString()}");
   
                var index = organizationList.IndexOf(org);
                throwALl.AppendLine($"------------------------------------------------");
                //                TimeZoneInfo cstZone = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time");
                //DateTime cstTime = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, cstZone);
                throwALl.AppendLine($"The date and time are {realtime} {crmZone.IsDaylightSavingTime(realtime)}.");
                throwALl.AppendLine($"CRM country Name:{countryList.ElementAt(index)}");
                throwALl.AppendLine($"CRM country no daylightsaving:{crmZone}");
                throwALl.AppendLine($"CRM country:{alocalZone2}");
                throwALl.AppendLine($"CRM country time:{countryTime}");
            }
            Debug.WriteLine(throwALl);
            Debug.WriteLine(throwALl.ToString());
            //throw new Exception(throwALl.ToString());

            //Test this in  test
            string timeString = DateTime.Now.ToShortTimeString();
            string companyName = "created : " + timeString;
           // Guid id = CreateLeadWithName(service, companyName);

        
            //Guid id2 = CreateIncomingWithName(service, companyName);
            ///CreateSMSWithName(service, companyName);
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
            //EntityReference role1EntityRef = GetAccount(service).ToEntityReference();
            //string Role1RoleName = "Legal owner";
            //EntityReference role2EntityRef = GetInstallationWithAlarmSystemId(service).ToEntityReference();
            //string Role2RoleName = "Installation";
            //var newConnectionId = AddConnection(role1EntityRef, Role1RoleName, role2EntityRef, Role2RoleName, service);
            //UpdateLead(service, id);

            //ProgramFetchQuery.Run(service);
            //ProgramFetchXMLPagingCookies.Run(service);
            //ProgramFetchXMLPagingCookies.RunQueryExpression(service);
            //ProgramFetchXMLPagingCookies.RunQueryExpressionXML(service);
            //ProgramFetchXMLPagingCookies.RunQueryExpressionXML(service);
            ProgramFetchXMLPagingCookies.Run(service);

           // int count = GetCountInstallationWithAlarmSystemId(service);
        }

        private static Guid AddConnection(EntityReference entity1, string role1, EntityReference entity2, string role2, OrganizationServiceProxy service)
        {
            var progress = new StringBuilder();
            try
            {
                progress.AppendLine("findRole1...");
                var findRole1 = GetRoleOrRolesForEntity(entity1.LogicalName, role1, service);
                progress.AppendLine($"Found! {findRole1.Count}\r\nfindRole2...");
                var findRole2 = GetRoleOrRolesForEntity(entity2.LogicalName, role2, service);
                progress.AppendLine($"Found! {findRole2.Count}");
                if (findRole1.Count != 1 || findRole2.Count != 1)
                {
                    throw new InvalidPluginExecutionException($"Tried to find unique roles failed. \r\n - Role {role1} for entity {entity1.LogicalName} return {findRole1.Count}"
                        + $"\r\n - Role {role2} for entity {entity2.LogicalName} return {findRole2.Count}");
                }
                //logic based on specific Connections..
                progress.AppendLine($"Should create connection?");
                var shouldCreateConnection = ProcessConnectionRules(entity1, role1, entity2, role2, service);
                progress.AppendLine($"==> {shouldCreateConnection}");
                if (!shouldCreateConnection)
                {
                    progress.AppendLine($":: No returning Guid.Empty");
                    return Guid.Empty;
                }
                else
                {
                    progress.AppendLine($":: Creating connection");
                    var connect = new Entity("connection");
                    connect["Record1Id".ToLower()] = entity1;
                    connect["Record1RoleId".ToLower()] = findRole1.First().ToEntityReference();
                    connect["Record2Id".ToLower()] = entity2;
                    connect["Record2RoleId".ToLower()] = findRole2.First().ToEntityReference();
                    var result = service.Create(connect);
                    progress.AppendLine($"==> DONE");
                    return result;
                }
            }
            catch (Exception error)
            {
                throw new Exception($"#AddConnection:{error.Message}\r\n{progress.ToString()}");
            }
        }

        private static bool ProcessConnectionRules(EntityReference entity1, string role1, EntityReference entity2, string role2, IOrganizationService service)
        {
            //MyPages connections...
            if (entity1.LogicalName == "contact" && entity2.LogicalName == "log_installation" ||
                entity1.LogicalName == "log_installation" && entity2.LogicalName == "contact")
            {
                var contactRef = (entity1.LogicalName == "contact" ? entity1 : entity2);
                var installationRef = (entity2.LogicalName == "log_installation" ? entity2 : entity1);
                var contactRole = (entity1.LogicalName == "contact" ? role1 : role2);
                var installationRole = (entity2.LogicalName == "log_installation" ? role2 : role1);


                var existingConnection = GetConnection(contactRef, contactRole, installationRef, installationRole, service);
                var connectionReference = existingConnection.Entities.FirstOrDefault();
                if (connectionReference != null)
                {   //do nothing, connection already exist                                     
                    return false;
                }
                //Register new Legal Owner / Pending Legal owner connection (will potentially replace "old" legal owner)
                if (contactRole == "Legal owner" || contactRole == "Pending Legal owner")
                {   //remove legal owner from existing contract
                    var existingConnectionRoleOnInstallation = GetConnectionsOfSpecificRoleForRecord(contactRole, installationRef.Id, service);
                    if (existingConnectionRoleOnInstallation != null && existingConnectionRoleOnInstallation.Entities.Count > 0)
                    {
                        var existingLegalOwnerConnection = existingConnectionRoleOnInstallation.Entities.FirstOrDefault();
                        var updateLegalOwnerConnection = new Entity("connection");
                        updateLegalOwnerConnection.Id = existingLegalOwnerConnection.Id;
                        updateLegalOwnerConnection["record2id"] = contactRef;
                        service.Update(updateLegalOwnerConnection);
                        return false; //no creation - only update
                    }
                }
            }
            return true;
        }

        public static EntityCollection GetConnection(EntityReference entity1, string role1, EntityReference entity2, string role2, IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("connection") { ColumnSet = new ColumnSet(true) };
            query.Criteria.AddCondition(new ConditionExpression("record1id", ConditionOperator.Equal, entity1.Id));
            query.Criteria.AddCondition(new ConditionExpression("record1roleidname", ConditionOperator.Equal, role1));
            query.Criteria.AddCondition(new ConditionExpression("record2id", ConditionOperator.Equal, entity2.Id));
            query.Criteria.AddCondition(new ConditionExpression("record2roleidname", ConditionOperator.Equal, role2));
            var allOfSpecificRole = service.RetrieveMultiple(query);
            return allOfSpecificRole;
        }

        public static EntityCollection GetConnectionsOfSpecificRoleForRecord(string role, Guid entityID, IOrganizationService service)
        {
            //QueryExpression query = new QueryExpression("connection") { ColumnSet = new ColumnSet("record1id", "record1roleidname", "record2id", "record2roleidname") }; // Comment: Using ColumnSet(true) is bad practice as you use more resouces than needed. 
            QueryExpression query = new QueryExpression("connection") { ColumnSet = new ColumnSet("record1id", "record2id", "record1roleid", "record2roleid", "statecode", "effectiveend", "effectivestart", "name") };


            query.Criteria.AddCondition(new ConditionExpression("record1id", ConditionOperator.Equal, entityID)); // Are you really interested in the field ModifedOnBehalfBy for instance?
            if (!String.IsNullOrEmpty(role))
            {
                query.Criteria.AddCondition(new ConditionExpression("record2roleidname", ConditionOperator.Equal, role));
            }
            var allOfSpecificRole = service.RetrieveMultiple(query);
            return allOfSpecificRole;
        }
        /// <summary>
        /// Get all Role(s) for a given entity type.
        /// </summary>
        /// <param name="entityname">name of the entity</param>
        /// <param name="role">The role to fetch. If empty, get all roles.</param>
        /// <param name="service"></param>
        /// <returns></returns>
        static List<Entity> GetRoleOrRolesForEntity(string entityname, string role, IOrganizationService service)
        {
            QueryExpression roleQuery = new QueryExpression
            {
                EntityName = "connectionrole",
                ColumnSet = new ColumnSet("connectionroleid", "name"),
                Distinct = true,
                LinkEntities =
                {
                    new LinkEntity
                    {
                        LinkToEntityName = "connectionroleobjecttypecode" ,
                        LinkToAttributeName = "connectionroleid",
                        LinkFromEntityName =  "connectionrole",
                        LinkFromAttributeName = "connectionroleid",
                        LinkCriteria = new FilterExpression
                        {
                            FilterOperator = LogicalOperator.And,
                            Conditions =
                            {
                                new ConditionExpression
                                {
                                        AttributeName = "associatedobjecttypecode",
                                        Operator = ConditionOperator.Equal,
                                        Values = { GetObjectTypeCode(entityname, service)}
                                }
                            }
                        }
                    }
                }
            };
            if (role.Length > 0)
                roleQuery.Criteria.AddCondition("name", ConditionOperator.Equal, role);
            return service.RetrieveMultiple(roleQuery).Entities.ToList();
        }

        public static int GetSameSystemID(IOrganizationService service)
        {
            QueryExpression roleQuery = new QueryExpression
            {
                EntityName = "connectionrole",
                ColumnSet = new ColumnSet("connectionroleid", "name"),
                Distinct = true
               
            };
            return service.RetrieveMultiple(roleQuery).Entities.ToList().Count();
        }

        public static int GetObjectTypeCode(string logicalName, IOrganizationService service)
        {
            //if (logicalName == "log_livingaddress") return 10117;
            //else if (logicalName == "contact") return 2;
            //else if (logicalName == "account") return 1;
            //else if (logicalName == "log_installation") return 10033;
            //else if (logicalName == "log_contract") return 10017;
            //else return QueryTypeCode(logicalName, service);
            return QueryTypeCode(logicalName, service);
        }

        static int QueryTypeCode(string logicalName, IOrganizationService service)
        {
            var typecode = ViewUtilities.Metadata(logicalName, service)?.ObjectTypeCode;
            return typecode == null ? throw new Exception($"TypeCode for entity '{logicalName}' not found in metadata.") : (int)typecode;
        }
        /// <summary>
        private static Guid CreateIncomingWithName(OrganizationServiceProxy service, string companyName)
        {
            var newLead = new Entity("log_incomingcontract");

            var user = GetUser(service);
            newLead["log_accountname"] = companyName;
            //Date cannot be in the future
            newLead["log_solddate"] = DateTime.Now;
            newLead["log_street1"] = companyName;
            newLead["log_installationstreet1"] = companyName;
            newLead["log_mainphone"] = "45512131";
            newLead["log_installationpostalcode"] = GetPostCode(service).ToEntityReference();
            newLead["log_salesperson"] = GetSalesPerson(service).ToEntityReference();

            return service.Create(newLead);
        }

        private static Entity GetAccount(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("account")
            {
                ColumnSet = new ColumnSet(true)
            };
            query.Criteria.AddCondition("accountid", ConditionOperator.Equal, accountId);
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var account = resultLise.Entities.FirstOrDefault();
            return account;

        }


        private static Entity GetAccountContact(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("account")
            {
                ColumnSet = new ColumnSet(true)
            };
            query.Criteria.AddCondition("accountid", ConditionOperator.Equal, accountId);
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
        //    query.Criteria.AddCondition("log_name", ConditionOperator.Equal, "SANO-306-10-15");
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
            //query.Criteria.AddCondition("log_employeenumber", ConditionOperator.Equal, "NO-61026");
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
           // query.Criteria.AddCondition("log_name", ConditionOperator.Equal, "1178");
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
            var resultLise = service.RetrieveMultiple(query);

            var post = resultLise.Entities.FirstOrDefault();
            return post;

        }

        private static Entity GetInstallationWithAlarmSystemId(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("log_installation");
            query.Criteria.AddCondition("log_hardwareid", ConditionOperator.Contains);
            query.Criteria.AddCondition("log_alarmsystemid", ConditionOperator.Contains);
            var resultLise = service.RetrieveMultiple(query);

            var installation = resultLise.Entities.FirstOrDefault();
            return installation;
        }

        private static int GetCountInstallationWithAlarmSystemId(OrganizationServiceProxy service)
        {
            var query = new QueryExpression
            {
                EntityName = "log_installation",
                ColumnSet = new ColumnSet("log_hardwareid2", "log_hardwareid", "log_alarmsystemid")
            }; 
            FilterExpression filter = new FilterExpression(LogicalOperator.And);
            filter.AddCondition("log_hardwareid", ConditionOperator.NotNull);
            filter.AddCondition("log_alarmsystemid", ConditionOperator.NotNull);
            filter.AddCondition("log_hardwareid2", ConditionOperator.NotNull);
            query.Criteria = filter;
            //    query.Criteria.AddCondition("log_alarmsystemid", ConditionOperator.Contains);
            var resultLise = service.RetrieveMultiple(query);
            foreach (var c in resultLise.Entities)
            {

                var jsonString = JsonConvert.SerializeObject(c);
                Debug.WriteLine(jsonString.ToString());
            }
            var count = resultLise.Entities.Count();
            Debug.WriteLine(count.ToString());
            return count;
        }

        //Get post code  which number is "1178"
        //Service degree is full in that area
        private static Entity GetSaleTYPE(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("sa_genericreason");
            var resultLise = service.RetrieveMultiple(query);
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
       //     newLead["log_address2_postalcode"] = GetPostCode(service).ToEntityReference();
            //newLead["log_postalcode"] = GetPostCode(service).ToEntityReference();
            newLead["address1_line1"] = "address1 vitaminveien 1, oslo";
            newLead["log_canoverwritecreditcheck"] = true;

            //GetTakdOverCASE
       //     newLead["log_takeovercase"] = GetTakdOverCASE(service).ToEntityReference();
            //home
            newLead["log_typeoflead"] = new OptionSetValue(182400000);
            newLead["log_movefrominstallation"] = GetInstallation(service).ToEntityReference();
            newLead["log_movetoinstallation"] = GetInstallation(service).ToEntityReference();

            //ireland eller spain spesifikk
           // newLead["sa_salestype"] = GetSaleTYPE(service).ToEntityReference();
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
            //newLead["log_takeovercase"] = GetTakdOverCASE(service).ToEntityReference();
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
            //entity["log_address2_postalcode"] = GetPostCode(service).ToEntityReference();
            //entity["log_postalcode"] = GetPostCode(service).ToEntityReference();
            entity["address1_line1"] = "address1 vitaminveien 1, oslo";
            entity["log_canoverwritecreditcheck"] = true;
           // entity["log_takeovercase"] = GetTakdOverCASE(service).ToEntityReference();
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
            qWorkorder.Criteria.AddCondition("log_installationid", ConditionOperator.Equal, accountId);

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
