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
 /*
 * Lead is a protential buyer. Lead Process is a class to show the customer acquisition process which
 * ldentifies potential buyers. And when the leads are considered qualified, get passed from marketing to create 
 * more entities. The purpose of it is to effective lead management is ultimately contribute to more sales* 
 */
    class LeadProcess
    {
        
        private const string currency = "NOK";
        private const string leadSourceName = "Autolease";
        public const string companyName = "Company 77n from Code";

        public static void Main(string[] args)
        {
            var credentials = CredentialCache.DefaultNetworkCredentials;
            var clientCredentials = new ClientCredentials();
            clientCredentials.UserName.UserName = credentials.Domain + "\\" + credentials.UserName;
            clientCredentials.UserName.Password = credentials.Password;
            var service = new OrganizationServiceProxy(new Uri("https://intmscrmtst.sectoralarm.net/SectorAlarmfrtstPLAYGROUND/XRMServices/2011/Organization.svc"), null, clientCredentials, null);

            Guid id = CreateLeadWithName(service, companyName);
            UpdateLead(service, id);
            QualifyLead(service, id);
            //ProgramFetchQuery.Run(service);
            ProgramApproveWO.Run(service);

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
        private static Entity GetUser( OrganizationServiceProxy service) {
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
            query.Criteria.AddCondition("log_name", ConditionOperator.Equal, "Domestic");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            return resultLise.Entities.FirstOrDefault();
        }


        //Get sales person which number is "GS-80064"
        private static Entity GetSalesPerson(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("log_employee");
            query.Criteria.AddCondition("log_employeenumber", ConditionOperator.Equal, "GS-80064");
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

        /*
        *  Lead create Process allow end user to choose which name they want to use 
        *  The purpose of we want the name as a parameter is that it can be easily track and create.
        */
        private static Guid CreateLeadWithName(OrganizationServiceProxy service, String name)
        {

            var leadSource = GetLeadSourceName(service, leadSourceName);
            var newLead = new Entity("lead")
            {
                LogicalName = "lead"
            };

            var user = GetUser(service);
            newLead["companyname"] = name;
            newLead["log_leadsourceid"] = leadSource.ToEntityReference();
            newLead["ownerid"] = user.ToEntityReference();
            return service.Create(newLead);
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
            var entity = new Entity("lead");
            entity["leadid"] = id;
            entity["log_contracttermsid"] = GetContractTerms(service).ToEntityReference();
            entity["log_dateofbirth"] = DateTime.Now.AddYears(-19);
            //Date cannot be in the future
            entity["log_solddate"] = DateTime.Now;
            entity["log_salespersonid"] = GetSalesPerson(service).ToEntityReference();
            entity["address2_line1"] = "address2 vitaminveien 1, oslo";
            entity["log_address2_postalcode"] = GetPostCode(service).ToEntityReference();
            entity["log_postalcode"] = GetPostCode(service).ToEntityReference();
            entity["address1_line1"] = "address1 vitaminveien 1, oslo";
            entity["log_canoverwritecreditcheck"] = true;
            service.Update(entity);

            var createEntity = new Entity("log_workorderproduct");
                createEntity["log_leadid"] = new EntityReference("lead", id);
                    service.Create(createEntity);

                Debug.WriteLine(entity.ToString());
            entity["log_convertleadflag"] = 1;

            service.Update(entity);
                Debug.WriteLine(entity.ToString());
            
        }
    }
}
