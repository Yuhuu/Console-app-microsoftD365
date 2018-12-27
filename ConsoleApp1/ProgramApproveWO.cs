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
    public class ProgramApproveWO
    {


        /*
         * 
          **/
        public static void Run(OrganizationServiceProxy service)
    {

            // Retrieve all accounts owned by the user with read access rights to the accounts and   
            // where the last name of the user is not Cannon.   
            string fetch = @"  
                <fetch mapping='logical'>
                <entity name = 'account'>
                <attribute name='name'/>
                <attribute name = 'primarycontactid' />
                <attribute name='log_dateofbirth'/>
                <attribute name = 'telephone1' />
                <attribute name='accountid'/>
                <order descending = 'false' attribute='name'/>
                <filter type = 'and' >
                <condition attribute='name' value='%Customer%' operator='like'/>
                </filter>
                </entity>
                </fetch> ";

            EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetch));
            foreach (var c in result.Entities)
            {
                Debug.WriteLine(c.Attributes["name"].ToString());
            }

        }
        private static Guid CreateWO(OrganizationServiceProxy service, String name)
        {

            var createEntity = new Entity("log_workorders");
            createEntity["log_name"] = name;
            createEntity["log_accountid"] = GetAccount(service).ToEntityReference();
            createEntity["log_contractid"] = GetContract(service).ToEntityReference();
            createEntity["log_installationid"] = GetInstallation(service).ToEntityReference();
            return service.Create(createEntity);
            //return null;
        }

        private static Entity GetAccount(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("account");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, "LeadProcess.companyName");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var user = resultLise.Entities.FirstOrDefault();
            return user;

        }

        private static Entity GetContract(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("log_contract");
            query.Criteria.AddCondition("log_name", ConditionOperator.Equal, "K-NOR-000011604");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var user = resultLise.Entities.FirstOrDefault();
            return user;

        }

        private static Entity GetInstallation(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("log_installation");
            query.Criteria.AddCondition("log_hardwareid", ConditionOperator.Equal, "Installation: vicky 68n from Code");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var user = resultLise.Entities.FirstOrDefault();
            return user;

        }

    }
}
