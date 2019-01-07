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


        public const string accountName = "Company 69n from Code";
        public static void Run(OrganizationServiceProxy service)
    {

            Guid id = CreateWO(service, accountName);
            ApproveWO(service,  id);

        }
        /*
         * Once we have approved the Work order we cannot add more Work Order Products. If the status is Open we can add as many work orders as we need.
         * 
          **/
        private static Guid CreateWO(OrganizationServiceProxy service, String name)
        {
            //log_approvedforinvoicen ? why not this one
            var createEntity = new Entity("log_workorders");
            createEntity["log_name"] = name;
            createEntity["log_accountid"] = GetAccount(service).ToEntityReference();
            createEntity["log_contractid"] = GetContract(service).ToEntityReference();
            createEntity["log_installationid"] = GetInstallation(service).ToEntityReference();
            return service.Create(createEntity);
        }

        private static Entity GetAccount(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("account");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, accountName);
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

        //Add two different work order products.One of type subscription, and one of type hardware.
        private static void AddProductsToWO(OrganizationServiceProxy service, Guid workorderID)
        {
            var createEntity = new Entity("log_workorderproduct");
            createEntity["log_workordersid"] = new EntityReference("log_workorders", workorderID);
            service.Create(createEntity);
            var updateWO = new Entity("log_workorders");
            updateWO["statuscode"] = new OptionSetValue(2);

            service.Update(updateWO);
        }
        private static void ApproveWO(OrganizationServiceProxy service, Guid workorderID)
        {

            var updateWO = new Entity("log_workorders");
            updateWO["log_workordersid"] = workorderID;
            updateWO["log_actualend"] = DateTime.Now;
            updateWO["log_sectormobilestatus"] = new OptionSetValue(182400006);
            //add actual installer
            updateWO["log_employeeid"] = GetInstallerPerson(service).ToEntityReference();
        
            service.Update(updateWO);
        //State settes til inactive and status set to 
        //Value: 182400002, Label: Scheduled
        //Value: 1, Label: Open
        //Value: 2, Label: Approved
        //Value: 182400000, Label: Cancelled
        //Value: 182400001, Label: No - show
            updateWO["statecode"] = new OptionSetValue(1);
            updateWO["statuscode"] = new OptionSetValue(2);

            service.Update(updateWO);
        }

        //Get sales person which number is "GS-80064"
        private static Entity GetInstallerPerson(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("log_employee");
            query.Criteria.AddCondition("log_employeenumber", ConditionOperator.Equal, "GS-80064");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var user = resultLise.Entities.FirstOrDefault();
            return user;

        }

    }
}
