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

     
        public static void Run(OrganizationServiceProxy service)
    {
             string accountName = DateTime.Now.ToShortTimeString();
            Guid id = CreateWO(service, accountName);
            AddProductWithTypeHardware(service, id);
            AddProductWithTypeContract(service, id);
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
            string accountName = DateTime.Now.ToShortTimeString();
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

        private static Entity GetProduct(OrganizationServiceProxy service, String productnumber)
        {
            var query = new QueryExpression("product");
            query.Criteria.AddCondition("productnumber", ConditionOperator.Equal, productnumber);
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var product = resultLise.Entities.FirstOrDefault();
            return product;
        }

        private static Entity GetPriceList(OrganizationServiceProxy service)
        {
            var query = new QueryExpression("pricelevel");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, "Eur");
            var resultLise = service.RetrieveMultiple(query);

            if (resultLise.Entities.Count == 0)
                throw new Exception("Entity Not found");
            var price = resultLise.Entities.FirstOrDefault();
            return price;
        }

        //Add two different work order products.One of type subscription, and one of type hardware.
        private static void AddProductWithTypeHardware(OrganizationServiceProxy service, Guid workorderID)
        {
            var createEntity = new Entity("log_workorderproduct");
            createEntity["log_workordersid"] = new EntityReference("log_workorders", workorderID);
            createEntity["log_hardwareproductid"] = GetProduct(service, "90001").ToEntityReference();
            createEntity["log_hardwareproductdefaultprice"] = new Money(Convert.ToDecimal(10.00));
            service.Create(createEntity);
        }


        private static void AddProductWithTypeContract(OrganizationServiceProxy service, Guid workorderID)
        {
            var createEntity = new Entity("log_workorderproduct");
            createEntity["log_hardwareproductid"] = GetProduct(service, "20041").ToEntityReference();
            createEntity["log_hardwareproductdefaultprice"] = new Money(Convert.ToDecimal(10.00));
            createEntity["log_workordersid"] = new EntityReference("log_workorders", workorderID);
            //Dette kan være tomt
            createEntity["log_contractproductid"] = GetProduct(service, "90500").ToEntityReference();
            createEntity["log_contractproductdefaultprice"] = new Money(Convert.ToDecimal(90.00));
            createEntity["log_pricelist"] = GetPriceList(service).ToEntityReference();
            service.Create(createEntity);
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
