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
    public class ProgramFetchQuery
    {


        /*
          * Retrieve all accounts owned by the user with read access rights to the accounts 
          * and the name contains customer
          * then it is printed out in Debug output.
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


}
}
