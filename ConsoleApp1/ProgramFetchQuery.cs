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
      *As all  know, we need to track, update, and track some more.So while you update your lead. be sure that you have chosen the 
      * right one to update.
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
                System.Console.WriteLine(c.Attributes["name"]);
            }
        }


}
}
