using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleAppCsharp.ConsoleApp1
{
    public static class ViewUtilities
    {
        public static List<EntityMetadata> entityMeta;
        //public static EntityMetadata[] Metadata(IOrganizationService service)
        //{
        //    GetMetadata(service, false);
        //    return entityMeta;
        //}
        //public static EntityMetadata[] Metadata(IOrganizationService service, bool force)
        //{
        //    GetMetadata(service, force);
        //    return entityMeta;
        //}
        public static EntityMetadata Metadata(string entityName, IOrganizationService service)
        {
            GetMetadata(service, false, entityName);
            return entityMeta.First(z => z.LogicalName == entityName);
        }
        //public static AttributeMetadata[] Attributes(IOrganizationService service)
        //{
        //    GetMetadata(service, false);
        //    return entityMeta.SelectMany(z => z.Attributes).ToArray();
        //}
        public static AttributeMetadata[] Attributes(string entityName, IOrganizationService service)
        {
            GetMetadata(service, false, entityName);
            return entityMeta.First(z => z.LogicalName == entityName).Attributes;
        }
        //public static AttributeMetadata[] Attributes(IOrganizationService service, bool force)
        //{
        //    GetMetadata(service, force);
        //    return entityMeta.SelectMany(z => z.Attributes).ToArray();
        //}
        static void GetMetadata(IOrganizationService service, bool forceReCaching, string entity)
        {
            //Check if metatdata already exist in memory or if we need to fetch from CRM
            if (forceReCaching || entityMeta == null || entityMeta.FirstOrDefault(z => z.LogicalName == entity) == null)
            {
                var qMetadata = new RetrieveEntityRequest()
                {
                    LogicalName = entity,
                    EntityFilters = EntityFilters.All
                };
                //var meta = (service.Execute(qMetadata) as RetrieveAllEntitiesResponse).EntityMetadata;
                //entityMeta.AddRange(meta);
                var meta = (service.Execute(qMetadata) as RetrieveEntityResponse);
                if (entityMeta == null)
                {
                    entityMeta = new List<EntityMetadata>();
                }
                entityMeta.Add(meta.EntityMetadata);
            }
        }
        public static decimal GetStep(int? precision)
        {
            var p = (int)precision * -1;
            return (decimal)Math.Pow(10, p);
        }
        public static string AttributeTypeToHtmlType(string attributetypename)
        {
            switch (attributetypename)
            {
                case "DateTimeType": return "datetime";
                case "BooleanType": return "radio";// "checkbox";             
                case "MoneyType":
                case "DoubleType":
                case "BigIntType":
                case "DecimalType":
                case "IntegerType": return "number";
                case "MemoType":
                case "EntityNameType":
                case "LookupType": //return "link for click, text m/ search for input?"
                case "StringType": return "text";
                case "StateType":
                case "StatusType":
                case "PicklistType": return "radio";
                case "ImageType": return "file";
                default: return attributetypename;
            }
        }
        public static string FormatTypeToHtmlType(string formatname)
        {
            switch (formatname)
            {
                case "Phone": return "tel";
                case "Url": return "url";
                case "Email": return "email";
                case "TickerSymbol": // "tickersymbol" - not used
                case "PhoneticGuide":// "yominame" - not used
                case "Text":
                case null:
                default: return "text";
            }
        }
        public static Dictionary<string, int> GetOptionSet(OptionSetMetadata options)
        {
            var optionSet = new Dictionary<string, int>();
            foreach (var o in options.Options)
            { // ToDo: Add multiple language support by checking for LocalizedLables language codes
                optionSet.Add(o.Label.LocalizedLabels.First().Label, (int)o.Value);
            }
            return optionSet;
        }
        public static string[] EntityColumns(string entity)
        {

            switch (entity)
            {
                case "sa_log": return new string[] { "sa_name", "sa_logdetails", "sa_logtype", "sa_installation" };
                case "account": return new string[] { "name", "emailaddress1", "telephone1", "log_mobilephone", "log_mobilephone2" }; // 182 400 000 Private, 01 SME
                //case "account": return new string[] { "log_customersegment", "name", "parentaccountid", "log_livingaddress", "accountnumber" }; // 182 400 000 Private, 01 SME
                //case "contact": return new string[] { "birthdate", "customertypecode", "emailaddress1", "emailaddress2", "telephone1", "telephone2", "firstname", "lastname", "sa_actualcontact", "address1_line1", "address1_line2", "address1_line3" }; // "log_livingaddress" };// Default Value 1, 182 400 000 Private, 01 SME, 02 Actual Contact
                case "contact": return new string[] { "mobilephone", "emailaddress1", "firstname", "lastname" };
                case "log_installation": return new string[] { "log_alarmsystemid", "log_hardwareid", "log_installationid", "log_installationdate", "log_street1", "log_postalcode", "log_city" };
                case "connection": return new string[] { "effectivestart", "effectiveend", "record1id", "record1roleid", "record2id", "record2roleid", "description", "statecode" };
                case "task": return new string[] { "subject" };
                default: throw new NotImplementedException();
            }
            //return new string[] { "#" };
        }
    }
}