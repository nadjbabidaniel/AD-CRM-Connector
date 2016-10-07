using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.ServiceModel.Description;


namespace AD_CRM
{
    public class CrmDataForAd
    {
        private static string soapOrgServiceUri = "https://adtocrmconnector.dev.anywhere24.com/XRMServices/2011/Organization.svc";
        private Uri serviceUri;
        private OrganizationServiceProxy proxy;
        private IOrganizationService _service;
        private OrganizationServiceContext orgSvcContext;

        public Entity SystemUserBasedOnId;
        public List<Entity> AD_GroupList;
        public List<Entity> AD_Groups__CrmsSecurityObject;
        public List<Entity> CRMsecurityobjectList;

        public EntityCollection Results;

        public CrmDataForAd(string user, string pass)
        {
            ClientCredentials credentials = new ClientCredentials();
            credentials.UserName.UserName = user;
            credentials.UserName.Password = pass;

            serviceUri = new Uri(soapOrgServiceUri);
            proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
            proxy.EnableProxyTypes();
            _service = (IOrganizationService)proxy;

            orgSvcContext = new OrganizationServiceContext(_service);

            //Populate list of all business units
            ListOfBusinessUnits();
        }

        public void DataForAdSync()
        {
            Guid id = new Guid("271cdd05-f97e-e611-80d1-00155d10673a");

            //using (OrganizationServiceContext orgSvcContext = new OrganizationServiceContext(_service))
            {
                SystemUserBasedOnId = (from campaing in orgSvcContext.CreateQuery("systemuser")
                                       where campaing.GetAttributeValue<Guid>("systemuserid") == id
                                       select campaing).FirstOrDefault();

                AD_GroupList = (from campaing in orgSvcContext.CreateQuery("a24_adgroup")
                                    //where campaing.GetAttributeValue<Guid>("systemuserid") == id
                                select campaing).ToList();

                AD_Groups__CrmsSecurityObject = (from campaing in orgSvcContext.CreateQuery("a24_a24_crmsecurityobject_a24_adgroup")
                                                     //where campaing.GetAttributeValue<Guid>("systemuserid") == id
                                                 select campaing).ToList();

                CRMsecurityobjectList = (from campaing in orgSvcContext.CreateQuery("a24_crmsecurityobject")
                                             //where campaing.GetAttributeValue<Guid>("systemuserid") == id
                                         select campaing).ToList();
            }
        }

        public Entity GetUserFromCRM(String fullAccountName)
        {

            Entity SystemUser = (from user in orgSvcContext.CreateQuery("systemuser")
                                                      where user.GetAttributeValue<String>("domainname").Equals(fullAccountName)
                                                      select user).FirstOrDefault();
            return SystemUser;
        }

        public void UpdateCrmEntity(Entity entity)
        {
            //entity.EntityState = EntityState.Changed;
            //_service.Update(entity);

            //orgSvcContext.ClearChanges();
            //orgSvcContext.Attach(entity);
            orgSvcContext.UpdateObject(entity);
            orgSvcContext.SaveChanges();

        }

        public void ListOfBusinessUnits()
        {
            QueryExpression businessUnitQuery = new QueryExpression
            {
                EntityName = "businessunit",
                ColumnSet = new ColumnSet("businessunitid", "name", "parentbusinessunitid"),
            };

            Results = _service.RetrieveMultiple(businessUnitQuery);
        }

        public void CreateNewCRMUser(DirectoryEntry ADuser)
        {
            string distinguishedName = ADuser.Properties["distinguishedName"].Value.ToString();
            string[] words = distinguishedName.Split(',');

            string OU = String.Empty;
            if (words[1].StartsWith("OU"))
            {
                OU = words[1].Substring(3);
            }

            if (!String.IsNullOrEmpty(OU))
            {
                Entity businessUnit = null;
                businessUnit = Results.Entities.ToList().FirstOrDefault(x => x.Attributes.Values.ToArray()[1].ToString().Equals(OU)); //Find BU with the same name

                if (businessUnit == null)
                {
                    //Take parent BU - he is only one with two attribute values, ie dont have parent BU
                    businessUnit = Results.Entities.ToList().FirstOrDefault(x => x.Attributes.Values.Count == 2);
                }

                String businesID = businessUnit.Attributes.Values.ToArray()[0].ToString();
                Entity user = CreateNewCrmSystemuUer(businesID);

                //var userCrmId = (from userTemp in orgSvcContext.CreateQuery("systemuser")
                //                 where userTemp.GetAttributeValue<String>("domainname").Equals(ADuser.Properties["userPrincipalName"].Value.ToString())
                //                 select userTemp).FirstOrDefault();

                //user["systemuserid"] = userCrmId.Id;

                Entity synchronizedUser = Synchronization(ADuser, user);

                //_service.Update(entity);        //UpdateCrmEntity(synchronizedUser); 
                //var _accountId = _service.Create(synchronizedUser);

            }
        }

        public Entity CreateNewCrmSystemuUer(string businesID)
        {
            Entity user = new Entity("systemuser");

            user["systemuserid"] = new Guid();
            user["domainname"] = String.Empty;
            user["firstname"] = String.Empty;
            user["lastname"] = String.Empty;
            user["title"] = String.Empty;
            user["address1_city"] = String.Empty;
            user["address1_line1"] = String.Empty;
            user["address1_postalcode"] = String.Empty;
            user["address1_stateorprovince"] = String.Empty;
            user["address1_country"] = String.Empty;
            user["address1_telephone1"] = String.Empty;
            user["mobilephone"] = String.Empty;
            user["address1_fax"] = String.Empty;

            Guid ownerId = new Guid(businesID);
            user["businessunitid"] = new EntityReference("businessunit", ownerId);

            return user;
        }

        public Entity Synchronization(DirectoryEntry ADuser, Entity crmUser)
        {

            if (ADuser.Properties["userPrincipalName"].Value != null)
            {
                string userPrincipalName = ADuser.Properties["userPrincipalName"].Value.ToString();
                string[] words = userPrincipalName.Split('@');
                string username = words[0];
                string[] domains = words[1].Split('.');
                string domain = domains[0];

                string fullDomainName = domain.ToUpper() + @"\" + username;

                var domainname = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("domainname"));
                crmUser.Attributes[domainname.Key] = fullDomainName;
            }

            if (ADuser.Properties["givenname"].Value != null)
            {
                var firstname = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("firstname"));
                crmUser.Attributes[firstname.Key] = ADuser.Properties["givenname"].Value.ToString();
            }

            if (ADuser.Properties["sn"].Value != null)
            {
                var lastname = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("lastname"));
                crmUser.Attributes[lastname.Key] = ADuser.Properties["sn"].Value.ToString();
            }

            if (ADuser.Properties["title"].Value != null)
            {
                var title = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("title"));
                crmUser.Attributes[title.Key] = ADuser.Properties["title"].Value.ToString();
            }

            if (ADuser.Properties["l"].Value != null)
            {
                var address1_city = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("address1_city"));
                crmUser.Attributes[address1_city.Key] = ADuser.Properties["l"].Value.ToString();
            }

            if (ADuser.Properties["streetAddress"].Value != null)
            {
                var address1_line1 = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("address1_line1"));
                crmUser.Attributes[address1_line1.Key] = ADuser.Properties["streetAddress"].Value.ToString();
            }

            if (ADuser.Properties["postalCode"].Value != null)
            {
                var address1_postalcode = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("address1_postalcode"));
                crmUser.Attributes[address1_postalcode.Key] = ADuser.Properties["postalCode"].Value.ToString();
            }

            if (ADuser.Properties["st"].Value != null)
            {
                var address1_stateorprovince = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("address1_stateorprovince"));
                crmUser.Attributes[address1_stateorprovince.Key] = ADuser.Properties["st"].Value.ToString();
            }

            if (ADuser.Properties["co"].Value != null)
            {
                var address1_country = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("address1_country"));
                crmUser.Attributes[address1_country.Key] = ADuser.Properties["co"].Value.ToString();
            }

            if (ADuser.Properties["telephoneNumber"].Value != null)
            {
                var address1_telephone1 = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("address1_telephone1"));
                crmUser.Attributes[address1_telephone1.Key] = ADuser.Properties["telephoneNumber"].Value.ToString();
            }

            if (ADuser.Properties["mobile"].Value != null)
            {
                var mobilephone = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("mobilephone"));
                crmUser.Attributes[mobilephone.Key] = ADuser.Properties["mobile"].Value.ToString();
            }

            if (ADuser.Properties["facsimileTelephoneNumber"].Value != null)
            {
                var address1_fax = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("address1_fax"));
                crmUser.Attributes[address1_fax.Key] = ADuser.Properties["facsimileTelephoneNumber"].Value.ToString();
            }




            return crmUser;
        }


    }
}
