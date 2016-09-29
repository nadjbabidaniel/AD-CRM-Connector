using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

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

        public Entity getUserFromCRM(String fullAccountName)
        {
                        
            Entity SystemUserBasedOnsAMAccountName = (from user in orgSvcContext.CreateQuery("systemuser")
                                   where user.GetAttributeValue<String>("domainname").Equals(fullAccountName)
                                                      select user).FirstOrDefault();
            return SystemUserBasedOnsAMAccountName;
        }


    }
}
