using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
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
        public List<Entity> CRMsecurityObjectList;

        public List<Entity> AllRoles;
        public List<Entity> AllTeams;


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

            //Populate list of all business units and Data model
            ListOfBusinessUnits();
            DataForAdSync();
        }

        public void DataForAdSync()
        {
            Guid id = new Guid("271cdd05-f97e-e611-80d1-00155d10673a");

            //using (OrganizationServiceContext orgSvcContext = new OrganizationServiceContext(_service))
            {
                SystemUserBasedOnId = (from campaing in orgSvcContext.CreateQuery("systemuser")
                                       where campaing.GetAttributeValue<Guid>("systemuserid") == id
                                       select campaing).FirstOrDefault();

                AD_GroupList = (from groupList in orgSvcContext.CreateQuery("a24_adgroup")
                                    //where campaing.GetAttributeValue<Guid>("systemuserid") == id
                                select groupList).ToList();

                AD_Groups__CrmsSecurityObject = (from GroupsObjects in orgSvcContext.CreateQuery("a24_a24_crmsecurityobject_a24_adgroup")
                                                 select GroupsObjects).ToList();

                CRMsecurityObjectList = (from crmSecObject in orgSvcContext.CreateQuery("a24_crmsecurityobject")
                                         select crmSecObject).ToList();

                AllRoles = (from role in orgSvcContext.CreateQuery("role")
                            select role).ToList();

                //list of all NOT default Teams
                AllTeams = (from team in orgSvcContext.CreateQuery("team")
                            where team.GetAttributeValue<bool>("isdefault") == false
                            select team).ToList();

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
                Entity user = CreateNewCrmSystemuUser(businesID);

                //var userCrmId = (from userTemp in orgSvcContext.CreateQuery("systemuser")
                //                 where userTemp.GetAttributeValue<String>("domainname").Equals(ADuser.Properties["userPrincipalName"].Value.ToString())
                //                 select userTemp).FirstOrDefault();

                //user["systemuserid"] = userCrmId.Id;

                Entity synchronizedUser = Synchronization(ADuser, user);

                //_service.Update(entity);        //UpdateCrmEntity(synchronizedUser); 
                //var _accountId = _service.Create(synchronizedUser);

            }
        }

        public Entity CreateNewCrmSystemuUser(string businesID)
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




        public void UpdateFromDataModel(DirectoryEntry ADuser, Entity CRMuser)
        {
            List<DirectoryEntry> userGroups = new List<DirectoryEntry>();                 //Show all users groups
            object obGroups = ADuser.Invoke("Groups");
            foreach (object ob in (IEnumerable)obGroups)
            {
                // Create object for each group.
                DirectoryEntry obGpEntry = new DirectoryEntry(ob);
                userGroups.Add(obGpEntry);
            }

            //////PART FOR READING USER ROLES:
            //List<String> listOfUserRoles = new List<String>();
            //QueryExpression queryExpression = new QueryExpression();
            //queryExpression.EntityName = "role"; //role entity name
            //ColumnSet cols = new ColumnSet();
            //cols.AddColumn("name"); //We only need role name
            //queryExpression.ColumnSet = cols;
            //ConditionExpression ce = new ConditionExpression();
            //ce.AttributeName = "systemuserid";
            //ce.Operator = ConditionOperator.Equal;
            //ce.Values.Add(CRMuser.Id);
            ////system roles
            //LinkEntity lnkEntityRole = new LinkEntity();
            //lnkEntityRole.LinkFromAttributeName = "roleid";
            //lnkEntityRole.LinkFromEntityName = "role"; //FROM
            //lnkEntityRole.LinkToEntityName = "systemuserroles";
            //lnkEntityRole.LinkToAttributeName = "roleid";
            ////system users
            //LinkEntity lnkEntitySystemusers = new LinkEntity();
            //lnkEntitySystemusers.LinkFromEntityName = "systemuserroles";
            //lnkEntitySystemusers.LinkFromAttributeName = "systemuserid";
            //lnkEntitySystemusers.LinkToEntityName = "systemuser";
            //lnkEntitySystemusers.LinkToAttributeName = "systemuserid";
            //lnkEntitySystemusers.LinkCriteria = new FilterExpression();
            //lnkEntitySystemusers.LinkCriteria.Conditions.Add(ce);
            //lnkEntityRole.LinkEntities.Add(lnkEntitySystemusers);
            //queryExpression.LinkEntities.Add(lnkEntityRole);
            //EntityCollection entColRoles = _service.RetrieveMultiple(queryExpression);
            //if (entColRoles != null && entColRoles.Entities.Count > 0)
            //{
            //    foreach (Entity entRole in entColRoles.Entities)
            //    {
            //        listOfUserRoles.Add(entRole.Attributes["name"].ToString().ToLower());
            //    }
            //}

            ////PART FOR READING USER TEAMS:
            List<String> listOfUserTeams = new List<String>();
            QueryExpression queryExpressionTeams = new QueryExpression();
            queryExpressionTeams.EntityName = "team"; //role entity name
            ColumnSet colsTeams = new ColumnSet();
            colsTeams.AddColumn("name"); //We only need role name
            queryExpressionTeams.ColumnSet = colsTeams;
            ConditionExpression ceTeams = new ConditionExpression();
            ceTeams.AttributeName = "systemuserid";
            ceTeams.Operator = ConditionOperator.Equal;
            ceTeams.Values.Add(CRMuser.Id);
            //system roles
            LinkEntity lnkEntityTeam = new LinkEntity();
            lnkEntityTeam.LinkFromAttributeName = "roleid";
            lnkEntityTeam.LinkFromEntityName = "role"; //FROM
            lnkEntityTeam.LinkToEntityName = "systemuserroles";
            lnkEntityTeam.LinkToAttributeName = "roleid";
            //system users
            LinkEntity lnkEntitySystemusersTeams = new LinkEntity();
            lnkEntitySystemusersTeams.LinkFromEntityName = "systemuserroles";
            lnkEntitySystemusersTeams.LinkFromAttributeName = "systemuserid";
            lnkEntitySystemusersTeams.LinkToEntityName = "systemuser";
            lnkEntitySystemusersTeams.LinkToAttributeName = "systemuserid";
            lnkEntitySystemusersTeams.LinkCriteria = new FilterExpression();
            lnkEntitySystemusersTeams.LinkCriteria.Conditions.Add(ceTeams);
            lnkEntityTeam.LinkEntities.Add(lnkEntitySystemusersTeams);
            queryExpressionTeams.LinkEntities.Add(lnkEntityTeam);
            EntityCollection entColTeams = _service.RetrieveMultiple(queryExpressionTeams);
            if (entColTeams != null && entColTeams.Entities.Count > 0)
            {
                foreach (Entity entTeam in entColTeams.Entities)
                {
                    listOfUserTeams.Add(entTeam.Attributes["name"].ToString().ToLower());
                }
            }





            foreach (var group in userGroups)
            {
                foreach (var CRMgroup in AD_GroupList)
                {
                    int index_adGroup = Array.IndexOf(CRMgroup.Attributes.Keys.ToArray(), "a24_name");
                    string value_adGroup = CRMgroup.Attributes.Values.ToArray()[index_adGroup].ToString();

                    int index_adGroupContained;
                    string value_adGroupContained = string.Empty;

                    if (value_adGroup.Equals(group.Properties["cn"].Value))
                    {
                        index_adGroupContained = Array.IndexOf(CRMgroup.Attributes.Keys.ToArray(), "a24_adgroupid");
                        value_adGroupContained = CRMgroup.Attributes.Values.ToArray()[index_adGroupContained].ToString();
                    }
                    List<String> valueListCrmSecurityObjectIdConn = new List<string>();
                    foreach (var item in AD_Groups__CrmsSecurityObject)
                    {

                        int index_adGroupsSecObjectsConn = Array.IndexOf(item.Attributes.Keys.ToArray(), "a24_adgroupid");
                        string value_adGroupsSecObjectsConn = item.Attributes.Values.ToArray()[index_adGroupsSecObjectsConn].ToString();

                        if (value_adGroupContained.Equals(value_adGroupsSecObjectsConn))
                        {
                            int index_crmsecurityobjectidConn = Array.IndexOf(item.Attributes.Keys.ToArray(), "a24_crmsecurityobjectid");
                            valueListCrmSecurityObjectIdConn.Add(item.Attributes.Values.ToArray()[index_crmsecurityobjectidConn].ToString());
                        }
                    }
                    foreach (var crmSecurityObject in CRMsecurityObjectList)
                    {
                        foreach (var a24_valueListCrmSecurityObjectIdConn in valueListCrmSecurityObjectIdConn)
                        {
                            int index_crmSecurityObject = Array.IndexOf(crmSecurityObject.Attributes.Keys.ToArray(), "a24_crmsecurityobjectid");
                            string value_crmSecurityObject = crmSecurityObject.Attributes.Values.ToArray()[index_crmSecurityObject].ToString();

                            if (value_crmSecurityObject.Equals(a24_valueListCrmSecurityObjectIdConn))
                            {

                                int nameIndex = Array.IndexOf(crmSecurityObject.Attributes.Keys.ToArray(), "a24_name");
                                string nameValue = crmSecurityObject.Attributes.Values.ToArray()[nameIndex].ToString();

                                int typeIndex = Array.IndexOf(crmSecurityObject.Attributes.Keys.ToArray(), "a24_type");
                                string typeValue = crmSecurityObject.Attributes.Values.ToArray()[typeIndex].ToString();

                                if (typeValue.Equals("False"))  //should be changed when model is changed with false -> role
                                {
                                    if (AllRoles.Count != 0)
                                    {
                                        var indexRoleName = Array.IndexOf(AllRoles.FirstOrDefault().Attributes.Keys.ToArray(), "name");
                                        var role = AllRoles.FirstOrDefault(x => x.Attributes.Values.ToArray()[indexRoleName].ToString().Equals(nameValue));

                                        if (role != null)
                                        {
                                            //_service.Associate("systemuser", CRMuser.Id,
                                            //    new Relationship("systemuserroles_association"),
                                            //    new EntityReferenceCollection() { new EntityReference("role", role.Id) });
                                        }
                                    }
                                }
                                else if (typeValue.Equals("True"))  //should be changed when model is changed with false -> team
                                {
                                    foreach (var item in AllTeams)
                                    {
                                        var indexTeam = Array.IndexOf(item.Attributes.Keys.ToArray(), "name");
                                        var valueTeam = item.Attributes.Values.ToArray()[indexTeam].ToString();

                                        if (nameValue.Equals(valueTeam))
                                        {
                                            //_service.Associate("systemuser", CRMuser.Id,
                                            //     new Relationship("teammembership_association"),
                                            //     new EntityReferenceCollection() { new EntityReference("team", item.Id) });
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }






        }


    }
}
