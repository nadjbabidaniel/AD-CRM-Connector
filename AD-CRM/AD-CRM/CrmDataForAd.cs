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
            //using (OrganizationServiceContext orgSvcContext = new OrganizationServiceContext(_service))
            {               
                AD_GroupList = (from groupList in orgSvcContext.CreateQuery("a24_adgroup")
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
            //_service.Update(entity);
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
                Entity CRMuser = CreateNewCrmSystemuUser(businesID);
                Entity synchronizedCRMuser = Synchronization(ADuser, CRMuser);

                string fullAccountName = @"A24XRMDOMAIN\" + ADuser.Properties["sAMAccountName"].Value.ToString();
                var userCrm = (from userTemp in orgSvcContext.CreateQuery("systemuser")
                                 where userTemp.GetAttributeValue<String>("domainname").Equals(fullAccountName)
                                 select userTemp).FirstOrDefault();

                //userCrm["systemuserid"] = userCrmId.Id;              
                //_service.Update(entity);        //UpdateCrmEntity(synchronizedUser); 
                try
                {
                    if (userCrm == null)
                    {
                        //var _accountId = _service.Create(synchronizedUser);
                    }
                }
                catch (Exception ex) { }

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

                crmUser.Attributes["domainname"] = fullDomainName;
            }

            if (ADuser.Properties["givenname"].Value != null)
            {
                crmUser.Attributes["firstname"] = ADuser.Properties["givenname"].Value.ToString();
            }

            if (ADuser.Properties["sn"].Value != null)
            {
                crmUser.Attributes["lastname"] = ADuser.Properties["sn"].Value.ToString();
            }

            if (ADuser.Properties["title"].Value != null)
            {
                crmUser.Attributes["title"] = ADuser.Properties["title"].Value.ToString();
            }

            if (ADuser.Properties["l"].Value != null)
            {
                crmUser.Attributes["address1_city"] = ADuser.Properties["l"].Value.ToString();
            }

            if (ADuser.Properties["streetAddress"].Value != null)
            {
                crmUser.Attributes["address1_line1"] = ADuser.Properties["streetAddress"].Value.ToString();
            }

            if (ADuser.Properties["postalCode"].Value != null)
            {
                crmUser.Attributes["address1_postalcode"] = ADuser.Properties["postalCode"].Value.ToString();
            }

            if (ADuser.Properties["st"].Value != null)
            {
                crmUser.Attributes["address1_stateorprovince"] = ADuser.Properties["st"].Value.ToString();
            }

            if (ADuser.Properties["co"].Value != null)
            {
                crmUser.Attributes["address1_country"] = ADuser.Properties["co"].Value.ToString();
            }

            if (ADuser.Properties["telephoneNumber"].Value != null)
            {               
                crmUser.Attributes["address1_telephone1"] = ADuser.Properties["telephoneNumber"].Value.ToString();
            }

            if (ADuser.Properties["mobile"].Value != null)
            {
                crmUser.Attributes["mobilephone"] = ADuser.Properties["mobile"].Value.ToString();
            }

            if (ADuser.Properties["facsimileTelephoneNumber"].Value != null)
            {
                crmUser.Attributes["address1_fax"] = ADuser.Properties["facsimileTelephoneNumber"].Value.ToString();
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
            List<String> listOfUserRoles = new List<String>();
            QueryExpression queryExpression = new QueryExpression();
            queryExpression.EntityName = "role"; //role entity name
            ColumnSet cols = new ColumnSet();
            cols.AddColumn("name"); //We only need role name
            queryExpression.ColumnSet = cols;
            ConditionExpression ce = new ConditionExpression();
            ce.AttributeName = "systemuserid";
            ce.Operator = ConditionOperator.Equal;
            ce.Values.Add(CRMuser.Id);
            //system roles
            LinkEntity lnkEntityRole = new LinkEntity();
            lnkEntityRole.LinkFromAttributeName = "roleid";
            lnkEntityRole.LinkFromEntityName = "role"; //FROM
            lnkEntityRole.LinkToEntityName = "systemuserroles";
            lnkEntityRole.LinkToAttributeName = "roleid";
            //system users
            LinkEntity lnkEntitySystemusers = new LinkEntity();
            lnkEntitySystemusers.LinkFromEntityName = "systemuserroles";
            lnkEntitySystemusers.LinkFromAttributeName = "systemuserid";
            lnkEntitySystemusers.LinkToEntityName = "systemuser";
            lnkEntitySystemusers.LinkToAttributeName = "systemuserid";
            lnkEntitySystemusers.LinkCriteria = new FilterExpression();
            lnkEntitySystemusers.LinkCriteria.Conditions.Add(ce);
            lnkEntityRole.LinkEntities.Add(lnkEntitySystemusers);
            queryExpression.LinkEntities.Add(lnkEntityRole);
            EntityCollection entColRoles = _service.RetrieveMultiple(queryExpression);
            if (entColRoles != null && entColRoles.Entities.Count > 0)
            {
                foreach (Entity entRole in entColRoles.Entities)
                {
                    listOfUserRoles.Add(entRole.Attributes["name"].ToString());
                }
            }

            ////PART FOR READING USER TEAMS:
            List<String> listOfUserTeams = new List<String>();
            QueryExpression queryExpressionTeams = new QueryExpression();
            queryExpressionTeams.EntityName = "team"; //role entity name
            ColumnSet colsTeams = new ColumnSet();
            colsTeams.AddColumn("name"); //We only need team name
            queryExpressionTeams.ColumnSet = colsTeams;
            ConditionExpression ceTeams = new ConditionExpression();
            ceTeams.AttributeName = "systemuserid";
            ceTeams.Operator = ConditionOperator.Equal;
            ceTeams.Values.Add(CRMuser.Id);
            //system roles
            LinkEntity lnkEntityTeam = new LinkEntity();
            lnkEntityTeam.LinkFromAttributeName = "teamid";
            lnkEntityTeam.LinkFromEntityName = "team"; //FROM
            lnkEntityTeam.LinkToEntityName = "teammembership";
            lnkEntityTeam.LinkToAttributeName = "teamid";
            //system users
            LinkEntity lnkEntitySystemusersTeams = new LinkEntity();
            lnkEntitySystemusersTeams.LinkFromEntityName = "teammembership";
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
                    listOfUserTeams.Add(entTeam.Attributes["name"].ToString());
                }
            }

            foreach (var group in userGroups)
            {
                var CRMgroup = AD_GroupList.FirstOrDefault(x => x.GetAttributeValue<string>("a24_name").Equals(group.Properties["cn"].Value.ToString()));

                if (CRMgroup != null)
                {
                    var valueListCrmSecurityObjectIdConn = AD_Groups__CrmsSecurityObject.Where(x => x.GetAttributeValue<Guid>("a24_adgroupid").ToString()
                                            .Equals(CRMgroup.GetAttributeValue<Guid>("a24_adgroupid").ToString()));

                    foreach (var a24_valueListCrmSecurityObjectIdConn in valueListCrmSecurityObjectIdConn)
                    {

                        var crmSecurityObject = CRMsecurityObjectList.FirstOrDefault(x => x.GetAttributeValue<Guid>("a24_crmsecurityobjectid").ToString()
                                                      .Equals(a24_valueListCrmSecurityObjectIdConn.GetAttributeValue<Guid>("a24_crmsecurityobjectid").ToString()));

                        string nameValue = crmSecurityObject.GetAttributeValue<string>("a24_name");
                        string typeValue = crmSecurityObject.GetAttributeValue<bool>("a24_type").ToString();

                        if (typeValue.Equals("False"))  //should be changed when model is changed with false -> role
                        {
                            if (AllRoles.Count > 0)
                            {
                                var role = AllRoles.FirstOrDefault(x => x.GetAttributeValue<string>("name").Equals(nameValue) &&
                                          x.GetAttributeValue<EntityReference>("businessunitid").Id.Equals(CRMuser.GetAttributeValue<EntityReference>("businessunitid").Id));  //ToLower() in case if its not case sensitive

                                if (role != null)
                                {
                                    string valueRole = role.GetAttributeValue<string>("name");
                                    if (nameValue.Equals(valueRole) && !listOfUserRoles.Contains(valueRole))
                                    {
                                        try
                                        {
                                            _service.Associate("systemuser", CRMuser.Id,
                                                new Relationship("systemuserroles_association"),
                                                new EntityReferenceCollection() { new EntityReference("role", role.Id) });
                                        }
                                        catch (Exception ex) { }
                                    }
                                }
                            }
                        }
                        else if (typeValue.Equals("True"))  //should be changed when model is changed with false -> team
                        {
                            if (AllTeams.Count > 0)
                            {
                                var team = AllTeams.FirstOrDefault(x => x.GetAttributeValue<string>("name").Equals(nameValue));   //ToLower() in case if its not case sensitive

                                if (team != null)
                                {
                                    string valueTeam = team.GetAttributeValue<string>("name");
                                    if (nameValue.Equals(valueTeam) && !listOfUserTeams.Contains(valueTeam))
                                    {
                                        try
                                        {
                                            _service.Associate("systemuser", CRMuser.Id,
                                                 new Relationship("teammembership_association"),
                                                 new EntityReferenceCollection() { new EntityReference("team", team.Id) });
                                        }
                                        catch (Exception ex) { }
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
