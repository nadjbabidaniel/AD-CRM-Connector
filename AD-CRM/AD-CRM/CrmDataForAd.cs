using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace AD_CRM
{
    internal enum SecurityObjectType
    {
        SecurityRole = 602370000,
        Team = 602370001
    };

    public class CRMDataForAD
    {
        #region fields

        public List<Entity> AD_GroupList;
        public List<Entity> AD_Groups__CrmsSecurityObject;
        public List<Entity> AllRoles;
        public List<Entity> AllTeams;
        public List<Entity> CRMsecurityObjectList;
        public EntityCollection Results;
        public Entity SystemUserBasedOnId;
        private static string _soapOrgServiceUri = "https://adtocrmconnector.dev.anywhere24.com/XRMServices/2011/Organization.svc";
        private OrganizationServiceContext _orgSvcContext;
        private IOrganizationService _service;
       
        #endregion fields

        #region constructor

        public CRMDataForAD(string user, string pass)
        {
            var credentials = new ClientCredentials();
            credentials.UserName.UserName = user;
            credentials.UserName.Password = pass;

            Uri serviceUri = new Uri(_soapOrgServiceUri);
            var proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
            proxy.EnableProxyTypes();
            _service = (IOrganizationService)proxy;

            _orgSvcContext = new OrganizationServiceContext(_service);

            //Populate list of all business units and Data model
            listOfBusinessUnits();
            dataForAdSync();
        }

        #endregion constructor

        #region public methods

        public void CompareOUandBU(DirectoryEntry adUser, Entity crmUser)
        {
            string distinguishedName = adUser.Properties["distinguishedName"].Value.ToString();
            string[] words = distinguishedName.Split(new[] { ",OU=" }, StringSplitOptions.None);

            string OU = words[1];

            EntityReference crmBusinessUnit = crmUser.GetAttributeValue<EntityReference>("businessunitid");

            if (!OU.Equals(crmBusinessUnit.Name))
            {
                //Do the logging part that UO name is not the same as BU name
            }
        }

        public void CreateNewCRMUser(DirectoryEntry adUser)
        {
          string domain = adUser.Properties["distinguishedName"].Value.ToString ();
          string[] splitDomain = domain.Split (new[] { ",DC=" }, StringSplitOptions.None);
          string DC = splitDomain[1] + @"\";
            string fullAccountName = DC.ToUpper() + adUser.Properties["sAMAccountName"].Value.ToString ();
            var userExistInCrm = (from userTemp in _orgSvcContext.CreateQuery("systemuser")
                           where userTemp.GetAttributeValue<String>("domainname").Equals(fullAccountName)
                           select userTemp).FirstOrDefault();
            if (userExistInCrm == null)
            {


                string distinguishedName = adUser.Properties["distinguishedName"].Value.ToString();
                string[] words = distinguishedName.Split(new[] { ",OU=" }, StringSplitOptions.None);

                string OU = words[1];

                Entity businessUnit = null;
                businessUnit = Results.Entities.ToList().FirstOrDefault(x => x.GetAttributeValue<string>("name").Equals(OU));    //Find BU with the same name  --You can test creation of new user in case existing BU

                if (businessUnit == null)
                {
                    //Take parent BU - he is only one with two attribute values, ie dont have parent BU
                    businessUnit = Results.Entities.ToList().FirstOrDefault(x => x.Attributes.Values.Count == 2);
                }

                Guid businesID = businessUnit.GetAttributeValue<Guid>("businessunitid");
                Entity CRMuser = createNewCrmSystemuUser(businesID);
                Synchronization(adUser, CRMuser);

                try
                {
                    //var _accountId = _service.Create(CRMuser);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.Message);
                    Console.WriteLine(ex.Message);
                }
            }
        }

        public Entity GetUserFromCRM(DirectoryEntry adUser)
        {
            string domain = adUser.Properties["distinguishedName"].Value.ToString ();
            string[] splitDomain = domain.Split (new[] { ",DC=" }, StringSplitOptions.None);
            string DC = splitDomain[1] + @"\";
            string fullAccountName = DC.ToUpper() + adUser.Properties["sAMAccountName"].Value.ToString ();

            Entity CRMuser = (from user in _orgSvcContext.CreateQuery("systemuser")
                              where user.GetAttributeValue<String>("domainname").Equals(fullAccountName)
                              select user).FirstOrDefault();
            return CRMuser;
        }

        public Entity Synchronization(DirectoryEntry adUser, Entity crmUser)
        {
            //if (adUser.Properties["userPrincipalName"].Value != null)
            //{
            //    string userPrincipalName = adUser.Properties["userPrincipalName"].Value.ToString();
            //    string[] words = userPrincipalName.Split('@');
            //    string username = words[0];
            //    string[] domains = words[1].Split('.');
            //    string domain = domains[0];

          if (adUser.Properties["distinguishedName"].Value != null)
            {
              string domain = adUser.Properties["distinguishedName"].Value.ToString ();
              string[] splitDomain = domain.Split (new[] { ",DC=" }, StringSplitOptions.None);
              string DC = splitDomain[1] + @"\";             
              string fullDomainName = DC.ToUpper() + adUser.Properties["sAMAccountName"].Value;

                crmUser.Attributes["domainname"] = fullDomainName;
            }

            if (adUser.Properties["givenname"].Value != null)
            {
                crmUser.Attributes["firstname"] = adUser.Properties["givenname"].Value.ToString();
            }

            if (adUser.Properties["sn"].Value != null)
            {
                crmUser.Attributes["lastname"] = adUser.Properties["sn"].Value.ToString();
            }

            if (adUser.Properties["title"].Value != null)
            {
                crmUser.Attributes["title"] = adUser.Properties["title"].Value.ToString();
            }

            if (adUser.Properties["l"].Value != null)
            {
                crmUser.Attributes["address1_city"] = adUser.Properties["l"].Value.ToString();
            }

            if (adUser.Properties["streetAddress"].Value != null)
            {
                crmUser.Attributes["address1_line1"] = adUser.Properties["streetAddress"].Value.ToString();
            }

            if (adUser.Properties["postalCode"].Value != null)
            {
                crmUser.Attributes["address1_postalcode"] = adUser.Properties["postalCode"].Value.ToString();
            }

            if (adUser.Properties["st"].Value != null)
            {
                crmUser.Attributes["address1_stateorprovince"] = adUser.Properties["st"].Value.ToString();
            }

            if (adUser.Properties["co"].Value != null)
            {
                crmUser.Attributes["address1_country"] = adUser.Properties["co"].Value.ToString();
            }

            if (adUser.Properties["telephoneNumber"].Value != null)
            {
                crmUser.Attributes["address1_telephone1"] = adUser.Properties["telephoneNumber"].Value.ToString();
            }

            if (adUser.Properties["mobile"].Value != null)
            {
                crmUser.Attributes["mobilephone"] = adUser.Properties["mobile"].Value.ToString();
            }

            if (adUser.Properties["facsimileTelephoneNumber"].Value != null)
            {
                crmUser.Attributes["address1_fax"] = adUser.Properties["facsimileTelephoneNumber"].Value.ToString();
            }

            crmUser.Attributes["a24_adsync_bit"] = true;

            return crmUser;
        }

        public void UpdateCrmUser(Entity crmUser)
        {
            //_service.Update(crmUser);
            _orgSvcContext.UpdateObject(crmUser);
            _orgSvcContext.SaveChanges();
        }

        public void UpdateFromDataModel(DirectoryEntry adUser, Entity crmUser)
        {
            List<DirectoryEntry> userGroups = new List<DirectoryEntry>();                 //Show all users groups
            object obGroups = adUser.Invoke("Groups");
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
            ce.Values.Add(crmUser.Id);

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
            ceTeams.Values.Add(crmUser.Id);

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
                var crmGroup = AD_GroupList.FirstOrDefault(x => x.GetAttributeValue<string>("a24_name").Equals(group.Properties["cn"].Value.ToString()));

                if (crmGroup != null)
                {
                    var valueListCrmSecurityObjectIdConn = AD_Groups__CrmsSecurityObject.Where(x => x.GetAttributeValue<Guid>("a24_adgroupid").ToString()
                                           .Equals(crmGroup.GetAttributeValue<Guid>("a24_adgroupid").ToString()));

                    foreach (var a24_valueListCrmSecurityObjectIdConn in valueListCrmSecurityObjectIdConn)
                    {
                        var crmSecurityObject = CRMsecurityObjectList.FirstOrDefault(x => x.GetAttributeValue<Guid>("a24_crmsecurityobjectid").ToString()
                                                     .Equals(a24_valueListCrmSecurityObjectIdConn.GetAttributeValue<Guid>("a24_crmsecurityobjectid").ToString()));

                        string nameValue = crmSecurityObject.GetAttributeValue<string>("a24_name");
                        OptionSetValue typeValue = crmSecurityObject.GetAttributeValue<OptionSetValue>("a24_type_opt");

                        checkSecurityObjectType(typeValue, nameValue, crmUser, listOfUserRoles, listOfUserTeams);
                    }
                }
            }
        }

        #endregion public methods

        #region private methods

        private void checkSecurityObjectType(OptionSetValue typeValue, string nameValue, Entity crmUser, List<string> listOfUserRoles, List<string> listOfUserTeams)
        {
            if (typeValue.Value == (int)(SecurityObjectType.SecurityRole))  //should be changed when model is changed with false -> role
            {
                if (AllRoles.Count > 0)
                {
                    var role = AllRoles.FirstOrDefault(x => x.GetAttributeValue<string>("name").Equals(nameValue) &&
                             x.GetAttributeValue<EntityReference>("businessunitid").Id.Equals(crmUser.GetAttributeValue<EntityReference>("businessunitid").Id));  //ToLower() in case if its not case sensitive

                    if (role != null)
                    {
                        string valueRole = role.GetAttributeValue<string>("name");
                        if (nameValue.Equals(valueRole) && !listOfUserRoles.Contains(valueRole))
                        {
                            try
                            {
                                _service.Associate("systemuser", crmUser.Id,
                                    new Relationship("systemuserroles_association"),
                                    new EntityReferenceCollection() { new EntityReference("role", role.Id) });
                            }
                            catch (Exception ex) { }
                        }
                    }
                }
            }
            else if (typeValue.Value == (int)(SecurityObjectType.Team))  //should be changed when model is changed with false -> team
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
                                _service.Associate("systemuser", crmUser.Id,
                                     new Relationship("teammembership_association"),
                                     new EntityReferenceCollection() { new EntityReference("team", team.Id) });
                            }
                            catch (Exception ex) { }
                        }
                    }
                }
            }
        }

        private Entity createNewCrmSystemuUser(Guid businesID)
        {
            Entity user = new Entity("systemuser");

            user["systemuserid"] = new Guid();
            user["islicensed"] = false;                      // No licence
            user["a24_overwriteadsync_bit"] = false;
            user["a24_adsync_bit"] = true;
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
            user["businessunitid"] = new EntityReference("businessunit", businesID);

            return user;
        }

        private void dataForAdSync()
        {
            //using (OrganizationServiceContext orgSvcContext = new OrganizationServiceContext(_service))
            {
                AD_GroupList = (from groupList in _orgSvcContext.CreateQuery("a24_adgroup")
                                select groupList).ToList();

                AD_Groups__CrmsSecurityObject = (from GroupsObjects in _orgSvcContext.CreateQuery("a24_a24_crmsecurityobject_a24_adgroup")
                                                 select GroupsObjects).ToList();

                CRMsecurityObjectList = (from crmSecObject in _orgSvcContext.CreateQuery("a24_crmsecurityobject")
                                         select crmSecObject).ToList();

                AllRoles = (from role in _orgSvcContext.CreateQuery("role")
                            select role).ToList();

                //list of all NOT default Teams
                AllTeams = (from team in _orgSvcContext.CreateQuery("team")
                            where team.GetAttributeValue<bool>("isdefault") == false
                            select team).ToList();

                var list = (from groupList in _orgSvcContext.CreateQuery("a24_logging") //Logging
                            select groupList).ToList();
            }
        }
        private void listOfBusinessUnits()
        {
            QueryExpression businessUnitQuery = new QueryExpression
            {
                EntityName = "businessunit",
                ColumnSet = new ColumnSet("businessunitid", "name", "parentbusinessunitid"),
            };

            Results = _service.RetrieveMultiple(businessUnitQuery);
        }

        #endregion private methods
    }
}