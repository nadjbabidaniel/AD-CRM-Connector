using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.Protocols;
using System.Linq;
using System.Net;
using System.Xml;
using Microsoft.Xrm.Sdk;

namespace AD_CRM
{
    internal class Program
    {
        #region fields

        private static string _username;
        private static string _password;
        private static bool _syncFirst;
        private static string _groupsFromXml;
        private static string _ldapPath;
        private static string _domain;

        private static List<DirectoryEntry> _allRelevantUsersAD = new List<DirectoryEntry>();
        private static List<DirectoryEntry> _allRelevantGroupsAD = new List<DirectoryEntry>();

        private static List<byte[]> _listUserIdsInByte;
        private static List<byte[]> _listGroupIdsInByte;

        private static CRMDataForAD _crm;

        #endregion fields

        private static void Main(string[] args)
        {
            //Reading ConfigFile.xml 
            XmlDocument doc = new XmlDocument();
            string Xml_Path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            Xml_Path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Xml_Path), "XmlData\\" + "ConfigFile.xml");
            doc.Load(Xml_Path);

            XmlNode activedirectory = doc.SelectSingleNode("/settings/activedirectory");
            string ldapConn = activedirectory.Attributes["ldap"].Value;
            _ldapPath = "LDAP://" + ldapConn;
            _username = activedirectory.Attributes["username"].Value;
            _password = activedirectory.Attributes["password"].Value;
            _domain = _username.Substring(0, _username.LastIndexOf(@"\") + 1);

            bool isSyncFirst;
            if (Boolean.TryParse(activedirectory.Attributes["syncFirst"].Value, out isSyncFirst))
            {
                _syncFirst = isSyncFirst;
            }
            else
            {
                _syncFirst = false;
            }

            XmlNodeList xmlADGroupsNodeList = doc.SelectNodes("/settings/organizations/organization/@adgroup");

            try
            {
                DirectoryEntry directoryEntry = new DirectoryEntry(_ldapPath, _username, _password);

                //// Create search for all groups from config file
                DirectorySearcher search = new DirectorySearcher(directoryEntry);
                foreach (XmlNode item in xmlADGroupsNodeList)
                {
                    _groupsFromXml += "(CN=" + item.Value + ")";
                }
                search.Filter = String.Format("(&(ObjectClass=Group)(|{0}))", _groupsFromXml);
                SearchResultCollection results = search.FindAll();

                //Create a list of users from received groups (AD relevant groups from ConfigFile)
                for (int i = 0; i < results.Count; i++)
                {
                    DirectoryEntry de = results[i].GetDirectoryEntry();
                    _allRelevantGroupsAD.Add(de);

                    var members = (IEnumerable)de.Invoke("members");
                    foreach (object member in members)
                    {
                        DirectoryEntry user = new DirectoryEntry(member);
                        _allRelevantUsersAD.Add(user);
                    }
                }
                //Get list of objectGUIDs from relevant users and populate the list with those values converted to Byte[]                     
                _listUserIdsInByte = getObjectGUID(_allRelevantUsersAD);

                //Get list of objectGUIDs from relevant groups and populate the list with those values converted to Byte[]
                _listGroupIdsInByte = getObjectGUID(_allRelevantGroupsAD);

                _crm = new CRMDataForAD(_username, _password, _domain);

                if (_syncFirst) // if _syncFirst is true, service will first sync AD data with CRM data and then listen to changes
                {
                    syncDataFirst();
                }
                listenToChanges(ldapConn);
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
                Console.WriteLine(e.Message);
                Console.ReadKey();
            }
        }

        #region notifiers

        private static void notifier_ObjectChanged(object sender, ObjectChangedEventArgs e)
        {
            if ((e.Result.DistinguishedName).Contains("OU=AD2CRM,OU=ComData,OU=Extern,OU=A24-UsersAndGroups,DC=a24xrmdomain,DC=info")) //Will be removed, filter for now only users and groups from our AD2CRM OU
            {
                Entity crmuser = null;
                object[] objectGuid = e.Result.Attributes["objectguid"].GetValues(typeof(Byte[]));
                byte[] bytefromGuid = (System.Byte[])objectGuid[0];

                bool contains = _listUserIdsInByte.Any(x => x.SequenceEqual(bytefromGuid));
                if (contains)
                {
                    DirectoryEntry adUser = new DirectoryEntry(_ldapPath + "/" + e.Result.DistinguishedName, _username, _password);
                    Console.WriteLine(e.Result.DistinguishedName + "------------------------------------------------------------------------------------------------------");

                    try
                    {
                        crmuser = _crm.GetUserFromCRM(adUser);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(ex.Message);
                        Console.WriteLine(ex.Message);
                    }
                    if (crmuser != null)
                    {
                        syncAndUpdateUser(adUser, crmuser);
                    }

                    isUserActive(adUser);

                    checkUserOUChange(bytefromGuid, adUser);

                    Console.WriteLine();
                    Console.WriteLine("====================================================================================================================================");
                    Console.WriteLine();
                }

                bool containsGroups = _listGroupIdsInByte.Any(x => x.SequenceEqual(bytefromGuid));
                if (containsGroups)
                {
                    DirectoryEntry adGroup = new DirectoryEntry(_ldapPath + "/" + e.Result.DistinguishedName, _username, _password);
                    Console.WriteLine(e.Result.DistinguishedName + "------------------------------------------------------------------------------------------------------");

                    checkUserAddedRemovedFromADGroup(adGroup, bytefromGuid);
                }
            }
        }

        #endregion notifiers

        #region private methods

        private static void syncDataFirst()
        {
            foreach (var adUser in _allRelevantUsersAD)
            {
                if (adUser.Properties["givenname"].Value.ToString().Equals("Andrea"))
                {
                    Entity crmUser = null;
                    try
                    {
                        crmUser = _crm.GetUserFromCRM(adUser);
                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Debug.WriteLine(e.Message);
                        Console.WriteLine(e.Message);
                    }
                    if (crmUser != null)
                    {
                        //  if a24_overwriteadsync_bit is set to FALSE, CRM data will be synced with AD relevant data and readonly. If TRUE, data will be disassociated from AD, and editable
                        if (crmUser.GetAttributeValue<bool>("a24_overwriteadsync_bit").Equals(null))
                        {
                            crmUser.Attributes["a24_overwriteadsync_bit"] = false; // set default value
                        }
                        if (crmUser.GetAttributeValue<bool>("a24_overwriteadsync_bit") == false)
                        {
                            Entity synchronizedUser = _crm.Synchronization(adUser, crmUser);
                            _crm.UpdateFromDataModel(adUser, crmUser);
                            _crm.UpdateCrmUser(crmUser);
                        }
                        else
                        {
                            crmUser.Attributes["a24_adsync_bit"] = false;
                        }

                        _crm.CompareOUandBU(adUser, crmUser);
                    }
                    else
                    {
                        try
                        {
                            _crm.CreateNewCRMUser(adUser);
                        }
                        catch (Exception e)
                        {
                            System.Diagnostics.Debug.WriteLine(e.Message);
                            Console.WriteLine(e.Message);
                        }
                    }
                }
            }
        }

        private static void listenToChanges(string ldapConn)
        {
            //Create Ldap connection for notifiers and remove domainName from user for loging
            _username = _username.Substring(_username.IndexOf(@"\") + 1);
            LdapConnection connection = new LdapConnection(ldapConn);
            var credentials = new NetworkCredential(_username, _password);
            connection.Credential = credentials;
            connection.Bind();

            using (ChangeNotifier notifier = new ChangeNotifier(connection))
            {
                //register some objects for notifications (limit 5)
                notifier.Register("dc=a24xrmdomain,dc=info", System.DirectoryServices.Protocols.SearchScope.Subtree);
                notifier.ObjectChanged += (notifier_ObjectChanged);

                Console.WriteLine("Waiting for changes...");
                Console.WriteLine();
                Console.ReadLine();
            }
        }

        private static void syncAndUpdateUser(DirectoryEntry adUser, Entity crmUser)
        {
            if (crmUser.GetAttributeValue<bool>("a24_overwriteadsync_bit").Equals(null))
            {
                crmUser.Attributes["a24_overwriteadsync_bit"] = false;
            }
            if (crmUser.GetAttributeValue<bool>("a24_overwriteadsync_bit") == false)
            {
                Entity synchronizedUser = _crm.Synchronization(adUser, crmUser);
                _crm.UpdateCrmUser(synchronizedUser);
            }
            else
            {
                crmUser.Attributes["a24_adsync_bit"] = false;
            }
        }

        private static Byte[] objectGuid(object o)
        {
            byte[] byteGuid = (Byte[])o;
            return byteGuid;
        }

        private static List<byte[]> getObjectGUID(List<DirectoryEntry> allEntities)
        {
            List<PropertyValueCollection> ids = allEntities.Select(x => x.Properties["objectGUID"]).ToList();

            List<byte[]> guidList = new List<byte[]>();
            byte[] bytesfromguid;

            foreach (var item in ids)
            {
                bytesfromguid = (byte[])item.Value;
                guidList.Add(bytesfromguid);
            }

            return guidList;
        }

        private static void checkUserOUChange(byte[] bytefromGuid, DirectoryEntry adUser)
        {
            // Check if user has changed his OU
            foreach (var item in _allRelevantUsersAD)
            {
                byte[] tempBute = objectGuid(item.Properties["objectguid"].Value);
                if (tempBute.SequenceEqual(bytefromGuid))
                {
                    if (!item.Properties["distinguishedName"].Value.ToString().Equals(adUser.Properties["distinguishedName"].Value.ToString()))
                        Console.WriteLine("User has changed his OU");
                    break;
                }
            }
        }

        private static void isUserActive(DirectoryEntry adUser)
        {
            //User active or deactivated
            int flags = (int)adUser.Properties["userAccountControl"].Value;
            var active = !Convert.ToBoolean(flags & 0x0002);
            if (active == false)
                Console.WriteLine("User deactivated"); //Logging part
        }

        private static void checkUserAddedRemovedFromADGroup(DirectoryEntry adGroup, byte[] bytefromGuid)
        {
            DirectoryEntry previousStateGroup = new DirectoryEntry(); // = allGroupListAD.FirstOrDefault(x=>x.Properties["objectguid"].Value.Equals(objectGuid[0]));

            foreach (var item in _allRelevantGroupsAD)  //Can be done and with /*listIdsInByteGroups*/ but than you have to read Entity
            {
                byte[] byteTemp = objectGuid(item.Properties["objectguid"].Value);
                if (byteTemp.SequenceEqual(bytefromGuid))
                    previousStateGroup = item;
            }

            List<DirectoryEntry> oldGroupUsersList = new List<DirectoryEntry>();
            List<DirectoryEntry> newGroupUsersList = new List<DirectoryEntry>();

            var membersOld = (IEnumerable)previousStateGroup.Invoke("members");
            foreach (object member in membersOld)
            {
                DirectoryEntry user = new DirectoryEntry(member);
                oldGroupUsersList.Add(user);
            }

            var membersNew = (IEnumerable)adGroup.Invoke("members");
            foreach (object member in membersNew)
            {
                DirectoryEntry user = new DirectoryEntry(member);
                newGroupUsersList.Add(user);
            }

            List<byte[]> _newGroupUsersListIdsInByte = getObjectGUID(newGroupUsersList);
            List<byte[]> _oldGroupUsersListIdsInByte = getObjectGUID(oldGroupUsersList);

            List<DirectoryEntry> usersAddedToGroups = new List<DirectoryEntry>(); //Users that are Added to Group
            List<DirectoryEntry> usersRemovedFromGroup = new List<DirectoryEntry>(); //Users that are Removed to Group

            //newGroupUsersList and _newGroupUsersListIdsInByte have same number of users, accordingly Id and its User
            for (int i = 0; i < _newGroupUsersListIdsInByte.Count; i++)
            {
                if (!(_oldGroupUsersListIdsInByte.Any(x => x.SequenceEqual(_newGroupUsersListIdsInByte[i]))))
                {
                    usersAddedToGroups.Add(newGroupUsersList[i]);
                }
            }

            for (int i = 0; i < _oldGroupUsersListIdsInByte.Count; i++)
            {
                if (!(_newGroupUsersListIdsInByte.Any(x => x.SequenceEqual(_oldGroupUsersListIdsInByte[i]))))
                {
                    usersRemovedFromGroup.Add(oldGroupUsersList[i]);
                }
            }

            //Update old List of groups with group that has new users
            byte[] guidAdGroup = objectGuid(adGroup.Properties["objectguid"].Value);

            foreach (var item in _allRelevantGroupsAD)
            {
                byte[] guidAllRelevantAdGroup = objectGuid(item.Properties["objectguid"].Value);

                if (guidAllRelevantAdGroup.SequenceEqual(guidAdGroup))
                {
                    _allRelevantGroupsAD.Remove(item);
                    _allRelevantGroupsAD.Add(adGroup);
                    break;
                }
            }
            addDataModelForNewADGroupUsers(usersAddedToGroups, usersRemovedFromGroup);
        }

        //Go through Added users and add Roles/Teams from our DataModel to them
        private static void addDataModelForNewADGroupUsers(List<DirectoryEntry> usersAddedToGroups, List<DirectoryEntry> usersRemovedFromGroup)
        {
            Entity crmUser = null;

            foreach (var adUser in usersAddedToGroups)
            {
                try
                {
                    crmUser = _crm.GetUserFromCRM(adUser);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.Message);
                    Console.WriteLine(ex.Message);
                }
                if (crmUser != null)
                {
                    if (crmUser.GetAttributeValue<bool>("a24_overwriteadsync_bit") == false)
                    {
                        Entity synchronizedUser = _crm.Synchronization(adUser, crmUser);
                        _crm.UpdateFromDataModel(adUser, crmUser);
                        _crm.UpdateCrmUser(synchronizedUser);
                    }
                    else
                    {
                        crmUser.Attributes["a24_adsync_bit"] = false;
                    }

                    _crm.CompareOUandBU(adUser, crmUser);
                }
                else
                {
                    _crm.CreateNewCRMUser(adUser);
                }

                //Add new user in list of all users                
                _allRelevantUsersAD.Add(adUser);
            }

            foreach (var ADuser in usersRemovedFromGroup)      //Need to check in case if he is relevant for some other relevant group !!!
            {                 
                _allRelevantUsersAD.Remove(_allRelevantUsersAD.FirstOrDefault(x => x.Guid == ADuser.Guid)); //Check once again what will happen in case that user is not in the list, but he should be alwaays there
            }

            _listUserIdsInByte = getObjectGUID(_allRelevantUsersAD);
        }

        #endregion private methods
    }
}
