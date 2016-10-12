using Microsoft.Xrm.Sdk;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.DirectoryServices;
using System.DirectoryServices.Protocols;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace AD_CRM
{
    class Program
    {
        private static String username;
        private static String password;
        private static bool syncFirst;
        private static String groupsFromXml;

        private static List<DirectoryEntry> allUsersListAD = new List<DirectoryEntry>();
        private static List<byte[]> listUserIdsInByte;

        private static List<DirectoryEntry> allGroupListAD = new List<DirectoryEntry>();
        private static List<byte[]> listGroupIdsInByte;

        private static CrmDataForAd CRM;

        static void Main(string[] args)
        {
            XmlDocument doc = new XmlDocument();
            string Xml_Path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            Xml_Path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Xml_Path), "XmlData\\" + "ConfigFile.xml");
            doc.Load(Xml_Path);

            XmlNode activedirectory = doc.SelectSingleNode("/settings/activedirectory");
            string LdapConn = activedirectory.Attributes["ldap"].Value;
            username = activedirectory.Attributes["username"].Value;
            password = activedirectory.Attributes["password"].Value;
            syncFirst = Boolean.Parse(activedirectory.Attributes["syncFirst"].Value);   //always true or false

            XmlNodeList xnList = doc.SelectNodes("/settings/organizations/organization/@adgroup");

            try
            {
                String ldapPath = "LDAP://" + LdapConn;

                DirectoryEntry directoryEntry = new DirectoryEntry(ldapPath, username, password);

                //// Create search for all groups from config file
                DirectorySearcher search = new DirectorySearcher(directoryEntry);
                foreach (XmlNode item in xnList)
                {
                    groupsFromXml += "(CN=" + item.Value + ")";
                }
                search.Filter = String.Format("(&(ObjectClass=Group)(|{0}))", groupsFromXml);
                SearchResultCollection results = search.FindAll();

                //Create a list of users from received groups           
                for (int i = 0; i < results.Count; i++)
                {
                    DirectoryEntry de = results[i].GetDirectoryEntry();
                    allGroupListAD.Add(de);

                    var members = (IEnumerable)de.Invoke("members");// Invoke("members");
                    foreach (object member in members)
                    {
                        DirectoryEntry user = new DirectoryEntry(member);
                        allUsersListAD.Add(user);
                    }
                }
                //Get list of objectGUIDs from relevant users and populate the list with those values converted to Byte[]                      //ADD TO FUNCTION
                listUserIdsInByte = getByteArrays(allUsersListAD);

                //Get list of objectGUIDs from relevant groups and populate the list with those values converted to Byte[]
                listGroupIdsInByte = getByteArrays(allGroupListAD);


                CRM = new CrmDataForAd(username, password);

                if (syncFirst)
                {
                    foreach (var ADuser in allUsersListAD)
                    {
                        Entity crmUser = CRM.GetUserFromCRM(ADuser);
                        if (crmUser != null)
                        {
                            // Add condition in case if flag allows this for RELEVANT DATA
                            //TODO try catch for no values also and for CRM access
                            //var overwriteactivedirectorysync = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("new_overwriteactivedirectorysync"));

                            if (crmUser.Attributes["a24_overwriteadsync"].ToString().Equals("false"))
                            {
                                Entity synchronizedUser = CRM.Synchronization(ADuser, crmUser);
                                CRM.UpdateFromDataModel(ADuser, crmUser);
                                CRM.UpdateCrmUser(synchronizedUser);
                            }

                            CRM.CompareOUandBU(ADuser, crmUser);
                        }
                        else
                        {
                            try
                            {
                                CRM.CreateNewCRMUser(ADuser);
                            }
                            catch { }
                        }

                    }
                }

                //Create Ldap connection for notifiers and remove domainName from user for loging
                username = username.Substring(username.IndexOf(@"\") + 1);

                LdapConnection connection = new LdapConnection(LdapConn);
                var credentials = new NetworkCredential(username, password);
                connection.Credential = credentials;
                connection.Bind();

                using (ChangeNotifier notifier = new ChangeNotifier(connection))
                {
                    //register some objects for notifications (limit 5)
                    notifier.Register("dc=a24xrmdomain,dc=info", System.DirectoryServices.Protocols.SearchScope.Subtree);

                    notifier.ObjectChanged += new EventHandler<ObjectChangedEventArgs>(notifier_ObjectChanged);

                    Console.WriteLine("Waiting for changes...");
                    Console.WriteLine();
                    Console.ReadLine();
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
                Console.WriteLine(e.Message);
                Console.ReadKey();
            }
        }

        static void notifier_ObjectChanged(object sender, ObjectChangedEventArgs e)
        {
            if ((e.Result.DistinguishedName).Contains("OU=AD2CRM,OU=ComData,OU=Extern,OU=A24-UsersAndGroups,DC=a24xrmdomain,DC=info"))
            {
                object[] objectGuid = e.Result.Attributes["objectguid"].GetValues(typeof(Byte[]));
                byte[] byteTest = (System.Byte[])objectGuid[0];


                bool contains = listUserIdsInByte.Any(x => x.SequenceEqual(byteTest));
                if (contains)
                {
                    DirectoryEntry ADuser = new DirectoryEntry("LDAP://XRMSERVER02.a24xrmdomain.info/" + e.Result.DistinguishedName, username, password);
                    Console.WriteLine(e.Result.DistinguishedName + "------------------------------------------------------------------------------------------------------");

                    Entity crmUser = CRM.GetUserFromCRM(ADuser);
                    if (crmUser != null)
                    {
                        // Add condition in case if flag allows this for RELEVANT DATA and change prefix publisher in CRM to a24-need help
                        //TODO try catch for no values also and for CRM access                       

                        if (crmUser.Attributes["a24_overwriteadsync"].ToString().Equals("false"))
                        {
                            Entity synchronizedUser = CRM.Synchronization(ADuser, crmUser);
                            CRM.UpdateCrmUser(synchronizedUser);
                        }
                    }
                    //User active or deactivated
                    int flags = (int)ADuser.Properties["userAccountControl"].Value;
                    var active = !Convert.ToBoolean(flags & 0x0002);
                    if (active == false) Console.WriteLine("User deactivated"); //Logging part

                    //Part to check if user has changed his OU
                    foreach (var item in allUsersListAD)
                    {
                        byte[] tempBute = ObjectGuid(item.Properties["objectguid"].Value);
                        if (tempBute.SequenceEqual(byteTest))
                        {
                            if (!item.Properties["distinguishedName"].Value.ToString().Equals(ADuser.Properties["distinguishedName"].Value.ToString())) Console.WriteLine("User has changed his OU"); break;
                        }
                    }

                    Console.WriteLine();
                    Console.WriteLine("====================================================================================================================================");
                    Console.WriteLine();
                    //Console.ReadKey();
                }

                bool containsGroups = listGroupIdsInByte.Any(x => x.SequenceEqual(byteTest));
                if (containsGroups)
                {
                    DirectoryEntry groupAD = new DirectoryEntry("LDAP://XRMSERVER02.a24xrmdomain.info/" + e.Result.DistinguishedName, username, password);
                    Console.WriteLine(e.Result.DistinguishedName + "------------------------------------------------------------------------------------------------------");

                    DirectoryEntry previousStateGroup = new DirectoryEntry(); // = allGroupListAD.FirstOrDefault(x=>x.Properties["objectguid"].Value.Equals(objectGuid[0]));

                    foreach (var item in allGroupListAD)  //Can be done and with /*listIdsInByteGroups*/ but than you have to read Entity
                    {
                        byte[] byteTemp = ObjectGuid(item.Properties["objectguid"].Value);
                        if (byteTemp.SequenceEqual(byteTest)) previousStateGroup = item;
                    }

                    List<DirectoryEntry> oldGroupUsersList = new List<DirectoryEntry>();
                    List<DirectoryEntry> newGroupUsersList = new List<DirectoryEntry>();

                    var membersOld = (IEnumerable)previousStateGroup.Invoke("members");// Invoke("members");
                    foreach (object member in membersOld)
                    {
                        DirectoryEntry user = new DirectoryEntry(member);
                        oldGroupUsersList.Add(user);
                    }

                    var membersNew = (IEnumerable)groupAD.Invoke("members");// Invoke("members");
                    foreach (object member in membersNew)
                    {
                        DirectoryEntry user = new DirectoryEntry(member);
                        newGroupUsersList.Add(user);
                    }

                    List<DirectoryEntry> UsersAddedToGroups = new List<DirectoryEntry>(); //Users that are Added to Group
                    List<DirectoryEntry> UsersRemovedFromGroup = new List<DirectoryEntry>(); //Users that are Removed to Group

                    //if (newGroupUsersList.Count >= oldGroupUsersList.Count)                             //    NEED TO CHECK LOGIC
                    {
                        foreach (var newUser in newGroupUsersList)
                        {
                            foreach (var oldUser in oldGroupUsersList)
                            {
                                byte[] newByteTest = ObjectGuid(newUser.Properties["objectguid"].Value);
                                byte[] oldByteTest = ObjectGuid(oldUser.Properties["objectguid"].Value);

                                if (newByteTest.SequenceEqual(oldByteTest)) continue;

                                UsersAddedToGroups.Add(newUser);
                            }
                        }
                    }
                    //else if (newGroupUsersList.Count < oldGroupUsersList.Count)
                    {
                        foreach (var oldUser in oldGroupUsersList)
                        {
                            foreach (var newUser in newGroupUsersList)
                            {
                                byte[] newByteTest = ObjectGuid(newUser.Properties["objectguid"].Value);
                                byte[] oldByteTest = ObjectGuid(oldUser.Properties["objectguid"].Value);

                                if (newByteTest.SequenceEqual(oldByteTest)) continue;

                                UsersRemovedFromGroup.Add(oldUser);
                            }
                        }
                    }
                    //Update old List of groups with group that have new users 
                    byte[] byteTempUpdate = ObjectGuid(groupAD.Properties["objectguid"].Value);

                    foreach (var item in allGroupListAD)
                    {
                        byte[] byteTempSearch = ObjectGuid(item.Properties["objectguid"].Value);

                        if (byteTempSearch.SequenceEqual(byteTempUpdate))
                        {
                            allGroupListAD.Remove(item);
                            allGroupListAD.Add(groupAD);
                            break;
                        }
                    }

                    //Go through Added users and add Roles/Teams from our DataModel to them
                    foreach (var ADuser in UsersAddedToGroups)
                    {
                        Entity crmUser = CRM.GetUserFromCRM(ADuser);
                        if (crmUser != null)
                        {
                            if (crmUser.Attributes["a24_overwriteadsync"].ToString().Equals("false"))
                            {
                                Entity synchronizedUser = CRM.Synchronization(ADuser, crmUser);
                                CRM.UpdateFromDataModel(ADuser, crmUser);
                                CRM.UpdateCrmUser(synchronizedUser);
                            }

                            CRM.CompareOUandBU(ADuser, crmUser);
                        }
                        else
                        {
                            CRM.CreateNewCRMUser(ADuser);
                        }

                        //Add new user in list of all users
                        allUsersListAD.Add(ADuser);
                    }

                    foreach (var ADuser in UsersRemovedFromGroup)
                    {
                        allUsersListAD.Remove(ADuser);           //Check once again what will happen in case that user is not in the list, but he should be alwaays there
                    }



                    listUserIdsInByte = getByteArrays(allUsersListAD);

                }
            }
        }

        static Byte[] ObjectGuid(object o)
        {
            byte[] byteTempUpdate = (System.Byte[])o;
            return byteTempUpdate;
        }

        static List<byte[]> getByteArrays(List<DirectoryEntry> allEntities)
        {
            List<PropertyValueCollection> ids = allEntities.Select(x => x.Properties["objectGUID"]).ToList();


            List<byte[]> temp = new List<byte[]>();
            byte[] byteTest;

            foreach (var item in ids)
            {
                byteTest = (byte[])item.Value;
                temp.Add(byteTest);
            }

            return temp;
        }
    }
}
