﻿using Microsoft.Xrm.Sdk;
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
        private static List<Entity> allUsersListCRM = new List<Entity>();
        private static List<DirectoryEntry> allUsersNotExistCRM = new List<DirectoryEntry>();
        private static List<byte[]> listIdsInByte;

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
                    var members = (IEnumerable)de.Invoke("members");// Invoke("members");

                    foreach (object member in members)
                    {
                        DirectoryEntry user = new DirectoryEntry(member);
                        allUsersListAD.Add(user);
                    }
                }
                ////////////////////////////////
                //if (allUsersListAD.Count > 0)
                //{
                //    var user = allUsersListAD[0];

                //    foreach (string propName in user.Properties.PropertyNames)    //Show all users properties
                //    {
                //        if (user.Properties[propName].Value != null)
                //        {
                //            Console.WriteLine(propName + " = " + user.Properties[propName].Value.ToString());
                //        }
                //        else
                //        {
                //            Console.WriteLine(propName + " = NULL");
                //        }
                //    }
                //}
                //Console.ReadKey();

                //Get list of objectGUIDs from relevant users and populate the list with its values converted to Byte[]
                List<PropertyValueCollection> ids = allUsersListAD.Select(x => x.Properties["objectGUID"]).ToList();
                listIdsInByte = new List<byte[]>();
                byte[] byteTest;

                foreach (var item in ids)
                {
                    byteTest = (System.Byte[])item.Value;
                    //int i = BitConverter.ToInt32(byteTest, 0);
                    listIdsInByte.Add(byteTest);
                }

                CrmDataForAd CRM = new CrmDataForAd(username, password);
                //CRM.DataForAdSync();

                if (syncFirst)
                {
                    foreach (var userAD in allUsersListAD)
                    {
                        string fullAccountName = @"A24XRMDOMAIN\" + userAD.Properties["sAMAccountName"].Value.ToString();
                        Entity crmUser = CRM.GetUserFromCRM(fullAccountName);
                        if (crmUser != null)
                        {
                            // Add condition in case if flag allows this for RELEVANT DATA and change prefix publisher in CRM to a24-need help
                            //TODO try catch for no values also and for CRM access
                            //var overwriteactivedirectorysync = crmUser.Attributes.FirstOrDefault(x => x.Key.Equals("new_overwriteactivedirectorysync"));

                            //Entity synchronizedUser = CRM.SynchronizationADuserWithCRMuser(userAD, crmUser);
                            //UpdateCrmEntity(crmUser);

                            allUsersListCRM.Add(crmUser);
                        }
                        else
                        {
                            allUsersNotExistCRM.Add(userAD);
                        }
                    }
                }

                foreach (var adUser in allUsersNotExistCRM)
                {
                    try
                    {
                        CRM.CreateNewCRMUser(adUser);
                    }
                    catch { }
                }


                //if (allUsersListCRM.Count > 0)
                //{
                //    var user = allUsersListCRM[0];

                //    foreach (var propName in user.Attributes)    //Show all users properties
                //    {
                //        Console.WriteLine(propName.Key + " = " + propName.Value);
                //    }
                //}

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
                byte[] byteTest2 = (System.Byte[])objectGuid[0];
                //string strTemp = BitConverter.ToString(byteTest2);

                //if (listIdsInByte[0].SequenceEqual(byteTest2))
                bool contains = listIdsInByte.Any(x => x.SequenceEqual(byteTest2));
                if (contains)
                {
                    DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://XRMSERVER02.a24xrmdomain.info/" + e.Result.DistinguishedName, username, password);
                    Console.WriteLine(e.Result.DistinguishedName + "-----------------------------");


                    Console.WriteLine();
                    Console.WriteLine("====================================================================================================================================");
                    Console.WriteLine();
                    //Console.ReadKey();
                }
            }
        }
    }
}
