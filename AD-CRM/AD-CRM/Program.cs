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
        private static String groupsFromXml;
        private static List<DirectoryEntry> allUsersList = new List<DirectoryEntry>();
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

            XmlNodeList xnList = doc.SelectNodes("/settings/organizations/organization/@adgroup");
           
            try
            {
                String ldapPath = "LDAP://" + LdapConn;

                DirectoryEntry directoryEntry = new DirectoryEntry(ldapPath, username, password);
               
                //// Create search for all groups from config file
                DirectorySearcher search = new DirectorySearcher(directoryEntry);
                foreach (XmlNode item in xnList)
                {
                    groupsFromXml += "(CN=" + item.Value +")";
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
                        allUsersList.Add(user);                                                                 
                    }
                }

                //Get list of objectGUIDs from relevant users and populate the list with its values converted to Byte[]
                List<PropertyValueCollection>  ids = allUsersList.Select(x => x.Properties["objectGUID"]).ToList();
                listIdsInByte = new List<byte[]>();
                byte[] byteTest;

                foreach (var item in ids)
                {
                    byteTest = (System.Byte[])item.Value;
                    //int i = BitConverter.ToInt32(byteTest, 0);
                    listIdsInByte.Add(byteTest);
                }

                CrmDataForAd CRM = new CrmDataForAd(username,password);
                CRM.DataForAdSync();

                Entity SystemUserBasedOnId = CRM.SystemUserBasedOnId;







                //Create Ldap connection for notifiers and remove domainName from user for loging
                username = username.Substring(username.IndexOf(@"\")+1);

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
            //if ((e.Result.DistinguishedName).Contains("OU=AD2CRM,OU=ComData,OU=Extern,OU=A24-UsersAndGroups,DC=a24xrmdomain,DC=info"))
            {                               
                object[] objectGuid = e.Result.Attributes["objectguid"].GetValues(typeof(Byte[]));
                byte[] byteTest2 = (System.Byte[])objectGuid[0];
                
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
