using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.Protocols;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace AD_CRM
{
    class Program
    {
        static void Main(string[] args)
        {
            //DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://93.184.191.103/xRMServer03.a24xrmdomain.info", @"A24XRMDOMAIN\Danijel.Nadjbabi", "Por4Xae3", AuthenticationTypes.Secure);

            XmlDocument doc = new XmlDocument();
            string Xml_Path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            Xml_Path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Xml_Path), "XmlData\\" + "ConfigFile.xml");
            doc.Load(Xml_Path);

            XmlNodeList xnList = doc.SelectNodes("/settings/organizations/organization/@adgroup");
            foreach (XmlNode item in xnList)
            {
                Console.WriteLine(item.Value);
            }

            try
            {
                List<DirectoryEntry> allUsersList = new List<DirectoryEntry>();

                //String ldapPath = "LDAP://XRMSERVER02.a24xrmdomain.info/OU=ComData,OU=Extern,OU=A24-UsersAndGroups,DC=a24xrmdomain,DC=info";
                String ldapPath = "LDAP://XRMSERVER02.a24xrmdomain.info";

                DirectoryEntry directoryEntry = new DirectoryEntry(ldapPath, @"A24XRMDOMAIN\Danijel.Nadjbabi", "Por4Xae3");
                //directoryEntry.Path = "LDAP://OU=ComData,OU=Extern,OU=A24-UsersAndGroups,DC=XRMSERVER02,DC=a24xrmdomain,DC=info";
                //directoryEntry = new DirectoryEntry("LDAP://DC=XRMSERVER02, DC=a24xrmdomain, DC=info", @"A24XRMDOMAIN\Danijel.Nadjbabi", "Por4Xae3");

                //// Search AD to see if the user already exists.
                DirectorySearcher search = new DirectorySearcher(directoryEntry);
                search.Filter = "(&(ObjectClass=Group)(|(CN=ComData2)(CN=A24-Member)))";

                //search.Filter = "(&(ObjectClass=user))";
                //SearchResult result = search.FindOne();
                //if (result != null)
                //{
                //    // Use the existing AD account.
                //    DirectoryEntry userADAccount = result.GetDirectoryEntry();

                //    var members = (IEnumerable)search.FindOne().GetDirectoryEntry().Invoke("members");

                //    foreach (object member in members)
                //    {
                //        DirectoryEntry de = new DirectoryEntry(member);
                //        Console.WriteLine(de.Name);
                //        Console.WriteLine(de.Path);
                //    }

                //    Console.WriteLine(userADAccount.Path);
                //    Console.ReadKey();
                //}


                ////////////////SearchResultCollection results = search.FindAll();
                ////////////////for (int i = 0; i < results.Count; i++)
                ////////////////{
                ////////////////    DirectoryEntry de = results[i].GetDirectoryEntry();

                ////////////////    var members = (IEnumerable)de.Invoke("members");// Invoke("members");

                ////////////////    foreach (object member in members)
                ////////////////    {
                ////////////////        DirectoryEntry user = new DirectoryEntry(member);
                ////////////////        allUsersList.Add(user);


                ////////////////        int flags = (int)user.Properties["userAccountControl"].Value;

                ////////////////        var active = !Convert.ToBoolean(flags & 0x0002);

                ////////////////        Console.WriteLine(user.Name);
                ////////////////        Console.WriteLine(active);
                ////////////////        Console.WriteLine(flags);
                ////////////////    }

                ////////////////    Console.WriteLine(de.Name + "-----------------------------");
                ////////////////}
                ////////////////Console.ReadKey();


                String ldapPath2 = "XRMSERVER02.a24xrmdomain.info";
                LdapConnection connection = new LdapConnection(ldapPath2);
                var credentials = new NetworkCredential(@"Danijel.Nadjbabi", "Por4Xae3");             
                connection.Credential = credentials;
                connection.Bind();

                string[] attribs = new string[3];
                attribs[0] = "name";
                attribs[1] = "description";
                attribs[2] = "objectGUID";
                SearchRequest request = new SearchRequest("DC=a24xrmdomain,DC=info", "(CN=Danije Nadjbabi | ComData)", System.DirectoryServices.Protocols.SearchScope.Subtree, attribs);
                SearchResponse searchResponse = (SearchResponse)connection.SendRequest(request);

                foreach (SearchResultEntry entry in searchResponse.Entries)
                {
                    Console.WriteLine(entry.DistinguishedName + "-----------------------------");
                    //Console.ReadKey();
                }

               
                    using (ChangeNotifier notifier = new ChangeNotifier(connection))
                    {
                    //register some objects for notifications (limit 5)
                    //notifier.Register("dc=a24xrmdomain,dc=info", System.DirectoryServices.Protocols.SearchScope.OneLevel);
                    notifier.Register("cn=Danije Nadjbabi | ComData,ou=ComData, ou=Extern,ou=A24-UsersAndGroups,dc=a24xrmdomain,dc=info", System.DirectoryServices.Protocols.SearchScope.Base);

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

            DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://XRMSERVER02.a24xrmdomain.info/" + e.Result.DistinguishedName, @"A24XRMDOMAIN\Danijel.Nadjbabi", "Por4Xae3");
            Console.WriteLine("DirectoryEntry_Name:" + directoryEntry.Name);

            //foreach (var item in directoryEntry.Properties.Values)
            //{
            //    Console.WriteLine(item);
            //}

            //if (directoryEntry.Properties["objectclass"].Equals("user"))
            //{
            int flags = (int)directoryEntry.Properties["userAccountControl"].Value;
            var active = !Convert.ToBoolean(flags & 0x0002);
            Console.WriteLine(active);
            //}

            Console.WriteLine(e.Result.DistinguishedName);
            foreach (string attrib in e.Result.Attributes.AttributeNames)
            {
                foreach (var item in e.Result.Attributes[attrib].GetValues(typeof(string)))
                {
                    Console.WriteLine("\t{0}: {1}", attrib, item);
                }
            }
            Console.WriteLine();
            Console.WriteLine("====================");
            Console.WriteLine();
        }
    }
}
