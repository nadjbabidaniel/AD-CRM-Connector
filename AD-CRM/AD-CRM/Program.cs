using System;
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
                DirectoryEntry directoryEntry;

                String ldapPath = "LDAP://XRMSERVER02.a24xrmdomain.info";

                directoryEntry = new DirectoryEntry(ldapPath, @"A24XRMDOMAIN\Danijel.Nadjbabi", "Por4Xae3");

                //// Search AD to see if the user already exists.
                DirectorySearcher search = new DirectorySearcher(directoryEntry);
                search.Filter = "(&(objectClass=user))";

                search.Filter = String.Format("(sAMAccountName={0})", "Gast");
                search.PropertiesToLoad.Add("samaccountname");
                search.PropertiesToLoad.Add("memberOf");
                SearchResult result = search.FindOne();

                if (result != null)
                {
                    // Use the existing AD account.
                    DirectoryEntry userADAccount = result.GetDirectoryEntry();
                    Console.WriteLine(userADAccount.Properties["memberOf"].Value);
                    Console.ReadKey();
                }              
            
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
                Console.WriteLine(e.Message);
                Console.ReadKey();
            }
        }
    }
}
