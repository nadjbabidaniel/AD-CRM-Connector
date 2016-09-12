using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.Protocols;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace AD_CRM
{
    class Program
    {
        static void Main(string[] args)
        {
            //DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://93.184.191.103/xRMServer03.a24xrmdomain.info", @"A24XRMDOMAIN\Danijel.Nadjbabi", "Por4Xae3", AuthenticationTypes.Secure);

            try
            {
                //var directoryEntry = new DirectoryEntry("LDAP://93.184.191.103/CN=Users;DC=A24XRMDOMAIN");
                //directoryEntry.Username = @"A24XRMDOMAIN\Danijel.Nadjbabi";
                //directoryEntry.Password = "Por4Xae3";
                //var test = directoryEntry.Properties;
                //DirectorySearcher ds = new DirectorySearcher(directoryEntry);
                //ds.SearchScope = SearchScope.Subtree;
                //ds.Filter = "(&(objectClass=group))";
                //var objSearchResults = ds.FindAll();
                //Console.WriteLine("Hello World: {0}", objSearchResults.ToString());
                //Console.ReadKey();


                DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://a24xrmdomain.info");
                directoryEntry.Path = "LDAP://OU=ComData,OU=Extern,OU=A24-UsersAndGroups,DC=a24xrmdomain,DC=info";
                Console.WriteLine("Hello World: {0}", directoryEntry.Children.ToString());
                Console.WriteLine("Hello World: {0}", directoryEntry.Guid.ToString());
                Console.WriteLine("Hello World: {0}", directoryEntry.Name.ToString());

                Console.WriteLine();

                using (DirectorySearcher ds = new DirectorySearcher(directoryEntry))
                {
                    ds.PropertiesToLoad.Add("name");
                    ds.PropertiesToLoad.Add("userPrincipalName");

                    ds.Filter = "(&(objectClass=user))";

                    SearchResultCollection results = ds.FindAll();
                    Console.WriteLine(results.Count);

                    foreach (SearchResult result in results)
                    {
                        Console.WriteLine("{0} - {1}",
                            result.Properties["name"][0].ToString(),
                            result.Properties["userPrincipalName"][0].ToString());
                    }
                }


                Console.ReadKey();


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
