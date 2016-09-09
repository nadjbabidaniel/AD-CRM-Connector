using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AD_CRM
{
    class Program
    {
        static void Main(string[] args)
        {
            //DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://93.184.191.103", @"A24XRMDOMAIN\Danijel.Nadjbabi", "Por4Xae3");

            try
            {               
                var directoryEntry = new DirectoryEntry("LDAP://93.184.191.103/CN=Users;DC=A24XRMDOMAIN");
                directoryEntry.Username = @"A24XRMDOMAIN\Danijel.Nadjbabi";
                directoryEntry.Password = "Por4Xae3";
                //var test = directoryEntry.Properties;

                DirectorySearcher ds = new DirectorySearcher(directoryEntry);
                ds.SearchScope = SearchScope.Subtree;
                ds.Filter = "(&(objectClass=group))";

                var objSearchResults = ds.FindAll();

                Console.WriteLine("Hello World: {0}", objSearchResults .ToString());
                Console.ReadKey();


            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
                Console.WriteLine("33333 World");
                Console.ReadKey();
            }
        }
    }
}
