using System;
using System.DirectoryServices;

namespace AD_CRM
{
  internal class Program
  {
    private static void Main (string[] args)
    {
      DirectoryEntry directoryEntry;

      String ldapPath = "LDAP://XRMSERVER02.a24xrmdomain.info";

      directoryEntry = new DirectoryEntry (ldapPath, @"A24XRMDOMAIN\Andrea.Borbelj"
      , "Nue5Gbe9");

      //// Search AD to see if the user already exists.
      DirectorySearcher search = new DirectorySearcher (directoryEntry);
      //search.SearchScope = SearchScope.Subtree;
      search.Filter = "(&(objectClass=user))";

      search.Filter = String.Format ("(sAMAccountName={0})", "Gast");
      search.PropertiesToLoad.Add ("samaccountname");
      search.PropertiesToLoad.Add ("memberOf");
      SearchResult result = search.FindOne ();

      if (result != null)
      {
        // Use the existing AD account.
        DirectoryEntry userADAccount = result.GetDirectoryEntry ();
        Console.WriteLine (userADAccount.Properties["memberOf"].Value);
      }
    }
  }
}