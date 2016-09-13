using System;
using System.IO;

namespace ADConnectorService
{
  public static class ErrorLog
  {
    public static void WriteErrorLog (Exception ex)
    {
      StreamWriter streamWriter;
      try
      {
        streamWriter = new StreamWriter (AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
        streamWriter.WriteLine (DateTime.Now + ": " + ex.Source.Trim () + "; " + ex.Message.Trim ());
        streamWriter.Flush ();
        streamWriter.Close ();
      }
      catch
      {
      }
    }

    public static void WriteErrorLog (string message)
    {
      StreamWriter streamWriter;
      try
      {
        streamWriter = new StreamWriter (AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
        streamWriter.WriteLine (DateTime.Now + ": " + message);
        streamWriter.Flush ();
        streamWriter.Close ();
      }
      catch
      {
      }
    }
  }
}