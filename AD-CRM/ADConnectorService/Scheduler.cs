
using System.ServiceProcess;
using System.Timers;

namespace ADConnectorService
{
  public partial class Scheduler : ServiceBase
  {
    private Timer timer;

    public Scheduler ()
    {
      InitializeComponent ();
    }

    protected override void OnStart (string[] args)
    {
      timer = new Timer ();
     timer.Interval = 30000; //every 30 secs
      timer.Elapsed += timer_Tick;
      timer.Enabled = true;
      ErrorLog.WriteErrorLog ("Test window service started");
    }

    private void timer_Tick (object sender, ElapsedEventArgs e)
    {
      //Write code here to do some job depends on your requirement
      ErrorLog.WriteErrorLog ("Timer ticked and some job has been done successfully");
    }

    protected override void OnStop ()
    {
      timer.Enabled = false;
      ErrorLog.WriteErrorLog ("Test window service stopped");
    }
  }
}
