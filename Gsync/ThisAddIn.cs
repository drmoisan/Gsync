using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Forms;

[assembly: log4net.Config.XmlConfigurator(ConfigFile = "log4net.config", Watch = true)]
[assembly: InternalsVisibleTo("Gsync.Test")]
namespace Gsync
{
    //[ExcludeFromCodeCoverage]
    public partial class ThisAddIn
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private AppGlobals _globals;
        private SynchronizationContext _uiContext;
        private int _uiThreadId;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.Startup += ApplicationStartupAsync;
            SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            _uiContext = SynchronizationContext.Current;
            _uiThreadId = Thread.CurrentThread.ManagedThreadId;
        }

        private async void ApplicationStartupAsync()
        {            
            // Initialize the application globals            
            _globals = await AppGlobals.CreateAsync(Application, _uiContext, _uiThreadId);
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonGsync();
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
