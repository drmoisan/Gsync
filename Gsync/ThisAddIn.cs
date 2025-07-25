using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Forms;
using log4net;
using log4net.Repository;
using log4net.Appender;
using System.Linq;


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
        private RibbonGsync _ribbon;

        static ThisAddIn() 
        { 
            var now = DateTime.Now;
            var logsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Gsync",
                "Logs"
            );
            Directory.CreateDirectory(logsPath);
            
            // Get the log4net repository
            ILoggerRepository repo = LogManager.GetRepository();

            // Set the file path for all_logs_file
            var allLogsAppender = repo.GetAppenders()
                .OfType<RollingFileAppender>()
                .FirstOrDefault(a => a.Name == "all_logs_file");
            if (allLogsAppender != null)
            {
                allLogsAppender.File = Path.Combine(logsPath, $"debug {now:yyyy-MM-dd-HH-mm}.log");
                allLogsAppender.ActivateOptions();
            }

            // Set the file path for method_calls_log_file
            var methodCallsAppender = repo.GetAppenders()
                .OfType<RollingFileAppender>()
                .FirstOrDefault(a => a.Name == "method_calls_log_file");
            if (methodCallsAppender != null)
            {
                methodCallsAppender.File = Path.Combine(logsPath, $"trace {now:yyyy-MM-dd-HH-mm}.log");
                methodCallsAppender.ActivateOptions();
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            logger.Debug($"{nameof(ThisAddIn_Startup)} called. Thread ID is: {Thread.CurrentThread.ManagedThreadId}");
            Application.Startup += ApplicationStartupAsync;
            SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            _uiContext = SynchronizationContext.Current;
            _uiThreadId = Thread.CurrentThread.ManagedThreadId;
            logger.Debug($"UI SynchronizationContext set. Thread ID: {_uiThreadId}");
        }

        private async void ApplicationStartupAsync()
        {            
            logger.Debug($"{nameof(ApplicationStartupAsync)} called.");
            // Initialize the application globals            
            _globals = await AppGlobals.CreateAsync(Application, _uiContext, _uiThreadId);
            _ribbon.Globals = _globals;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {            
            logger.Debug($"{nameof(CreateRibbonExtensibilityObject)} called.  Thread ID is: {Thread.CurrentThread.ManagedThreadId}");
            _ribbon = new RibbonGsync();
            return _ribbon;
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
