using Gsync.Ribbon;
using Gsync.Utilities.Interfaces;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Gsync
{
    [ComVisible(true)]
    public class RibbonGsync : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public RibbonGsync() { }

        protected IApplicationGlobals _globals;
        internal IApplicationGlobals Globals 
        { 
            get => _globals;
            set 
            { 
                _globals = value; 
                Dev = new DevelopmentMethods(_globals);
            }
        }

        internal DevelopmentMethods Dev { get; set; }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Outlook.Explorer")
            {
                return GetResourceText("Gsync.Ribbon.RibbonGsync.xml");
            }
            return null;
        }

        #endregion

        #region Ribbon Callbacks
        
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #region Development Methods

        public void LoopInboxItems_Click(Office.IRibbonControl control) => Dev.LoopInbox();

        #endregion Development Methods

        #endregion Ribbon Callbacks

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion Helpers
    }
}
