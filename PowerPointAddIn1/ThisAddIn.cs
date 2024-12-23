using System.Windows;
using Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        private RibbonExtensibility _ribbon;
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new RibbonExtensibility();
            return _ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var taskPane = CustomTaskPanes.Add(new TaskPaneWinForm(), "Task pane");
            taskPane.Width = 400;
            taskPane.Visible = false;
            _ribbon.TaskPane = taskPane;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public override void BeginInit()
        {
            DpiHelper.SetThreadDpiAwareness(DpiAwarenessContext.PerMonitorAware);
            _ = new Application
            {
                ShutdownMode = ShutdownMode.OnExplicitShutdown
            };
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
