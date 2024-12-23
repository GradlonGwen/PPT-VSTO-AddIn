using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;

namespace PowerPointAddIn1
{
    [ComVisible(true)]
    public class RibbonExtensibility : IRibbonExtensibility
    {
        public string GetCustomUI(string RibbonID)
        {
            return ribbon;
        }

        public void OnLoad(IRibbonUI ribbonUi)
        {
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            return null;
        }

        public void OnButtonClicked(IRibbonControl control)
        {
            DpiHelper.SetThreadDpiAwareness(DpiAwarenessContext.PerMonitorAware);
            new Window1().ShowDialog();
        }

        public void OnTaskPaneClicked(IRibbonControl control)
        {
            // SetThreadDpiAwareness has no effect here
            //DpiHelper.SetThreadDpiAwareness(DpiAwarenessContext.PerMonitorAware);
            TaskPane.Visible = true;
        }

        private string ribbon = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI onLoad=""OnLoad"" xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" xmlns:p=""PowerPointAddIn1"">
	<ribbon startFromScratch=""false"">
		<tabs>
			<tab label=""VSTO-test"" id=""Power-user"" keytip=""z"">
<group autoScale=""true"" label=""Group1"" id=""groupHelp"">
<button enabled=""true"" getImage=""GetImage"" label=""open window"" id=""btnOpenWindow"" onAction=""OnButtonClicked""/>
<button enabled=""true"" getImage=""GetImage"" label=""open task pane"" id=""btnOpenTaskPane"" onAction=""OnTaskPaneClicked""/>
</group>
</tab>
</tabs>
</ribbon>
</customUI>
";

        public CustomTaskPane TaskPane { get; set; }
    }

    
}
