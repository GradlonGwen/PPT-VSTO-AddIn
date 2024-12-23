using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace PowerPointAddIn1
{
    public partial class TaskPaneWinForm : UserControl
    {
        public TaskPaneWinForm()
        {
            InitializeComponent();

            var wpfHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                Margin = new Padding(0)
            };
            wpfHost.Child = new HostedWpfTaskPane();
            Controls.Add(wpfHost);
        }
    }
}
