using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1
{
    public class DpiHelper
    {
        [DllImport("User32.dll")]
        private static extern DpiAwarenessContext SetThreadDpiAwarenessContext(DpiAwarenessContext dpiAwarenessContext);
        
        // A user having the version 10.0.10240 was having the issue:
        // Unable to find an entry point named 'SetThreadDpiAwarenessContext' in DLL 'User32.dll'.
        private static Version ThreadDpiAwarenessMinimalVersion => Version.Parse("10.0.10241");
        private static Version ThreadDpiAwarenessPerMonitorV2MinimalVersion => Version.Parse("10.0.15063");
        public static void SetThreadDpiAwareness(DpiAwarenessContext dpiAwarenessContext)
        {
            Version osVersion = Environment.OSVersion.Version;

            if (dpiAwarenessContext == DpiAwarenessContext.PerMonitorAwareV2 &&
                osVersion.CompareTo(ThreadDpiAwarenessMinimalVersion) < 0)
            {
                // V2 is preferred, but if the OS does not support it, fallback to v1 version.
                dpiAwarenessContext = DpiAwarenessContext.PerMonitorAware;
            }

            if (osVersion.CompareTo(ThreadDpiAwarenessMinimalVersion) >= 0)
            {
                SetThreadDpiAwarenessContext(dpiAwarenessContext);
            }
        }
    }
}
