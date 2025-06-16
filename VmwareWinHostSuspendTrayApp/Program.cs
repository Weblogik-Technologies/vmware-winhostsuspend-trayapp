
using System;
using System.Windows.Forms;
using Microsoft.Win32;

namespace LogikVmwareWinHostSuspendTrayApp
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayAppContext());
        }
    }
}
