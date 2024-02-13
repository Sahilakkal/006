using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelHierarchyConversion_InterOp
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (!IsRunAsAdmin())
            {
                // Restart the application with administrative privileges
                ElevatePermissions();
                //MessageBox.Show("Admin : true");
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ExcelHierarchyCon());
        }

        static bool IsRunAsAdmin()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);

            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        static void ElevatePermissions()
        {
            // Get the path to the current executable
            string executablePath = Process.GetCurrentProcess().MainModule.FileName;

            // Start a new process with administrative rights
            ProcessStartInfo startInfo = new ProcessStartInfo(executablePath)
            {
                Verb = "runas", // Run as administrator
                UseShellExecute = true
            };

            try
            {
                Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error restarting application with elevated permissions: {ex.Message}");
            }
        }
    }
}
