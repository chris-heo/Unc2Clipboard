using System;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Security.Principal;
using System.Diagnostics;
using Microsoft.Win32;

namespace Unc2Clipboard
{
    class Program
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;

        // https://gist.github.com/ambyte/01664dc7ee576f69042c
        [DllImport("mpr.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int WNetGetConnection(
            [MarshalAs(UnmanagedType.LPTStr)] string localName,
            [MarshalAs(UnmanagedType.LPTStr)] StringBuilder remoteName,
            ref int length);
        /// <summary>
        /// Given a path, returns the UNC path or the original. (No exceptions
        /// are raised by this function directly). For example, "P:\2008-02-29"
        /// might return: "\\networkserver\Shares\Photos\2008-02-09"
        /// </summary>
        /// <param name="originalPath">The path to convert to a UNC Path</param>
        /// <returns>A UNC path. If a network drive letter is specified, the
        /// drive letter is converted to a UNC or network path. If the 
        /// originalPath cannot be converted, it is returned unchanged.</returns>
        public static string GetUNCPath(string originalPath)
        {
            StringBuilder sb = new StringBuilder(512);
            int size = sb.Capacity;

            // look for the {LETTER}: combination ...
            if (originalPath.Length > 2 && originalPath[1] == ':')
            {
                // don't use char.IsLetter here - as that can be misleading
                // the only valid drive letters are a-z && A-Z.
                char c = originalPath[0];
                if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
                {
                    int error = WNetGetConnection(originalPath.Substring(0, 2), sb, ref size);
                    if (error == 0)
                    {
                        string path = Path.GetFullPath(originalPath)
                            .Substring(Path.GetPathRoot(originalPath).Length);
                        return Path.Combine(sb.ToString().TrimEnd(), path);
                    }
                }
            }

            return originalPath;
        }


        // https://stackoverflow.com/questions/133379/elevating-process-privilege-programmatically
        private static bool IsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static bool ensureAdmin(string arguments)
        {
            if (IsAdministrator())
                return true;
            
            // Restart program and run as admin
            string exeName = Process.GetCurrentProcess().MainModule.FileName;
            ProcessStartInfo startInfo = new ProcessStartInfo(exeName, arguments);
            startInfo.Verb = "runas";
            Process.Start(startInfo);
            return false;
        }

        static bool install(string target, string appkey, string commandtext, string command)
        {
            RegistryKey keymenu = null;
            RegistryKey keycmd = null;
            try
            {
                keymenu = Registry.ClassesRoot.CreateSubKey(target + "\\shell\\" + appkey);
                keymenu.SetValue("", commandtext);

                keycmd = Registry.ClassesRoot.CreateSubKey(target + "\\shell\\" + appkey + "\\command");
                keycmd.SetValue("", command);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error when accessing the registry:");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (keymenu != null)
                    keymenu.Close();
                if (keycmd != null)
                    keycmd.Close();
            }
            
            return true;
        }

        static void uninstall(string target, string appkey)
        {
            Registry.ClassesRoot.DeleteSubKeyTree(target + "\\shell\\" + appkey);
        }

        [STAThread]
        static void Main(string[] args)
        {
            string commandtext = "Copy as UNC path";

            if (args.Length != 1)
            {
                string exename = AppDomain.CurrentDomain.FriendlyName;
                Console.WriteLine("{0} - Copy a file path as UNC path into the clipboard via Explorer's context menu", exename);
                Console.WriteLine("Usage:");
                Console.WriteLine("  Put this executable in the folder of your liking (should be accessible for all users)");
                Console.WriteLine("  Install with argument --install, uninstall with --uninstall");
                Console.WriteLine("  You will have to confirm a UAC request in order to write the keys into the registry");
                Console.WriteLine("  After that, you're set. Just right click on a file or folder and select \"{0}\"", commandtext);
                Console.ReadKey(); // wait for the user
                return;
            }

            string appkey = "CopyAsUnc";

            // installation targets for the registry
            string[] targets = new string[] { "*", "Directory" };

            if (args[0].ToUpper() == "--INSTALL")
            {
                if (ensureAdmin(args[0]) == false)
                {
                    Console.WriteLine("The application must be running as Administrator for (un-)installing.");
                    return;
                }

                string cmd = string.Format("\"{0}\" \"%1\"", Process.GetCurrentProcess().MainModule.FileName);

                foreach (string target in targets)
                {
                    install(target, appkey, commandtext, cmd);
                }
                
                Console.WriteLine("installed.");
                return;
            }
            else if(args[0].ToUpper() == "--UNINSTALL")
            {
                if (ensureAdmin(args[0]) == false)
                {
                    Console.WriteLine("The application must be running as Administrator for (un-)installing.");
                    return;
                }

                foreach (string target in targets)
                {
                    uninstall(target, appkey);
                }
            }
            else
            {
                // prevent the console window from showing
                var handle = GetConsoleWindow();
                ShowWindow(handle, SW_HIDE);
                Clipboard.SetText(GetUNCPath(args[0]));
            }
        }
    }
}
