
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace HealthcareSanatoriumInterface
{
    internal static class LauncherProgram
    {
        [STAThread]
        private static void Main(string[] args)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string hostPath = ResolveHost(baseDir);

            if (string.IsNullOrWhiteSpace(hostPath) || !File.Exists(hostPath))
            {
                MessageBox.Show(
                    "Не найден исполняемый файл приложения.\nПроверьте установку пакета.",
                    "Запуск",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                Environment.Exit(1);
                return;
            }

            int exitCode = RunHost(hostPath, args);
            Environment.Exit(exitCode);
        }

        private static string ResolveHost(string baseDir)
        {
            string x64 = Path.Combine(baseDir, "HealthcareSanatoriumInterface.x64.exe");
            string x86 = Path.Combine(baseDir, "HealthcareSanatoriumInterface.x86.exe");

            bool hasX64 = File.Exists(x64);
            bool hasX86 = File.Exists(x86);
            bool provider64 = Environment.Is64BitOperatingSystem && HasAnyProvider(RegistryView.Registry64);
            bool provider32 = HasAnyProvider(RegistryView.Registry32);
            bool access64 = Environment.Is64BitOperatingSystem && IsAccessInstalled(RegistryView.Registry64);
            bool access32 = IsAccessInstalled(RegistryView.Registry32);

            if (hasX64 && provider64 && access64)
            {
                return x64;
            }

            if (hasX86 && provider32 && access32)
            {
                return x86;
            }

            if (hasX64 && provider64)
            {
                return x64;
            }

            if (hasX86 && provider32)
            {
                return x86;
            }

            if (hasX64 && access64)
            {
                return x64;
            }

            if (hasX86 && access32)
            {
                return x86;
            }

            if (hasX64)
            {
                return x64;
            }

            if (hasX86)
            {
                return x86;
            }

            return null;
        }

        private static bool HasAnyProvider(RegistryView view)
        {
            return IsProviderInstalled("Microsoft.ACE.OLEDB.16.0", view) ||
                   IsProviderInstalled("Microsoft.ACE.OLEDB.12.0", view) ||
                   IsProviderInstalled("Microsoft.Jet.OLEDB.4.0", view);
        }

        private static bool IsAccessInstalled(RegistryView view)
        {
            try
            {
                using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, view))
                {
                    if (root.OpenSubKey(@"Access.Application\CLSID") != null) return true;
                    if (root.OpenSubKey(@"Access.Application.16\CLSID") != null) return true;
                    if (root.OpenSubKey(@"Access.Application.15\CLSID") != null) return true;
                    if (root.OpenSubKey(@"Access.Application.14\CLSID") != null) return true;
                }

                using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                {
                    using (RegistryKey key = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE"))
                    {
                        if (key != null && !string.IsNullOrWhiteSpace(Convert.ToString(key.GetValue(null)))) return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool IsProviderInstalled(string providerName, RegistryView view)
        {
            try
            {
                using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, view))
                using (RegistryKey key = root.OpenSubKey(providerName))
                {
                    return key != null;
                }
            }
            catch
            {
                return false;
            }
        }

        private static int RunHost(string hostPath, string[] args)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = hostPath;
            startInfo.WorkingDirectory = Path.GetDirectoryName(hostPath);
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = false;
            startInfo.Arguments = BuildArguments(args);
            startInfo.WindowStyle = ProcessWindowStyle.Normal;

            using (Process process = Process.Start(startInfo))
            {
                if (process == null)
                {
                    return 1;
                }

                process.WaitForExit();
                return process.ExitCode;
            }
        }

        private static string BuildArguments(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                return string.Empty;
            }

            List<string> parts = new List<string>();
            for (int i = 0; i < args.Length; i++)
            {
                parts.Add(QuoteArgument(args[i]));
            }

            return string.Join(" ", parts.ToArray());
        }

        private static string QuoteArgument(string argument)
        {
            if (string.IsNullOrEmpty(argument))
            {
                return "\"\"";
            }

            if (argument.IndexOfAny(new[] { ' ', '\t', '"' }) < 0)
            {
                return argument;
            }

            return "\"" + argument.Replace("\"", "\\\"") + "\"";
        }
    }
}
