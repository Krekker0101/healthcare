using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace HealthcareSanatoriumInterface
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string dbPath = ResolveDatabasePath(baseDir);
            string assemblyPath = Path.Combine(baseDir, "HealthcareModules.dll");

            if (args.Length > 0)
            {
                if (args[0] == "--self-test")
                {
                    RunWithExitCode(() => RunSelfTest(assemblyPath, dbPath));
                    return;
                }

                if (args[0] == "--pdf-self-test")
                {
                    RunWithExitCode(() => RunPdfSelfTest(assemblyPath, dbPath));
                    return;
                }

                if (args[0] == "--generate-manual-pdf")
                {
                    string target = args.Length > 1 ? args[1] : Path.Combine(baseDir, "UserGuide.pdf");
                    RunWithExitCode(() => GenerateManualPdf(assemblyPath, target));
                    return;
                }

                if (args[0].StartsWith("--generate-manual-pdf=", StringComparison.OrdinalIgnoreCase))
                {
                    string target = args[0].Substring("--generate-manual-pdf=".Length);
                    RunWithExitCode(() => GenerateManualPdf(assemblyPath, target));
                    return;
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (Form form = CreateMainForm(assemblyPath, dbPath))
            {
                Application.Run(form);
            }
        }

        private static int RunSelfTest(string assemblyPath, string dbPath)
        {
            try
            {
                int result = InvokeIntMethod(assemblyPath, "PluginBootstrap", "RunSelfTest", dbPath);
                TryDeleteFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "self-test-error.txt"));
                return result;
            }
            catch (Exception ex)
            {
                File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "self-test-error.txt"), ex.ToString());
                return 1;
            }
        }

        private static int RunPdfSelfTest(string assemblyPath, string dbPath)
        {
            try
            {
                int result = InvokeIntMethod(assemblyPath, "PluginBootstrap", "RunPdfSelfTest", dbPath);
                TryDeleteFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pdf-self-test-error.txt"));
                return result;
            }
            catch (Exception ex)
            {
                File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pdf-self-test-error.txt"), ex.ToString());
                return 1;
            }
        }

        private static void GenerateManualPdf(string assemblyPath, string targetPath)
        {
            InvokeVoidMethod(assemblyPath, "PluginBootstrap", "GenerateManualPdf", targetPath);
        }

        private static Form CreateMainForm(string assemblyPath, string dbPath)
        {
            return (Form)InvokeMethod(assemblyPath, "PluginBootstrap", "CreateMainForm", dbPath);
        }

        private static string ResolveDatabasePath(string baseDir)
        {
            string[] candidates = new[]
            {
                Path.Combine(baseDir, "HealthcareSanatoriumSystem.accdb"),
                Path.Combine(baseDir, "HealthcareSanatoriumSystem_FIXED.accdb"),
                Path.Combine(baseDir, "HealthcareSanatoriumSystem.mdb"),
                Path.Combine(baseDir, "db.accdb"),
                Path.Combine(baseDir, "assest", "db.accdb")
            };

            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return Path.Combine(baseDir, "HealthcareSanatoriumSystem.accdb");
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        private static void RunWithExitCode(Func<int> action)
        {
            Environment.Exit(action());
        }

        private static void RunWithExitCode(Action action)
        {
            action();
            Environment.Exit(0);
        }

        private static int InvokeIntMethod(string assemblyPath, string typeName, string methodName, params object[] args)
        {
            return (int)InvokeMethod(assemblyPath, typeName, methodName, args);
        }

        private static void InvokeVoidMethod(string assemblyPath, string typeName, string methodName, params object[] args)
        {
            InvokeMethod(assemblyPath, typeName, methodName, args);
        }

        private static object InvokeMethod(string assemblyPath, string typeName, string methodName, params object[] args)
        {
            Assembly assembly = Assembly.LoadFrom(assemblyPath);
            Type pluginType = assembly.GetType("HealthcareSanatoriumInterface." + typeName, throwOnError: true);
            MethodInfo method = pluginType.GetMethod(methodName, BindingFlags.Public | BindingFlags.Static);
            if (method == null)
            {
                throw new MissingMethodException(pluginType.FullName, methodName);
            }

            return method.Invoke(null, args);
        }
    }
}
