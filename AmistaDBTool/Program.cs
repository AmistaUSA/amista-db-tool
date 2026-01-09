using System;
using System.IO;
using System.Security;
using System.Windows.Forms;
using Microsoft.Extensions.Configuration;

namespace AmistaDBTool;

static class Program
{
    // Allowed directories for config files (security boundary)
    private static readonly string[] AllowedConfigDirectories = new[]
    {
        AppDomain.CurrentDomain.BaseDirectory,
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        Path.GetTempPath() // Allow temp folder for installer connection tests
    };

    [STAThread]
    static int Main(string[] args)
    {
        // Check for CLI mode
        if (args.Length > 0 && args[0] == "--test-connection")
        {
            string configPath = args.Length > 1 ? args[1] : null;
            string validationError = ValidateConfigPath(configPath);
            if (validationError != null)
            {
                Console.WriteLine($"Error: {validationError}");
                Console.WriteLine("Press Enter to close...");
                Console.ReadLine();
                return 1;
            }
            return RunConnectionTest(configPath);
        }

        AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
        Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

        try
        {
            SecureLogger.LogStartup();
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
            return 0;
        }
        catch (Exception ex)
        {
            LogCrash(ex);
            return 1;
        }
    }

    /// <summary>
    /// Validates the config file path for security.
    /// Returns null if valid, or an error message if invalid.
    /// </summary>
    static string ValidateConfigPath(string configPath)
    {
        if (string.IsNullOrWhiteSpace(configPath))
            return "Configuration file path not provided.";

        // Check for path traversal attempts
        if (configPath.Contains("..") || configPath.Contains("..\\") || configPath.Contains("../"))
            return "Invalid path: path traversal not allowed.";

        // Get full resolved path
        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(configPath);
        }
        catch (Exception)
        {
            return "Invalid path format.";
        }

        // Verify file exists
        if (!File.Exists(fullPath))
            return "Configuration file not found.";

        // Verify extension is .json
        string extension = Path.GetExtension(fullPath).ToLowerInvariant();
        if (extension != ".json")
            return "Configuration file must be a .json file.";

        // Verify the file is within allowed directories
        bool isInAllowedDirectory = false;
        foreach (string allowedDir in AllowedConfigDirectories)
        {
            if (!string.IsNullOrEmpty(allowedDir))
            {
                string normalizedAllowed = Path.GetFullPath(allowedDir).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                string normalizedPath = fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

                if (normalizedPath.StartsWith(normalizedAllowed, StringComparison.OrdinalIgnoreCase))
                {
                    isInAllowedDirectory = true;
                    break;
                }
            }
        }

        if (!isInAllowedDirectory)
        {
            SecureLogger.LogWarning($"Path traversal attempt blocked: {configPath}");
            return "Configuration file must be in the application directory or user data folder.";
        }

        return null; // Valid
    }

    static int RunConnectionTest(string configPath)
    {
        try
        {
            // Path already validated by caller, just load the config
            var builder = new ConfigurationBuilder()
                .AddJsonFile(configPath, optional: false, reloadOnChange: false);
            var config = builder.Build();

            Console.WriteLine("Testing connection...");
            bool success = false;
            
            // Use a specific logger for CLI that writes to Console
            var connector = new SapConnector(config, (msg) => Console.WriteLine(msg));
            
            try 
            {
                var company = connector.GetCompany();
                if (company != null && company.Connected)
                {
                    Console.WriteLine("SUCCESS: Connection established.");
                    success = true;
                    connector.Disconnect();
                }
                else
                {
                   Console.WriteLine("FAILURE: Connected property is false.");
                }
            }
            catch (SapConnectionException ex)
            {
                // SapConnectionException already has sanitized message
                Console.WriteLine($"FAILURE: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Log full details, show generic message
                SecureLogger.LogError($"CLI connection test failed: {ex.Message}");
                Console.WriteLine("FAILURE: Connection test failed. Check logs for details.");
            }

            if (!success)
            {
                Console.WriteLine("Press Enter to close...");
                Console.ReadLine();
                return 1;
            }
            return 0;
        }
        catch (Exception ex)
        {
            SecureLogger.LogError($"CLI critical error: {ex.Message}");
            Console.WriteLine("Critical Error: An unexpected error occurred. Check logs for details.");
            Console.WriteLine("Press Enter to close...");
            Console.ReadLine();
            return 1;
        }
    }

    static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        LogCrash(e.ExceptionObject as Exception);
    }

    static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
    {
        LogCrash(e.Exception);
    }

    static void LogCrash(Exception? ex)
    {
        // Log full details securely
        SecureLogger.LogCrash(ex);
        var logDir = SecureLogger.GetLogDirectory();

        // Show generic message to user - no internal details
        MessageBox.Show($"An unexpected error occurred and the application needs to close.\n\n" +
                        $"Error details have been saved to:\n{logDir}\n\n" +
                        "Please contact support if this problem persists.",
                        "Application Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
    }
}
