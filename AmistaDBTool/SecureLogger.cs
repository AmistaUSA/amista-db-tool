using System;
using System.IO;

namespace AmistaDBTool
{
    /// <summary>
    /// Secure logging utility that stores logs in a protected location with rotation.
    /// </summary>
    public static class SecureLogger
    {
        private static readonly object _lock = new object();
        private static readonly string LogDirectory;
        private static readonly string DebugLogPath;
        private static readonly string CrashLogPath;

        private const long MaxLogSizeBytes = 5 * 1024 * 1024; // 5 MB
        private const int MaxBackupCount = 3;

        static SecureLogger()
        {
            // Store logs in user's local app data (secure, user-specific location)
            LogDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AmistaDBTool",
                "Logs"
            );

            // Ensure directory exists
            try
            {
                Directory.CreateDirectory(LogDirectory);
            }
            catch
            {
                // Fallback to temp directory if LocalAppData fails
                LogDirectory = Path.Combine(Path.GetTempPath(), "AmistaDBTool", "Logs");
                Directory.CreateDirectory(LogDirectory);
            }

            DebugLogPath = Path.Combine(LogDirectory, "debug.log");
            CrashLogPath = Path.Combine(LogDirectory, "crash.log");
        }

        /// <summary>
        /// Gets the directory where logs are stored.
        /// </summary>
        public static string GetLogDirectory() => LogDirectory;

        /// <summary>
        /// Logs a debug message with timestamp.
        /// </summary>
        public static void LogDebug(string message)
        {
            WriteLog(DebugLogPath, "DEBUG", message);
        }

        /// <summary>
        /// Logs an error message with timestamp.
        /// </summary>
        public static void LogError(string message)
        {
            WriteLog(DebugLogPath, "ERROR", message);
        }

        /// <summary>
        /// Logs a crash with full exception details.
        /// </summary>
        public static void LogCrash(Exception? ex)
        {
            var message = $"Exception: {ex?.Message}\nStack Trace:\n{ex?.StackTrace}";
            WriteLog(CrashLogPath, "CRASH", message);
        }

        /// <summary>
        /// Logs application startup.
        /// </summary>
        public static void LogStartup()
        {
            WriteLog(DebugLogPath, "INFO", "Application starting...");
        }

        private static void WriteLog(string logPath, string level, string message)
        {
            lock (_lock)
            {
                try
                {
                    // Check if rotation is needed
                    RotateLogIfNeeded(logPath);

                    // Write the log entry
                    var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    var logEntry = $"[{timestamp}] [{level}] {message}{Environment.NewLine}";
                    File.AppendAllText(logPath, logEntry);
                }
                catch
                {
                    // Silently fail - logging should never crash the app
                }
            }
        }

        private static void RotateLogIfNeeded(string logPath)
        {
            try
            {
                if (!File.Exists(logPath)) return;

                var fileInfo = new FileInfo(logPath);
                if (fileInfo.Length < MaxLogSizeBytes) return;

                // Rotate existing backups
                for (int i = MaxBackupCount - 1; i >= 1; i--)
                {
                    var oldBackup = $"{logPath}.{i}";
                    var newBackup = $"{logPath}.{i + 1}";

                    if (File.Exists(newBackup))
                        File.Delete(newBackup);

                    if (File.Exists(oldBackup))
                        File.Move(oldBackup, newBackup);
                }

                // Move current log to .1
                var firstBackup = $"{logPath}.1";
                if (File.Exists(firstBackup))
                    File.Delete(firstBackup);

                File.Move(logPath, firstBackup);
            }
            catch
            {
                // Silently fail rotation - continue logging to current file
            }
        }
    }
}
