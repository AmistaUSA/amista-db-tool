namespace AmistaDBTool;

static class Program
{
    [STAThread]
    static void Main()
    {
        AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
        Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

        try
        {
            SecureLogger.LogStartup();
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
        catch (Exception ex)
        {
            LogCrash(ex);
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
        SecureLogger.LogCrash(ex);
        var logDir = SecureLogger.GetLogDirectory();
        MessageBox.Show($"Application crashed. Error: {ex?.Message}\n\nLogs saved to: {logDir}",
                        "Application Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
    }
}
