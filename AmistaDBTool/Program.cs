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
            File.AppendAllText("debug.log", $"[{DateTime.Now}] Application Starting...\n");
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
        string message = $"CRASH at {DateTime.Now}: {ex?.Message}\n{ex?.StackTrace}\n\n";
        File.AppendAllText("crash_log.txt", message);
        MessageBox.Show($"Application Crashed! see crash_log.txt. Error: {ex?.Message}");
    }
}
