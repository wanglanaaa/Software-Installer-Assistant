using System;
using System.Windows.Forms;
using System.Threading;

namespace SoftwareInstaller.UI
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Add global exception handlers
            Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        private static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            // 记录异常，然后显示它。
            MessageBox.Show("Unhandled UI Exception: " + e.Exception.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            // 记录异常，然后显示它。
            MessageBox.Show("Unhandled Application Exception: " + (e.ExceptionObject as Exception)?.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}