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
            // TODO: 将异常记录到日志文件
            // File.AppendAllText("error.log", e.Exception.ToString());
            MessageBox.Show("发生了一个未处理的UI异常: \n" + e.Exception.ToString(), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception? ex = e.ExceptionObject as Exception;
            string errorMessage = ex?.ToString() ?? "一个未知的应用程序错误发生了。";
            // TODO: 将异常记录到日志文件
            // File.AppendAllText("error.log", errorMessage);
            MessageBox.Show("发生了一个未处理的应用程序异常: \n" + errorMessage, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}