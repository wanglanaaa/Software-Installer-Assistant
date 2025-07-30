using System.Windows.Forms;

namespace SoftwareInstaller.Utils
{
    public static class ErrorHandler
    {
        public static void ShowError(string message)
        {
            MessageBox.Show(message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
