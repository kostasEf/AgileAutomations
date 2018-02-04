using System.Windows.Forms;
using AgileAutomations.Interface;

namespace AgileAutomations.Helper
{
    public class FileBrowser : IFileBrowser
    {
        public string GetFullName()
        {
            var openFileDialog = new OpenFileDialog
            {
                InitialDirectory = "c:\\",
                Filter = "excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true
            };

            return openFileDialog.ShowDialog() == DialogResult.OK ? openFileDialog.FileName : string.Empty;
        }
    }
}