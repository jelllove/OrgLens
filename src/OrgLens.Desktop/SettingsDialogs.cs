using System.Runtime.CompilerServices;
using System.Windows.Forms;

[assembly: InternalsVisibleTo("OrgLens.Preview")]

namespace OrgLens.Desktop
{
    internal interface ISettingsDialogs
    {
        bool Confirm(IWin32Window owner, string message, string title);
        void Information(IWin32Window owner, string message, string title);
        void Error(IWin32Window owner, string message, string title);
        string ChooseOpenFile(IWin32Window owner, string currentPath);
        string ChooseSaveFile(IWin32Window owner, string currentPath);
    }

    internal sealed class SettingsDialogs : ISettingsDialogs
    {
        public string ChooseOpenFile(IWin32Window owner, string currentPath)
        {
            using (var dialog = new OpenFileDialog
            {
                Title = "Open OrgLens settings · Preview only",
                Filter = "OrgLens settings (*.json)|*.json|All files (*.*)|*.*",
                CheckFileExists = true, Multiselect = false, RestoreDirectory = true,
                FileName = currentPath ?? ""
            })
                return dialog.ShowDialog(owner) == DialogResult.OK ? dialog.FileName : null;
        }

        public string ChooseSaveFile(IWin32Window owner, string currentPath)
        {
            using (var dialog = new SaveFileDialog
            {
                Title = "Save OrgLens settings · Does not apply to Outlook",
                Filter = "OrgLens settings (*.json)|*.json", DefaultExt = "json", AddExtension = true,
                OverwritePrompt = true, RestoreDirectory = true, FileName = currentPath ?? "OrgLens-settings.json"
            })
                return dialog.ShowDialog(owner) == DialogResult.OK ? dialog.FileName : null;
        }

        public bool Confirm(IWin32Window owner, string message, string title)
        {
            return MessageBox.Show(owner, message, title, MessageBoxButtons.YesNo,
                MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes;
        }

        public void Information(IWin32Window owner, string message, string title)
        {
            MessageBox.Show(owner, message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void Error(IWin32Window owner, string message, string title)
        {
            MessageBox.Show(owner, message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
