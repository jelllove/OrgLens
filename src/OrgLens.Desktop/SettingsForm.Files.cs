using System.IO;
using OrgLens.Core;

namespace OrgLens.Desktop
{
    public sealed partial class SettingsForm
    {
        private void OpenSettings()
        {
            if (busy) return;
            RunOperation("Opening settings…", "Could not open settings", () =>
            {
                string selected = dialogs.ChooseOpenFile(this, filePath);
                if (selected == null)
                {
                    SetStatus("Open canceled · Current settings kept.", false);
                    return;
                }
                string fullPath = Path.GetFullPath(selected);
                // Read and validate before touching controls, account state, dirty baseline or the selected path.
                var configuration = ConfigurationFile.Load(fullPath);
                if (!ConfirmDiscard("Open a different settings file"))
                {
                    SetStatus("Open canceled · Unsaved settings kept.", false);
                    return;
                }
                BindDocument(configuration, fullPath);
                SetStatus("Settings loaded · Preview only. Click Apply explicitly to change the selected Inbox.", false);
            });
        }

        private void SaveSettings(bool saveAs)
        {
            if (busy) return;
            RunOperation("Saving settings…", "Could not save settings", () =>
            {
                var configuration = editor.ReadConfiguration();
                string selected = saveAs || filePath == null ? dialogs.ChooseSaveFile(this, filePath) : filePath;
                if (selected == null)
                {
                    SetStatus("Save canceled · Current settings kept.", false);
                    return;
                }
                string fullPath = Path.GetFullPath(selected);
                ConfigurationFile.Save(fullPath, configuration);
                filePath = fullPath;
                cleanSignature = editor.EditSignature;
                SetStatus("Settings file saved · Outlook rules unchanged. Apply is a separate action.", false);
            });
        }

        private void BindDocument(OrgLensConfiguration configuration, string path)
        {
            editor.SetConfiguration(configuration);
            filePath = path;
            cleanSignature = editor.EditSignature;
            appliedSignature = null;
            UpdatePreview();
            UpdateActions();
        }

        private bool ConfirmDiscard(string action)
        {
            return !HasUnsavedChanges || dialogs.Confirm(this,
                action + " and discard unsaved settings edits?\r\n\r\nChoose No to keep editing, then Save or Save as. "
                + "Saving a file never applies rules to Outlook.", "Unsaved OrgLens settings");
        }
    }
}
