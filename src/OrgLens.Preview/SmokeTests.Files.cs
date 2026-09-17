using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OrgLens.Core;
using OrgLens.Desktop;

namespace OrgLens.Preview
{
    internal static partial class SmokeTests
    {
        private static void TestSettingsFiles()
        {
            string directory = Path.Combine(Environment.CurrentDirectory, ".orglens-ui-smoke-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                string first = Path.Combine(directory, "first.json");
                string second = Path.Combine(directory, "second.json");
                var service = new ProbeService { ReturnNoAccounts = true };
                var dialogs = new TestDialogs { SavePath = first };
                using (var form = new SettingsForm(service, dialogs))
                {
                    Show(form);
                    var editor = Find<GroupSettingsControl>(form, "GroupedSettings");
                    Find<NumericUpDown>(form, "BossReadFontSize").Value = 24;
                    Find<NumericUpDown>(form, "BossUnreadFontSize").Value = 27;
                    Find<NumericUpDown>(form, "CustomUnreadFontSize").Value = 16;
                    Find<NumericUpDown>(form, "ToMeFontSize").Value = 13;
                    Find<CheckBox>(form, "ToMeIncludeRead").Checked = true;
                    Find<TextBox>(form, "CustomAddresses").Text = "USER@example.com;user@example.com,\r\nother@example.com";
                    Assert(form.HasUnsavedChanges, "Offline editing must track unsaved settings.");
                    Click(form, "SaveSettingsAs");
                    Assert(File.Exists(first) && form.SelectedFilePath == first && !form.HasUnsavedChanges,
                        "Save as must write the explicitly chosen file and clear its dirty baseline.");
                    Assert(Find<Label>(form, "SettingsFilePath").Text.Contains(first)
                        && Find<Label>(form, "SettingsDocumentState").Text.Contains("preview only"),
                        "The selected path and unapplied state must remain visible after export.");
                    var saved = ConfigurationFile.Load(first);
                    Assert(saved.BossRead.FontSize == 24 && saved.BossUnread.FontSize == 27
                        && saved.CustomUnread.FontSize == 16 && saved.ToMe.FontSize == 13 && !saved.ToMeUnreadOnly
                        && saved.CustomAddresses.SequenceEqual(new[] { "user@example.com", "other@example.com" }),
                        "Every shared style, custom sender and TO ME scope must round-trip through real JSON.");
                    string json = File.ReadAllText(first);
                    Assert(!json.Contains("accountId") && !json.Contains("Jamie") && !json.Contains("sample-avery")
                        && !json.Contains("credentials"), "User settings must not contain discovered people, accounts or credentials.");
                    Find<NumericUpDown>(form, "BossReadFontSize").Value = 25;
                    int chooserCalls = dialogs.SaveCount;
                    Click(form, "SaveSettings");
                    Assert(dialogs.SaveCount == chooserCalls && ConfigurationFile.Load(first).BossRead.FontSize == 25
                        && !form.HasUnsavedChanges, "Save must atomically update the same chosen file without asking for a new path.");
                    Find<NumericUpDown>(form, "BossReadFontSize").Value = 26;
                    string signature = editor.EditSignature;
                    dialogs.SavePath = null;
                    Click(form, "SaveSettingsAs");
                    Assert(form.SelectedFilePath == first && form.HasUnsavedChanges && editor.EditSignature == signature,
                        "Cancelled Save as must not change the document or baseline.");
                    dialogs.SavePath = Path.Combine(directory, "missing-folder", "cannot-save.json");
                    Click(form, "SaveSettingsAs");
                    Assert(form.SelectedFilePath == first && form.HasUnsavedChanges
                        && ConfigurationFile.Load(first).BossRead.FontSize == 25 && dialogs.Errors.Last().Contains("Could not save"),
                        "Failed Save as must preserve previous file contents, path and dirty state.");
                    using (var locked = new FileStream(first, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    {
                        Click(form, "SaveSettings");
                        Assert(form.HasUnsavedChanges && form.SelectedFilePath == first,
                            "Failed save of a locked file must not clear the unsaved warning.");
                    }
                    Assert(!Directory.GetFiles(directory, ".orglens-*.tmp").Any(), "Failed atomic save must clean up its staging file.");

                    var other = new OrgLensConfiguration { CustomAddresses = { "new@example.com" }, ToMeUnreadOnly = true };
                    other.BossRead.FontSize = 33;
                    ConfigurationFile.Save(second, other);
                    dialogs.OpenPath = second;
                    dialogs.ConfirmResult = false;
                    Click(form, "OpenSettings");
                    Assert(form.SelectedFilePath == first && editor.EditSignature == signature && form.HasUnsavedChanges,
                        "Rejected discard prompt must preserve the current document even after choosing a valid file.");
                    dialogs.ConfirmResult = true;
                    dialogs.OpenPath = null;
                    Click(form, "OpenSettings");
                    Assert(form.SelectedFilePath == first && editor.EditSignature == signature && form.HasUnsavedChanges,
                        "Cancelling Open must preserve dirty state.");
                    TestInvalidImports(form, dialogs, directory, json, signature, first);
                    dialogs.OpenPath = second;
                    Click(form, "OpenSettings");
                    Assert(form.SelectedFilePath == second && !form.HasUnsavedChanges
                        && Find<NumericUpDown>(form, "BossReadFontSize").Value == 33
                        && Find<MailPreviewControl>(form, "MessagePreview").Rows[1].Font.SizeInPoints == 33
                        && !Find<CheckBox>(form, "ToMeIncludeRead").Checked,
                        "Successful import must replace all controls, path and preview, without applying.");
                    Assert(service.ApplyCount == 0 && !service.Calls.Any(call => call.StartsWith("apply:")),
                        "Open / Save must never invoke Apply, including in demo mode.");
                    int beforeClose = dialogs.ConfirmationMessages.Count;
                    form.Close();
                    Assert(dialogs.ConfirmationMessages.Count == beforeClose, "Closing a saved/imported document needs no unsaved-file warning.");
                }

                // A brand-new service and form must import from disk, not accidentally pass via in-memory state.
                var freshService = new ProbeService();
                var freshDialogs = new TestDialogs { OpenPath = first };
                using (var fresh = new SettingsForm(freshService, freshDialogs))
                {
                    Show(fresh);
                    Click(fresh, "OpenSettings");
                    var editor = Find<GroupSettingsControl>(fresh, "GroupedSettings");
                    Assert(editor.ReadConfiguration().BossRead.FontSize == 25
                        && editor.ReadConfiguration().CustomAddresses.Count == 2
                        && !editor.ReadConfiguration().ToMeUnreadOnly && fresh.SelectedFilePath == first
                        && freshService.ApplyCount == 0, "Saved settings must round-trip across fresh service/form instances.");
                    Find<NumericUpDown>(fresh, "BossReadFontSize").Value = 29;
                    string signature = editor.EditSignature;
                    freshService.NextApplyError = new OrgLensException("Simulated native rollback.");
                    Click(fresh, "ApplyRules");
                    Assert(fresh.HasUnsavedChanges && fresh.SelectedFilePath == first && editor.EditSignature == signature,
                        "Apply rollback must preserve file dirtiness and selected path.");
                    Click(fresh, "ApplyRules");
                    Assert(fresh.HasUnsavedChanges && freshService.ApplyCount == 1
                        && ConfigurationFile.Load(first).BossRead.FontSize == 25,
                        "Applying must not silently save or clear edits to a selected user file.");
                    Click(fresh, "SaveSettings");
                    Assert(!fresh.HasUnsavedChanges && ConfigurationFile.Load(first).BossRead.FontSize == 29,
                        "Only successful Save should clear an edited user file's warning.");
                    Click(fresh, "RemoveRules");
                    Assert(editor.EditSignature == signature && fresh.SelectedFilePath == first && !fresh.HasUnsavedChanges,
                        "Removing rules must not reset a saved document or mark it edited.");
                    Find<ComboBox>(fresh, "AccountSelector").SelectedIndex = 1;
                    Assert(fresh.SelectedFilePath == null && editor.ReadConfiguration().BossRead.FontSize == 11
                        && freshService.ApplyCount == 1, "Account switch must detach the old file and load the new account without applying.");
                    fresh.Close();
                }
            }
            finally
            {
                Directory.Delete(directory, true);
            }
        }

        private static void TestInvalidImports(SettingsForm form, TestDialogs dialogs, string directory,
            string validJson, string expectedSignature, string expectedPath)
        {
            string[] invalid =
            {
                "{ not json",
                validJson.TrimEnd().Substring(0, validJson.TrimEnd().Length - 1),
                validJson + "{}",
                Regex.Replace(validJson, "\"customAddresses\"\\s*:\\s*\\[[^\\]]*\\]", "\"customAddresses\": {}"),
                validJson.Replace("\"version\": 2", "\"version\": 99"),
                validJson.Replace("\"fontSize\": 24", "\"fontSize\": 128"),
                validJson.Replace("\"fontSize\": 24", "\"fontSize\": 10.5"),
                validJson.Replace("user@example.com", "not-an-address")
            };
            Assert(invalid.All(json => json != validJson),
                "Invalid import fixtures must actually alter the relevant JSON fields.");
            for (int i = 0; i < invalid.Length; i++)
            {
                string path = Path.Combine(directory, "invalid-" + i + ".json");
                File.WriteAllText(path, invalid[i]);
                dialogs.OpenPath = path;
                int errors = dialogs.Errors.Count;
                string preview = Find<MailPreviewControl>(form, "MessagePreview").AccessibleDescription;
                Click(form, "OpenSettings");
                Assert(dialogs.Errors.Count == errors + 1 && form.SelectedFilePath == expectedPath && form.HasUnsavedChanges
                    && Find<GroupSettingsControl>(form, "GroupedSettings").EditSignature == expectedSignature
                    && Find<MailPreviewControl>(form, "MessagePreview").AccessibleDescription == preview,
                    "Invalid JSON/version/font/email import must leave settings, preview, dirty baseline and file path unchanged: " + i);
            }
            dialogs.OpenPath = Path.Combine(directory, "does-not-exist.json");
            int previousErrors = dialogs.Errors.Count;
            Click(form, "OpenSettings");
            Assert(dialogs.Errors.Count == previousErrors + 1 && form.SelectedFilePath == expectedPath && form.HasUnsavedChanges,
                "Missing import file must show a targeted I/O error and leave the current document unchanged.");
        }
    }
}
