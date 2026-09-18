using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Security;
using System.Windows.Forms;
using OrgLens.Core;

namespace OrgLens.Desktop
{
    public sealed partial class SettingsForm
    {
        private void LoadAccounts()
        {
            bool success = RunOperation("Loading accounts…", "Could not load accounts", () =>
            {
                var accounts = service.GetAccounts();
                if (IsDisposed || Disposing) return;
                suppressEvents = true;
                try
                {
                    accountSelector.Items.Clear();
                    foreach (var account in accounts) accountSelector.Items.Add(account);
                    accountSelector.SelectedIndex = -1;
                }
                finally { suppressEvents = false; }
                if (accounts.Count == 0)
                    SetNotice("No supported accounts. Open, edit and save settings here; Refresh to retry Outlook.");
                SetStatus(accounts.Count == 0 ? "Settings editor ready · No account available for Apply." : "Accounts loaded.", false);
            }, showLoading: true);
            if (IsDisposed || Disposing) return;
            if (!success)
                SetNotice("Accounts unavailable. Settings editing / files still work. Refresh to retry.");
            else if (accountSelector.Items.Count > 0)
                accountSelector.SelectedIndex = 0;
        }

        private void AccountChanged(object sender, EventArgs e)
        {
            if (suppressEvents || busy) return;
            var selected = accountSelector.SelectedItem as MailboxAccount;
            if (selected == activeAccount) return;
            if (!ConfirmDiscard("Switch accounts"))
            {
                RestoreAccountSelection();
                return;
            }
            if (selected == null) return;
            OrgLensConfiguration loaded = null;
            bool success = RunOperation("Loading this account's settings…", "Could not load account settings",
                () => loaded = ConfigurationValidator.Normalize(service.LoadConfiguration(selected.Id)), showLoading: true);
            if (IsDisposed || Disposing) return;
            if (!success && activeAccount != null)
            {
                RestoreAccountSelection();
                return;
            }
            activeAccount = selected;
            managers = null;
            accountConfigLoaded = success;
            if (success) BindDocument(loaded, null);
            else
            {
                // Initial failure leaves the independent document editable, not a stale account configuration.
                appliedSignature = null;
                SetStatus("Account settings unavailable · Current document is preview only; nothing was applied.", true);
            }
            RefreshHierarchy();
        }

        private void RestoreAccountSelection()
        {
            suppressEvents = true;
            try { accountSelector.SelectedItem = activeAccount; }
            finally { suppressEvents = false; }
        }

        private void RefreshHierarchy()
        {
            if (busy) return;
            if (activeAccount == null)
            {
                LoadAccounts();
                return;
            }
            // Refresh updates directory membership only; it must never reload over user edits or an imported file.
            managers = null;
            editor.ShowMembers(null, "Discovering this account's manager chain…");
            bool success = RunOperation("Refreshing manager hierarchy…", "Could not load manager hierarchy", () =>
            {
                var result = service.DiscoverManagers(activeAccount.Id);
                if (IsDisposed || Disposing) return;
                if (result == null || result.Managers == null)
                    throw new OrgLensException("The directory hierarchy is unavailable. Refresh to retry.");
                managers = result.Managers;
                editor.ShowMembers(managers);
                SetNotice((managers.Count == 0 ? "No managers returned. " : managers.Count + " BOSS members. ") + result.Notice);
                SetStatus("Hierarchy refreshed · Settings and unsaved edits were preserved.", false);
            }, showLoading: true);
            if (IsDisposed || Disposing) return;
            if (!success)
            {
                editor.ShowMembers(null, "Directory lookup failed. Refresh, or disable both BOSS styles to apply CUSTOM / TO ME.");
                SetNotice("Directory unavailable. Refresh, or disable both BOSS styles; settings and file operations still work.");
            }
            if (!accountConfigLoaded && success)
                SetStatus("Saved account settings could not be loaded. Current document is preview only; edit or Open settings.", true);
            UpdateActions();
        }

        private void EditorChanged(object sender, EventArgs e)
        {
            SetStatus(HasUnsavedChanges ? "Unsaved edits · Save settings to keep a file copy. Apply is a separate action."
                : "No unsaved file edits · Preview reflects the current settings.", false);
            UpdatePreview();
            UpdateActions();
        }

        private void UpdatePreview()
        {
            try
            {
                preview.SetConfiguration(editor.ReadConfiguration());
            }
            catch (OrgLensException error)
            {
                preview.AccessibleDescription = "Preview shows the last valid settings. " + error.Message;
                SetStatus("Preview paused · " + error.Message, true);
            }
        }

        private void ApplyRules()
        {
            if (busy || activeAccount == null) return;
            OrgLensConfiguration configuration = null;
            if (!RunOperation("Validating settings…", "Could not apply OrgLens rules", () =>
            {
                configuration = editor.ReadConfiguration();
                if (managers == null && configuration.BossEnabled)
                    throw new OrgLensException("The BOSS hierarchy is unavailable. Refresh it, or disable BOTH BOSS styles to apply CUSTOM / TO ME.");
            })) return;
            string target = activeAccount.SmtpAddress + " · Inbox";
            string message = service.IsDemo
                ? "Simulate Apply for " + target + "?\r\n\r\nOnly this demo's in-memory settings will change. Outlook has not been changed and will not be accessed."
                : "Apply formatting to " + target + "?\r\n\r\nThis replaces only OrgLens-owned formatting rules. All other rules are retained. "
                    + "Outlook built-in formatting can take precedence. No messages are modified.";
            if (!dialogs.Confirm(this, message, service.IsDemo ? "Simulate Apply" : "Apply to selected Inbox")) return;
            RunOperation("Applying grouped formatting…", "Could not apply OrgLens rules", () =>
            {
                service.ApplyConfiguration(activeAccount.Id, configuration, managers);
                appliedSignature = editor.EditSignature;
                // With no selected user file, the account cache is the durable copy.
                if (filePath == null) cleanSignature = editor.EditSignature;
                accountConfigLoaded = true;
                string result = service.IsDemo ? "Demo settings applied in memory. Outlook has not been changed."
                    : "OrgLens formatting applied to " + target + ". All other formatting rules were retained.";
                SetStatus(result, false);
                dialogs.Information(this, result, service.IsDemo ? "Demo only" : "OrgLens applied");
            });
        }

        private void RemoveRules()
        {
            if (busy || activeAccount == null) return;
            string target = activeAccount.SmtpAddress + " · Inbox";
            string message = (service.IsDemo ? "Simulate removal from " : "Remove only OrgLens-owned formatting rules from ")
                + target + "?\r\n\r\nAll original and other formatting rules will be retained. "
                + "Your preferences, selected settings file and unsaved edits will be kept."
                + (service.IsDemo ? "\r\nOutlook will not be accessed or changed." : "");
            if (!dialogs.Confirm(this, message, service.IsDemo ? "Simulate removal" : "Remove OrgLens rules")) return;
            RunOperation("Removing OrgLens-owned rules…", "Could not remove OrgLens rules", () =>
            {
                service.RemoveRules(activeAccount.Id);
                appliedSignature = null;
                string result = service.IsDemo ? "Demo rules removed in memory. Preferences kept. Outlook has not been changed."
                    : "OrgLens-owned rules removed from " + target + ". Preferences and all other rules were retained.";
                SetStatus(result, false);
                dialogs.Information(this, result, service.IsDemo ? "Demo only" : "OrgLens removed");
            });
        }

        private void UpdateActions()
        {
            bool valid;
            try { editor.ReadConfiguration(); valid = true; }
            catch (OrgLensException) { valid = false; }
            accountSelector.Enabled = !busy && accountSelector.Items.Count > 0;
            refreshButton.Enabled = !busy;
            applyButton.Enabled = !busy && activeAccount != null && valid && (managers != null || !editor.BossEnabled);
            removeButton.Enabled = !busy && activeAccount != null;
            closeButton.Enabled = !busy;
            openButton.Enabled = !busy;
            saveButton.Enabled = !busy;
            saveAsButton.Enabled = !busy;
            editor.Enabled = !busy;
            fileLabel.Text = filePath == null ? "SETTINGS FILE · No file selected — Open or Save as to choose a JSON file."
                : "SETTINGS FILE · " + filePath;
            tooltips.SetToolTip(fileLabel, filePath ?? "Settings files contain only styles, custom email addresses and TO ME scope.");
            fileLabel.AccessibleDescription = filePath ?? "No selected settings file";
            documentLabel.Text = (HasUnsavedChanges ? "Unsaved file edits · " : "No unsaved file edits · ")
                + (appliedSignature == editor.EditSignature ? service.IsDemo ? "Simulated only; Outlook unchanged."
                    : "Applied to this account."
                    : "Loaded / preview only — click " + (service.IsDemo ? "Simulate apply." : "Apply."));
            string target = activeAccount == null ? "No account selected" : activeAccount.SmtpAddress + " · Inbox only";
            scopeLabel.Text = target + ". Only OrgLens-owned rules change; original rules are retained. No messages are modified.\r\n"
                + (activeAccount != null && managers == null && editor.BossEnabled
                    ? "Apply unavailable: refresh the directory or turn OFF both BOSS READ and UNREAD. Open / Save never applies."
                    : "Open / Save never applies. BOSS > CUSTOM > TO ME. Outlook built-in formatting can take precedence.");
        }

        private bool RunOperation(string progress, string errorTitle, Action action, bool showLoading = false)
        {
            if (IsDisposed || Disposing) return false;
            busy = true;
            UseWaitCursor = true;
            UpdateActions();
            SetStatus(progress, false);
            statusLabel.Update();
            try
            {
                // Only immutable presentation data crosses threads; action stays on the caller's Outlook STA.
                using (var indicator = showLoading
                    ? new LoadingProgress(Handle, Bounds, Screen.FromHandle(Handle).WorkingArea, DeviceDpi, progress)
                    : null)
                {
                    loadingProgress = indicator;
                    action();
                    return !IsDisposed && !Disposing;
                }
            }
            catch (OrgLensException error) { ReportError(errorTitle, error.Message); return false; }
            catch (COMException error)
            {
                ReportError(errorTitle, "Check classic Outlook and retry.\r\n" + error.Message
                    + "\r\nError: 0x" + error.ErrorCode.ToString("X8"));
                return false;
            }
            catch (IOException error) { ReportError(errorTitle, error.Message); return false; }
            catch (UnauthorizedAccessException error) { ReportError(errorTitle, "Access denied. " + error.Message); return false; }
            catch (SecurityException error) { ReportError(errorTitle, error.Message); return false; }
            catch (SerializationException error) { ReportError(errorTitle, "Invalid settings JSON. " + error.Message); return false; }
            catch (ArgumentException error) { ReportError(errorTitle, error.Message); return false; }
            catch (NotSupportedException error) { ReportError(errorTitle, error.Message); return false; }
            finally
            {
                loadingProgress = null;
                busy = false;
                if (!IsDisposed && !Disposing)
                {
                    UseWaitCursor = false;
                    UpdateActions();
                }
            }
        }

        private void ReportError(string title, string detail)
        {
            if (IsDisposed || Disposing) return;
            SetStatus(title + ". No successful change is confirmed.", true);
            dialogs.Error(this, detail, title);
        }

        private void SetStatus(string text, bool error)
        {
            statusLabel.Text = text;
            statusLabel.ForeColor = error ? System.Drawing.Color.FromArgb(185, 28, 28) : Muted;
            statusLabel.AccessibleDescription = text;
        }

        private void SetNotice(string text)
        {
            noticeLabel.Text = text;
            noticeLabel.AccessibleDescription = text;
            tooltips.SetToolTip(noticeLabel, text);
        }
    }
}
