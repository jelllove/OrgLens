using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using OrgLens.Core;
using OrgLens.Desktop;

namespace OrgLens.Preview
{
    internal static partial class SmokeTests
    {
        internal static void Run()
        {
            Assert(Thread.CurrentThread.GetApartmentState() == ApartmentState.STA, "The UI must run on STA.");
            TestPalette();
            TestColorPickers();
            TestModelessWindow();
            TestGroupedEditingAndAccountSafety();
            TestDirectoryAndOperationFailures();
            TestLoadingFailuresAndNoAccounts();
            TestSettingsFiles();
            TestRealScopeWithFakeService();
            TestModalEntryPoint();
            TestDemoStorageCopies();
        }

        private static void TestPalette()
        {
            var colors = (MailColor[])Enum.GetValues(typeof(MailColor));
            Assert(colors.Length == 17, "Automatic plus the 16 native Outlook colors must be offered.");
            foreach (var color in colors)
            {
                var expected = color == MailColor.Automatic ? SystemColors.WindowText : Color.FromName(color.ToString());
                Assert(OutlookPalette.ToColor(color).ToArgb() == expected.ToArgb(), "Incorrect palette RGB: " + color);
            }
            ExpectOrgLensError(() => OutlookPalette.ToColor((MailColor)99));
        }

        private static void TestGroupedEditingAndAccountSafety()
        {
            var service = new ProbeService();
            var dialogs = new TestDialogs();
            using (var form = new SettingsForm(service, dialogs))
            {
                Show(form);
                var editor = Find<GroupSettingsControl>(form, "GroupedSettings");
                var preview = Find<MailPreviewControl>(form, "MessagePreview");
                var accounts = Find<ComboBox>(form, "AccountSelector");
                Assert(accounts.Items.Count == 2 && accounts.SelectedIndex == 0, "First account should load automatically.");
                Assert(service.Calls.Take(3).SequenceEqual(new[] { "accounts", "load:sample-avery", "discover:sample-avery" }),
                    "Configuration must load BEFORE the separate directory request.");
                Assert(Count<GroupBox>(form) == 3 && Count<GroupStyleEditor>(form) == 4,
                    "Exactly three groups and four shared style rows are required.");
                Assert(Count<DataGridView>(form) == 0, "Per-manager editing must be removed.");
                Assert(Count<NumericUpDown>(form) == 4, "All four style rows need font-size editors.");
                var members = Find<TextBox>(form, "BossMembers");
                Assert(members.ReadOnly && members.Text.Contains("Jamie Rivera") && members.Lines.Length == 3,
                    "Discovered bosses should appear only in a read-only member summary.");
                Assert(Find<Label>(form, "DemoBanner").Visible && Find<Label>(form, "HierarchySource").Text.Contains("Fictional"),
                    "The demo must clearly identify its fictional source.");
                foreach (string name in new[] { "BossRead", "BossUnread", "CustomUnread", "ToMe" })
                {
                    var size = Find<NumericUpDown>(form, name + "FontSize");
                    Assert(size.Minimum == 1 && size.Maximum == 127 && size.DecimalPlaces == 0,
                        "Font sizes must be whole points from 1 through 127.");
                    Assert(Find<ComboBox>(form, name + "Color").Items.Count == 17, "Only native palette colors may be selected.");
                }
                Assert(preview.Rows.Count == 9 && preview.Rows[0].Kind == RuleKind.BossUnread
                    && preview.Rows[1].Kind == RuleKind.BossRead && preview.Rows[2].Kind == RuleKind.CustomUnread
                    && preview.Rows[3].Kind == null && preview.Rows[4].Kind == RuleKind.ToMe
                    && preview.Rows[5].Kind == RuleKind.ToMe && preview.Rows[6].Kind == null
                    && preview.Rows[7].Kind == null && preview.Rows[8].Kind == null,
                    "Preview must cover read/unread BOSS and CUSTOM, direct To/Cc, read scope, and excluded DL/BCC.");

                var fontSize = Find<NumericUpDown>(form, "BossUnreadFontSize");
                int normalHeight = preview.Rows[0].Height;
                fontSize.Value = 127;
                Assert(preview.Rows[0].Font.SizeInPoints == 127 && preview.Rows[0].Height > normalHeight
                    && preview.Rows[0].Height > preview.Rows[0].Font.Height + form.Font.Height * 2,
                    "127-point text must enlarge preview rows rather than clipping vertically.");
                Assert(preview.AutoScrollMinSize.Width > preview.Width, "Large preview text must be horizontally scrollable.");
                fontSize.Value = 1;
                Assert(preview.Rows[0].Font.SizeInPoints == 1, "Minimum font size must preview accurately.");
                fontSize.Value = 26;
                Find<ComboBox>(form, "BossUnreadColor").SelectedItem = MailColor.Maroon;
                Find<CheckBox>(form, "BossUnreadBold").Checked = false;
                Find<CheckBox>(form, "BossUnreadItalic").Checked = true;
                Assert(form.HasUnsavedChanges && preview.Rows[0].Font.Italic && !preview.Rows[0].Font.Bold
                    && preview.Rows[0].Color == OutlookPalette.ToColor(MailColor.Maroon),
                    "Style changes must update the live preview and dirty state.");
                var custom = Find<TextBox>(form, "CustomAddresses");
                custom.Text = " NEW@example.com;new@example.com,\r\nsecond@example.com";
                Assert(editor.ReadConfiguration().CustomAddresses.SequenceEqual(new[] { "new@example.com", "second@example.com" }),
                    "Custom addresses must validate, normalize and deduplicate all supported separators.");
                custom.Text = "";
                Assert(preview.Rows[2].Kind == null, "Removing all custom addresses must remove CUSTOM matching.");
                custom.Text = "casey.reed@example.com";
                Find<CheckBox>(form, "ToMeIncludeRead").Checked = true;
                Assert(!editor.ReadConfiguration().ToMeUnreadOnly && preview.Rows[6].Kind == RuleKind.ToMe
                    && preview.Rows[7].Kind == null && preview.Rows[8].Kind == null,
                    "Including read messages must affect direct To/Cc only, never DL/BCC.");
                Find<CheckBox>(form, "BossUnreadEnabled").Checked = false;
                Assert(preview.Rows[0].Kind == RuleKind.ToMe, "Disabling BOSS should reveal the next eligible style.");
                Find<CheckBox>(form, "BossUnreadEnabled").Checked = true;
                Assert(preview.Rows[0].Kind == RuleKind.BossUnread, "BOSS must take precedence over TO ME.");
                string signature = editor.EditSignature;
                int confirmations = dialogs.ConfirmationMessages.Count;
                Click(form, "RefreshHierarchy");
                Assert(editor.EditSignature == signature && form.HasUnsavedChanges
                    && dialogs.ConfirmationMessages.Count == confirmations,
                    "Refresh must preserve all shared edits without a discard prompt.");
                dialogs.ConfirmResult = false;
                Click(form, "ApplyRules");
                Assert(service.ApplyCount == 0 && form.HasUnsavedChanges, "Cancelling Apply must not mutate or clear edits.");
                dialogs.ConfirmResult = true;
                Click(form, "ApplyRules");
                Assert(service.ApplyCount == 1 && service.LastAppliedAccount == "sample-avery" && !form.HasUnsavedChanges,
                    "Explicit Apply without a user file should cache only the selected account.");
                Assert(dialogs.InformationMessages.Last().Contains("Outlook has not been changed"),
                    "Demo Apply must never claim a real Outlook mutation.");

                fontSize.Value = 31;
                dialogs.ConfirmResult = false;
                form.Close();
                Assert(!form.IsDisposed && form.Visible, "Cancelling close must preserve edits.");
                accounts.SelectedIndex = 1;
                Assert(accounts.SelectedIndex == 0 && fontSize.Value == 31, "Cancelled account switch must restore the selection.");
                dialogs.ConfirmResult = true;
                service.NextLoadError = new COMException("Deliberate account load failure.", unchecked((int)0x80004005));
                accounts.SelectedIndex = 1;
                Assert(accounts.SelectedIndex == 0 && fontSize.Value == 31 && form.HasUnsavedChanges,
                    "Failed account load must roll back selection, controls and dirty baseline.");
                accounts.SelectedIndex = 1;
                Assert(fontSize.Value == 12 && !form.HasUnsavedChanges && members.Text.Contains("No managers")
                    && Find<Button>(form, "ApplyRules").Enabled, "A valid empty chain is not a directory failure.");
                fontSize.Value = 19;
                Click(form, "ApplyRules");
                Assert(service.LastAppliedAccount == "sample-quinn", "No-manager accounts may apply valid grouped settings.");
                accounts.SelectedIndex = 0;
                Assert(fontSize.Value == 26 && !Find<CheckBox>(form, "BossUnreadBold").Checked,
                    "Switching back must load that account's last successful configuration.");
                form.Size = form.MinimumSize;
                form.PerformLayout();
                Assert(Find<Button>(form, "ApplyRules").Width > 100 && fontSize.Width >= 50,
                    "Minimum window size must retain usable editors and actions.");
                form.Close();
            }
        }

        private static void TestDirectoryAndOperationFailures()
        {
            var service = new ProbeService { NextDiscoveryError = new OrgLensException("Directory temporarily unavailable.") };
            var dialogs = new TestDialogs();
            using (var form = new SettingsForm(service, dialogs))
            {
                Show(form);
                var editor = Find<GroupSettingsControl>(form, "GroupedSettings");
                Assert(!Find<Button>(form, "ApplyRules").Enabled && Find<Button>(form, "RemoveRules").Enabled
                    && Find<TextBox>(form, "CustomAddresses").Text.Contains("casey")
                    && Find<Button>(form, "OpenSettings").Enabled && Find<Button>(form, "SaveSettingsAs").Enabled,
                    "Directory failure must keep loaded settings, file editing and Remove available.");
                Find<NumericUpDown>(form, "CustomUnreadFontSize").Value = 18;
                Find<CheckBox>(form, "BossReadEnabled").Checked = false;
                Assert(!Find<Button>(form, "ApplyRules").Enabled, "One enabled BOSS style still requires a directory.");
                Find<CheckBox>(form, "BossUnreadEnabled").Checked = false;
                Assert(Find<Button>(form, "ApplyRules").Enabled, "CUSTOM / TO ME can apply with both BOSS styles disabled.");
                service.NextApplyError = new OrgLensException("Deliberate transactional Apply failure.");
                string signature = editor.EditSignature;
                Click(form, "ApplyRules");
                Assert(form.HasUnsavedChanges && service.ApplyCount == 0 && editor.EditSignature == signature,
                    "Failed Apply must retain all edits and the prior dirty state.");
                Click(form, "ApplyRules");
                Assert(service.ApplyCount == 1 && service.LastAppliedWithoutHierarchy,
                    "Hierarchy failure must be passed as null, never an invented empty chain.");
                Find<CheckBox>(form, "BossReadEnabled").Checked = true;
                service.NextDiscoveryError = new COMException("Deliberate directory error.", unchecked((int)0x80004005));
                signature = editor.EditSignature;
                Click(form, "RefreshHierarchy");
                Assert(editor.EditSignature == signature && form.HasUnsavedChanges
                    && dialogs.Errors.Last().Contains("0x80004005"), "COM failure must preserve edits and expose its error code.");
                Click(form, "RefreshHierarchy");
                Assert(Find<Button>(form, "ApplyRules").Enabled && editor.EditSignature == signature && form.HasUnsavedChanges,
                    "Directory retry must recover without resetting configuration.");
                var custom = Find<TextBox>(form, "CustomAddresses");
                custom.Text = "not-an-email";
                Assert(!Find<Button>(form, "ApplyRules").Enabled
                    && Find<Label>(form, "AddressValidation").Text.Contains("Invalid")
                    && Find<MailPreviewControl>(form, "MessagePreview").AccessibleDescription.Contains("last valid"),
                    "Invalid edited emails must show a clear validation error and never apply.");
                int previousRemovals = service.RemoveCount;
                dialogs.ConfirmResult = false;
                Click(form, "RemoveRules");
                Assert(service.RemoveCount == previousRemovals, "Cancelled removal must not call the backend.");
                dialogs.ConfirmResult = true;
                signature = editor.EditSignature;
                Click(form, "RemoveRules");
                Assert(service.RemoveCount == previousRemovals + 1 && editor.EditSignature == signature && form.HasUnsavedChanges,
                    "Remove needs only an account and must preserve even invalid unsaved preferences.");
                service.NextRemoveError = new COMException("Deliberate removal failure.");
                Click(form, "RemoveRules");
                Assert(editor.EditSignature == signature && form.HasUnsavedChanges
                    && Find<Label>(form, "OperationStatus").Text.Contains("Could not remove"),
                    "Failed removal must preserve state and never report success.");
                form.Close();
            }
        }

        private static void TestLoadingFailuresAndNoAccounts()
        {
            var dialogs = new TestDialogs();
            var service = new ProbeService { NextAccountsError = new OrgLensException("Deliberate account failure.") };
            using (var form = new SettingsForm(service, dialogs))
            {
                Show(form);
                Assert(dialogs.Errors.Count == 1 && !Find<Button>(form, "ApplyRules").Enabled
                    && Find<GroupSettingsControl>(form, "GroupedSettings").Enabled,
                    "Account errors must leave the settings editor available.");
                Click(form, "RefreshHierarchy");
                Assert(Find<Button>(form, "ApplyRules").Enabled, "Refresh should retry failed account loading.");
                form.Close();
            }
            service = new ProbeService { NextLoadError = new System.IO.IOException("Cannot read account cache.") };
            using (var form = new SettingsForm(service, dialogs))
            {
                Show(form);
                Assert(service.DiscoverCount == 1 && Find<Button>(form, "RemoveRules").Enabled
                    && Find<Label>(form, "OperationStatus").Text.Contains("could not be loaded"),
                    "Initial cache failure must be explicit and separate from discovery.");
                form.Close();
            }
            service = new ProbeService { ReturnNoAccounts = true };
            using (var form = new SettingsForm(service, dialogs))
            {
                Show(form);
                Assert(!Find<Button>(form, "ApplyRules").Enabled && !Find<Button>(form, "RemoveRules").Enabled
                    && Find<Label>(form, "HierarchyNotice").Text.Contains("No supported accounts"),
                    "No accounts must disable only native operations, not the settings document.");
                Assert(service.DiscoverCount == 0 && Find<Button>(form, "SaveSettingsAs").Enabled, "Offline file editing remains usable.");
                form.Close();
            }
        }

        private static void TestDemoStorageCopies()
        {
            var service = new DemoService();
            string account = service.GetAccounts()[0].Id;
            var people = service.DiscoverManagers(account).Managers;
            var config = service.LoadConfiguration(account);
            config.BossRead.Color = MailColor.Red;
            config.CustomAddresses.Add("extra@example.com");
            service.ApplyConfiguration(account, config, people);
            config.BossRead.Color = MailColor.White;
            config.CustomAddresses.Clear();
            Assert(service.LoadConfiguration(account).BossRead.Color == MailColor.Red
                && service.LoadConfiguration(account).CustomAddresses.Count == 2, "Demo Apply must deep-copy styles and addresses.");
            var loaded = service.LoadConfiguration(account);
            loaded.BossRead.Color = MailColor.Aqua;
            loaded.CustomAddresses.Clear();
            Assert(service.LoadConfiguration(account).BossRead.Color == MailColor.Red, "Demo Load must return detached copies.");
            service.RemoveRules(account);
            Assert(service.LoadConfiguration(account).BossRead.Color == MailColor.Red, "Demo Remove must retain user preferences.");
            Assert(new DemoService().LoadConfiguration(account).BossRead.Color == MailColor.Navy,
                "A new demo instance must not read any real profile or persist across services.");
            loaded.BossRead.FontSize = 128;
            ExpectOrgLensError(() => service.ApplyConfiguration(account, loaded, people));
            Assert(service.LoadConfiguration(account).BossRead.Color == MailColor.Red, "Rejected demo Apply must not update cached preferences.");
        }

        private static void TestModalEntryPoint()
        {
            bool loaded = false;
            using (var form = new SettingsForm(new DemoService()))
            {
                form.Shown += (sender, args) => form.BeginInvoke((MethodInvoker)(() =>
                {
                    loaded = Count<GroupStyleEditor>(form) == 4 && Find<TextBox>(form, "BossMembers").Text.Contains("Jamie");
                    form.Close();
                }));
                form.ShowDialog();
            }
            Assert(loaded, "The public SettingsForm(IOrgLensService) constructor must support modal STA hosting.");
        }

        private static void TestRealScopeWithFakeService()
        {
            var service = new ProbeService { DemoMode = false };
            var dialogs = new TestDialogs();
            using (var form = new SettingsForm(service, dialogs))
            {
                Show(form);
                Assert(!Find<Label>(form, "DemoBanner").Visible
                    && Find<Label>(form, "HierarchySource").Text.Contains("Exchange address book"), "Real mode must identify its directory source.");
                Click(form, "ApplyRules");
                string confirmation = dialogs.ConfirmationMessages.Last();
                Assert(confirmation.Contains("avery.morgan@example.com") && confirmation.Contains("Inbox")
                    && confirmation.Contains("only OrgLens-owned") && confirmation.Contains("All other rules are retained"),
                    "Real Apply confirmation must identify the selected Inbox and narrowly scoped change.");
                Find<ComboBox>(form, "AccountSelector").SelectedIndex = 1;
                Click(form, "RemoveRules");
                Assert(service.LastRemovedAccount == "sample-quinn"
                    && dialogs.ConfirmationMessages.Last().Contains("preferences"), "Removal must target only the active account and keep preferences.");
                form.Close();
            }
        }

        private static void Show(Form form) { form.Show(); Application.DoEvents(); }
        private static void Click(Control form, string name) { Find<Button>(form, name).PerformClick(); }

        private static T Find<T>(Control parent, string name) where T : Control
        {
            T result = parent.Controls.Find(name, true).OfType<T>().SingleOrDefault();
            Assert(result != null, "Missing control: " + name);
            return result;
        }

        private static int Count<T>(Control parent) where T : Control
        {
            return parent.Controls.Cast<Control>().Sum(child => (child is T ? 1 : 0) + Count<T>(child));
        }

        private static void ExpectOrgLensError(Action action)
        {
            bool rejected = false;
            try { action(); }
            catch (OrgLensException) { rejected = true; }
            Assert(rejected, "Invalid configuration should raise a targeted OrgLensException.");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private sealed class TestDialogs : ISettingsDialogs
        {
            internal bool ConfirmResult = true;
            internal string OpenPath;
            internal string SavePath;
            internal int OpenCount;
            internal int SaveCount;
            internal readonly List<string> ConfirmationMessages = new List<string>();
            internal readonly List<string> InformationMessages = new List<string>();
            internal readonly List<string> Errors = new List<string>();
            public bool Confirm(IWin32Window owner, string message, string title)
            {
                ConfirmationMessages.Add(message);
                return ConfirmResult;
            }
            public void Information(IWin32Window owner, string message, string title) { InformationMessages.Add(message); }
            public void Error(IWin32Window owner, string message, string title) { Errors.Add(title + ": " + message); }
            public string ChooseOpenFile(IWin32Window owner, string currentPath) { OpenCount++; return OpenPath; }
            public string ChooseSaveFile(IWin32Window owner, string currentPath) { SaveCount++; return SavePath; }
        }

        private sealed class ProbeService : IOrgLensService
        {
            private readonly DemoService demo = new DemoService();
            private readonly int threadId = Thread.CurrentThread.ManagedThreadId;
            internal readonly List<string> Calls = new List<string>();
            internal Exception NextAccountsError;
            internal Exception NextDiscoveryError;
            internal Exception NextLoadError;
            internal Exception NextApplyError;
            internal Exception NextRemoveError;
            internal bool ReturnNoAccounts;
            internal bool DemoMode = true;
            internal int DiscoverCount;
            internal int ApplyCount;
            internal int RemoveCount;
            internal string LastAppliedAccount;
            internal string LastRemovedAccount;
            internal bool LastAppliedWithoutHierarchy;

            public bool IsDemo { get { return DemoMode; } }
            public IReadOnlyList<MailboxAccount> GetAccounts()
            {
                Record("accounts");
                ThrowOnce(ref NextAccountsError);
                return ReturnNoAccounts ? new MailboxAccount[0] : demo.GetAccounts();
            }
            public HierarchyResult DiscoverManagers(string accountId)
            {
                Record("discover:" + accountId);
                DiscoverCount++;
                ThrowOnce(ref NextDiscoveryError);
                return demo.DiscoverManagers(accountId);
            }
            public OrgLensConfiguration LoadConfiguration(string accountId)
            {
                Record("load:" + accountId);
                ThrowOnce(ref NextLoadError);
                return demo.LoadConfiguration(accountId);
            }
            public void ApplyConfiguration(string accountId, OrgLensConfiguration configuration, IReadOnlyList<ManagerPerson> people)
            {
                Record("apply:" + accountId);
                ThrowOnce(ref NextApplyError);
                demo.ApplyConfiguration(accountId, configuration, people);
                ApplyCount++;
                LastAppliedAccount = accountId;
                LastAppliedWithoutHierarchy = people == null;
            }
            public void RemoveRules(string accountId)
            {
                Record("remove:" + accountId);
                ThrowOnce(ref NextRemoveError);
                demo.RemoveRules(accountId);
                RemoveCount++;
                LastRemovedAccount = accountId;
            }
            private void Record(string operation)
            {
                Assert(Thread.CurrentThread.ManagedThreadId == threadId
                    && Thread.CurrentThread.GetApartmentState() == ApartmentState.STA, "All service calls must stay on the original STA UI thread.");
                Calls.Add(operation);
            }
            private static void ThrowOnce(ref Exception error)
            {
                if (error == null) return;
                var exception = error;
                error = null;
                throw exception;
            }
        }
    }
}
