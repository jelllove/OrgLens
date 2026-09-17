using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OrgLens.Desktop;

namespace OrgLens.Preview
{
    internal static partial class SmokeTests
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowEnabled(IntPtr handle);

        private static void TestModelessWindow()
        {
            var dialogs = new TestDialogs();
            var service = new ProbeService();
            var created = new List<SettingsForm>();
            using (var host = new Form())
            using (var window = new ModelessSettingsWindow(() =>
            {
                var form = new SettingsForm(service, dialogs);
                created.Add(form);
                return form;
            }))
            using (var timeout = new Timer { Interval = 3000 })
            {
                Show(host);
                bool timedOut = false;
                timeout.Tick += (sender, args) =>
                {
                    timedOut = true;
                    timeout.Stop();
                    if (created.Count > 0) created[created.Count - 1].Close();
                };
                timeout.Start();
                window.Show();
                timeout.Stop();
                Assert(!timedOut && created.Count == 1, "Opening settings must return without a modal message loop.");
                var first = created[0];
                Assert(first.Visible && !first.Modal && !first.IsDisposed && !first.TopMost,
                    "The settings window must remain alive, modeless and not forced on top.");
                Assert(IsWindowEnabled(host.Handle), "Opening OrgLens must not disable the host's native window.");
                bool hostResponded = false;
                host.BeginInvoke((MethodInvoker)(() => hostResponded = true));
                Application.DoEvents();
                Assert(hostResponded, "The host message loop must keep processing input.");
                int calls = service.Calls.Count;
                Find<NumericUpDown>(first, "BossUnreadFontSize").Value = 21;
                first.WindowState = FormWindowState.Minimized;
                window.Show();
                Assert(created.Count == 1 && first.WindowState == FormWindowState.Normal
                    && service.Calls.Count == calls && first.HasUnsavedChanges,
                    "Reopening must restore the same window without reloading or discarding edits.");
                dialogs.ConfirmResult = false;
                first.Close();
                window.Show();
                Assert(!first.IsDisposed && created.Count == 1,
                    "Cancelling the ordinary Close prompt must retain the same settings window.");
                dialogs.ConfirmResult = true;
                first.Close();
                Assert(first.IsDisposed, "A normal modeless Close must dispose the window.");
                window.Show();
                Assert(created.Count == 2 && created[1].Visible, "Closing and reopening must create a fresh window.");
                Find<NumericUpDown>(created[1], "BossUnreadFontSize").Value = 23;
                dialogs.ConfirmResult = false;
                int confirmations = dialogs.ConfirmationMessages.Count;
                window.Dispose();
                Assert(created[1].IsDisposed && dialogs.ConfirmationMessages.Count == confirmations,
                    "Host shutdown must dispose even a dirty window without leaving a cancelled orphan.");
                bool rejected = false;
                try { window.Show(); }
                catch (ObjectDisposedException) { rejected = true; }
                Assert(rejected, "A disconnected window host must not reopen against stale Outlook state.");
                Assert(IsWindowEnabled(host.Handle), "Closing/disconnecting must leave the host enabled.");
                host.Close();
            }
        }
    }
}
