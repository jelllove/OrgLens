using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Security;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using OrgLens.Core;
using OrgLens.Desktop;

namespace OrgLens.Preview
{
    internal static partial class SmokeTests
    {
        private delegate bool EnumerateWindow(IntPtr window, IntPtr parameter);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumerateWindow callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr window, uint command);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr window, out WindowRectangle rectangle);
        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr window, int index);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr window, IntPtr dc, uint flags);

        [StructLayout(LayoutKind.Sequential)]
        private struct WindowRectangle
        {
            internal int Left, Top, Right, Bottom;
        }

        private static void TestLoadingProgress()
        {
            var service = new ProbeService();
            var dialogs = new TestDialogs();
            var progressThreads = new HashSet<uint>();
            using (var form = new SettingsForm(service, dialogs))
            {
                bool shown = false;
                bool painted = false;
                form.Shown += (sender, args) => shown = true;
                Find<MailPreviewControl>(form, "MessagePreview").Paint += (sender, args) => painted = true;
                service.BeforeCall = operation =>
                {
                    Assert(shown && painted, "The settings contents must paint before initial service work.");
                    string phase = operation == "accounts" ? "Loading accounts"
                        : operation.StartsWith("discover:") ? "Refreshing manager hierarchy" : "Loading this account";
                    IntPtr window = WaitForLoadingWindow(form, phase);
                    uint process;
                    uint progressThread = GetWindowThreadProcessId(window, out process);
                    Assert(progressThread != GetWindowThreadProcessId(form.Handle, out process),
                        "Only the progress UI must use a separate UI thread.");
                    progressThreads.Add(progressThread);
                    Assert(!Find<Button>(form, "RefreshHierarchy").Enabled
                        && !Find<Button>(form, "CloseSettings").Enabled,
                        "Busy operations must not allow conflicting refresh or close actions.");
                    if (operation == "accounts" || operation.StartsWith("discover:"))
                        AssertProgressAnimates(window);
                };
                Show(form);
                AssertNoLoadingWindows(progressThreads);
                string signature = Find<GroupSettingsControl>(form, "GroupedSettings").EditSignature;
                for (int attempt = 0; attempt < 3; attempt++)
                {
                    Click(form, "RefreshHierarchy");
                    AssertNoLoadingWindows(progressThreads);
                }
                Assert(signature == Find<GroupSettingsControl>(form, "GroupedSettings").EditSignature,
                    "Repeated animated refresh must preserve settings.");
                form.Close();
                Assert(form.IsDisposed, "Settings must close normally after progress completes.");
                AssertNoLoadingWindows(progressThreads);
            }
        }

        private static void TestLoadingProgressFailures()
        {
            var threads = new HashSet<uint>();
            var dialogs = new TestDialogs { BeforeError = () => AssertNoLoadingWindows(threads) };
            var service = new ProbeService { NextAccountsError = new OrgLensException("Account lookup failed.") };
            using (var form = new SettingsForm(service, dialogs))
            {
                service.BeforeCall = operation =>
                {
                    string phase = operation == "accounts" ? "Loading accounts"
                        : operation.StartsWith("discover:") ? "Refreshing manager hierarchy" : "Loading this account";
                    uint process;
                    threads.Add(GetWindowThreadProcessId(WaitForLoadingWindow(form, phase), out process));
                };
                Show(form);
                Assert(dialogs.Errors.Count == 1, "Initial account failure must be reported.");
                AssertNoLoadingWindows(threads);
                Click(form, "RefreshHierarchy");
                Assert(Find<Button>(form, "ApplyRules").Enabled, "Account retry must recover after progress closes.");
                foreach (var error in new Exception[]
                {
                    new OrgLensException("Directory failed."), new COMException("Directory COM failed."),
                    new IOException("Directory IO failed."), new UnauthorizedAccessException("Access failed."),
                    new SecurityException("Security failed."), new SerializationException("Serialization failed."),
                    new ArgumentException("Argument failed."), new NotSupportedException("Unsupported directory.")
                })
                {
                    service.NextDiscoveryError = error;
                    Click(form, "RefreshHierarchy");
                    AssertNoLoadingWindows(threads);
                    Assert(Find<Button>(form, "RefreshHierarchy").Enabled,
                        "Every reported error must restore refresh controls.");
                }
                service.NextDiscoveryError = new InvalidOperationException("Unexpected directory failure.");
                bool propagated = false;
                try { Click(form, "RefreshHierarchy"); }
                catch (InvalidOperationException error) { propagated = error.Message == "Unexpected directory failure."; }
                Assert(propagated, "Unexpected service errors must still propagate after progress cleanup.");
                AssertNoLoadingWindows(threads);
                Click(form, "RefreshHierarchy");
                Assert(Find<Button>(form, "ApplyRules").Enabled, "Refresh must recover after unexpected errors.");
                form.Close();
                AssertNoLoadingWindows(threads);
            }

            foreach (var probe in new[]
            {
                new ProbeService { ReturnNoAccounts = true },
                new ProbeService { NextLoadError = new IOException("Saved settings failed.") },
                new ProbeService { NextDiscoveryError = new OrgLensException("Initial directory failed.") }
            })
            using (var form = new SettingsForm(probe, dialogs))
            {
                Show(form);
                AssertNoLoadingWindows(threads);
                form.Close();
                AssertNoLoadingWindows(threads);
            }
        }

        private static void TestLoadingProgressShutdown()
        {
            var threads = new HashSet<uint>();
            foreach (string shutdownAt in new[] { "accounts", "load:", "discover:" })
            {
                var service = new ProbeService();
                SettingsForm form = null;
                using (var host = new ModelessSettingsWindow(() => form = new SettingsForm(service, new TestDialogs())))
                {
                    service.BeforeCall = operation =>
                    {
                        if (!operation.StartsWith(shutdownAt)) return;
                        uint process;
                        string phase = operation == "accounts" ? "Loading accounts"
                            : operation.StartsWith("discover:") ? "Refreshing manager hierarchy" : "Loading this account";
                        threads.Add(GetWindowThreadProcessId(WaitForLoadingWindow(form, phase), out process));
                        form.Close();
                        Assert(!form.IsDisposed, "An ordinary close cannot interrupt a directory operation.");
                        host.Dispose();
                        Assert(form.IsDisposed, "Host shutdown must dispose settings even during service work.");
                        AssertNoLoadingWindows(threads);
                    };
                    host.Show();
                    Application.DoEvents();
                    Assert(form.IsDisposed, "Shutdown during deferred initialization must finish without orphan controls.");
                    AssertNoLoadingWindows(threads);
                }
            }

            var unopened = new ProbeService();
            using (var form = new SettingsForm(unopened, new TestDialogs()))
            {
                form.Show();
                form.Dispose();
                Application.DoEvents();
                Assert(unopened.Calls.Count == 0, "Disposal before deferred loading must not access the service.");
                AssertNoLoadingWindows(threads);
            }

            foreach (IntPtr invalidOwner in new[] { IntPtr.Zero, new IntPtr(-1) })
            using (var progress = new LoadingProgress(invalidOwner, new Rectangle(0, 0, 800, 600),
                new Rectangle(0, 0, 1920, 1080), 96, "Invalid owner regression"))
            {
                var observation = Stopwatch.StartNew();
                while (observation.ElapsedMilliseconds < 500)
                {
                    Application.DoEvents();
                    Assert(LoadingWindows().Count == 0, "An invalid owner must never create an ownerless loading window.");
                    Thread.Sleep(5);
                }
            }
            AssertNoLoadingWindows(threads);
        }

        private static List<IntPtr> LoadingWindows()
        {
            var result = new List<IntPtr>();
            uint currentProcess = (uint)Process.GetCurrentProcess().Id;
            EnumWindows((window, parameter) =>
            {
                uint process;
                GetWindowThreadProcessId(window, out process);
                if (process != currentProcess) return true;
                var text = new StringBuilder(200);
                GetWindowText(window, text, text.Capacity);
                if (text.ToString().StartsWith("OrgLens loading · ")) result.Add(window);
                return true;
            }, IntPtr.Zero);
            return result;
        }

        private static IntPtr WaitForLoadingWindow(SettingsForm owner, string phase)
        {
            IntPtr foreground = GetForegroundWindow();
            var timeout = Stopwatch.StartNew();
            IntPtr window = IntPtr.Zero;
            while (timeout.ElapsedMilliseconds < 3000)
            {
                window = LoadingWindows().SingleOrDefault(candidate => IsWindowVisible(candidate));
                if (window != IntPtr.Zero) break;
                Thread.Sleep(15);
            }
            Assert(window != IntPtr.Zero, "A visible animated loading panel is required during " + phase + ".");
            var text = new StringBuilder(200);
            GetWindowText(window, text, text.Capacity);
            Assert(text.ToString().Contains(phase), "Progress must describe the current operation without a fake percentage.");
            Assert(GetWindow(window, 4) == owner.Handle, "Progress must be natively owned by the settings window.");
            int style = GetWindowLong(window, -20);
            Assert((style & 0x08000000) != 0 && (style & 0x00000080) != 0 && (style & 0x00040000) == 0,
                "Progress must not activate or create an extra taskbar entry.");
            Assert(foreground == GetForegroundWindow(), "Opening progress must not steal focus.");
            WindowRectangle bounds;
            Assert(GetWindowRect(window, out bounds) && bounds.Right - bounds.Left < owner.Width
                && bounds.Bottom - bounds.Top < owner.Height, "Progress must be a compact panel, not another full app.");
            return window;
        }

        private static void AssertProgressAnimates(IntPtr window)
        {
            using (var before = CaptureProgress(window))
            {
                bool changed = false;
                for (int sample = 0; sample < 6 && !changed; sample++)
                {
                    // Deliberately block the original service STA without pumping messages.
                    Thread.Sleep(120);
                    using (var after = CaptureProgress(window))
                    {
                        for (int y = 0; y < before.Height && !changed; y += 2)
                        for (int x = 0; x < before.Width && !changed; x += 2)
                            changed = before.GetPixel(x, y) != after.GetPixel(x, y);
                    }
                }
                Assert(changed, "Progress must actually animate while the original service/UI STA is blocked.");
            }
        }

        private static Bitmap CaptureProgress(IntPtr window)
        {
            WindowRectangle bounds;
            Assert(GetWindowRect(window, out bounds), "Progress must retain its native window while loading.");
            var bitmap = new Bitmap(bounds.Right - bounds.Left, bounds.Bottom - bounds.Top);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                IntPtr dc = graphics.GetHdc();
                try { Assert(PrintWindow(window, dc, 0), "The native progress surface must paint."); }
                finally { graphics.ReleaseHdc(dc); }
            }
            return bitmap;
        }

        private static void AssertNoLoadingWindows(HashSet<uint> progressThreads)
        {
            var timeout = Stopwatch.StartNew();
            while (timeout.ElapsedMilliseconds < 3000)
            {
                // Cleanup may send native owner notifications; production deliberately never joins this thread.
                Application.DoEvents();
                using (var process = Process.GetCurrentProcess())
                {
                    bool alive = process.Threads.Cast<ProcessThread>().Any(thread => progressThreads.Contains((uint)thread.Id));
                    if (LoadingWindows().Count == 0 && !alive) return;
                }
                Thread.Sleep(15);
            }
            Assert(false, "Completed/failed/closed settings must leave no loading windows or animation threads.");
        }
    }
}
