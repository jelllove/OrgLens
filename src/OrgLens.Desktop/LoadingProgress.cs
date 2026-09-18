using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace OrgLens.Desktop
{
    internal sealed class LoadingProgress : IDisposable
    {
        private readonly IntPtr owner;
        private readonly Rectangle ownerBounds;
        private readonly Rectangle workingArea;
        private readonly int dpi;
        private readonly string phase;
        private int stopped;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindow(IntPtr handle);

        private bool CanShow { get { return Volatile.Read(ref stopped) == 0 && IsWindow(owner); } }

        internal LoadingProgress(IntPtr owner, Rectangle ownerBounds, Rectangle workingArea, int dpi, string phase)
        {
            this.owner = owner;
            this.ownerBounds = ownerBounds;
            this.workingArea = workingArea;
            this.dpi = dpi;
            this.phase = phase;
            var thread = new Thread(Run) { IsBackground = true, Name = "OrgLens loading indicator" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        public void Dispose()
        {
            // Never wait/Join or invoke controls here: native ownership can message the busy Outlook STA.
            Interlocked.Exchange(ref stopped, 1);
        }

        private void Run()
        {
            if (!CanShow) return;
            try
            {
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException, true);
                using (var window = new ProgressWindow(this))
                {
                    if (CanShow) Application.Run(window);
                }
            }
            catch (ExternalException error) { Trace.TraceError("OrgLens loading display failed: " + error.Message); }
            catch (InvalidOperationException error) { Trace.TraceError("OrgLens loading display failed: " + error.Message); }
        }

        private sealed class ProgressWindow : Form
        {
            private readonly LoadingProgress progress;
            private readonly System.Windows.Forms.Timer animation;
            private readonly Font phaseFont;
            private readonly Font detailFont;
            private readonly float scale;
            private int frame;

            [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
            private static extern IntPtr SetWindowLongPtr(IntPtr handle, int index, IntPtr value);

            [DllImport("user32.dll", EntryPoint = "SetWindowLongW", SetLastError = true)]
            private static extern int SetWindowLong32(IntPtr handle, int index, int value);

            [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
            private static extern void SetLastError(uint error);

            internal ProgressWindow(LoadingProgress progress)
            {
                this.progress = progress;
                scale = progress.dpi / 96F;
                Text = "OrgLens loading · " + progress.phase;
                AccessibleName = progress.phase;
                AccessibleDescription = "Please wait. This Outlook operation cannot be interrupted.";
                FormBorderStyle = FormBorderStyle.None;
                ShowInTaskbar = false;
                StartPosition = FormStartPosition.Manual;
                AutoScaleMode = AutoScaleMode.None;
                BackColor = Color.White;
                ClientSize = new Size((int)(420 * scale), (int)(108 * scale));
                Location = new Point(
                    Math.Max(progress.workingArea.Left, Math.Min(progress.ownerBounds.Left + (progress.ownerBounds.Width - Width) / 2,
                        progress.workingArea.Right - Width)),
                    Math.Max(progress.workingArea.Top, Math.Min(progress.ownerBounds.Top + (progress.ownerBounds.Height - Height) / 2,
                        progress.workingArea.Bottom - Height)));
                phaseFont = new Font("Segoe UI", 10.5F, FontStyle.Bold, GraphicsUnit.Point);
                detailFont = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint
                    | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                animation = new System.Windows.Forms.Timer { Interval = 60 };
                animation.Tick += (sender, args) =>
                {
                    if (!progress.CanShow)
                    {
                        Close();
                        return;
                    }
                    frame = (frame + 1) % 12;
                    Invalidate();
                };
                animation.Start();
            }

            protected override bool ShowWithoutActivation { get { return true; } }

            protected override CreateParams CreateParams
            {
                get
                {
                    var parameters = base.CreateParams;
                    parameters.ExStyle |= 0x08000000 | 0x00000080; // WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW
                    return parameters;
                }
            }

            protected override void SetVisibleCore(bool value)
            {
                if (value)
                {
                    try
                    {
                        if (!progress.CanShow)
                        {
                            Application.ExitThread();
                            return;
                        }
                        SetNativeOwner();
                        if (!progress.CanShow)
                        {
                            Application.ExitThread();
                            return;
                        }
                    }
                    catch (Win32Exception error)
                    {
                        Trace.TraceError("OrgLens loading display failed: " + error.Message);
                        Application.ExitThread();
                        return;
                    }
                }
                base.SetVisibleCore(value);
            }

            private void SetNativeOwner()
            {
                // Form.Show(owner) would access the other thread's managed Form. Set only the native owner.
                IntPtr handle = Handle;
                // Zero is also a valid previous owner, so clear and check the native last-error value.
                SetLastError(0);
                IntPtr previous = IntPtr.Size == 8 ? SetWindowLongPtr(handle, -8, progress.owner)
                    : new IntPtr(SetWindowLong32(handle, -8, progress.owner.ToInt32()));
                int error = Marshal.GetLastWin32Error();
                if (previous == IntPtr.Zero && error != 0) throw new Win32Exception(error);
            }

            protected override void WndProc(ref Message message)
            {
                if (message.Msg == 0x0021) // WM_MOUSEACTIVATE
                {
                    message.Result = new IntPtr(3); // MA_NOACTIVATE
                    return;
                }
                base.WndProc(ref message);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var border = new Pen(Color.FromArgb(203, 213, 225)))
                    e.Graphics.DrawRectangle(border, 0, 0, Width - 1, Height - 1);
                for (int dot = 0; dot < 12; dot++)
                {
                    double angle = dot * Math.PI / 6;
                    int distance = (dot - frame + 12) % 12;
                    using (var brush = new SolidBrush(Color.FromArgb(45 + (11 - distance) * 19, 37, 99, 235)))
                        e.Graphics.FillEllipse(brush, (float)(40 + 16 * Math.Cos(angle)) * scale,
                            (float)(49 + 16 * Math.Sin(angle)) * scale, 6 * scale, 6 * scale);
                }
                TextRenderer.DrawText(e.Graphics, progress.phase, phaseFont,
                    new Rectangle((int)(80 * scale), (int)(27 * scale), Width - (int)(96 * scale), (int)(27 * scale)),
                    Color.FromArgb(30, 41, 59), TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
                TextRenderer.DrawText(e.Graphics, "Please wait — this may take a moment.", detailFont,
                    new Rectangle((int)(80 * scale), (int)(57 * scale), Width - (int)(96 * scale), (int)(25 * scale)),
                    Color.FromArgb(100, 116, 139), TextFormatFlags.Left | TextFormatFlags.NoPrefix);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    animation?.Dispose();
                    phaseFont?.Dispose();
                    detailFont?.Dispose();
                }
                base.Dispose(disposing);
            }
        }
    }
}
