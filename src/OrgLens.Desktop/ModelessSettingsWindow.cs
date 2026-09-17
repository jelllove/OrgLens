using System;
using System.Windows.Forms;

namespace OrgLens.Desktop
{
    public sealed class ModelessSettingsWindow : IDisposable
    {
        private readonly Func<SettingsForm> createForm;
        private SettingsForm form;
        private bool disposed;

        public ModelessSettingsWindow(Func<SettingsForm> createForm)
        {
            this.createForm = createForm ?? throw new ArgumentNullException(nameof(createForm));
        }

        public void Show()
        {
            if (disposed) throw new ObjectDisposedException(nameof(ModelessSettingsWindow));
            if (form == null)
            {
                form = createForm();
                form.Disposed += FormDisposed;
                bool shown = false;
                try
                {
                    form.Show();
                    shown = true;
                }
                finally
                {
                    if (!shown) ReleaseForm();
                }
                return;
            }
            if (form.WindowState == FormWindowState.Minimized)
                form.WindowState = FormWindowState.Normal;
            form.Activate();
        }

        private void FormDisposed(object sender, EventArgs e)
        {
            var closed = (SettingsForm)sender;
            closed.Disposed -= FormDisposed;
            if (ReferenceEquals(form, closed)) form = null;
        }

        private void ReleaseForm()
        {
            var closing = form;
            form = null;
            if (closing == null) return;
            closing.Disposed -= FormDisposed;
            // Host shutdown cannot be cancelled by this window's unsaved-edit prompt.
            closing.Dispose();
        }

        public void Dispose()
        {
            disposed = true;
            ReleaseForm();
        }
    }
}
