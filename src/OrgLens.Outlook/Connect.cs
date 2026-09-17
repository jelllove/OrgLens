using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OrgLens.Core;
using OrgLens.Desktop;
using Office = Microsoft.Office.Core;
using OutlookApi = Microsoft.Office.Interop.Outlook;

[assembly: ComVisible(false)]

namespace OrgLens.Outlook
{
    [ComVisible(true)]
    [Guid("20FA2F90-F84D-40E7-A36D-B2F484B2B58E")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonCallbacks
    {
        [DispId(1)]
        void OpenSettings([In, MarshalAs(UnmanagedType.Interface)] Office.IRibbonControl control);
    }

    [ComVisible(true)]
    [Guid("D7E2D48A-9466-4D83-856F-AC197BC23A98")]
    [ProgId("OrgLens.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonCallbacks))]
    public sealed class Connect : IDTExtensibility2, Office.IRibbonExtensibility, IRibbonCallbacks
    {
        private OutlookApi.Application application;
        private ModelessSettingsWindow settingsWindow;

        public void OnConnection(object application, ConnectMode connectMode, object addInInstance, ref Array custom)
        {
            this.application = (OutlookApi.Application)application;
        }

        public void OnDisconnection(DisconnectMode removeMode, ref Array custom)
        {
            Disconnect();
        }

        private void Disconnect()
        {
            application = null;
            var window = settingsWindow;
            settingsWindow = null;
            window?.Dispose();
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { Disconnect(); }

        public string GetCustomUI(string ribbonId)
        {
            if (ribbonId != "Microsoft.Outlook.Explorer") return null;
            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                "<ribbon><tabs><tab id=\"OrgLensTab\" label=\"OrgLens\">" +
                "<group id=\"OrgLensFormatting\" label=\"Inbox formatting\">" +
                "<button id=\"OrgLensSettings\" label=\"Formatting rules\" size=\"large\" " +
                "imageMso=\"ConditionalFormattingHighlightCellsRules\" onAction=\"OpenSettings\" " +
                "screentip=\"OrgLens\" supertip=\"Format BOSS, custom senders, and messages addressed to you.\"/>" +
                "</group></tab></tabs></ribbon></customUI>";
        }

        public void OpenSettings(Office.IRibbonControl control)
        {
            if (application == null)
            {
                MessageBox.Show("OrgLens is not connected to Outlook. Restart classic Outlook.",
                    "OrgLens", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                if (settingsWindow == null)
                    settingsWindow = new ModelessSettingsWindow(
                        () => new SettingsForm(new OutlookOrgLensService(application)));
                settingsWindow.Show();
            }
            catch (COMException error)
            {
                MessageBox.Show("Outlook could not open OrgLens. " + error.Message,
                    "OrgLens", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (OrgLensException error)
            {
                MessageBox.Show(error.Message, "OrgLens", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
