using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using OrgLens.Core;

namespace OrgLens.Desktop
{
    public sealed partial class SettingsForm : Form
    {
        private static readonly Color Ink = Color.FromArgb(30, 41, 59);
        private static readonly Color Muted = Color.FromArgb(100, 116, 139);
        private static readonly Color Accent = Color.FromArgb(37, 99, 235);
        private readonly IOrgLensService service;
        private readonly ISettingsDialogs dialogs;
        private readonly List<Font> ownedFonts = new List<Font>();
        private readonly ToolTip tooltips = new ToolTip { AutoPopDelay = 20000 };
        private readonly Icon windowIcon;
        private readonly ComboBox accountSelector;
        private readonly GroupSettingsControl editor;
        private readonly MailPreviewControl preview;
        private readonly Label noticeLabel;
        private readonly Label fileLabel;
        private readonly Label documentLabel;
        private readonly Label statusLabel;
        private readonly Label scopeLabel;
        private readonly Button refreshButton;
        private readonly Button applyButton;
        private readonly Button removeButton;
        private readonly Button closeButton;
        private readonly Button openButton;
        private readonly Button saveButton;
        private readonly Button saveAsButton;
        private MailboxAccount activeAccount;
        private IReadOnlyList<ManagerPerson> managers;
        private string filePath;
        private string cleanSignature;
        private string appliedSignature;
        private bool suppressEvents;
        private bool busy;
        private bool initialized;
        private bool accountConfigLoaded;
        private LoadingProgress loadingProgress;

        public SettingsForm(IOrgLensService service) : this(service, new SettingsDialogs()) { }

        internal SettingsForm(IOrgLensService service, ISettingsDialogs dialogs)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("OrgLens settings must be opened on an STA UI thread.");
            this.service = service ?? throw new ArgumentNullException(nameof(service));
            this.dialogs = dialogs ?? throw new ArgumentNullException(nameof(dialogs));
            Name = "OrgLensSettings";
            Text = service.IsDemo ? "OrgLens · Sample preview" : "OrgLens · Inbox formatting";
            Font = OwnFont(9.5F, FontStyle.Regular);
            ForeColor = Ink;
            BackColor = Color.White;
            ClientSize = new Size(1260, 940);
            MinimumSize = new Size(1100, 780);
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            windowIcon = OrgLensImages.CreateWindowIcon();
            Icon = windowIcon;
            ShowIcon = true;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, Padding = new Padding(22, 14, 22, 12),
                ColumnCount = 1, RowCount = 9
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            foreach (int height in new[] { 58, service.IsDemo ? 32 : 0, 44, 44, 44, -1, 42, 46, 24 })
                root.RowStyles.Add(new RowStyle(height < 0 ? SizeType.Percent : SizeType.Absolute, height < 0 ? 100 : height));
            Controls.Add(root);

            var header = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, Margin = Padding.Empty };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112));
            header.Controls.Add(Label("OrgLens  /  Inbox formatting", OwnFont(21, FontStyle.Bold), Ink), 0, 0);
            openButton = Button("&Open settings", "OpenSettings", false);
            openButton.Click += (sender, args) => OpenSettings();
            saveButton = Button("&Save", "SaveSettings", false);
            saveButton.Click += (sender, args) => SaveSettings(false);
            saveAsButton = Button("Save &as…", "SaveSettingsAs", false);
            saveAsButton.Click += (sender, args) => SaveSettings(true);
            header.Controls.Add(openButton, 1, 0);
            header.Controls.Add(saveButton, 2, 0);
            header.Controls.Add(saveAsButton, 3, 0);
            root.Controls.Add(header, 0, 0);

            var banner = Label("SAMPLE-ONLY DEMO  ·  Fictional people and messages. Outlook is never accessed or changed.",
                OwnFont(9.5F, FontStyle.Bold), Color.FromArgb(30, 64, 175));
            banner.Name = "DemoBanner";
            banner.BackColor = Color.FromArgb(239, 246, 255);
            banner.Padding = new Padding(10, 0, 0, 0);
            banner.Visible = service.IsDemo;
            root.Controls.Add(banner, 0, 1);

            var document = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, Margin = Padding.Empty };
            document.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            document.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            fileLabel = Label("", Font, Muted);
            fileLabel.Name = "SettingsFilePath";
            fileLabel.AutoEllipsis = true;
            documentLabel = Label("", Font, Accent);
            documentLabel.Name = "SettingsDocumentState";
            document.Controls.Add(fileLabel, 0, 0);
            document.Controls.Add(documentLabel, 0, 1);
            root.Controls.Add(document, 0, 2);

            var accountRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, Margin = Padding.Empty };
            accountRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            accountRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            accountRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 176));
            accountRow.Controls.Add(Label("Account", Font, Ink), 0, 0);
            accountSelector = new ComboBox
            {
                Name = "AccountSelector", AccessibleName = "Mailbox account", Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 9, 12, 0)
            };
            accountSelector.SelectedIndexChanged += AccountChanged;
            accountRow.Controls.Add(accountSelector, 1, 0);
            refreshButton = Button("&Refresh hierarchy", "RefreshHierarchy", false);
            refreshButton.Click += (sender, args) => RefreshHierarchy();
            accountRow.Controls.Add(refreshButton, 2, 0);
            root.Controls.Add(accountRow, 0, 3);

            var directory = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, Margin = Padding.Empty };
            var source = Label(service.IsDemo ? "SOURCE · Fictional example.com directory"
                : "SOURCE · Exchange address book · Entire manager chain", Font, Muted);
            source.Name = "HierarchySource";
            noticeLabel = Label("No hierarchy loaded. Settings can be edited, opened and saved independently.", Font, Muted);
            noticeLabel.Name = "HierarchyNotice";
            noticeLabel.AutoEllipsis = true;
            directory.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
            directory.RowStyles.Add(new RowStyle(SizeType.Percent, 55));
            directory.Controls.Add(source, 0, 0);
            directory.Controls.Add(noticeLabel, 0, 1);
            root.Controls.Add(directory, 0, 4);

            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Margin = new Padding(0, 8, 0, 0) };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            editor = new GroupSettingsControl { Font = Font, Margin = Padding.Empty };
            editor.ConfigurationChanged += EditorChanged;
            body.Controls.Add(editor, 0, 0);
            var previewHost = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Margin = new Padding(14, 0, 0, 0)
            };
            previewHost.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            previewHost.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var previewTitle = Label("LIVE PREVIEW  ·  Fictional messages only\r\nScroll to compare read, unread, To / Cc and excluded cases.", Font, Muted);
            previewTitle.Name = "PreviewHeading";
            previewHost.Controls.Add(previewTitle, 0, 0);
            preview = new MailPreviewControl { Dock = DockStyle.Fill, Font = Font, Margin = Padding.Empty };
            previewHost.Controls.Add(preview, 0, 1);
            body.Controls.Add(previewHost, 1, 0);
            root.Controls.Add(body, 0, 5);

            scopeLabel = Label("", Font, Muted);
            scopeLabel.Name = "ScopeNotice";
            root.Controls.Add(scopeLabel, 0, 6);
            var actions = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, Margin = Padding.Empty };
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 208));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 184));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
            removeButton = Button(service.IsDemo ? "Simulate &remove" : "&Remove OrgLens rules", "RemoveRules", false);
            removeButton.Click += (sender, args) => RemoveRules();
            applyButton = Button(service.IsDemo ? "Simulate &apply" : "&Apply to Inbox", "ApplyRules", true);
            applyButton.Click += (sender, args) => ApplyRules();
            closeButton = Button("&Close", "CloseSettings", false);
            closeButton.Click += (sender, args) => Close();
            actions.Controls.Add(removeButton, 0, 0);
            actions.Controls.Add(applyButton, 2, 0);
            actions.Controls.Add(closeButton, 3, 0);
            root.Controls.Add(actions, 0, 7);
            CancelButton = closeButton;
            statusLabel = Label("Ready · Open or save settings without Outlook.", Font, Muted);
            statusLabel.Name = "OperationStatus";
            root.Controls.Add(statusLabel, 0, 8);
            cleanSignature = editor.EditSignature;
            UpdatePreview();
            UpdateActions();
        }

        internal bool HasUnsavedChanges { get { return editor.EditSignature != cleanSignature; } }
        internal string SelectedFilePath { get { return filePath; } }

        protected override void OnShown(EventArgs e)
        {
            if (!initialized)
            {
                initialized = true;
                BeginInvoke((MethodInvoker)(() =>
                {
                    if (IsDisposed || Disposing || !Visible) return;
                    // Paint all settings controls before blocking the Outlook STA on directory work.
                    Refresh();
                    LoadAccounts();
                }));
            }
            base.OnShown(e);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (busy || !ConfirmDiscard("Close OrgLens")) e.Cancel = true;
            base.OnFormClosing(e);
        }

        private static Label Label(string text, Font font, Color color)
        {
            return new Label
            {
                Text = text, Font = font, ForeColor = color, Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft, Margin = Padding.Empty, UseMnemonic = false
            };
        }

        private Button Button(string text, string name, bool primary)
        {
            var button = new Button
            {
                Text = text, Name = name, Dock = DockStyle.Fill, Margin = new Padding(0, 6, 8, 6),
                FlatStyle = FlatStyle.Flat, BackColor = primary ? Accent : Color.White,
                ForeColor = primary ? Color.White : Ink, Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false, Font = OwnFont(9.5F, primary ? FontStyle.Bold : FontStyle.Regular)
            };
            button.FlatAppearance.BorderColor = primary ? Accent : Color.FromArgb(203, 213, 225);
            return button;
        }

        private Font OwnFont(float size, FontStyle style)
        {
            var font = new Font("Segoe UI", size, style, GraphicsUnit.Point);
            ownedFonts.Add(font);
            return font;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) loadingProgress?.Dispose();
            base.Dispose(disposing);
            if (!disposing) return;
            tooltips.Dispose();
            windowIcon?.Dispose();
            foreach (var font in ownedFonts) font.Dispose();
            ownedFonts.Clear();
        }
    }
}
