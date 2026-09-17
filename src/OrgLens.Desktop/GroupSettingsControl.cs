using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OrgLens.Core;

namespace OrgLens.Desktop
{
    internal sealed class GroupSettingsControl : UserControl
    {
        private readonly GroupStyleEditor bossRead = new GroupStyleEditor("BossRead", "READ");
        private readonly GroupStyleEditor bossUnread = new GroupStyleEditor("BossUnread", "UNREAD");
        private readonly GroupStyleEditor custom = new GroupStyleEditor("CustomUnread", "UNREAD");
        private readonly GroupStyleEditor toMe = new GroupStyleEditor("ToMe", "Enabled");
        private readonly TextBox members;
        private readonly TextBox addresses;
        private readonly CheckBox includeRead;
        private readonly Label validation;
        private bool binding;

        internal event EventHandler ConfigurationChanged;

        internal GroupSettingsControl()
        {
            Name = "GroupedSettings";
            Dock = DockStyle.Fill;
            AutoScroll = true;
            var stack = new TableLayoutPanel
            {
                Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1,
                Padding = new Padding(0, 0, 10, 0)
            };
            stack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Controls.Add(stack);

            var boss = Group("BossGroup", "01  BOSS  ·  Your entire manager chain");
            Add(boss, Note("All discovered members share these two styles. Priority: BOSS > CUSTOM > TO ME."));
            members = new TextBox
            {
                Name = "BossMembers", AccessibleName = "Discovered BOSS group members, read only",
                ReadOnly = true, Multiline = true, ScrollBars = ScrollBars.Vertical, Height = 58,
                Dock = DockStyle.Top, BackColor = Color.FromArgb(248, 250, 252),
                Text = "No directory hierarchy loaded.", Margin = new Padding(0, 4, 0, 4)
            };
            Add(boss, members);
            Add(boss, bossRead);
            Add(boss, bossUnread);
            stack.Controls.Add(boss);

            var customGroup = Group("CustomGroup", "02  CUSTOM  ·  Senders you choose");
            Add(customGroup, Note("UNREAD only. Add, edit or remove email addresses below.\r\nSeparate with new lines, commas or semicolons; duplicates are removed."));
            addresses = new TextBox
            {
                Name = "CustomAddresses", AccessibleName = "Custom sender email addresses",
                Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical,
                Height = 64, Dock = DockStyle.Top, Margin = new Padding(0, 4, 0, 4)
            };
            Add(customGroup, addresses);
            validation = Note("");
            validation.Name = "AddressValidation";
            Add(customGroup, validation);
            Add(customGroup, custom);
            stack.Controls.Add(customGroup);

            var recipient = Group("ToMeGroup", "03  TO ME  ·  Explicit To / Cc");
            Add(recipient, Note("Your own address in To or Cc. Not distribution-list-only or BCC.\r\nUnread only by default; optionally include read messages."));
            includeRead = new CheckBox
            {
                Name = "ToMeIncludeRead", AccessibleName = "Include read messages in To me formatting",
                Text = "Include read messages (both READ and UNREAD)", AutoSize = true,
                Dock = DockStyle.Top, Margin = new Padding(0, 5, 0, 4)
            };
            Add(recipient, includeRead);
            Add(recipient, toMe);
            stack.Controls.Add(recipient);

            foreach (var editor in new[] { bossRead, bossUnread, custom, toMe })
                editor.RuleStyleChanged += Changed;
            addresses.TextChanged += Changed;
            includeRead.CheckedChanged += Changed;
            SetConfiguration(new OrgLensConfiguration());
        }

        internal bool BossEnabled { get { return bossRead.IsRuleEnabled || bossUnread.IsRuleEnabled; } }

        internal string EditSignature
        {
            get
            {
                return string.Join("\n", bossRead.EditSignature, bossUnread.EditSignature,
                    custom.EditSignature, toMe.EditSignature, includeRead.Checked, addresses.Text);
            }
        }

        internal void SetConfiguration(OrgLensConfiguration config)
        {
            binding = true;
            try
            {
                bossRead.SetStyle(config.BossRead);
                bossUnread.SetStyle(config.BossUnread);
                custom.SetStyle(config.CustomUnread);
                toMe.SetStyle(config.ToMe);
                includeRead.Checked = !config.ToMeUnreadOnly;
                addresses.Text = string.Join(Environment.NewLine, config.CustomAddresses);
            }
            finally { binding = false; }
            ValidateAddresses();
        }

        internal OrgLensConfiguration ReadConfiguration()
        {
            return ConfigurationValidator.Normalize(new OrgLensConfiguration
            {
                BossRead = bossRead.ReadStyle(), BossUnread = bossUnread.ReadStyle(),
                CustomUnread = custom.ReadStyle(), ToMe = toMe.ReadStyle(),
                ToMeUnreadOnly = !includeRead.Checked,
                CustomAddresses = ConfigurationValidator.ParseAddresses(addresses.Text)
            });
        }

        internal void ShowMembers(IReadOnlyList<ManagerPerson> people, string unavailable = null)
        {
            members.Text = people == null ? unavailable ?? "Hierarchy unavailable. Refresh to retry."
                : people.Count == 0 ? "No managers returned by the directory. The BOSS group is empty."
                : string.Join(Environment.NewLine, people.Select(person =>
                    person.DisplayName + "  ·  " + person.RoleLabel + "  ·  " + person.SmtpAddress));
        }

        private void Changed(object sender, EventArgs e)
        {
            if (binding) return;
            ValidateAddresses();
            ConfigurationChanged?.Invoke(this, EventArgs.Empty);
        }

        private void ValidateAddresses()
        {
            try
            {
                int count = ConfigurationValidator.ParseAddresses(addresses.Text).Count;
                validation.Text = count + " unique sender" + (count == 1 ? "" : "s") +
                    (count == 0 ? " · Add an address to activate CUSTOM matching." : " · UNREAD only");
                validation.ForeColor = Color.FromArgb(100, 116, 139);
            }
            catch (OrgLensException error)
            {
                validation.Text = error.Message;
                validation.ForeColor = Color.FromArgb(185, 28, 28);
            }
        }

        private static GroupBox Group(string name, string text)
        {
            var group = new GroupBox
            {
                Name = name, Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top, Padding = new Padding(12, 6, 12, 8), Margin = new Padding(0, 0, 0, 10)
            };
            var content = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1 };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            group.Controls.Add(content);
            return group;
        }

        private static void Add(GroupBox group, Control control)
        {
            var content = (TableLayoutPanel)group.Controls[0];
            control.Dock = DockStyle.Top;
            content.Controls.Add(control, 0, content.RowCount++);
        }

        private static Label Note(string text)
        {
            return new Label
            {
                Text = text, AutoSize = true, Dock = DockStyle.Top, UseMnemonic = false,
                ForeColor = Color.FromArgb(100, 116, 139), Margin = new Padding(0, 2, 0, 2)
            };
        }
    }
}
