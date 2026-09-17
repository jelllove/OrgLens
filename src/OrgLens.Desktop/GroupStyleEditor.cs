using System;
using System.Drawing;
using System.Windows.Forms;
using OrgLens.Core;

namespace OrgLens.Desktop
{
    internal sealed class GroupStyleEditor : UserControl
    {
        private readonly CheckBox enabled;
        private readonly ComboBox color;
        private readonly CheckBox bold;
        private readonly CheckBox italic;
        private readonly NumericUpDown size;

        internal event EventHandler RuleStyleChanged;
        internal bool IsRuleEnabled { get { return enabled.Checked; } }

        internal GroupStyleEditor(string name, string scope)
        {
            Name = name + "Style";
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Dock = DockStyle.Top;
            Margin = new Padding(0, 4, 0, 4);
            var row = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 6 };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 58));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 54));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 65));
            enabled = Check(name + "Enabled", scope);
            color = new OutlookColorComboBox
            {
                Name = name + "Color", AccessibleName = name + " Outlook color", Dock = DockStyle.Fill,
                Margin = new Padding(0, 4, 6, 4)
            };
            bold = Check(name + "Bold", "Bold");
            italic = Check(name + "Italic", "Italic");
            size = new NumericUpDown
            {
                Name = name + "FontSize", AccessibleName = name + " font size in points",
                Minimum = 1, Maximum = 127, DecimalPlaces = 0, Value = 11, Dock = DockStyle.Fill,
                Margin = new Padding(0, 4, 0, 4)
            };
            row.Controls.Add(enabled, 0, 0);
            row.Controls.Add(color, 1, 0);
            row.Controls.Add(bold, 2, 0);
            row.Controls.Add(italic, 3, 0);
            row.Controls.Add(new Label
            {
                Text = "Size pt", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft,
                Margin = Padding.Empty
            }, 4, 0);
            row.Controls.Add(size, 5, 0);
            Controls.Add(row);
            enabled.CheckedChanged += Changed;
            color.SelectedIndexChanged += Changed;
            bold.CheckedChanged += Changed;
            italic.CheckedChanged += Changed;
            size.ValueChanged += Changed;
        }

        internal RuleStyle ReadStyle()
        {
            if (size.Value != decimal.Truncate(size.Value))
                throw new OrgLensException("Font size must be a whole number from 1 to 127 points.");
            return new RuleStyle
            {
                Enabled = enabled.Checked, Color = (MailColor)color.SelectedItem,
                Bold = bold.Checked, Italic = italic.Checked, FontSize = (int)size.Value
            };
        }

        internal void SetStyle(RuleStyle style)
        {
            enabled.Checked = style.Enabled;
            color.SelectedItem = style.Color;
            bold.Checked = style.Bold;
            italic.Checked = style.Italic;
            size.Value = style.FontSize;
        }

        internal string EditSignature
        {
            get { return string.Join("|", enabled.Checked, color.SelectedItem, bold.Checked, italic.Checked, size.Value); }
        }

        private static CheckBox Check(string name, string text)
        {
            return new CheckBox
            {
                Name = name, AccessibleName = name, Text = text, AutoSize = true,
                Dock = DockStyle.Fill, Margin = new Padding(0, 4, 2, 4)
            };
        }

        private void Changed(object sender, EventArgs e)
        {
            RuleStyleChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
