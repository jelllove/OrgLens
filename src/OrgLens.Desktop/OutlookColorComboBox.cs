using System;
using System.Drawing;
using System.Windows.Forms;
using OrgLens.Core;

namespace OrgLens.Desktop
{
    internal sealed class OutlookColorComboBox : ComboBox
    {
        internal OutlookColorComboBox()
        {
            DropDownStyle = ComboBoxStyle.DropDownList;
            DrawMode = DrawMode.OwnerDrawFixed;
            foreach (MailColor color in Enum.GetValues(typeof(MailColor))) Items.Add(color);
            SelectedIndex = 0;
            UpdateItemHeight();
        }

        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            UpdateItemHeight();
        }

        protected override void OnDpiChangedAfterParent(EventArgs e)
        {
            base.OnDpiChangedAfterParent(e);
            UpdateItemHeight();
        }

        private int ScalePixels(int pixels) { return Math.Max(1, (int)Math.Round(pixels * DeviceDpi / 96.0)); }
        private void UpdateItemHeight() { ItemHeight = Font.Height + ScalePixels(6); }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            base.OnDrawItem(e);
            e.DrawBackground();
            if (e.Index >= 0)
            {
                int inset = ScalePixels(4);
                int height = Math.Max(1, Math.Min(Font.Height, e.Bounds.Height - inset));
                var swatch = new Rectangle(e.Bounds.Left + inset,
                    e.Bounds.Top + (e.Bounds.Height - height) / 2, ScalePixels(24), height);
                using (var brush = new SolidBrush(OutlookPalette.ToColor((MailColor)Items[e.Index])))
                    e.Graphics.FillRectangle(brush, swatch);
                using (var border = new Pen(SystemColors.GrayText))
                    e.Graphics.DrawRectangle(border, swatch);
                var textBounds = Rectangle.FromLTRB(swatch.Right + ScalePixels(6), e.Bounds.Top,
                    e.Bounds.Right, e.Bounds.Bottom);
                var textColor = !Enabled || (e.State & DrawItemState.Disabled) != 0
                    ? SystemColors.GrayText : e.ForeColor;
                TextRenderer.DrawText(e.Graphics, GetItemText(Items[e.Index]), e.Font, textBounds, textColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine
                    | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            }
            e.DrawFocusRectangle();
        }
    }
}
