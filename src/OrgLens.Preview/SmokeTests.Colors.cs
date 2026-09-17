using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using OrgLens.Core;
using OrgLens.Desktop;

namespace OrgLens.Preview
{
    internal static partial class SmokeTests
    {
        private static void TestColorPickers()
        {
            using (var form = new SettingsForm(new ProbeService(), new TestDialogs()))
            {
                Show(form);
                var draw = typeof(ComboBox).GetMethod("OnDrawItem", BindingFlags.Instance | BindingFlags.NonPublic);
                foreach (string name in new[] { "BossRead", "BossUnread", "CustomUnread", "ToMe" })
                {
                    var picker = Find<ComboBox>(form, name + "Color");
                    Assert(picker.DrawMode == DrawMode.OwnerDrawFixed,
                        "Every color picker must draw actual color swatches, not just names.");
                    Assert(picker.DropDownStyle == ComboBoxStyle.DropDownList
                        && picker.Items.Count == 17 && !string.IsNullOrEmpty(picker.AccessibleName),
                        "Keep the native named, keyboard-selectable Outlook palette.");
                    for (int index = 0; index < picker.Items.Count; index++)
                    {
                        picker.SelectedIndex = index;
                        var color = (MailColor)picker.SelectedItem;
                        Assert(picker.Text == color.ToString(), "Color names must remain readable and accessible.");
                        foreach (var state in new[] { DrawItemState.None, DrawItemState.Selected | DrawItemState.Focus,
                            DrawItemState.ComboBoxEdit, DrawItemState.Disabled })
                        {
                            using (var bitmap = new Bitmap(220, picker.ItemHeight))
                            using (var graphics = Graphics.FromImage(bitmap))
                            {
                                var bounds = new Rectangle(Point.Empty, bitmap.Size);
                                draw.Invoke(picker, new object[] {
                                    new DrawItemEventArgs(graphics, picker.Font, bounds, index, state)
                                });
                                int inset = Math.Max(1, (int)Math.Round(4 * picker.DeviceDpi / 96.0));
                                Assert(bitmap.GetPixel(inset + 8, bitmap.Height / 2).ToArgb()
                                    == OutlookPalette.ToColor(color).ToArgb(),
                                    "Both selected display and dropdown must show the actual swatch: " + color + "/" + state);
                                if (color == MailColor.White)
                                    Assert(bitmap.GetPixel(inset, bitmap.Height / 2).ToArgb() != Color.White.ToArgb(),
                                        "White swatches need a visible border.");
                            }

                        }
                    }
                    using (var bitmap = new Bitmap(picker.Width, picker.Height))
                    {
                        picker.SelectedItem = MailColor.Red;
                        picker.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                        bool redVisible = false;
                        for (int y = 0; y < bitmap.Height; y++)
                            for (int x = 0; x < Math.Min(40, bitmap.Width); x++)
                                redVisible |= bitmap.GetPixel(x, y).ToArgb() == Color.Red.ToArgb();
                        Assert(redVisible, "The real collapsed native combo must visibly paint its selected color.");
                    }
                    using (var largeFont = new Font(picker.Font.FontFamily, picker.Font.Size * 2))
                    {
                        var original = picker.Font;
                        picker.Font = largeFont;
                        Assert(picker.ItemHeight > picker.Font.Height, "Swatches and labels must follow font scaling.");
                        picker.Font = original;
                    }
                }
                form.Close();
            }
        }
    }
}
