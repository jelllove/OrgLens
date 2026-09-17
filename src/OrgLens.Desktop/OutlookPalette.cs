using System.Drawing;
using OrgLens.Core;

namespace OrgLens.Desktop
{
    internal static class OutlookPalette
    {
        internal static Color ToColor(MailColor color)
        {
            switch (color)
            {
                case MailColor.Automatic: return SystemColors.WindowText;
                case MailColor.Black: return Color.FromArgb(0, 0, 0);
                case MailColor.Maroon: return Color.FromArgb(128, 0, 0);
                case MailColor.Green: return Color.FromArgb(0, 128, 0);
                case MailColor.Olive: return Color.FromArgb(128, 128, 0);
                case MailColor.Navy: return Color.FromArgb(0, 0, 128);
                case MailColor.Purple: return Color.FromArgb(128, 0, 128);
                case MailColor.Teal: return Color.FromArgb(0, 128, 128);
                case MailColor.Gray: return Color.FromArgb(128, 128, 128);
                case MailColor.Silver: return Color.FromArgb(192, 192, 192);
                case MailColor.Red: return Color.FromArgb(255, 0, 0);
                case MailColor.Lime: return Color.FromArgb(0, 255, 0);
                case MailColor.Yellow: return Color.FromArgb(255, 255, 0);
                case MailColor.Blue: return Color.FromArgb(0, 0, 255);
                case MailColor.Fuchsia: return Color.FromArgb(255, 0, 255);
                case MailColor.Aqua: return Color.FromArgb(0, 255, 255);
                case MailColor.White: return Color.FromArgb(255, 255, 255);
                default: throw new OrgLensException("Choose a supported Outlook font color.");
            }
        }
    }
}
