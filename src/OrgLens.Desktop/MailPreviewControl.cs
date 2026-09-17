using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OrgLens.Core;

namespace OrgLens.Desktop
{
    internal sealed class MailPreviewControl : ScrollableControl
    {
        private readonly List<PreviewRow> rows = new List<PreviewRow>();
        private readonly Color muted = Color.FromArgb(100, 116, 139);
        private int contentHeight;
        internal IReadOnlyList<PreviewRow> Rows { get { return rows; } }

        internal MailPreviewControl()
        {
            Name = "MessagePreview";
            AccessibleName = "Fictional messages illustrating grouped formatting";
            BackColor = Color.White;
            AutoScroll = true;
            DoubleBuffered = true;
            SetStyle(ControlStyles.ResizeRedraw, true);
        }

        internal void SetConfiguration(OrgLensConfiguration configuration)
        {
            ClearRows();
            bool custom = configuration.CustomAddresses.Count > 0;
            Add(configuration, "Jamie Rivera · Milestone planning", "BOSS · UNREAD · also directly To you", true, false, false, true);
            Add(configuration, "Morgan Chen · Project review", "BOSS · READ", true, false, true, false);
            Add(configuration, "Casey Reed · Partner update", custom ? "CUSTOM member · UNREAD" : "CUSTOM sample · add an email to enable matching",
                false, custom, false, false);
            Add(configuration, "Casey Reed · Previous update", "CUSTOM member · READ · not directly To / Cc you", false, custom, true, false);
            Add(configuration, "Riley Park · A question for you", "TO you explicitly · UNREAD", false, false, false, true);
            Add(configuration, "Robin Lane · Keeping you informed", "CC you explicitly · UNREAD", false, false, false, true);
            Add(configuration, "Riley Park · Your earlier reply", "TO you explicitly · READ", false, false, true, true);
            Add(configuration, "Team digest · News for everyone", "Distribution list only · UNREAD · excluded", false, false, false, false);
            Add(configuration, "Alex Kim · Private copy", "BCC only · UNREAD · excluded", false, false, false, false);
            AccessibleDescription = string.Join("; ", rows.Select(row =>
                row.Detail + ": " + row.StyleDescription)) + ". All people and messages here are fictional; no mailbox messages are read.";
            MeasureRows();
            Invalidate();
        }

        private void Add(OrgLensConfiguration config, string subject, string detail,
            bool boss, bool custom, bool read, bool direct)
        {
            var kind = RuleDefinition.MatchKind(config, boss, custom, read, direct);
            var style = kind.HasValue ? RuleDefinition.StyleFor(config, kind.Value) : null;
            rows.Add(new PreviewRow
            {
                Subject = subject, Detail = detail, Kind = kind,
                Color = style == null ? Color.FromArgb(51, 65, 85) : OutlookPalette.ToColor(style.Color),
                Font = new Font(Font.FontFamily, style?.FontSize ?? Font.Size,
                    (style != null && style.Bold ? FontStyle.Bold : FontStyle.Regular) |
                    (style != null && style.Italic ? FontStyle.Italic : FontStyle.Regular), GraphicsUnit.Point),
                StyleDescription = style == null ? "No OrgLens style"
                    : kind + " · " + style.Color + " · " + style.FontSize + " pt"
                        + (style.Bold ? " · bold" : "") + (style.Italic ? " · italic" : "")
            });
        }

        private void MeasureRows()
        {
            contentHeight = 0;
            int width = 0;
            foreach (var row in rows)
            {
                row.Height = Math.Max(82, row.Font.Height + Font.Height * 2 + 22);
                contentHeight += row.Height;
                width = Math.Max(width, TextRenderer.MeasureText(row.Subject, row.Font).Width + 32);
                width = Math.Max(width, TextRenderer.MeasureText(row.Detail, Font).Width + 32);
            }
            AutoScrollMinSize = new Size(width, contentHeight);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            int y = AutoScrollPosition.Y;
            int x = AutoScrollPosition.X;
            int width = Math.Max(ClientSize.Width, AutoScrollMinSize.Width);
            foreach (var row in rows)
            {
                if (y + row.Height >= 0 && y <= ClientSize.Height)
                {
                    using (var pen = new Pen(Color.FromArgb(226, 232, 240)))
                        e.Graphics.DrawLine(pen, x, y + row.Height - 1, x + width, y + row.Height - 1);
                    Draw(e.Graphics, row.Subject, row.Font, row.Color,
                        new Rectangle(x + 12, y + 8, width - 24, row.Font.Height + 6));
                    Draw(e.Graphics, row.Detail, Font, muted,
                        new Rectangle(x + 12, y + row.Font.Height + 15, width - 24, Font.Height + 2));
                    Draw(e.Graphics, row.StyleDescription, Font, muted,
                        new Rectangle(x + 12, y + row.Font.Height + Font.Height + 17, width - 24, Font.Height + 2));
                }
                y += row.Height;
            }
        }

        private static void Draw(Graphics graphics, string text, Font font, Color color, Rectangle bounds)
        {
            TextRenderer.DrawText(graphics, text, font, bounds, color,
                TextFormatFlags.Left | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
        }

        private void ClearRows()
        {
            foreach (var row in rows) row.Font.Dispose();
            rows.Clear();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) ClearRows();
            base.Dispose(disposing);
        }
    }

    internal sealed class PreviewRow
    {
        internal string Subject;
        internal string Detail;
        internal string StyleDescription;
        internal RuleKind? Kind;
        internal Font Font;
        internal Color Color;
        internal int Height;
    }
}
