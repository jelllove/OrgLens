using System;
using OutlookApi = Microsoft.Office.Interop.Outlook;

namespace OrgLens.Outlook
{
    internal sealed class NativeRuleSnapshot
    {
        public int Index { get; private set; }
        public string Name { get; private set; }
        public string Filter { get; private set; }
        public bool Enabled { get; private set; }
        public string FontName { get; private set; }
        public int FontSize { get; private set; }
        public bool Bold { get; private set; }
        public bool Italic { get; private set; }
        public bool Underline { get; private set; }
        public bool Strikethrough { get; private set; }
        public OutlookApi.OlColor Color { get; private set; }
        public OutlookApi.OlCategoryColor ExtendedColor { get; private set; }

        public static NativeRuleSnapshot Capture(int index, OutlookApi.AutoFormatRule rule)
        {
            using (var scope = new ComScope())
            {
                var font = scope.Own(rule.Font);
                string filter = rule.Filter;
                return new NativeRuleSnapshot
                {
                    Index = index, Name = rule.Name, Filter = filter,
                    // Older versions could leave enabled match-all rules behind; never reactivate those.
                    Enabled = rule.Enabled && !string.IsNullOrWhiteSpace(filter),
                    FontName = font.Name, FontSize = font.Size, Bold = font.Bold, Italic = font.Italic,
                    Underline = font.Underline, Strikethrough = font.Strikethrough,
                    Color = font.Color, ExtendedColor = font.ExtendedColor
                };
            }
        }

        public void Restore(OutlookApi.AutoFormatRule rule)
        {
            rule.Enabled = false;
            rule.Filter = Filter;
            using (var scope = new ComScope())
            {
                var font = scope.Own(rule.Font);
                font.Name = FontName;
                font.Size = FontSize;
                font.Bold = Bold;
                font.Italic = Italic;
                font.Underline = Underline;
                font.Strikethrough = Strikethrough;
                font.Color = Color;
                if (ExtendedColor != OutlookApi.OlCategoryColor.olCategoryColorNone) font.ExtendedColor = ExtendedColor;
            }
            rule.Enabled = Enabled;
        }
    }
}
