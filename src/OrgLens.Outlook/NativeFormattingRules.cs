using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OrgLens.Core;
using OutlookApi = Microsoft.Office.Interop.Outlook;

namespace OrgLens.Outlook
{
    internal sealed class NativeFormattingRules : IFormattingRules
    {
        private readonly OutlookApi.AutoFormatRules rules;
        private readonly List<ExpectedCondition> expected = new List<ExpectedCondition>();

        public NativeFormattingRules(OutlookApi.AutoFormatRules rules)
        {
            this.rules = rules;
        }

        public IReadOnlyList<ExistingRule> Read()
        {
            var result = new List<ExistingRule>();
            for (int i = 1; i <= rules.Count; i++)
            {
                using (var scope = new ComScope())
                {
                    var rule = scope.Own(rules[i]);
                    result.Add(new ExistingRule(rule.Name, rule.Standard));
                }
            }
            return result;
        }

        public void Remove(int oneBasedIndex)
        {
            rules.Remove(oneBasedIndex);
            expected.RemoveAll(item => item.Index == oneBasedIndex);
            foreach (var item in expected)
                if (item.Index > oneBasedIndex) item.Index--;
        }

        public void Add(string name, string filter, RuleStyle style)
        {
            using (var scope = new ComScope())
            {
                var rule = scope.Own(rules.Add(name));
                // A newly added rule has an empty filter, which would match every message.
                rule.Enabled = false;
                rule.Filter = filter;
                var font = scope.Own(rule.Font);
                font.Color = (OutlookApi.OlColor)(int)style.Color;
                font.Bold = style.Bold;
                font.Italic = style.Italic;
                font.Size = style.FontSize;
                rule.Enabled = style.Enabled;
                expected.Add(new ExpectedCondition(rules.Count, name, filter, style.Enabled));
            }
        }

        public void Save()
        {
            rules.Save();
            foreach (var wanted in expected)
            {
                using (var scope = new ComScope())
                {
                    var actual = scope.Own(rules[wanted.Index]);
                    VerifyCondition(actual, wanted.Name, wanted.Filter, wanted.Enabled);
                }
            }
        }

        public void Verify(IReadOnlyList<FormattingRule> planned)
        {
            foreach (var wanted in planned)
            {
                using (var scope = new ComScope())
                {
                    var actual = scope.Own(rules[wanted.Name]);
                    VerifyCondition(actual, wanted.Name, wanted.Filter, wanted.Style.Enabled);
                }
            }
        }

        private void VerifyCondition(OutlookApi.AutoFormatRule actual, string name, string filter, bool enabled)
        {
            if (actual.Name == name && actual.Enabled == enabled &&
                string.Equals(actual.Filter, filter, StringComparison.Ordinal)) return;
            try { DisableOwnedRules(); }
            catch (COMException error)
            {
                throw new OrgLensException("Outlook lost an OrgLens condition and could not disable the affected rules. " +
                    "Switch to your original view under View > Change View before continuing.", error);
            }
            throw new OrgLensException("Outlook did not retain an OrgLens condition. OrgLens rules were disabled " +
                "rather than left matching every message. No successful Apply is confirmed.");
        }

        public IReadOnlyList<NativeRuleSnapshot> CaptureOwned()
        {
            var result = new List<NativeRuleSnapshot>();
            for (int i = 1; i <= rules.Count; i++)
            {
                using (var scope = new ComScope())
                {
                    var rule = scope.Own(rules[i]);
                    if (!rule.Standard && RuleDefinition.IsOwned(rule.Name)) result.Add(NativeRuleSnapshot.Capture(i, rule));
                }
            }
            return result;
        }

        public void RestoreOwned(IReadOnlyList<NativeRuleSnapshot> snapshot)
        {
            var existing = Read();
            for (int i = existing.Count - 1; i >= 0; i--)
                if (!existing[i].Standard && RuleDefinition.IsOwned(existing[i].Name)) Remove(i + 1);
            expected.Clear();
            foreach (var saved in snapshot)
            {
                using (var scope = new ComScope())
                {
                    var rule = scope.Own(saved.Index > rules.Count ? rules.Add(saved.Name) : rules.Insert(saved.Name, saved.Index));
                    saved.Restore(rule);
                    expected.Add(new ExpectedCondition(saved.Index, saved.Name, saved.Filter, saved.Enabled));
                }
            }
            Save();
        }

        private void DisableOwnedRules()
        {
            for (int i = 1; i <= rules.Count; i++)
            {
                using (var scope = new ComScope())
                {
                    var rule = scope.Own(rules[i]);
                    if (!rule.Standard && RuleDefinition.IsOwned(rule.Name)) rule.Enabled = false;
                }
            }
            rules.Save();
        }

        private sealed class ExpectedCondition
        {
            public ExpectedCondition(int index, string name, string filter, bool enabled)
            {
                Index = index; Name = name; Filter = filter; Enabled = enabled;
            }
            public int Index { get; set; }
            public string Name { get; }
            public string Filter { get; }
            public bool Enabled { get; }
        }
    }
}
