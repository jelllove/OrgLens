using System;
using System.Collections.Generic;

namespace OrgLens.Core
{
    public sealed class ExistingRule
    {
        public ExistingRule(string name, bool standard)
        {
            Name = name;
            Standard = standard;
        }
        public string Name { get; }
        public bool Standard { get; }
    }

    public interface IFormattingRules
    {
        IReadOnlyList<ExistingRule> Read();
        void Remove(int oneBasedIndex);
        void Add(string name, string filter, RuleStyle style);
        void Save();
    }

    public static class RuleReconciler
    {
        public static void Replace(IFormattingRules target, IReadOnlyList<FormattingRule> desired)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            RuleDefinition.Validate(desired);
            if (desired.Count == 0)
                throw new OrgLensException("There are no formatting groups. Use Remove to remove OrgLens rules.");

            RemoveOwnedWithoutSaving(target);
            for (int i = 0; i < desired.Count; i++)
                target.Add(desired[i].Name, desired[i].Filter, desired[i].Style);
            target.Save();
        }

        public static void RemoveOwned(IFormattingRules target)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            RemoveOwnedWithoutSaving(target);
            target.Save();
        }

        private static void RemoveOwnedWithoutSaving(IFormattingRules target)
        {
            var existing = target.Read();
            for (int i = existing.Count - 1; i >= 0; i--)
                if (!existing[i].Standard && RuleDefinition.IsOwned(existing[i].Name)) target.Remove(i + 1);
        }
    }
}
