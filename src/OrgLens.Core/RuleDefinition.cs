using System;
using System.Collections.Generic;
using System.Linq;

namespace OrgLens.Core
{
    public static class RuleDefinition
    {
        public const string ViewName = "OrgLens - Inbox";
        public const string RulePrefix = "OrgLens.v1:";
        public const string ReadProperty = "urn:schemas:httpmail:read";
        public const string ToMeProperty = "http://schemas.microsoft.com/mapi/proptag/0x0057000B";
        public const string CcMeProperty = "http://schemas.microsoft.com/mapi/proptag/0x0058000B";
        public const string NeverMatch = "(\"urn:schemas:httpmail:read\" IS NULL AND \"urn:schemas:httpmail:read\" IS NOT NULL)";

        public static IReadOnlyList<FormattingRule> Build(OrgLensConfiguration configuration,
            IReadOnlyList<ManagerPerson> managers, IReadOnlyList<ManagerPerson> customSenders = null)
        {
            var config = ConfigurationValidator.Normalize(configuration);
            if (managers == null && config.BossEnabled)
                throw new OrgLensException("The BOSS hierarchy is unavailable. Refresh it, or disable both BOSS " +
                    "styles before applying only Custom / To me rules.");
            string boss = config.BossEnabled ? SenderGroup(managers ?? new ManagerPerson[0]) : NeverMatch;
            var custom = customSenders ?? config.CustomAddresses.Select(address =>
                new ManagerPerson(address, address, address, "")).ToList();
            string customGroup = config.CustomUnread.Enabled ? SenderGroup(custom) : NeverMatch;
            string bossUnread = And(boss, ReadState(false));
            string bossRead = And(boss, ReadState(true));
            string customUnread = And(customGroup, ReadState(false));
            string customExclusive = config.BossUnread.Enabled ? And(customUnread, Not(bossUnread)) : customUnread;
            string toMe = "(" + EqualsValue(ToMeProperty, "1") + " OR " + EqualsValue(CcMeProperty, "1") + ")";
            if (config.ToMeUnreadOnly) toMe = And(toMe, ReadState(false));
            if (config.BossUnread.Enabled) toMe = And(toMe, Not(bossUnread));
            if (config.BossRead.Enabled) toMe = And(toMe, Not(bossRead));
            if (config.CustomUnread.Enabled) toMe = And(toMe, Not(customUnread));
            return new[]
            {
                new FormattingRule(RuleKind.BossUnread, bossUnread, config.BossUnread),
                new FormattingRule(RuleKind.BossRead, bossRead, config.BossRead),
                new FormattingRule(RuleKind.CustomUnread, customExclusive, config.CustomUnread),
                new FormattingRule(RuleKind.ToMe, toMe, config.ToMe)
            };
        }

        public static RuleKind? MatchKind(OrgLensConfiguration config, bool fromBoss, bool fromCustom,
            bool isRead, bool directlyToOrCcMe)
        {
            if (fromBoss && !isRead && config.BossUnread.Enabled) return RuleKind.BossUnread;
            if (fromBoss && isRead && config.BossRead.Enabled) return RuleKind.BossRead;
            if (fromCustom && !isRead && config.CustomUnread.Enabled) return RuleKind.CustomUnread;
            if (directlyToOrCcMe && config.ToMe.Enabled && (!config.ToMeUnreadOnly || !isRead)) return RuleKind.ToMe;
            return null;
        }

        public static RuleStyle StyleFor(OrgLensConfiguration config, RuleKind kind)
        {
            switch (kind)
            {
                case RuleKind.BossUnread: return config.BossUnread;
                case RuleKind.BossRead: return config.BossRead;
                case RuleKind.CustomUnread: return config.CustomUnread;
                case RuleKind.ToMe: return config.ToMe;
                default: throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }

        public static string NameFor(ManagerPerson person)
        {
            if (person == null) throw new ArgumentNullException(nameof(person));
            string identity = !string.IsNullOrWhiteSpace(person.SmtpAddress) ? person.SmtpAddress
                : !string.IsNullOrWhiteSpace(person.LegacyAddress) ? person.LegacyAddress : person.EntryId;
            if (string.IsNullOrWhiteSpace(identity))
                throw new OrgLensException("A manager has no stable identity for a formatting rule.");
            return RulePrefix + StableIdentifier.Hash(identity.Trim().ToLowerInvariant());
        }

        public static bool IsOwned(string name)
        {
            if (IsGroupedName(name)) return true;
            if (name == null || !name.StartsWith(RulePrefix, StringComparison.Ordinal)) return false;
            var suffix = name.Substring(RulePrefix.Length);
            return suffix.Length == 64 && suffix.All(c => (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F'));
        }

        public static bool IsGroupedName(string name)
        {
            return name == "OrgLens.v2:BossUnread" || name == "OrgLens.v2:BossRead" ||
                name == "OrgLens.v2:CustomUnread" || name == "OrgLens.v2:ToMe";
        }

        public static string FilterFor(ManagerPerson person)
        {
            if (person == null) throw new ArgumentNullException(nameof(person));
            var clauses = new List<string>();
            if (!string.IsNullOrWhiteSpace(person.SmtpAddress))
            {
                string smtp = ConfigurationValidator.NormalizeAddress(person.SmtpAddress);
                Add(clauses, "0x5D01001F", smtp);
                Add(clauses, "0x5D01001E", smtp);
                Add(clauses, "0x0C1F001F", smtp);
                Add(clauses, "0x0C1F001E", smtp);
            }
            if (!string.IsNullOrWhiteSpace(person.LegacyAddress))
            {
                if (person.LegacyAddress.Any(char.IsControl))
                    throw new OrgLensException("Invalid Exchange address for " + person.DisplayName + ".");
                Add(clauses, "0x0C1F001F", person.LegacyAddress);
                Add(clauses, "0x0C1F001E", person.LegacyAddress);
            }
            if (clauses.Count == 0)
                throw new OrgLensException("No sender address is published for " + person.DisplayName +
                    ". Disable the BOSS group or ask your directory administrator to fix the entry.");
            // AutoFormatRule.Filter expects DASL without the @SQL= prefix used by Items.Restrict.
            return "(" + string.Join(" OR ", clauses) + ")";
        }

        public static void Validate(IReadOnlyList<FormattingRule> rules)
        {
            if (rules == null) throw new ArgumentNullException(nameof(rules));
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (var rule in rules)
            {
                if (rule == null) throw new OrgLensException("A formatting rule is missing.");
                ConfigurationValidator.ValidateStyle(rule.Style);
                if (!IsGroupedName(rule.Name) || !names.Add(rule.Name))
                    throw new OrgLensException("The formatting plan contains an unknown or duplicate group.");
                if (string.IsNullOrWhiteSpace(rule.Filter))
                    throw new OrgLensException("A formatting filter is empty; it would match every message.");
            }
        }

        private static string SenderGroup(IEnumerable<ManagerPerson> senders)
        {
            var filters = senders.Select(FilterFor).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            return filters.Count == 0 ? NeverMatch : "(" + string.Join(" OR ", filters) + ")";
        }

        private static void Add(ICollection<string> clauses, string property, string value)
        {
            clauses.Add(EqualsValue("http://schemas.microsoft.com/mapi/proptag/" + property,
                "'" + value.Replace("'", "''") + "'"));
        }

        // A missing property must be false, not SQL UNKNOWN, when higher-priority rules are negated.
        private static string EqualsValue(string property, string literal)
        {
            return "(\"" + property + "\" IS NOT NULL AND \"" + property + "\" = " + literal + ")";
        }
        private static string ReadState(bool read) { return EqualsValue(ReadProperty, read ? "1" : "0"); }
        private static string And(string left, string right) { return "(" + left + " AND " + right + ")"; }
        private static string Not(string filter) { return "(NOT " + filter + ")"; }
    }
}
