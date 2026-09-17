using System;
using System.Collections.Generic;
using System.Linq;
using OrgLens.Core;

namespace OrgLens.Tests
{
    internal static partial class Program
    {
        private const string ReadTag = "urn:schemas:httpmail:read";
        private const string ToTag = "http://schemas.microsoft.com/mapi/proptag/0x0057000B";
        private const string CcTag = "http://schemas.microsoft.com/mapi/proptag/0x0058000B";
        private static readonly string[] SenderTags =
        {
            "http://schemas.microsoft.com/mapi/proptag/0x5D01001F",
            "http://schemas.microsoft.com/mapi/proptag/0x5D01001E",
            "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F",
            "http://schemas.microsoft.com/mapi/proptag/0x0C1F001E"
        };

        private static void CheckDaslEvaluator()
        {
            Check("independent DASL evaluator implements SQL three-valued boolean tables", () =>
            {
                var values = new[] { SqlTruth.False, SqlTruth.True, SqlTruth.Unknown };
                var and = new[,]
                {
                    { SqlTruth.False, SqlTruth.False, SqlTruth.False },
                    { SqlTruth.False, SqlTruth.True, SqlTruth.Unknown },
                    { SqlTruth.False, SqlTruth.Unknown, SqlTruth.Unknown }
                };
                var or = new[,]
                {
                    { SqlTruth.False, SqlTruth.True, SqlTruth.Unknown },
                    { SqlTruth.True, SqlTruth.True, SqlTruth.True },
                    { SqlTruth.Unknown, SqlTruth.True, SqlTruth.Unknown }
                };
                var not = new[] { SqlTruth.True, SqlTruth.False, SqlTruth.Unknown };
                for (int left = 0; left < values.Length; left++)
                {
                    for (int right = 0; right < values.Length; right++)
                    {
                        var message = new Dictionary<string, object>
                        {
                            ["left"] = left == 2 ? null : (object)left,
                            ["right"] = right == 2 ? null : (object)right
                        };
                        Equal(and[left, right],
                            DaslExpression.Parse("\"left\" = 1 AND \"right\" = 1").Evaluate(message));
                        Equal(or[left, right],
                            DaslExpression.Parse("\"left\" = 1 OR \"right\" = 1").Evaluate(message));
                        Equal(not[left], DaslExpression.Parse("NOT \"left\" = 1").Evaluate(message));
                        Equal(left == 2 ? SqlTruth.True : SqlTruth.False,
                            DaslExpression.Parse("\"left\" IS NULL").Evaluate(message));
                        Equal(left == 2 ? SqlTruth.False : SqlTruth.True,
                            DaslExpression.Parse("\"left\" IS NOT NULL").Evaluate(message));
                    }
                }
                var absent = new Dictionary<string, object>();
                Equal(SqlTruth.Unknown, DaslExpression.Parse("NOT \"absent\" = 1").Evaluate(absent));
                Equal(SqlTruth.True, DaslExpression.Parse("\"absent\" IS NULL").Evaluate(absent));
                Equal(SqlTruth.False,
                    DaslExpression.Parse("\"absent\" IS NOT NULL AND \"absent\" = 1").Evaluate(absent));
                True(!DaslExpression.Parse("\"absent\" = 1").Matches(absent));
            });
            Check("DASL evaluator respects precedence, parentheses, and escaped quoted strings", () =>
            {
                var message = new Dictionary<string, object>
                {
                    ["yes"] = 1, ["no"] = 0, ["quote\"property"] = "O'NEIL AND 'X'"
                };
                True(DaslExpression.Parse("\"yes\" = 1 OR \"no\" = 1 AND \"no\" = 1").Matches(message));
                True(!DaslExpression.Parse("(\"yes\" = 1 OR \"no\" = 1) AND \"no\" = 1").Matches(message));
                True(DaslExpression.Parse("NOT (\"no\" = 1 OR NOT \"yes\" = 1)").Matches(message));
                True(DaslExpression.Parse("\"quote\"\"property\" = 'o''neil AND ''x'''").Matches(message));
            });
            Check("DASL evaluator rejects unsupported or trailing syntax instead of executing it", () =>
            {
                foreach (var invalid in new[]
                {
                    "", "\"x\" = 2", "\"x\" = 10", "\"x\" = true", "\"x\" LIKE '%x%'",
                    "\"x\" = 'unterminated", "(\"x\" = 1", "\"x\" = 1)",
                    "\"x\" = 1; System.IO.File.Delete('anything')", "\"x\" IS NOTHING",
                    "@SQL=\"x\" = 1", "\"x\" = 1 OR"
                })
                    Throws<FormatException>(() => DaslExpression.Parse(invalid));
            });
        }

        private static void CheckGroupedRules()
        {
            Check("all managers share exactly four ordered native formatting groups", () =>
            {
                var config = DistinctConfiguration();
                var rules = RuleDefinition.Build(config, new[] { Person("boss"), Person("chief") });
                Equal(4, rules.Count);
                var expected = new[] { RuleKind.BossUnread, RuleKind.BossRead, RuleKind.CustomUnread, RuleKind.ToMe };
                for (int i = 0; i < expected.Length; i++)
                {
                    Equal(expected[i], rules[i].Kind);
                    Equal("OrgLens.v2:" + expected[i], rules[i].Name);
                    AssertStyle(RuleDefinition.StyleFor(config, expected[i]), rules[i].Style);
                    True(!ReferenceEquals(RuleDefinition.StyleFor(config, expected[i]), rules[i].Style));
                    True(!rules[i].Filter.Contains("@SQL="));
                    DaslExpression.Parse(rules[i].Filter);
                }
                Throws<ArgumentOutOfRangeException>(() => RuleDefinition.StyleFor(config, (RuleKind)999));
            });
            Check("generated DASL and MatchKind enforce priority across every boolean and missing-property case", () =>
            {
                var boss = Person("boss");
                var chief = Person("chief");
                var custom = Person("custom");
                var shared = Person("shared");
                var stranger = Person("stranger");
                var scenarios = new List<SenderScenario>();
                foreach (var person in new[] { boss, chief, custom, shared, stranger })
                {
                    bool fromBoss = person == boss || person == chief || person == shared;
                    bool fromCustom = person == custom || person == shared;
                    foreach (var property in SenderTags)
                        scenarios.Add(new SenderScenario(person.EntryId + " SMTP " + property, fromBoss, fromCustom,
                            new Dictionary<string, object> { [property] = person.SmtpAddress.ToUpperInvariant() }));
                    foreach (var property in SenderTags.Skip(2))
                        scenarios.Add(new SenderScenario(person.EntryId + " Exchange " + property, fromBoss, fromCustom,
                            new Dictionary<string, object> { [property] = person.LegacyAddress.ToUpperInvariant() }));
                }
                scenarios.Add(new SenderScenario("absent sender", false, false, new Dictionary<string, object>()));
                scenarios.Add(new SenderScenario("null sender", false, false,
                    SenderTags.ToDictionary(property => property, property => (object)null)));
                scenarios.Add(new SenderScenario("overlap across sender properties", true, true,
                    new Dictionary<string, object>
                    {
                        [SenderTags[0]] = boss.SmtpAddress, [SenderTags[2]] = custom.LegacyAddress
                    }));

                int cases = 0;
                for (int enabled = 0; enabled < 32; enabled++)
                {
                    var config = new OrgLensConfiguration
                    {
                        CustomAddresses = new List<string> { custom.SmtpAddress, shared.SmtpAddress },
                        ToMeUnreadOnly = (enabled & 16) != 0
                    };
                    config.BossUnread.Enabled = (enabled & 1) != 0;
                    config.BossRead.Enabled = (enabled & 2) != 0;
                    config.CustomUnread.Enabled = (enabled & 4) != 0;
                    config.ToMe.Enabled = (enabled & 8) != 0;
                    var rules = RuleDefinition.Build(config, new[] { boss, chief, shared }, new[] { custom, shared });
                    var expressions = rules.Select(rule => DaslExpression.Parse(rule.Filter)).ToArray();
                    foreach (var scenario in scenarios)
                    {
                        // -2 is explicit NULL; -1 is absent. Both must remain UNKNOWN in unguarded equality.
                        for (int read = -2; read <= 1; read++)
                        {
                            for (int to = -1; to <= 1; to++)
                            {
                                for (int cc = -1; cc <= 1; cc++)
                                {
                                    var message = new Dictionary<string, object>(scenario.Properties);
                                    if (read == -2) message[ReadTag] = null;
                                    else if (read >= 0) message[ReadTag] = read;
                                    if (to >= 0) message[ToTag] = to;
                                    if (cc >= 0) message[CcTag] = cc;
                                    bool direct = to == 1 || cc == 1;
                                    RuleKind? expected = null;
                                    if (scenario.Boss && read == 0 && config.BossUnread.Enabled)
                                        expected = RuleKind.BossUnread;
                                    else if (scenario.Boss && read == 1 && config.BossRead.Enabled)
                                        expected = RuleKind.BossRead;
                                    else if (scenario.Custom && read == 0 && config.CustomUnread.Enabled)
                                        expected = RuleKind.CustomUnread;
                                    else if (direct && config.ToMe.Enabled && (!config.ToMeUnreadOnly || read == 0))
                                        expected = RuleKind.ToMe;
                                    string context = "flags=" + enabled + ", " + scenario.Name +
                                        ", read=" + read + ", To=" + to + ", Cc=" + cc;
                                    var matches = rules.Where((rule, index) =>
                                        rule.Style.Enabled && expressions[index].Matches(message)).ToArray();
                                    Equal(expected.HasValue ? 1 : 0, matches.Length, context);
                                    if (expected.HasValue) Equal(expected.Value, matches[0].Kind, context);
                                    if (read >= 0)
                                        Equal(expected, RuleDefinition.MatchKind(config, scenario.Boss,
                                            scenario.Custom, read == 1, direct), context);
                                    cases++;
                                }
                            }
                        }
                    }
                }
                Equal(38016, cases);
                Console.WriteLine("  Evaluated " + cases + " independent DASL message/configuration combinations.");
            });
            Check("To me uses explicit native To/Cc flags and never recipient-name substrings", () =>
            {
                Equal(ReadTag, RuleDefinition.ReadProperty);
                Equal(ToTag, RuleDefinition.ToMeProperty);
                Equal(CcTag, RuleDefinition.CcMeProperty);
                var config = new OrgLensConfiguration();
                var rules = RuleDefinition.Build(config, new ManagerPerson[0]);
                var expression = DaslExpression.Parse(rules.Single(rule => rule.Kind == RuleKind.ToMe).Filter);
                True(expression.Properties.Contains(ToTag));
                True(expression.Properties.Contains(CcTag));
                True(expression.Properties.All(property =>
                    property == ReadTag || property == ToTag || property == CcTag));
                var namesOnly = new Dictionary<string, object>
                {
                    [ReadTag] = 0, [ToTag] = 0, [CcTag] = 0,
                    ["urn:schemas:httpmail:displayto"] = "All staff; me@example.com",
                    ["urn:schemas:httpmail:displaycc"] = "me@example.com (distribution list)",
                    ["urn:schemas:httpmail:displaybcc"] = "me@example.com",
                    ["http://schemas.microsoft.com/mapi/proptag/0x0E04001F"] = "me@example.com",
                    ["http://schemas.microsoft.com/mapi/proptag/0x0059000B"] = 1
                };
                True(!expression.Matches(namesOnly), "DL names, BCC, and PR_MESSAGE_RECIP_ME are not direct To/Cc.");
                namesOnly[ToTag] = 1;
                True(expression.Matches(namesOnly));
                namesOnly[ToTag] = 0;
                namesOnly[CcTag] = 1;
                True(expression.Matches(namesOnly));
                namesOnly[ReadTag] = 1;
                True(!expression.Matches(namesOnly));
                config.ToMeUnreadOnly = false;
                expression = DaslExpression.Parse(RuleDefinition.Build(config, new ManagerPerson[0])[3].Filter);
                True(expression.Matches(namesOnly));
            });
            Check("SMTP and Exchange DN sender matching is exact across ANSI and Unicode properties", () =>
            {
                var boss = Person("boss");
                var expression = DaslExpression.Parse(RuleDefinition.FilterFor(boss));
                Equal(4, expression.Properties.Count);
                foreach (var property in SenderTags)
                {
                    True(expression.Matches(new Dictionary<string, object> { [property] = boss.SmtpAddress }));
                    True(!expression.Matches(new Dictionary<string, object> { [property] = "not" + boss.SmtpAddress }));
                    True(!expression.Matches(new Dictionary<string, object> { [property] = boss.SmtpAddress + ".invalid" }));
                    True(!expression.Matches(new Dictionary<string, object> { [property] = null }));
                }
                foreach (var property in SenderTags.Skip(2))
                {
                    True(expression.Matches(new Dictionary<string, object> { [property] = boss.LegacyAddress }));
                    True(!expression.Matches(new Dictionary<string, object> { [property] = boss.LegacyAddress + "-other" }));
                }
                True(!expression.Matches(new Dictionary<string, object>()));
            });
            Check("SMTP-only and Exchange-only managers match without requiring both addresses", () =>
            {
                var smtp = DaslExpression.Parse(RuleDefinition.FilterFor(
                    new ManagerPerson("smtp", "SMTP", "boss@example.com", "")));
                var legacy = DaslExpression.Parse(RuleDefinition.FilterFor(
                    new ManagerPerson("dn", "Exchange", "", "/o=Example/cn=Boss")));
                True(smtp.Matches(new Dictionary<string, object> { [SenderTags[0]] = "BOSS@EXAMPLE.COM" }));
                True(legacy.Matches(new Dictionary<string, object> { [SenderTags[2]] = "/O=EXAMPLE/CN=BOSS" }));
                True(!legacy.Matches(new Dictionary<string, object> { [SenderTags[0]] = "/o=Example/cn=Boss" }));
            });
            Check("DASL apostrophe escaping preserves SMTP and Exchange values without broadening matches", () =>
            {
                var person = new ManagerPerson("quote", "Quoted", "o'neil@example.com",
                    "/o=Example/cn=O'Neil OR '1'='1");
                var filter = RuleDefinition.FilterFor(person);
                True(filter.Contains("o''neil@example.com"));
                True(filter.Contains("O''Neil OR ''1''=''1"));
                var expression = DaslExpression.Parse(filter);
                True(expression.Matches(new Dictionary<string, object> { [SenderTags[0]] = person.SmtpAddress }));
                True(expression.Matches(new Dictionary<string, object> { [SenderTags[2]] = person.LegacyAddress }));
                True(!expression.Matches(new Dictionary<string, object> { [SenderTags[2]] = "/o=Example/cn=O" }));
                True(!expression.Matches(new Dictionary<string, object> { [SenderTags[0]] = "stranger@example.com" }));
                True(!expression.Matches(new Dictionary<string, object>()));
            });
            Check("custom address list works without directory resolution and only matches unread mail", () =>
            {
                var config = new OrgLensConfiguration
                {
                    CustomAddresses = new List<string> { " FIRST@example.com ", "second@example.com", "o'neil@example.com" }
                };
                var expression = DaslExpression.Parse(RuleDefinition.Build(config, new ManagerPerson[0])[2].Filter);
                foreach (string address in new[] { "first@example.com", "second@example.com", "o'neil@example.com" })
                {
                    var message = new Dictionary<string, object> { [SenderTags[0]] = address, [ReadTag] = 0 };
                    True(expression.Matches(message));
                    message[ReadTag] = 1;
                    True(!expression.Matches(message));
                }
                True(!expression.Matches(new Dictionary<string, object>
                {
                    [SenderTags[0]] = "notfirst@example.com", [ReadTag] = 0
                }));
            });
            Check("unavailable hierarchy rejects either enabled BOSS style before target mutation", () =>
            {
                foreach (bool unread in new[] { false, true })
                {
                    var config = new OrgLensConfiguration();
                    config.BossUnread.Enabled = unread;
                    config.BossRead.Enabled = !unread;
                    var target = SeededRules();
                    Throws<OrgLensException>(() => RuleReconciler.Replace(target, RuleDefinition.Build(config, null)));
                    Equal(0, target.Mutations);
                    Equal(0, target.ReadCalls);
                }
            });
            Check("known empty hierarchy and custom list match none rather than every message", () =>
            {
                var rules = RuleDefinition.Build(new OrgLensConfiguration(), new ManagerPerson[0]);
                for (int read = -1; read <= 1; read++)
                {
                    var message = new Dictionary<string, object> { [ToTag] = 1 };
                    if (read >= 0) message[ReadTag] = read;
                    foreach (var rule in rules.Take(3))
                    {
                        True(!string.IsNullOrWhiteSpace(rule.Filter));
                        Equal(SqlTruth.False, DaslExpression.Parse(rule.Filter).Evaluate(message));
                    }
                }
                var never = DaslExpression.Parse(RuleDefinition.NeverMatch);
                Equal(SqlTruth.False, never.Evaluate(new Dictionary<string, object>()));
                Equal(SqlTruth.False, never.Evaluate(new Dictionary<string, object> { [ReadTag] = 1 }));
            });
            Check("disabled BOSS permits unavailable hierarchy and disabled groups retain safe filters", () =>
            {
                var config = new OrgLensConfiguration { CustomAddresses = new List<string> { "custom@example.com" } };
                config.BossRead.Enabled = false;
                config.BossUnread.Enabled = false;
                var rules = RuleDefinition.Build(config, null);
                True(DaslExpression.Parse(rules[2].Filter).Matches(new Dictionary<string, object>
                {
                    [SenderTags[0]] = "custom@example.com", [ReadTag] = 0
                }));
                config.CustomUnread.Enabled = false;
                config.ToMe.Enabled = false;
                var target = new FakeRules();
                rules = RuleDefinition.Build(config, new[] { new ManagerPerson("missing", "Missing", "", "") });
                RuleReconciler.Replace(target, rules);
                Equal(4, target.Items.Count);
                foreach (var rule in rules)
                {
                    True(!rule.Style.Enabled);
                    True(!string.IsNullOrWhiteSpace(rule.Filter));
                    DaslExpression.Parse(rule.Filter);
                    True(!target.Added[rule.Name].Style.Enabled);
                }
            });
            Check("invalid sender addresses cannot create match-all filters", () =>
            {
                foreach (var person in new[]
                {
                    new ManagerPerson("empty", "Empty", "", ""),
                    new ManagerPerson("invalid", "Invalid", "not-an-email", ""),
                    new ManagerPerson("control", "Control", "", "/o=Example/\ncn=Boss")
                })
                    Throws<OrgLensException>(() => RuleDefinition.FilterFor(person));
                Throws<ArgumentNullException>(() => RuleDefinition.FilterFor(null));
            });
            Check("generated plans own independent copies of every font style", () =>
            {
                var config = DistinctConfiguration();
                var rules = RuleDefinition.Build(config, new[] { Person("boss") });
                var before = config.Copy();
                foreach (var rule in rules)
                {
                    rule.Style.Bold = !rule.Style.Bold;
                    rule.Style.Italic = !rule.Style.Italic;
                    rule.Style.Enabled = !rule.Style.Enabled;
                    rule.Style.Color = MailColor.White;
                    rule.Style.FontSize = 45;
                }
                AssertConfiguration(before, config);
            });
        }

        private static void CheckReconciliation()
        {
            Check("replace explicitly saves the rule collection once after all changes", () =>
            {
                var target = SeededRules();
                RuleReconciler.Replace(target, RuleDefinition.Build(new OrgLensConfiguration(), new[] { Person("boss") }));
                Equal(1, target.SaveCalls);
                Equal(target.Snapshot(), target.SavedSnapshot);
            });
            Check("remove explicitly saves the rule collection", () =>
            {
                var target = SeededRules();
                RuleReconciler.RemoveOwned(target);
                Equal(1, target.SaveCalls);
                Equal(target.Snapshot(), target.SavedSnapshot);
            });
            Check("native collection save failures are not reported as successful applies", () =>
            {
                var target = SeededRules();
                target.SaveError = new InvalidOperationException("Fixture collection save failure.");
                Throws<InvalidOperationException>(() => RuleReconciler.Replace(target,
                    RuleDefinition.Build(new OrgLensConfiguration(), new[] { Person("boss") })));
            });
            Check("legacy v1 identity survives manager rename and casing", () =>
            {
                Equal(RuleDefinition.NameFor(Person("boss")), RuleDefinition.NameFor(
                    new ManagerPerson("different-entry", "New name", " BOSS@EXAMPLE.COM ", "")));
                True(RuleDefinition.IsOwned(RuleDefinition.NameFor(new ManagerPerson("id", "DN", "", "/o=Example"))));
                True(RuleDefinition.IsOwned(RuleDefinition.NameFor(new ManagerPerson("id", "Entry", "", ""))));
                Throws<OrgLensException>(() => RuleDefinition.NameFor(new ManagerPerson("", "Missing", "", "")));
                Throws<ArgumentNullException>(() => RuleDefinition.NameFor(null));
            });
            Check("ownership accepts only exact v2 groups and strict legacy v1 hashes", () =>
            {
                string legacy = RuleDefinition.NameFor(Person("boss"));
                True(RuleDefinition.IsOwned(legacy));
                foreach (RuleKind kind in Enum.GetValues(typeof(RuleKind)))
                    True(RuleDefinition.IsOwned("OrgLens.v2:" + kind));
                foreach (string other in new[]
                {
                    null, "", "Unread messages", "OrgLens.v1:My custom rule", legacy.ToLowerInvariant(),
                    legacy + "0", legacy.Substring(0, legacy.Length - 1), "OrgLens.v1:" + new string('G', 64),
                    "OrgLens.v2:BossUnread ", "orglens.v2:BossRead", "OrgLens.v2:Boss", "OrgLens.v2:ToMe:Mine"
                })
                    True(!RuleDefinition.IsOwned(other), other);
            });
            Check("replace migrates legacy rules to four groups while preserving native and unrelated rules", () =>
            {
                var target = SeededRules();
                target.Items.Add(new ExistingRule("OrgLens.v2:BossUnread", false));
                target.Items.Add(new ExistingRule("OrgLens.v2:BossUnread", false));
                target.Items.Add(new ExistingRule("OrgLens.v2:My custom rule", false));
                target.Items.Add(new ExistingRule(RuleDefinition.NameFor(Person("standard")), true));
                target.Items.Add(new ExistingRule("OrgLens.v2:ToMe", true));
                var preserved = target.Items.Where(item => item.Standard || !RuleDefinition.IsOwned(item.Name)).ToArray();
                var desired = RuleDefinition.Build(DistinctConfiguration(), new[] { Person("boss"), Person("chief") });
                RuleReconciler.Replace(target, desired);
                Equal(preserved.Length + 4, target.Items.Count);
                for (int i = 0; i < preserved.Length; i++) True(ReferenceEquals(preserved[i], target.Items[i]));
                for (int i = 0; i < desired.Count; i++)
                {
                    Equal(desired[i].Name, target.Items[preserved.Length + i].Name);
                    Equal(desired[i].Filter, target.Added[desired[i].Name].Filter);
                    AssertStyle(desired[i].Style, target.Added[desired[i].Name].Style);
                }
            });
            Check("repeated apply is idempotent including filters, order, and all font attributes", () =>
            {
                var target = SeededRules();
                var config = DistinctConfiguration();
                var desired = RuleDefinition.Build(config, new[] { Person("boss"), Person("chief") });
                RuleReconciler.Replace(target, desired);
                string before = target.Snapshot();
                RuleReconciler.Replace(target, RuleDefinition.Build(config, new[] { Person("boss"), Person("chief") }));
                Equal(6, target.Items.Count);
                Equal(before, target.Snapshot());
            });
            Check("remove deletes only owned nonstandard v1 and v2 rules in descending index order", () =>
            {
                var target = SeededRules();
                target.Items.Add(new ExistingRule("OrgLens.v2:BossRead", false));
                target.Items.Add(new ExistingRule(RuleDefinition.NameFor(Person("standard")), true));
                target.Items.Add(new ExistingRule("OrgLens.v2:ToMe", true));
                target.Items.Add(new ExistingRule("OrgLens.v1:My custom rule", false));
                RuleReconciler.RemoveOwned(target);
                Equal(5, target.Items.Count);
                Equal("4,3", string.Join(",", target.RemovedIndexes));
                Equal("Unread messages", target.Items[0].Name);
                Equal("My customers", target.Items[1].Name);
                True(target.Items[2].Standard);
                True(target.Items[3].Standard);
                Equal("OrgLens.v1:My custom rule", target.Items[4].Name);
                string before = target.Snapshot();
                RuleReconciler.RemoveOwned(target);
                Equal(before, target.Snapshot());
            });
            Check("invalid formatting plans are rejected before any target read or mutation", () =>
            {
                var good = new FormattingRule(RuleKind.BossUnread, RuleDefinition.NeverMatch, new RuleStyle());
                var invalidPlans = new IReadOnlyList<FormattingRule>[]
                {
                    new FormattingRule[0],
                    new FormattingRule[] { good, null },
                    new[] { good, good },
                    new[] { good, new FormattingRule((RuleKind)999, RuleDefinition.NeverMatch, new RuleStyle()) },
                    new[] { good, new FormattingRule(RuleKind.ToMe, " ", new RuleStyle()) },
                    new[] { good, new FormattingRule(RuleKind.ToMe, null, new RuleStyle()) },
                    new[] { good, new FormattingRule(RuleKind.ToMe, RuleDefinition.NeverMatch,
                        new RuleStyle { Color = (MailColor)999 }) },
                    new[] { good, new FormattingRule(RuleKind.ToMe, RuleDefinition.NeverMatch,
                        new RuleStyle { FontSize = 128, Enabled = false }) }
                };
                foreach (var plan in invalidPlans)
                {
                    var target = SeededRules();
                    string before = target.Snapshot();
                    Throws<OrgLensException>(() => RuleReconciler.Replace(target, plan));
                    Equal(0, target.Mutations);
                    Equal(0, target.ReadCalls);
                    Equal(before, target.Snapshot());
                }
                var nullTarget = SeededRules();
                Throws<ArgumentNullException>(() => RuleReconciler.Replace(nullTarget, null));
                Equal(0, nullTarget.Mutations);
                Equal(0, nullTarget.ReadCalls);
                Throws<ArgumentNullException>(() => RuleReconciler.Replace(null, new[] { good }));
                Throws<ArgumentNullException>(() => RuleReconciler.RemoveOwned(null));
                Throws<ArgumentNullException>(() => new FormattingRule(RuleKind.ToMe, RuleDefinition.NeverMatch, null));
            });
        }

        private static FakeRules SeededRules()
        {
            var result = new FakeRules();
            result.Items.Add(new ExistingRule("Unread messages", true));
            result.Items.Add(new ExistingRule("My customers", false));
            result.Items.Add(new ExistingRule(RuleDefinition.NameFor(Person("former-boss")), false));
            return result;
        }

        private sealed class SenderScenario
        {
            internal SenderScenario(string name, bool boss, bool custom, Dictionary<string, object> properties)
            {
                Name = name;
                Boss = boss;
                Custom = custom;
                Properties = properties;
            }
            internal string Name { get; }
            internal bool Boss { get; }
            internal bool Custom { get; }
            internal Dictionary<string, object> Properties { get; }
        }

        private sealed class AddedRule
        {
            internal string Filter { get; set; }
            internal RuleStyle Style { get; set; }
        }

        private sealed class FakeRules : IFormattingRules
        {
            internal List<ExistingRule> Items { get; } = new List<ExistingRule>();
            internal Dictionary<string, AddedRule> Added { get; } = new Dictionary<string, AddedRule>();
            internal List<int> RemovedIndexes { get; } = new List<int>();
            internal int Mutations { get; private set; }
            internal int ReadCalls { get; private set; }
            internal int SaveCalls { get; private set; }
            internal string SavedSnapshot { get; private set; }
            internal Exception SaveError { get; set; }
            public IReadOnlyList<ExistingRule> Read() { ReadCalls++; return Items.ToArray(); }
            public void Remove(int index)
            {
                Added.Remove(Items[index - 1].Name);
                Items.RemoveAt(index - 1);
                RemovedIndexes.Add(index);
                Mutations++;
            }
            public void Add(string name, string filter, RuleStyle style)
            {
                Items.Add(new ExistingRule(name, false));
                Added[name] = new AddedRule { Filter = filter, Style = style.Copy() };
                Mutations++;
            }
            public void Save()
            {
                SaveCalls++;
                if (SaveError != null) throw SaveError;
                SavedSnapshot = Snapshot();
            }
            internal string Snapshot()
            {
                return string.Join("\n", Items.Select(item =>
                {
                    AddedRule added;
                    return item.Name + "|" + item.Standard + (Added.TryGetValue(item.Name, out added)
                        ? "|" + added.Filter + "|" + StyleSnapshot(added.Style) : "");
                }));
            }
        }
    }
}
