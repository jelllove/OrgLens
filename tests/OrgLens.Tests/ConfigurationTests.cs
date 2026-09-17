using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OrgLens.Core;

namespace OrgLens.Tests
{
    internal static partial class Program
    {
        private static void CheckConfiguration()
        {
            Check("v2 defaults define shared read/unread BOSS, unread custom, and unread To me styles", () =>
            {
                var config = new OrgLensConfiguration();
                Equal(2, OrgLensConfiguration.CurrentVersion);
                Equal(2, config.Version);
                AssertStyle(new RuleStyle { Enabled = true, Color = MailColor.Navy, Bold = false,
                    Italic = false, FontSize = 11 }, config.BossRead);
                AssertStyle(new RuleStyle { Enabled = true, Color = MailColor.Blue, Bold = true,
                    Italic = false, FontSize = 12 }, config.BossUnread);
                AssertStyle(new RuleStyle { Enabled = true, Color = MailColor.Purple, Bold = true,
                    Italic = false, FontSize = 11 }, config.CustomUnread);
                AssertStyle(new RuleStyle { Enabled = true, Color = MailColor.Teal, Bold = false,
                    Italic = false, FontSize = 11 }, config.ToMe);
                True(config.ToMeUnreadOnly);
                True(config.BossEnabled);
                Equal(0, config.CustomAddresses.Count);
                Equal(null, RuleDefinition.MatchKind(config, false, false, true, true));
                Equal((RuleKind?)RuleKind.ToMe, RuleDefinition.MatchKind(config, false, false, false, true));
            });
            Check("style and configuration copies isolate all attributes and custom addresses", () =>
            {
                var source = DistinctConfiguration();
                var original = DistinctConfiguration();
                var copy = source.Copy();
                AssertConfiguration(source, copy);
                AssertIndependent(source, copy);
                ChangeConfiguration(copy);
                AssertConfiguration(original, source);
                var style = source.BossUnread.Copy();
                AssertStyle(source.BossUnread, style);
                style.Enabled = !style.Enabled;
                style.Color = MailColor.Black;
                style.Bold = !style.Bold;
                style.Italic = !style.Italic;
                style.FontSize = 37;
                AssertStyle(original.BossUnread, source.BossUnread);
            });
            Check("normalization trims and deduplicates address casing without mutating input", () =>
            {
                var config = DistinctConfiguration();
                config.CustomAddresses = new List<string>
                {
                    " FIRST@Example.COM ", "second@example.com", "first@example.com", "SECOND@EXAMPLE.COM"
                };
                var before = config.Copy();
                var normalized = ConfigurationValidator.Normalize(config);
                Equal("first@example.com,second@example.com", string.Join(",", normalized.CustomAddresses));
                AssertConfiguration(before, config);
                AssertIndependent(config, normalized);
                ChangeConfiguration(normalized);
                AssertConfiguration(before, config);
            });
            Check("custom email input accepts semicolon, comma, CRLF, and newline separators with deduplication", () =>
            {
                var addresses = ConfigurationValidator.ParseAddresses(
                    " First@Example.com ; second@example.com,\r\nFIRST@example.COM\n o'neil@example.com ; , \r\n");
                Equal("first@example.com,second@example.com,o'neil@example.com", string.Join(",", addresses));
                Equal(0, ConfigurationValidator.ParseAddresses(" ;,\r\n  ").Count);
                Equal(0, ConfigurationValidator.ParseAddresses("").Count);
                Throws<ArgumentNullException>(() => ConfigurationValidator.ParseAddresses(null));
            });
            Check("custom email validation rejects blanks, display names, controls, and malformed addresses", () =>
            {
                foreach (string address in new[]
                {
                    null, "", "   ", "not-an-email", "boss@@example.com", "a b@example.com",
                    "Boss <boss@example.com>", "<boss@example.com>", "boss@example.com\n",
                    "boss@example.com\0", "boss@example.com,other@example.com",
                    new string('a', 245) + "@example.com"
                })
                {
                    Throws<OrgLensException>(() => ConfigurationValidator.NormalizeAddress(address));
                    var config = new OrgLensConfiguration { CustomAddresses = new List<string> { address } };
                    config.CustomUnread.Enabled = false;
                    Throws<OrgLensException>(() => ConfigurationValidator.Normalize(config));
                }
                Throws<OrgLensException>(() => ConfigurationValidator.ParseAddresses(
                    "valid@example.com;not-an-email;other@example.com"));
            });
            Check("font sizes accept every integer from 1 through 127 and all supported colors", () =>
            {
                for (int size = 1; size <= 127; size++)
                    ConfigurationValidator.ValidateStyle(new RuleStyle { FontSize = size });
                foreach (MailColor color in Enum.GetValues(typeof(MailColor)))
                    ConfigurationValidator.ValidateStyle(new RuleStyle { Color = color });
            });
            Check("invalid font sizes and colors are rejected even for disabled styles", () =>
            {
                foreach (int size in new[] { int.MinValue, -1, 0, 128, int.MaxValue })
                    Throws<OrgLensException>(() => ConfigurationValidator.ValidateStyle(
                        new RuleStyle { FontSize = size, Enabled = false }));
                foreach (var color in new[] { (MailColor)(-1), (MailColor)17, (MailColor)999 })
                    Throws<OrgLensException>(() => ConfigurationValidator.ValidateStyle(
                        new RuleStyle { Color = color, Enabled = false }));
                Throws<OrgLensException>(() => ConfigurationValidator.ValidateStyle(null));
            });
            Check("normalization requires v2 and every style and list", () =>
            {
                Throws<OrgLensException>(() => ConfigurationValidator.Normalize(null));
                var invalid = new Action<OrgLensConfiguration>[]
                {
                    c => c.Version = 0, c => c.Version = 1, c => c.Version = 3,
                    c => c.BossRead = null, c => c.BossUnread = null,
                    c => c.CustomUnread = null, c => c.ToMe = null, c => c.CustomAddresses = null,
                    c => c.BossRead.FontSize = 0, c => c.BossUnread.FontSize = 128,
                    c => c.CustomUnread.Color = (MailColor)999, c => c.ToMe.Color = (MailColor)(-1)
                };
                foreach (var change in invalid)
                {
                    var config = new OrgLensConfiguration();
                    change(config);
                    var target = SeededRules();
                    Throws<OrgLensException>(() => RuleReconciler.Replace(target,
                        RuleDefinition.Build(config, new[] { Person("boss") })));
                    Equal(0, target.ReadCalls);
                    Equal(0, target.Mutations);
                }
            });
        }

        private static void CheckPersistence()
        {
            Check("chosen Unicode and spaced file path round-trips every v2 setting and style", () =>
                WithConfigurationFolder(folder =>
                {
                    string nested = Path.Combine(folder, "设置 with spaces");
                    Directory.CreateDirectory(nested);
                    string path = Path.Combine(nested, "我的 OrgLens settings.json");
                    var original = DistinctConfiguration();
                    ConfigurationFile.Save(path, original);
                    True(File.Exists(path));
                    AssertConfiguration(original, ConfigurationFile.Load(path));
                    Equal(1, Directory.GetFiles(nested).Length);
                    Equal(0, Directory.GetFiles(folder).Length);
                    AssertNoStagingFiles(nested);
                }));
            Check("saved and separately loaded configurations are independent full objects", () =>
                WithConfigurationFolder(folder =>
                {
                    string path = Path.Combine(folder, "independent.json");
                    var config = DistinctConfiguration();
                    ConfigurationFile.Save(path, config);
                    var first = ConfigurationFile.Load(path);
                    var second = ConfigurationFile.Load(path);
                    AssertIndependent(config, first);
                    AssertIndependent(first, second);
                    var original = DistinctConfiguration();
                    ChangeConfiguration(config);
                    ChangeConfiguration(first);
                    AssertConfiguration(original, second);
                    AssertConfiguration(original, ConfigurationFile.Load(path));
                }));
            Check("save overwrites the selected existing file without changing another file", () =>
                WithConfigurationFolder(folder =>
                {
                    string selected = Path.Combine(folder, "selected.json");
                    string other = Path.Combine(folder, "not selected.json");
                    var initial = new OrgLensConfiguration();
                    ConfigurationFile.Save(selected, initial);
                    ConfigurationFile.Save(other, initial);
                    byte[] otherBytes = File.ReadAllBytes(other);
                    var replacement = DistinctConfiguration();
                    ConfigurationFile.Save(selected, replacement);
                    AssertConfiguration(replacement, ConfigurationFile.Load(selected));
                    AssertConfiguration(initial, ConfigurationFile.Load(other));
                    BytesEqual(otherBytes, File.ReadAllBytes(other));
                    Equal(2, Directory.GetFiles(folder).Length);
                    AssertNoStagingFiles(folder);
                }));
            Check("defaults persist as explicit complete configuration rather than relying on deserialization defaults", () =>
                WithConfigurationFolder(folder =>
                {
                    string path = Path.Combine(folder, "defaults.json");
                    var defaults = new OrgLensConfiguration();
                    ConfigurationFile.Save(path, defaults);
                    AssertConfiguration(defaults, ConfigurationFile.Load(path));
                    string text = File.ReadAllText(path);
                    foreach (string field in ConfigurationJsonFields().Keys)
                        True(Regex.IsMatch(text, "\"" + field + "\"\\s*:"), field);
                }));
            Check("save and open normalize duplicate address casing without altering the source object", () =>
                WithConfigurationFolder(folder =>
                {
                    string path = Path.Combine(folder, "duplicates.json");
                    var config = DistinctConfiguration();
                    config.CustomAddresses = new List<string>
                    {
                        " FIRST@Example.COM ", "first@example.com", " Second@Example.COM ", "SECOND@example.com"
                    };
                    var before = config.Copy();
                    ConfigurationFile.Save(path, config);
                    AssertConfiguration(before, config);
                    Equal("first@example.com,second@example.com",
                        string.Join(",", ConfigurationFile.Load(path).CustomAddresses));
                    var fields = ConfigurationJsonFields();
                    fields["customAddresses"] = "[\" One@Example.com \",\"one@example.COM\",\"two@example.com\"]";
                    File.WriteAllText(path, JsonObject(fields), new UTF8Encoding(false));
                    Equal("one@example.com,two@example.com",
                        string.Join(",", ConfigurationFile.Load(path).CustomAddresses));
                }));
            Check("export is portable v2 and imports reject organization, account IDs, and credential fields", () =>
                WithConfigurationFolder(folder =>
                {
                    string path = Path.Combine(folder, "portable.json");
                    var fields = ConfigurationJsonFields();
                    fields["accountId"] = "\"fixture-account-id\"";
                    fields["managers"] = "[{\"entryId\":\"fixture-manager-id\",\"legacyAddress\":\"fixture-dn\"}]";
                    fields["credentials"] = "{\"accessToken\":\"fixture-token-not-a-credential\"}";
                    RejectJson(folder, "unrecognized-fields.json", JsonObject(fields));
                    ConfigurationFile.Save(path, DistinctConfiguration());
                    Equal(2, ConfigurationFile.Load(path).Version);
                    string text = File.ReadAllText(path);
                    var allowed = new HashSet<string>(ConfigurationJsonFields().Keys)
                    {
                        "enabled", "color", "bold", "italic", "fontSize"
                    };
                    foreach (Match property in Regex.Matches(text, "\"([^\"]+)\"\\s*:"))
                        True(allowed.Contains(property.Groups[1].Value), property.Value);
                    foreach (string forbidden in new[]
                    {
                        "managers", "managerChain", "hierarchy", "accountId", "entryId", "legacyAddress",
                        "credentials", "password", "accessToken", "fixture-account-id",
                        "fixture-manager-id", "fixture-dn", "fixture-token-not-a-credential"
                    })
                        True(text.IndexOf(forbidden, StringComparison.OrdinalIgnoreCase) < 0, forbidden);
                }));
            string validJson = JsonObject(ConfigurationJsonFields());
            var invalidJson = new Dictionary<string, string>
            {
                ["empty file"] = "", ["whitespace only"] = " ", ["incomplete object"] = "{",
                ["truncated halfway"] = validJson.Substring(0, validJson.Length / 2),
                ["missing final brace"] = validJson.Substring(0, validJson.Length - 1),
                ["trailing garbage"] = validJson + "garbage", ["trailing object"] = validJson + "{}",
                ["trailing boolean"] = validJson + " true", ["array root"] = "[]",
                ["null root"] = "null", ["boolean root"] = "true",
                ["duplicate version"] = validJson.Replace("\"version\":2", "\"version\":2,\"version\":2")
            };
            foreach (var invalid in invalidJson)
                Check("rejects " + invalid.Key + " JSON without resetting the chosen file", () =>
                    WithConfigurationFolder(folder =>
                        RejectJson(folder, invalid.Key + ".json", invalid.Value)));
            Check("unsupported or missing configuration versions fail instead of silently migrating file content", () =>
                WithConfigurationFolder(folder =>
                {
                    foreach (string version in new[] { "-1", "0", "1", "3", "999", "null", "\"invalid\"", "2.5" })
                    {
                        var fields = ConfigurationJsonFields();
                        fields["version"] = version;
                        RejectJson(folder, "version.json", JsonObject(fields));
                    }
                    var missing = ConfigurationJsonFields();
                    missing.Remove("version");
                    RejectJson(folder, "missing-version.json", JsonObject(missing));
                }));
            Check("every top-level configuration field is required and rejects null", () =>
                WithConfigurationFolder(folder =>
                {
                    foreach (string field in ConfigurationJsonFields().Keys)
                    {
                        var missing = ConfigurationJsonFields();
                        missing.Remove(field);
                        RejectJson(folder, "missing-" + field + ".json", JsonObject(missing));
                        var nullField = ConfigurationJsonFields();
                        nullField[field] = "null";
                        RejectJson(folder, "null-" + field + ".json", JsonObject(nullField));
                    }
                }));
            Check("all five attributes of all four persisted styles are required and reject null", () =>
                WithConfigurationFolder(folder =>
                {
                    foreach (string group in new[] { "bossRead", "bossUnread", "customUnread", "toMe" })
                    {
                        foreach (string attribute in StyleJsonFields().Keys)
                        {
                            var style = StyleJsonFields();
                            style.Remove(attribute);
                            var fields = ConfigurationJsonFields();
                            fields[group] = JsonObject(style);
                            RejectJson(folder, "missing-" + group + "-" + attribute + ".json", JsonObject(fields));
                            style = StyleJsonFields();
                            style[attribute] = "null";
                            fields[group] = JsonObject(style);
                            RejectJson(folder, "null-" + group + "-" + attribute + ".json", JsonObject(fields));
                        }
                    }
                }));
            Check("persisted fonts reject noninteger and out-of-range values for every style", () =>
                WithConfigurationFolder(folder =>
                {
                    foreach (string group in new[] { "bossRead", "bossUnread", "customUnread", "toMe" })
                    {
                        foreach (string size in new[] { "-1", "0", "128", "2147483648", "11.5", "\"large\"" })
                        {
                            var style = StyleJsonFields();
                            style["fontSize"] = size;
                            var fields = ConfigurationJsonFields();
                            fields[group] = JsonObject(style);
                            RejectJson(folder, "invalid-font-" + group + ".json", JsonObject(fields));
                        }
                    }
                }));
            Check("persisted colors and email lists reject invalid values", () =>
                WithConfigurationFolder(folder =>
                {
                    foreach (string color in new[] { "\"NotAColor\"", "\"blue\"", "\"999\"", "999", "null" })
                    {
                        var style = StyleJsonFields();
                        style["color"] = color;
                        var fields = ConfigurationJsonFields();
                        fields["toMe"] = JsonObject(style);
                        RejectJson(folder, "invalid-color.json", JsonObject(fields));
                    }
                    foreach (string addresses in new[]
                    {
                        "[null]", "[\"\"]", "[\"   \"]", "[\"not-an-email\"]",
                        "[\"Name <boss@example.com>\"]", "[\"boss@example.com\\n\"]",
                        "[\"valid@example.com\",\"invalid\"]", "\"boss@example.com\"", "{}"
                    })
                    {
                        var fields = ConfigurationJsonFields();
                        fields["customAddresses"] = addresses;
                        RejectJson(folder, "invalid-addresses.json", JsonObject(fields));
                    }
                }));
            Check("oversized input files fail without truncation or default fallback", () =>
                WithConfigurationFolder(folder =>
                {
                    string path = Path.Combine(folder, "oversized.json");
                    byte[] bytes = Encoding.UTF8.GetBytes(new string(' ', ConfigurationFile.MaximumBytes + 1));
                    File.WriteAllBytes(path, bytes);
                    Throws<OrgLensException>(() => ConfigurationFile.Load(path));
                    BytesEqual(bytes, File.ReadAllBytes(path));
                }));
            Check("missing file and directory IO errors propagate instead of returning defaults", () =>
                WithConfigurationFolder(folder =>
                {
                    Throws<FileNotFoundException>(() => ConfigurationFile.Load(Path.Combine(folder, "missing.json")));
                    string absentDirectory = Path.Combine(folder, "does not exist", "settings.json");
                    Throws<DirectoryNotFoundException>(() => ConfigurationFile.Load(absentDirectory));
                    Throws<DirectoryNotFoundException>(() => ConfigurationFile.Save(absentDirectory,
                        new OrgLensConfiguration()));
                    Equal(0, Directory.GetFiles(folder).Length);
                }));
            Check("failed validation before save preserves the entire previously selected file", () =>
                WithConfigurationFolder(folder =>
                {
                    string path = Path.Combine(folder, "existing.json");
                    var previous = DistinctConfiguration();
                    ConfigurationFile.Save(path, previous);
                    byte[] before = File.ReadAllBytes(path);
                    foreach (var invalidate in new Action<OrgLensConfiguration>[]
                    {
                        c => c.Version = 1, c => c.BossRead.FontSize = 0,
                        c => c.BossUnread.FontSize = 128, c => c.CustomUnread.Color = (MailColor)999,
                        c => c.ToMe = null, c => c.CustomAddresses.Add("not-an-email")
                    })
                    {
                        var config = new OrgLensConfiguration();
                        invalidate(config);
                        Throws<OrgLensException>(() => ConfigurationFile.Save(path, config));
                        BytesEqual(before, File.ReadAllBytes(path));
                        AssertConfiguration(previous, ConfigurationFile.Load(path));
                        AssertNoStagingFiles(folder);
                    }
                    Throws<OrgLensException>(() => ConfigurationFile.Save(path, null));
                    BytesEqual(before, File.ReadAllBytes(path));
                    AssertNoStagingFiles(folder);
                }));
            Check("oversized save rolls back staging without truncating an existing file", () =>
                WithConfigurationFolder(folder =>
                {
                    string path = Path.Combine(folder, "existing.json");
                    var previous = DistinctConfiguration();
                    ConfigurationFile.Save(path, previous);
                    byte[] before = File.ReadAllBytes(path);
                    var large = new OrgLensConfiguration
                    {
                        CustomAddresses = Enumerable.Range(0, 6500).Select(index =>
                            new string('x', 50) + index.ToString("D5") + "@" + new string('a', 60) + "." +
                            new string('b', 50) + ".example.com").ToList()
                    };
                    Throws<OrgLensException>(() => ConfigurationFile.Save(path, large));
                    BytesEqual(before, File.ReadAllBytes(path));
                    AssertConfiguration(previous, ConfigurationFile.Load(path));
                    AssertNoStagingFiles(folder);
                }));
            Check("failed replacement IO does not truncate the existing file or leak staging files", () =>
                WithConfigurationFolder(folder =>
                {
                    string path = Path.Combine(folder, "locked.json");
                    ConfigurationFile.Save(path, DistinctConfiguration());
                    byte[] before = File.ReadAllBytes(path);
                    using (var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        Throws<IOException>(() => ConfigurationFile.Save(path, new OrgLensConfiguration()));
                        Equal((long)before.Length, locked.Length);
                    }
                    BytesEqual(before, File.ReadAllBytes(path));
                    AssertConfiguration(DistinctConfiguration(), ConfigurationFile.Load(path));
                    AssertNoStagingFiles(folder);
                }));
            Check("disposing a pending save rolls back and cleans its temporary staging file", () =>
                WithConfigurationFolder(folder =>
                {
                    string path = Path.Combine(folder, "rollback.json");
                    ConfigurationFile.Save(path, DistinctConfiguration());
                    byte[] before = File.ReadAllBytes(path);
                    using (var pending = ConfigurationFile.PrepareSave(path, new OrgLensConfiguration()))
                    {
                        Equal(1, Directory.GetFiles(folder, ".orglens-*.tmp").Length);
                        BytesEqual(before, File.ReadAllBytes(path));
                        pending.Dispose();
                        AssertNoStagingFiles(folder);
                    }
                    BytesEqual(before, File.ReadAllBytes(path));
                    AssertConfiguration(DistinctConfiguration(), ConfigurationFile.Load(path));
                    AssertNoStagingFiles(folder);
                }));
            Check("disposing an uncommitted new-file save leaves no target or staging file", () =>
                WithConfigurationFolder(folder =>
                {
                    string path = Path.Combine(folder, "not yet created.json");
                    using (ConfigurationFile.PrepareSave(path, DistinctConfiguration()))
                    {
                        True(!File.Exists(path));
                        Equal(1, Directory.GetFiles(folder, ".orglens-*.tmp").Length);
                    }
                    True(!File.Exists(path));
                    Equal(0, Directory.GetFiles(folder).Length);
                }));
            Check("pending commit writes a captured configuration once and survives disposal", () =>
                WithConfigurationFolder(folder =>
                {
                    string path = Path.Combine(folder, "commit.json");
                    ConfigurationFile.Save(path, new OrgLensConfiguration());
                    var config = DistinctConfiguration();
                    var expected = config.Copy();
                    using (var pending = ConfigurationFile.PrepareSave(path, config))
                    {
                        ChangeConfiguration(config);
                        AssertConfiguration(new OrgLensConfiguration(), ConfigurationFile.Load(path));
                        pending.Commit();
                        AssertConfiguration(expected, ConfigurationFile.Load(path));
                        AssertNoStagingFiles(folder);
                        Throws<InvalidOperationException>(() => pending.Commit());
                    }
                    AssertConfiguration(expected, ConfigurationFile.Load(path));
                    AssertNoStagingFiles(folder);
                }));
            Check("failed pending commit preserves old content and disposal removes the staged file", () =>
                WithConfigurationFolder(folder =>
                {
                    string path = Path.Combine(folder, "commit locked.json");
                    ConfigurationFile.Save(path, DistinctConfiguration());
                    byte[] before = File.ReadAllBytes(path);
                    using (var pending = ConfigurationFile.PrepareSave(path, new OrgLensConfiguration()))
                    {
                        using (var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            Throws<IOException>(() => pending.Commit());
                            Equal((long)before.Length, locked.Length);
                        }
                        Equal(1, Directory.GetFiles(folder, ".orglens-*.tmp").Length);
                        BytesEqual(before, File.ReadAllBytes(path));
                    }
                    BytesEqual(before, File.ReadAllBytes(path));
                    AssertNoStagingFiles(folder);
                }));
        }

        private static OrgLensConfiguration DistinctConfiguration()
        {
            return new OrgLensConfiguration
            {
                BossRead = new RuleStyle { Enabled = false, Color = MailColor.Maroon, Bold = true,
                    Italic = true, FontSize = 1 },
                BossUnread = new RuleStyle { Enabled = true, Color = MailColor.Lime, Bold = false,
                    Italic = true, FontSize = 127 },
                CustomUnread = new RuleStyle { Enabled = false, Color = MailColor.Fuchsia, Bold = true,
                    Italic = false, FontSize = 23 },
                ToMe = new RuleStyle { Enabled = true, Color = MailColor.Silver, Bold = false,
                    Italic = false, FontSize = 42 },
                ToMeUnreadOnly = false,
                CustomAddresses = new List<string> { "first@example.com", "o'neil@example.com" }
            };
        }

        private static void ChangeConfiguration(OrgLensConfiguration config)
        {
            foreach (var style in new[] { config.BossRead, config.BossUnread, config.CustomUnread, config.ToMe })
            {
                style.Enabled = !style.Enabled;
                style.Bold = !style.Bold;
                style.Italic = !style.Italic;
                style.Color = style.Color == MailColor.White ? MailColor.Black : MailColor.White;
                style.FontSize = style.FontSize == 127 ? 1 : style.FontSize + 1;
            }
            config.ToMeUnreadOnly = !config.ToMeUnreadOnly;
            config.CustomAddresses.Clear();
            config.CustomAddresses.Add("changed@example.com");
        }

        private static void AssertStyle(RuleStyle expected, RuleStyle actual)
        {
            Equal(expected.Enabled, actual.Enabled, "Enabled");
            Equal(expected.Color, actual.Color, "Color");
            Equal(expected.Bold, actual.Bold, "Bold");
            Equal(expected.Italic, actual.Italic, "Italic");
            Equal(expected.FontSize, actual.FontSize, "FontSize");
        }

        private static string StyleSnapshot(RuleStyle style)
        {
            return style.Enabled + "|" + style.Color + "|" + style.Bold + "|" + style.Italic + "|" + style.FontSize;
        }

        private static void AssertConfiguration(OrgLensConfiguration expected, OrgLensConfiguration actual)
        {
            Equal(expected.Version, actual.Version);
            AssertStyle(expected.BossRead, actual.BossRead);
            AssertStyle(expected.BossUnread, actual.BossUnread);
            AssertStyle(expected.CustomUnread, actual.CustomUnread);
            AssertStyle(expected.ToMe, actual.ToMe);
            Equal(expected.ToMeUnreadOnly, actual.ToMeUnreadOnly);
            True(expected.CustomAddresses.SequenceEqual(actual.CustomAddresses), "CustomAddresses");
        }

        private static void AssertIndependent(OrgLensConfiguration left, OrgLensConfiguration right)
        {
            True(!ReferenceEquals(left, right));
            True(!ReferenceEquals(left.BossRead, right.BossRead));
            True(!ReferenceEquals(left.BossUnread, right.BossUnread));
            True(!ReferenceEquals(left.CustomUnread, right.CustomUnread));
            True(!ReferenceEquals(left.ToMe, right.ToMe));
            True(!ReferenceEquals(left.CustomAddresses, right.CustomAddresses));
        }

        private static Dictionary<string, string> StyleJsonFields()
        {
            return new Dictionary<string, string>
            {
                ["enabled"] = "true", ["color"] = "\"Blue\"", ["bold"] = "true",
                ["italic"] = "false", ["fontSize"] = "11"
            };
        }

        private static Dictionary<string, string> ConfigurationJsonFields()
        {
            string style = JsonObject(StyleJsonFields());
            return new Dictionary<string, string>
            {
                ["version"] = "2", ["bossRead"] = style, ["bossUnread"] = style, ["customUnread"] = style,
                ["toMe"] = style, ["toMeUnreadOnly"] = "true", ["customAddresses"] = "[]"
            };
        }

        private static string JsonObject(Dictionary<string, string> fields)
        {
            return "{" + string.Join(",", fields.Select(field => "\"" + field.Key + "\":" + field.Value)) + "}";
        }

        private static void RejectJson(string folder, string name, string text)
        {
            string path = Path.Combine(folder, name);
            File.WriteAllText(path, text, new UTF8Encoding(false));
            byte[] before = File.ReadAllBytes(path);
            var error = Throws<OrgLensException>(() => ConfigurationFile.Load(path), name + ": " + text);
            True(!string.IsNullOrWhiteSpace(error.Message), name);
            BytesEqual(before, File.ReadAllBytes(path));
            AssertNoStagingFiles(folder);
        }

        private static void WithConfigurationFolder(Action<string> test)
        {
            string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".orglens-tests-" +
                Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            try { test(folder); }
            finally { Directory.Delete(folder, true); }
        }

        private static void AssertNoStagingFiles(string folder)
        {
            Equal(0, Directory.GetFiles(folder, ".orglens-*.tmp").Length, "Leaked pending-save staging files");
        }

        private static void BytesEqual(byte[] expected, byte[] actual)
        {
            True(expected.SequenceEqual(actual), "File bytes changed unexpectedly");
        }
    }
}
