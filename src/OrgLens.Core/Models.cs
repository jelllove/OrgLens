using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace OrgLens.Core
{
    public enum MailColor
    {
        Automatic = 0,
        Black = 1,
        Maroon = 2,
        Green = 3,
        Olive = 4,
        Navy = 5,
        Purple = 6,
        Teal = 7,
        Gray = 8,
        Silver = 9,
        Red = 10,
        Lime = 11,
        Yellow = 12,
        Blue = 13,
        Fuchsia = 14,
        Aqua = 15,
        White = 16
    }

    public sealed class MailboxAccount
    {
        public MailboxAccount(string id, string displayName, string smtpAddress)
        {
            Id = id;
            DisplayName = displayName;
            SmtpAddress = smtpAddress;
        }

        public string Id { get; }
        public string DisplayName { get; }
        public string SmtpAddress { get; }
        public override string ToString() { return DisplayName + " <" + SmtpAddress + ">"; }
    }

    public sealed class ManagerPerson
    {
        public ManagerPerson(string entryId, string displayName, string smtpAddress,
            string legacyAddress, int level = 0)
        {
            EntryId = entryId;
            DisplayName = displayName;
            SmtpAddress = smtpAddress;
            LegacyAddress = legacyAddress;
            Level = level;
        }

        public string EntryId { get; }
        public string DisplayName { get; }
        public string SmtpAddress { get; }
        public string LegacyAddress { get; }
        public int Level { get; }
        public string RoleLabel
        {
            get { return Level == 1 ? "Direct manager / BOSS" : "Manager +" + Level; }
        }

        public ManagerPerson AtLevel(int level)
        {
            return new ManagerPerson(EntryId, DisplayName, SmtpAddress, LegacyAddress, level);
        }
    }

    [DataContract]
    public sealed class RuleStyle
    {
        [DataMember(Name = "enabled", IsRequired = true)]
        public bool Enabled { get; set; } = true;
        public MailColor Color { get; set; } = MailColor.Blue;
        [DataMember(Name = "bold", IsRequired = true)]
        public bool Bold { get; set; } = true;
        [DataMember(Name = "italic", IsRequired = true)]
        public bool Italic { get; set; }
        [DataMember(Name = "fontSize", IsRequired = true)]
        public int FontSize { get; set; } = 11;
        [DataMember(Name = "color", IsRequired = true)]
        public string ColorName
        {
            get { return Color.ToString(); }
            set
            {
                MailColor parsed;
                if (!Enum.TryParse(value, false, out parsed) || !Enum.IsDefined(typeof(MailColor), parsed))
                    throw new SerializationException("Unknown Outlook color: " + value);
                Color = parsed;
            }
        }

        public RuleStyle Copy()
        {
            return new RuleStyle
            {
                Enabled = Enabled, Color = Color, Bold = Bold, Italic = Italic, FontSize = FontSize
            };
        }
    }

    [DataContract]
    public sealed class OrgLensConfiguration
    {
        public const int CurrentVersion = 2;
        [DataMember(Name = "version", IsRequired = true)]
        public int Version { get; set; } = CurrentVersion;
        [DataMember(Name = "bossRead", IsRequired = true)]
        public RuleStyle BossRead { get; set; } = new RuleStyle { Color = MailColor.Navy, Bold = false };
        [DataMember(Name = "bossUnread", IsRequired = true)]
        public RuleStyle BossUnread { get; set; } = new RuleStyle { Color = MailColor.Blue, FontSize = 12 };
        [DataMember(Name = "customUnread", IsRequired = true)]
        public RuleStyle CustomUnread { get; set; } = new RuleStyle { Color = MailColor.Purple };
        [DataMember(Name = "toMe", IsRequired = true)]
        public RuleStyle ToMe { get; set; } = new RuleStyle { Color = MailColor.Teal, Bold = false };
        [DataMember(Name = "toMeUnreadOnly", IsRequired = true)]
        public bool ToMeUnreadOnly { get; set; } = true;
        [DataMember(Name = "customAddresses", IsRequired = true)]
        public List<string> CustomAddresses { get; set; } = new List<string>();

        public bool BossEnabled { get { return BossRead.Enabled || BossUnread.Enabled; } }

        public OrgLensConfiguration Copy()
        {
            return new OrgLensConfiguration
            {
                Version = Version,
                BossRead = BossRead?.Copy(),
                BossUnread = BossUnread?.Copy(),
                CustomUnread = CustomUnread?.Copy(),
                ToMe = ToMe?.Copy(),
                ToMeUnreadOnly = ToMeUnreadOnly,
                CustomAddresses = CustomAddresses?.ToList()
            };
        }
    }

    public enum RuleKind { BossUnread, BossRead, CustomUnread, ToMe }

    public sealed class FormattingRule
    {
        public FormattingRule(RuleKind kind, string filter, RuleStyle style)
        {
            Kind = kind;
            Filter = filter;
            Style = style ?? throw new ArgumentNullException(nameof(style));
        }
        public RuleKind Kind { get; }
        public string Name { get { return "OrgLens.v2:" + Kind; } }
        public string Filter { get; }
        public RuleStyle Style { get; }
    }

    public sealed class HierarchyResult
    {
        public HierarchyResult(IReadOnlyList<ManagerPerson> managers, string notice)
        {
            Managers = managers;
            Notice = notice;
        }

        public IReadOnlyList<ManagerPerson> Managers { get; }
        public string Notice { get; }
    }

    public interface IOrgLensService
    {
        bool IsDemo { get; }
        IReadOnlyList<MailboxAccount> GetAccounts();
        HierarchyResult DiscoverManagers(string accountId);
        OrgLensConfiguration LoadConfiguration(string accountId);
        void ApplyConfiguration(string accountId, OrgLensConfiguration configuration,
            IReadOnlyList<ManagerPerson> managers);
        void RemoveRules(string accountId);
    }

    public sealed class OrgLensException : Exception
    {
        public OrgLensException(string message) : base(message) { }
        public OrgLensException(string message, Exception innerException) : base(message, innerException) { }
    }
}
