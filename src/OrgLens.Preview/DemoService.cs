using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OrgLens.Core;

namespace OrgLens.Preview
{
    internal sealed class DemoService : IOrgLensService
    {
        private readonly int uiThreadId = Thread.CurrentThread.ManagedThreadId;
        private readonly IReadOnlyList<MailboxAccount> accounts = new List<MailboxAccount>
        {
            new MailboxAccount("sample-avery", "Avery Morgan · Sample account", "avery.morgan@example.com"),
            new MailboxAccount("sample-quinn", "Quinn Lee · No-manager example", "quinn.lee@example.com")
        }.AsReadOnly();
        private readonly IReadOnlyList<ManagerPerson> managers = new List<ManagerPerson>
        {
            new ManagerPerson("sample-jamie", "Jamie Rivera", "jamie.rivera@example.com", "", 1),
            new ManagerPerson("sample-morgan", "Morgan Chen", "morgan.chen@example.com", "", 2),
            new ManagerPerson("sample-taylor", "Taylor Ellis", "taylor.ellis@example.com", "", 3)
        }.AsReadOnly();
        private readonly Dictionary<string, OrgLensConfiguration> settings =
            new Dictionary<string, OrgLensConfiguration>(StringComparer.Ordinal);
        private readonly HashSet<string> appliedAccounts = new HashSet<string>(StringComparer.Ordinal);

        public bool IsDemo { get { return true; } }

        public IReadOnlyList<MailboxAccount> GetAccounts()
        {
            RequireUiThread();
            return accounts;
        }

        public HierarchyResult DiscoverManagers(string accountId)
        {
            ValidateAccount(accountId);
            return accountId == accounts[0].Id
                ? new HierarchyResult(managers,
                    "3 manager levels found. This fictional chain ends at Taylor Ellis.")
                : new HierarchyResult(new List<ManagerPerson>().AsReadOnly(),
                    "This fictional account has no manager configured. Try Avery Morgan to explore sample styles.");
        }

        public OrgLensConfiguration LoadConfiguration(string accountId)
        {
            ValidateAccount(accountId);
            OrgLensConfiguration stored;
            return settings.TryGetValue(accountId, out stored) ? stored.Copy() : new OrgLensConfiguration
            {
                CustomAddresses = new List<string> { "casey.reed@example.com" }
            };
        }

        public void ApplyConfiguration(string accountId, OrgLensConfiguration configuration,
            IReadOnlyList<ManagerPerson> people)
        {
            ValidateAccount(accountId);
            var normalized = ConfigurationValidator.Normalize(configuration);
            if (people != null && people.Any(person => accountId != accounts[0].Id
                || !managers.Any(sample => sample.SmtpAddress == person.SmtpAddress)))
            {
                throw new OrgLensException("The demo accepts only the selected account's fictional directory members.");
            }
            RuleDefinition.Validate(RuleDefinition.Build(normalized, people));
            settings[accountId] = normalized;
            appliedAccounts.Add(accountId);
        }

        public void RemoveRules(string accountId)
        {
            ValidateAccount(accountId);
            appliedAccounts.Remove(accountId);
        }

        private void ValidateAccount(string accountId)
        {
            RequireUiThread();
            if (!accounts.Any(account => account.Id == accountId))
            {
                throw new OrgLensException("This is not a sample account. The demo never connects to Outlook.");
            }
        }

        private void RequireUiThread()
        {
            if (Thread.CurrentThread.ManagedThreadId != uiThreadId
                || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                throw new InvalidOperationException("Demo service calls must run on the original STA UI thread.");
            }
        }
    }
}
