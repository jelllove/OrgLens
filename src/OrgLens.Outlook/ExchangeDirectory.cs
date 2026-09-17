using System;
using System.Collections.Generic;
using OrgLens.Core;
using OutlookApi = Microsoft.Office.Interop.Outlook;

namespace OrgLens.Outlook
{
    internal sealed class ExchangeDirectory
    {
        private readonly OutlookApi.Application application;

        public ExchangeDirectory(OutlookApi.Application application)
        {
            this.application = application;
        }

        public IReadOnlyList<MailboxAccount> GetAccounts()
        {
            var result = new List<MailboxAccount>();
            using (var scope = new ComScope())
            {
                var session = scope.Own(application.Session);
                var accounts = scope.Own(session.Accounts);
                for (int i = 1; i <= accounts.Count; i++)
                {
                    var account = scope.Own(accounts[i]);
                    if (account.AccountType != OutlookApi.OlAccountType.olExchange) continue;
                    var store = scope.Own(account.DeliveryStore);
                    if (store != null)
                        result.Add(new MailboxAccount(store.StoreID, account.DisplayName, account.SmtpAddress));
                }
            }
            return result;
        }

        public HierarchyResult Discover(string accountId)
        {
            using (var scope = new ComScope())
            {
                var session = scope.Own(application.Session);
                if (session.Offline)
                    throw new OrgLensException("Outlook is offline. Connect to Exchange and refresh the hierarchy.");
                var account = scope.Own(FindAccount(session, accountId));
                var recipient = scope.Own(account.CurrentUser);
                var entry = scope.Own(recipient.AddressEntry);
                if (entry == null)
                    throw new OrgLensException("The selected account has no Exchange address-book entry.");
                var user = scope.Own(entry.GetExchangeUser());
                if (user == null)
                    throw new OrgLensException("The selected account could not be resolved as an Exchange user.");
                return HierarchyDiscovery.Discover(ReadPerson(user), person =>
                {
                    using (var managerScope = new ComScope())
                    {
                        var currentEntry = managerScope.Own(session.GetAddressEntryFromID(person.EntryId));
                        var currentUser = managerScope.Own(currentEntry.GetExchangeUser());
                        if (currentUser == null)
                            throw new OrgLensException("An entry in the manager chain is no longer an Exchange user. " +
                                "Refresh the address book and try again.");
                        var manager = managerScope.Own(currentUser.GetExchangeUserManager());
                        return manager == null ? null : ReadPerson(manager);
                    }
                });
            }
        }

        internal static OutlookApi.Account FindAccount(OutlookApi.NameSpace session, string accountId)
        {
            using (var scope = new ComScope())
            {
                var accounts = scope.Own(session.Accounts);
                for (int i = 1; i <= accounts.Count; i++)
                {
                    var account = accounts[i];
                    bool matched = false;
                    try
                    {
                        using (var storeScope = new ComScope())
                        {
                            var store = storeScope.Own(account.DeliveryStore);
                            matched = account.AccountType == OutlookApi.OlAccountType.olExchange &&
                                store != null && string.Equals(store.StoreID, accountId, StringComparison.Ordinal);
                            if (matched) return account;
                        }
                    }
                    finally
                    {
                        if (!matched)
                        {
                            using (var accountScope = new ComScope()) accountScope.Own(account);
                        }
                    }
                }
            }
            throw new OrgLensException("The selected Exchange account is no longer available. Reopen OrgLens.");
        }

        private static ManagerPerson ReadPerson(OutlookApi.ExchangeUser user)
        {
            return new ManagerPerson(user.ID, user.Name, user.PrimarySmtpAddress, user.Address);
        }
    }
}
