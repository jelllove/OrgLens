using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using OrgLens.Core;
using OutlookApi = Microsoft.Office.Interop.Outlook;

namespace OrgLens.Outlook
{
    public sealed class OutlookOrgLensService : IOrgLensService
    {
        private readonly OutlookApi.Application application;
        private readonly ExchangeDirectory directory;
        private readonly int threadId;
        private readonly ProfileConfigurationStore configurations = new ProfileConfigurationStore();

        public OutlookOrgLensService(OutlookApi.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            directory = new ExchangeDirectory(application);
            threadId = Thread.CurrentThread.ManagedThreadId;
        }

        public bool IsDemo { get { return false; } }

        public IReadOnlyList<MailboxAccount> GetAccounts()
        {
            CheckThread();
            return directory.GetAccounts();
        }

        public HierarchyResult DiscoverManagers(string accountId)
        {
            CheckThread();
            return directory.Discover(accountId);
        }

        public OrgLensConfiguration LoadConfiguration(string accountId)
        {
            CheckThread();
            return configurations.Load(accountId);
        }

        public void ApplyConfiguration(string accountId, OrgLensConfiguration configuration,
            IReadOnlyList<ManagerPerson> managers)
        {
            CheckThread();
            var normalized = ConfigurationValidator.Normalize(configuration);
            var rules = RuleDefinition.Build(normalized, managers);
            using (var scope = new ComScope())
            {
                var inbox = GetInbox(scope, accountId);
                if (normalized.CustomUnread.Enabled && normalized.CustomAddresses.Count > 0)
                    rules = RuleDefinition.Build(normalized, managers, ResolveCustomSenders(normalized.CustomAddresses));
                using (var pending = configurations.Prepare(accountId, normalized))
                {
                    var views = scope.Own(inbox.Views);
                    var view = scope.Own(FindView(views));
                    bool created = view == null;
                    if (created)
                    {
                        var original = scope.Own(inbox.CurrentView);
                        RequireTable(original);
                        view = scope.Own(CreatePrivateView(views, original, RuleDefinition.ViewName));
                    }
                    var table = RequireTable(view);
                    var nativeRules = scope.Own(table.AutoFormatRules);
                    var adapter = new NativeFormattingRules(nativeRules);
                    var snapshot = adapter.CaptureOwned();
                    try
                    {
                        RuleReconciler.Replace(adapter, rules);
                        // View.Save() clears custom conditions on affected Outlook builds. Save the collection only.
                        pending.Commit();
                    }
                    catch (COMException error)
                    {
                        RollBack(view, adapter, snapshot, created, error);
                        throw new OrgLensException("Outlook could not save the formatting rules. " +
                            "Previous rules were restored; old empty conditions remain disabled. " + error.Message, error);
                    }
                    catch (OrgLensException error)
                    {
                        RollBack(view, adapter, snapshot, created, error);
                        throw new OrgLensException("Outlook did not confirm the saved conditions. " +
                            "Previous rules were restored; old empty conditions remain disabled. " + error.Message, error);
                    }
                    catch (IOException error)
                    {
                        RollBack(view, adapter, snapshot, created, error);
                        throw new OrgLensException("The account's settings file could not be saved. " +
                            "Previous rules were restored; old empty conditions remain disabled. " + error.Message, error);
                    }
                    catch (UnauthorizedAccessException error)
                    {
                        RollBack(view, adapter, snapshot, created, error);
                        throw new OrgLensException("OrgLens cannot write the account's settings file. " +
                            "Previous rules were restored; old empty conditions remain disabled. " + error.Message, error);
                    }
                    try
                    {
                        var explorer = scope.Own(application.ActiveExplorer());
                        if (explorer == null)
                        {
                            explorer = scope.Own(inbox.GetExplorer());
                            explorer.Display();
                        }
                        else
                        {
                            var currentFolder = scope.Own(explorer.CurrentFolder);
                            if (currentFolder == null || currentFolder.EntryID != inbox.EntryID || currentFolder.StoreID != inbox.StoreID)
                                explorer.CurrentFolder = inbox;
                        }
                        // View.Apply() also drops conditions on affected builds; select the saved view instead.
                        explorer.CurrentView = view.Name;
                        object displayedView = explorer.CurrentView;
                        var displayed = scope.Own((OutlookApi.View)displayedView);
                        var displayedRules = scope.Own(RequireTable(displayed).AutoFormatRules);
                        new NativeFormattingRules(displayedRules).Verify(rules);
                    }
                    catch (COMException error)
                    {
                        throw new OrgLensException("Rules were saved, but Outlook could not display the view. " +
                            "Open this account's Inbox and choose View > Change View > " +
                            RuleDefinition.ViewName + ". " + error.Message, error);
                    }
                }
            }
        }

        private IReadOnlyList<ManagerPerson> ResolveCustomSenders(IReadOnlyList<string> addresses)
        {
            var senders = new List<ManagerPerson>();
            using (var scope = new ComScope())
            {
                var session = scope.Own(application.Session);
                foreach (string address in addresses)
                {
                    using (var recipientScope = new ComScope())
                    {
                        var recipient = recipientScope.Own(session.CreateRecipient(address));
                        recipient.Resolve();
                        string legacyAddress = "";
                        if (recipient.Resolved)
                        {
                            var entry = recipientScope.Own(recipient.AddressEntry);
                            if (entry != null && entry.Type == "EX")
                            {
                                legacyAddress = entry.Address;
                                var user = recipientScope.Own(entry.GetExchangeUser());
                                if (user != null && !string.IsNullOrWhiteSpace(user.PrimarySmtpAddress) &&
                                    !string.Equals(user.PrimarySmtpAddress, address, StringComparison.OrdinalIgnoreCase))
                                    senders.Add(new ManagerPerson(user.ID, address, user.PrimarySmtpAddress, legacyAddress));
                            }
                        }
                        // External SMTP senders need not resolve in the corporate address book.
                        senders.Add(new ManagerPerson(address, address, address, legacyAddress));
                    }
                }
            }
            return senders;
        }

        public void RemoveRules(string accountId)
        {
            CheckThread();
            using (var scope = new ComScope())
            {
                var inbox = GetInbox(scope, accountId);
                var views = scope.Own(inbox.Views);
                var view = scope.Own(FindView(views));
                if (view == null)
                    throw new OrgLensException("This account has no OrgLens view or saved OrgLens rules.");
                var table = RequireTable(view);
                var rules = scope.Own(table.AutoFormatRules);
                var adapter = new NativeFormattingRules(rules);
                var snapshot = adapter.CaptureOwned();
                try
                {
                    RuleReconciler.RemoveOwned(adapter);
                }
                catch (COMException error)
                {
                    RollBack(view, adapter, snapshot, false, error);
                    throw new OrgLensException("Outlook could not remove the rules. The previous view was restored. " +
                        error.Message, error);
                }
                try
                {
                    RefreshDisplayedViews(inbox);
                }
                catch (COMException error)
                {
                    throw new OrgLensException("OrgLens rules were removed, but an open Outlook window could not " +
                        "refresh. Switch away from and back to the OrgLens view to update its display. " +
                        error.Message, error);
                }
            }
        }

        private void RefreshDisplayedViews(OutlookApi.MAPIFolder inbox)
        {
            using (var scope = new ComScope())
            {
                var explorers = scope.Own(application.Explorers);
                for (int i = 1; i <= explorers.Count; i++)
                {
                    using (var explorerScope = new ComScope())
                    {
                        var explorer = explorerScope.Own(explorers[i]);
                        var folder = explorerScope.Own(explorer.CurrentFolder);
                        if (folder == null || folder.EntryID != inbox.EntryID || folder.StoreID != inbox.StoreID)
                            continue;
                        object currentView = explorer.CurrentView;
                        var displayed = explorerScope.Own((OutlookApi.View)currentView);
                        if (displayed.Name == RuleDefinition.ViewName)
                            explorer.CurrentView = RuleDefinition.ViewName;
                    }
                }
            }
        }

        private OutlookApi.MAPIFolder GetInbox(ComScope scope, string accountId)
        {
            var session = scope.Own(application.Session);
            var account = scope.Own(ExchangeDirectory.FindAccount(session, accountId));
            var store = scope.Own(account.DeliveryStore);
            return scope.Own(store.GetDefaultFolder(OutlookApi.OlDefaultFolders.olFolderInbox));
        }

        private static OutlookApi.View FindView(OutlookApi.Views views)
        {
            for (int i = 1; i <= views.Count; i++)
            {
                var view = views[i];
                if (string.Equals(view.Name, RuleDefinition.ViewName, StringComparison.Ordinal))
                {
                    if (view.SaveOption != OutlookApi.OlViewSaveOption.olViewSaveOptionThisFolderOnlyMe)
                    {
                        using (var scope = new ComScope()) scope.Own(view);
                        throw new OrgLensException("A shared or all-folders view already uses the OrgLens name. " +
                            "Rename that view before applying OrgLens. No rules were changed.");
                    }
                    return view;
                }
                using (var scope = new ComScope()) scope.Own(view);
            }
            return null;
        }

        private static OutlookApi.TableView RequireTable(OutlookApi.View view)
        {
            if (view == null || view.ViewType != OutlookApi.OlViewType.olTableView)
                throw new OrgLensException("OrgLens requires an Inbox table view. Select Compact, Single, " +
                    "or Preview under Outlook's View > Change View and try again.");
            return (OutlookApi.TableView)view;
        }

        internal static OutlookApi.View CreatePrivateView(OutlookApi.Views views, OutlookApi.View original, string name)
        {
            using (var scope = new ComScope())
            {
                var sourceRules = scope.Own(RequireTable(original).AutoFormatRules);
                var snapshot = new NativeFormattingRules(sourceRules).CaptureAll();
                string layout = original.XML;
                var created = scope.Own(views.Add(name, OutlookApi.OlViewType.olTableView,
                    OutlookApi.OlViewSaveOption.olViewSaveOptionThisFolderOnlyMe));
                try
                {
                    // Copy() can save stale source XML and discard its native conditions. Initialize only
                    // the new view's layout, then restore native rules after the last View.Save().
                    created.XML = layout;
                    created.Save();
                    var persisted = scope.Own(views[name]);
                    var rules = scope.Own(RequireTable(persisted).AutoFormatRules);
                    new NativeFormattingRules(rules).RestoreCopy(snapshot);
                    var verified = scope.Own(views[name]);
                    var savedRules = scope.Own(RequireTable(verified).AutoFormatRules);
                    new NativeFormattingRules(savedRules).VerifyCopy(snapshot);
                    return views[name];
                }
                catch (COMException error)
                {
                    RemoveIncompleteView(views, name, error);
                    throw;
                }
                catch (OrgLensException error)
                {
                    RemoveIncompleteView(views, name, error);
                    throw;
                }
            }
        }

        private static void RemoveIncompleteView(OutlookApi.Views views, string name, Exception originalError)
        {
            try
            {
                using (var scope = new ComScope())
                {
                    var failed = scope.Own(views[name]);
                    if (failed != null) failed.Delete();
                }
            }
            catch (COMException cleanupError)
            {
                throw new OrgLensException("Outlook could not create or remove the incomplete OrgLens view. " +
                    "Your source view was not edited.", new AggregateException(originalError, cleanupError));
            }
        }

        private static void RollBack(OutlookApi.View view, NativeFormattingRules adapter,
            IReadOnlyList<NativeRuleSnapshot> snapshot, bool created, Exception originalError)
        {
            try
            {
                if (created) view.Delete();
                else
                {
                    adapter.RestoreOwned(snapshot);
                }
            }
            catch (COMException rollbackError)
            {
                throw new OrgLensException("Outlook failed to save the rules AND failed to restore the view. " +
                    "Your original non-OrgLens view is unchanged. Switch to it under View > Change View. " +
                    "Inspect the OrgLens view before retrying.",
                    new AggregateException(originalError, rollbackError));
            }
            catch (OrgLensException rollbackError)
            {
                throw new OrgLensException("Outlook could not restore verified conditions. OrgLens rules were left disabled " +
                    "where possible. Your original non-OrgLens view was not changed.",
                    new AggregateException(originalError, rollbackError));
            }
        }

        private void CheckThread()
        {
            if (Thread.CurrentThread.ManagedThreadId != threadId ||
                Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new OrgLensException("Outlook operations must run on the add-in's original STA UI thread.");
        }
    }
}
