using System;
using System.Collections.Generic;

namespace OrgLens.Core
{
    public static class HierarchyDiscovery
    {
        public static HierarchyResult Discover(ManagerPerson self, Func<ManagerPerson, ManagerPerson> getManager)
        {
            if (self == null) throw new ArgumentNullException(nameof(self));
            if (getManager == null) throw new ArgumentNullException(nameof(getManager));
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<ManagerPerson>();
            var current = self;
            visited.Add(Identity(current));

            while (true)
            {
                var manager = getManager(current);
                if (manager == null)
                {
                    return new HierarchyResult(result, result.Count == 0
                        ? "No manager is published for this account in the Exchange address book. " +
                          "Ask your directory administrator to check the manager relationship."
                        : "Found " + result.Count + " manager(s). The address book publishes no further manager. " +
                          "This is the published chain, which may not include every level in your organization.");
                }

                if (!visited.Add(Identity(manager)))
                    throw new OrgLensException("The address book contains a circular manager relationship. " +
                        "No rules were changed. Ask your directory administrator to correct the hierarchy.");

                // A corrupt directory must not keep Outlook's UI thread busy indefinitely.
                if (result.Count >= 100)
                    throw new OrgLensException("The published manager chain exceeds the 100-level safety limit. " +
                        "No partial hierarchy will be applied.");

                result.Add(manager.AtLevel(result.Count + 1));
                current = manager;
            }
        }

        private static string Identity(ManagerPerson person)
        {
            if (!string.IsNullOrWhiteSpace(person.EntryId)) return person.EntryId;
            if (!string.IsNullOrWhiteSpace(person.LegacyAddress)) return person.LegacyAddress;
            if (!string.IsNullOrWhiteSpace(person.SmtpAddress)) return person.SmtpAddress;
            throw new OrgLensException("An address-book entry has no stable identity. No rules were changed.");
        }
    }
}
