using System;
using System.Runtime.InteropServices;
using System.Threading;
using OrgLens.Outlook;
using OutlookApi = Microsoft.Office.Interop.Outlook;

namespace OrgLens.Tests
{
    internal static class OutlookReadOnly
    {
        internal static int Run()
        {
            using (var timeout = new Timer(state =>
            {
                Console.Error.WriteLine("Read-only verification timed out waiting for Outlook. " +
                    "Dismiss any Outlook sign-in/security dialogs and retry. No settings were written.");
                Environment.Exit(3);
            }, null, TimeSpan.FromSeconds(45), Timeout.InfiniteTimeSpan))
            {
                Console.WriteLine("LIVE READ-ONLY: attaching to the existing Outlook process (45-second timeout).");
                object running = Marshal.GetActiveObject("Outlook.Application");
                try
                {
                    var service = new OutlookOrgLensService((OutlookApi.Application)running);
                    Console.WriteLine("LIVE READ-ONLY: reading configured Exchange accounts.");
                    var accounts = service.GetAccounts();
                    Console.WriteLine("LIVE READ-ONLY: Exchange account count = " + accounts.Count);
                    if (accounts.Count == 0)
                    {
                        Console.WriteLine("No Exchange account is configured; manager discovery was not verified.");
                        return 2;
                    }
                    Console.WriteLine("LIVE READ-ONLY: discovering the primary account's manager chain.");
                    var hierarchy = service.DiscoverManagers(accounts[0].Id);
                    var configuration = service.LoadConfiguration(accounts[0].Id);
                    Console.WriteLine("LIVE READ-ONLY: primary account manager levels = " + hierarchy.Managers.Count);
                    Console.WriteLine("LIVE READ-ONLY: loaded formatting configurations = " +
                        (configuration == null ? 0 : 1));
                    Console.WriteLine("No names, addresses, messages, or tokens were logged. No Outlook settings were written.");
                    return 0;
                }
                finally { Marshal.ReleaseComObject(running); }
            }
        }
    }
}
