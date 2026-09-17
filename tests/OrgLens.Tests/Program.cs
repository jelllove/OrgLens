using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using OrgLens.Core;
using OrgLens.Outlook;

namespace OrgLens.Tests
{
    internal static partial class Program
    {
        private static int passed;
        private static int failed;

        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length == 1 && args[0] == "--installer-process-fixture")
            {
                Console.WriteLine("Ready: harmless installer process-detection fixture.");
                System.Threading.Thread.Sleep(TimeSpan.FromMinutes(2));
                return 0;
            }
            if (args.Length == 1 && args[0] == "--outlook-read-only") return OutlookReadOnly.Run();
            if (args.Length == 1 && args[0] == "--rule-persistence")
            {
                CheckReconciliation();
                Console.WriteLine(passed + " passed, " + failed + " failed rule persistence checks.");
                return failed == 0 ? 0 : 1;
            }
            if (args.Length != 0)
                throw new ArgumentException("Usage: OrgLens.Tests.exe [--outlook-read-only | --rule-persistence | --installer-process-fixture]");
            CheckHierarchy();
            CheckDaslEvaluator();
            CheckGroupedRules();
            CheckReconciliation();
            CheckConfiguration();
            CheckPersistence();
            CheckAddInInterfaces();
            CheckIconAssets();
            Console.WriteLine((failed == 0 ? "PASS: " : "FAIL: ") + passed + " passed, " +
                failed + " failed OrgLens checks.");
            return failed == 0 ? 0 : 1;
        }

        private static void CheckHierarchy()
        {
            var self = Person("self");
            var boss = Person("boss");
            var chief = Person("chief");
            Check("discovers all published levels with shared BOSS membership", () =>
            {
                var chain = HierarchyDiscovery.Discover(self, p => p.EntryId == "self" ? boss :
                    p.EntryId == "boss" ? chief : null);
                Equal(2, chain.Managers.Count);
                Equal(1, chain.Managers[0].Level);
                Equal("Direct manager / BOSS", chain.Managers[0].RoleLabel);
                Equal(2, chain.Managers[1].Level);
                Equal("Manager +2", chain.Managers[1].RoleLabel);
                Equal("chief", chain.Managers[1].EntryId);
                Equal(0, boss.Level);
                True(chain.Notice.Contains("published chain"));
            });
            Check("empty hierarchy is explicit", () =>
            {
                var chain = HierarchyDiscovery.Discover(self, p => null);
                Equal(0, chain.Managers.Count);
                True(chain.Notice.Contains("No manager"));
            });
            Check("rejects self manager", () =>
                Throws<OrgLensException>(() => HierarchyDiscovery.Discover(self, p => self)));
            Check("rejects cycles through earlier managers", () =>
                Throws<OrgLensException>(() => HierarchyDiscovery.Discover(self,
                    p => p.EntryId == "self" ? boss : p.EntryId == "boss" ? chief : boss)));
            Check("hierarchy cycle identities are case insensitive", () =>
                Throws<OrgLensException>(() => HierarchyDiscovery.Discover(self, p => Person("SELF"))));
            Check("hierarchy falls back to Exchange and SMTP identities", () =>
            {
                foreach (var person in new[]
                {
                    new ManagerPerson("", "DN identity", "", "/o=Example/cn=Self"),
                    new ManagerPerson("", "SMTP identity", "self@example.com", "")
                })
                    Throws<OrgLensException>(() => HierarchyDiscovery.Discover(person, p =>
                        new ManagerPerson("", "Renamed", p.SmtpAddress.ToUpperInvariant(),
                            p.LegacyAddress.ToUpperInvariant())));
            });
            Check("identity-free directory entries are rejected", () =>
            {
                var missing = new ManagerPerson("", "Missing", "", "");
                Throws<OrgLensException>(() => HierarchyDiscovery.Discover(missing, p => null));
                Throws<OrgLensException>(() => HierarchyDiscovery.Discover(self, p => missing));
            });
            Check("directory errors propagate instead of truncating", () =>
            {
                int calls = 0;
                var error = new InvalidOperationException("Directory unavailable");
                Equal(error, Throws<InvalidOperationException>(() => HierarchyDiscovery.Discover(self,
                    p => ++calls == 1 ? boss : throw error)));
                Equal(2, calls);
            });
            Check("hierarchy null arguments fail explicitly", () =>
            {
                Throws<ArgumentNullException>(() => HierarchyDiscovery.Discover(null, p => null));
                Throws<ArgumentNullException>(() => HierarchyDiscovery.Discover(self, null));
            });
            Check("100 levels are supported", () =>
            {
                int counter = 0;
                var chain = HierarchyDiscovery.Discover(self,
                    p => ++counter <= 100 ? Person("manager" + counter) : null);
                Equal(100, chain.Managers.Count);
                Equal(100, chain.Managers[99].Level);
                Equal(101, counter);
            });
            Check("oversized chains fail within the safety bound", () =>
            {
                int counter = 0;
                Throws<OrgLensException>(() => HierarchyDiscovery.Discover(self, p => Person("m" + ++counter)));
                Equal(101, counter);
            });
        }

        private static void CheckAddInInterfaces()
        {
            Check("Ribbon XML exposes manager settings", () =>
            {
                string xml = new Connect().GetCustomUI("Microsoft.Outlook.Explorer");
                var document = XDocument.Parse(xml);
                XNamespace ns = "http://schemas.microsoft.com/office/2009/07/customui";
                var button = document.Descendants(ns + "button").Single();
                Equal("OpenSettings", (string)button.Attribute("onAction"));
                Equal("GetRibbonImage", (string)button.Attribute("getImage"));
                Equal(null, (string)button.Attribute("imageMso"));
                Equal(null, new Connect().GetCustomUI("Microsoft.Outlook.Mail.Read"));
            });
            Check("COM exposes add-in, Ribbon, and callback interfaces without Outlook", () =>
            {
                var unknown = Marshal.GetIUnknownForObject(new Connect());
                try
                {
                    foreach (var id in new[] { "B65AD801-ABAF-11D0-BB8B-00A0C90F2744",
                        "000C0396-0000-0000-C000-000000000046", "20FA2F90-F84D-40E7-A36D-B2F484B2B58E" })
                    {
                        var guid = new Guid(id);
                        IntPtr pointer;
                        Equal(0, Marshal.QueryInterface(unknown, ref guid, out pointer));
                        Marshal.Release(pointer);
                    }
                }
                finally { Marshal.Release(unknown); }
            });
            Check("service contract exposes full configuration rather than per-manager styles", () =>
            {
                var service = typeof(IOrgLensService);
                Equal(typeof(OrgLensConfiguration),
                    service.GetMethod("LoadConfiguration", new[] { typeof(string) }).ReturnType);
                Equal(typeof(void), service.GetMethod("ApplyConfiguration", new[]
                {
                    typeof(string), typeof(OrgLensConfiguration), typeof(IReadOnlyList<ManagerPerson>)
                }).ReturnType);
                True(service.GetMethod("GetAccounts") != null);
                True(service.GetMethod("DiscoverManagers") != null);
                True(service.GetMethod("RemoveRules") != null);
                Equal(null, service.GetMethod("LoadRules"));
                Equal(null, service.GetMethod("ApplyRules"));
            });
        }

        private static ManagerPerson Person(string name)
        {
            return new ManagerPerson(name, name, name + "@example.com", "/o=Example/ou=Users/cn=" + name);
        }

        private static void Check(string name, Action test)
        {
            try { test(); }
            catch (Exception error)
            {
                failed++;
                Console.Error.WriteLine("FAIL " + name + Environment.NewLine + error);
                return;
            }
            passed++;
            Console.WriteLine("PASS " + name);
        }

        private static void True(bool condition, string context = null)
        {
            if (!condition) throw new InvalidOperationException("Assertion failed. " + context);
        }

        private static void Equal<T>(T expected, T actual, string context = null)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
                throw new InvalidOperationException("Expected " + expected + ", actual " + actual + ". " + context);
        }

        private static T Throws<T>(Action action, string context = null) where T : Exception
        {
            try { action(); }
            catch (T error) { return error; }
            throw new InvalidOperationException("Expected exception " + typeof(T).Name + ". " + context);
        }
    }
}
