using System;
using System.IO;
using OrgLens.Core;

namespace OrgLens.Outlook
{
    internal sealed class ProfileConfigurationStore
    {
        public OrgLensConfiguration Load(string accountId)
        {
            try { return ConfigurationFile.Load(PathFor(accountId)); }
            catch (FileNotFoundException) { return new OrgLensConfiguration(); }
            catch (DirectoryNotFoundException) { return new OrgLensConfiguration(); }
        }

        public PendingConfigurationSave Prepare(string accountId, OrgLensConfiguration configuration)
        {
            string path = PathFor(accountId);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            return ConfigurationFile.PrepareSave(path, configuration);
        }

        private static string PathFor(string accountId)
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OrgLens", "Profiles", StableIdentifier.Hash(accountId) + ".json");
        }
    }
}
