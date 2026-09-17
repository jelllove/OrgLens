using System;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OrgLens.Core
{
    public static class ConfigurationFile
    {
        public const int MaximumBytes = 1024 * 1024;

        public static OrgLensConfiguration Load(string path)
        {
            using (var stream = File.OpenRead(path))
            {
                if (stream.Length > MaximumBytes)
                    throw new OrgLensException("Settings files must be no larger than 1 MB.");
                try
                {
                    using (var text = new StreamReader(stream, new UTF8Encoding(false, true)))
                    using (var reader = new JsonTextReader(text) { MaxDepth = 16, DateParseHandling = DateParseHandling.None })
                    {
                        var document = JObject.Load(reader, new JsonLoadSettings
                        {
                            DuplicatePropertyNameHandling = DuplicatePropertyNameHandling.Error
                        });
                        if (reader.Read())
                            throw new OrgLensException("Unexpected content after the settings object.");
                        ValidateShape(document);
                        return ConfigurationValidator.Normalize(document.ToObject<OrgLensConfiguration>(Serializer()));
                    }
                }
                catch (JsonException error)
                {
                    throw new OrgLensException("This is not a valid OrgLens v2 settings file. " + error.Message, error);
                }
                catch (DecoderFallbackException error)
                {
                    throw new OrgLensException("The settings file contains invalid text encoding.", error);
                }
            }
        }

        public static void Save(string path, OrgLensConfiguration configuration)
        {
            using (var pending = PrepareSave(path, configuration)) pending.Commit();
        }

        public static PendingConfigurationSave PrepareSave(string path, OrgLensConfiguration configuration)
        {
            var normalized = ConfigurationValidator.Normalize(configuration);
            string target = Path.GetFullPath(path);
            string temporary = Path.Combine(Path.GetDirectoryName(target), ".orglens-" + Guid.NewGuid().ToString("N") + ".tmp");
            var pending = new PendingConfigurationSave(temporary, target);
            bool prepared = false;
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    using (var text = new StreamWriter(stream, new UTF8Encoding(false), 1024, true))
                    using (var writer = new JsonTextWriter(text) { Formatting = Formatting.Indented })
                    {
                        Serializer().Serialize(writer, normalized);
                        writer.Flush();
                    }
                    if (stream.Length > MaximumBytes)
                        throw new OrgLensException("The settings exceed the 1 MB file limit.");
                    stream.Flush(true);
                }
                prepared = true;
                return pending;
            }
            finally
            {
                if (!prepared) pending.Dispose();
            }
        }

        private static JsonSerializer Serializer()
        {
            return JsonSerializer.Create(new JsonSerializerSettings
            {
                TypeNameHandling = TypeNameHandling.None,
                MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Error,
                DateParseHandling = DateParseHandling.None,
                CheckAdditionalContent = true,
                MaxDepth = 16
            });
        }

        private static void ValidateShape(JObject document)
        {
            RequireType(document["version"], JTokenType.Integer, "version");
            RequireType(document["toMeUnreadOnly"], JTokenType.Boolean, "toMeUnreadOnly");
            RequireType(document["customAddresses"], JTokenType.Array, "customAddresses");
            foreach (var address in document["customAddresses"])
                RequireType(address, JTokenType.String, "customAddresses entry");
            foreach (string group in new[] { "bossRead", "bossUnread", "customUnread", "toMe" })
            {
                var style = document[group];
                RequireType(style, JTokenType.Object, group);
                RequireType(style["enabled"], JTokenType.Boolean, group + ".enabled");
                RequireType(style["bold"], JTokenType.Boolean, group + ".bold");
                RequireType(style["italic"], JTokenType.Boolean, group + ".italic");
                RequireType(style["fontSize"], JTokenType.Integer, group + ".fontSize");
                RequireType(style["color"], JTokenType.String, group + ".color");
            }
        }

        private static void RequireType(JToken token, JTokenType type, string name)
        {
            if (token == null || token.Type != type)
                throw new OrgLensException("Settings field '" + name + "' must have JSON type " + type + ".");
        }
    }

    public sealed class PendingConfigurationSave : IDisposable
    {
        private readonly string temporary;
        private readonly string target;
        private bool committed;

        internal PendingConfigurationSave(string temporary, string target)
        {
            this.temporary = temporary;
            this.target = target;
        }

        public void Commit()
        {
            if (committed) throw new InvalidOperationException("The settings have already been saved.");
            if (File.Exists(target)) File.Replace(temporary, target, null);
            else File.Move(temporary, target);
            committed = true;
        }

        public void Dispose()
        {
            if (!committed && File.Exists(temporary)) File.Delete(temporary);
        }
    }
}
