using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;

namespace OrgLens.Core
{
    public static class ConfigurationValidator
    {
        public static OrgLensConfiguration Normalize(OrgLensConfiguration configuration)
        {
            if (configuration == null) throw new OrgLensException("The settings file contains no configuration.");
            if (configuration.Version != OrgLensConfiguration.CurrentVersion)
                throw new OrgLensException("Unsupported settings version: " + configuration.Version + ".");
            ValidateStyle(configuration.BossRead);
            ValidateStyle(configuration.BossUnread);
            ValidateStyle(configuration.CustomUnread);
            ValidateStyle(configuration.ToMe);
            if (configuration.CustomAddresses == null)
                throw new OrgLensException("The settings file has no customAddresses list.");
            var copy = configuration.Copy();
            copy.CustomAddresses = configuration.CustomAddresses.Select(NormalizeAddress)
                .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            return copy;
        }

        public static void ValidateStyle(RuleStyle style)
        {
            if (style == null) throw new OrgLensException("A required rule style is missing.");
            if (!Enum.IsDefined(typeof(MailColor), style.Color))
                throw new OrgLensException("Choose a supported Outlook font color.");
            if (style.FontSize < 1 || style.FontSize > 127)
                throw new OrgLensException("Font size must be a whole number from 1 to 127 points.");
        }

        public static List<string> ParseAddresses(string text)
        {
            if (text == null) throw new ArgumentNullException(nameof(text));
            return text.Split(new[] { '\r', '\n', ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(NormalizeAddress).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        public static string NormalizeAddress(string value)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Any(char.IsControl))
                throw new OrgLensException("Enter a plain email address, not a blank value or a display name.");
            string address = value.Trim();
            try
            {
                var parsed = new MailAddress(address);
                if (!string.Equals(parsed.Address, address, StringComparison.OrdinalIgnoreCase) ||
                    parsed.DisplayName.Length != 0 || address.IndexOf('@') <= 0 || address.Length > 254)
                    throw new OrgLensException("Enter a plain email address: " + address);
            }
            catch (FormatException error)
            {
                throw new OrgLensException("Invalid email address: " + address, error);
            }
            return address.ToLowerInvariant();
        }
    }
}
