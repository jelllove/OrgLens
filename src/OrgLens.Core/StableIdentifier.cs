using System;
using System.Security.Cryptography;
using System.Text;

namespace OrgLens.Core
{
    public static class StableIdentifier
    {
        public static string Hash(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("An identity is required.", nameof(value));
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(value))).Replace("-", "");
        }
    }
}
