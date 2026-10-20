using System;
using System.Linq;
using System.Net;
using System.Net.Sockets;

namespace OfficeTranslate.Core
{
    public static class EndpointSecurity
    {
        public static bool RequiresPlaintextKeyWarning(string baseUrl, string apiKey) =>
            RequiresPlaintextKeyWarning(baseUrl, apiKey, Dns.GetHostAddresses);

        internal static bool RequiresPlaintextKeyWarning(string baseUrl, string apiKey, Func<string, IPAddress[]> resolve)
        {
            if (string.IsNullOrWhiteSpace(apiKey) || !Uri.TryCreate(baseUrl, UriKind.Absolute, out var uri) ||
                uri.Scheme != Uri.UriSchemeHttp) return false;
            var host = uri.DnsSafeHost.TrimEnd('.');
            if (host.Equals("localhost", StringComparison.OrdinalIgnoreCase) ||
                host.EndsWith(".localhost", StringComparison.OrdinalIgnoreCase) ||
                host.Equals(Environment.MachineName, StringComparison.OrdinalIgnoreCase)) return false;
            if (IPAddress.TryParse(host, out var address)) return !IsLocalAddress(address);
            try
            {
                // A LAN DNS name is exempt only when every resolved address is local.
                var addresses = resolve(host);
                return addresses.Length == 0 || !addresses.All(IsLocalAddress);
            }
            catch (SocketException) { return true; }
            catch (ArgumentException) { return true; }
        }

        private static bool IsLocalAddress(IPAddress address)
        {
            if (address.IsIPv4MappedToIPv6) address = address.MapToIPv4();
            if (IPAddress.IsLoopback(address)) return true;
            var bytes = address.GetAddressBytes();
            if (address.AddressFamily == AddressFamily.InterNetwork)
                return bytes[0] == 10 || (bytes[0] == 172 && bytes[1] >= 16 && bytes[1] <= 31) ||
                    (bytes[0] == 192 && bytes[1] == 168) || (bytes[0] == 169 && bytes[1] == 254);
            return address.IsIPv6LinkLocal || (bytes[0] & 0xfe) == 0xfc;
        }
    }
}
