using System.Net.Http.Headers;
using Microsoft.Extensions.Primitives;

namespace word_dynamics_api.Helpers;

public static class HeadersHelper
{
    public static List<string> HeadersToCopy = new List<string>(new [] {
        "Retry-After"
    }); // Currently we are only interested in this header

    public static void CopyHeaders(this HttpResponseHeaders source, IHeaderDictionary target) {
        foreach(var header in source) {
            if (HeadersToCopy.Contains(header.Key, StringComparer.InvariantCultureIgnoreCase)) {   
                target.Append(header.Key, new StringValues(header.Value?.ToArray()));
            }
        }
    }
}
