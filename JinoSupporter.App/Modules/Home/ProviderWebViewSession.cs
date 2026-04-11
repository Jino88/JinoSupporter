using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;

namespace JinoSupporter.App.Modules.Home;

internal static class ProviderWebViewSession
{
    private static readonly Dictionary<string, CoreWebView2Environment> Environments = new(StringComparer.OrdinalIgnoreCase);

    public static async Task<CoreWebView2Environment> GetEnvironmentAsync(string providerKey)
    {
        if (Environments.TryGetValue(providerKey, out CoreWebView2Environment? environment))
        {
            return environment;
        }

        string userDataFolderPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "JinoWorkHost",
            "WebView2",
            providerKey);

        Directory.CreateDirectory(userDataFolderPath);
        environment = await CoreWebView2Environment.CreateAsync(userDataFolder: userDataFolderPath);
        Environments[providerKey] = environment;
        return environment;
    }
}
