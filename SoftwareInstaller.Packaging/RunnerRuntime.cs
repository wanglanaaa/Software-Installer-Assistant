using System.Diagnostics;
using System.Text;
#if NETFRAMEWORK
using Newtonsoft.Json;
#else
using System.Text.Json;
#endif

namespace SoftwareInstaller.Packaging;

// Shared runtime operations keep the generated Framework runner independent of the maker.
public static class RunnerRuntime
{
    public const long MaxBundleBytes = 3500L * 1024 * 1024;
    public static string ExecutablePath
    {
        get { using var process = Process.GetCurrentProcess(); return process.MainModule!.FileName; }
    }

    public static T ReadJson<T>(Stream input)
    {
#if NETFRAMEWORK
        using var reader = new StreamReader(input, Encoding.UTF8, true, 4096, true);
        using var json = new JsonTextReader(reader) { MaxDepth = 64 };
        return JsonSerializer.CreateDefault().Deserialize<T>(json)!;
#else
        return JsonSerializer.Deserialize<T>(input)!;
#endif
    }

    public static string Json(object value)
    {
#if NETFRAMEWORK
        return JsonConvert.SerializeObject(value, Formatting.Indented);
#else
        return JsonSerializer.Serialize(value, new JsonSerializerOptions { WriteIndented = true });
#endif
    }

    public static async Task WriteTextAsync(string path, string text, bool append = false)
    {
        using var writer = new StreamWriter(path, append, new UTF8Encoding(false));
        await writer.WriteAsync(text).ConfigureAwait(false);
    }
}
