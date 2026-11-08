#if NETFRAMEWORK
using System.Reflection;

namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit { }
}

namespace SoftwareInstaller.Runner
{
    internal static class ApplicationConfiguration
    {
        public static void Initialize()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
        }
    }

    internal static class EmbeddedDependencies
    {
        public static void Register() => AppDomain.CurrentDomain.AssemblyResolve += (_, args) =>
        {
            if (new AssemblyName(args.Name).Name != "Newtonsoft.Json") return null;
            using var stream = typeof(EmbeddedDependencies).Assembly.GetManifestResourceStream("SoftwareInstaller.Newtonsoft.Json.dll")!;
            using var bytes = new MemoryStream();
            stream.CopyTo(bytes);
            return Assembly.Load(bytes.ToArray());
        };
    }
}
#endif
