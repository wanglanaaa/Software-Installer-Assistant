using System.Diagnostics;
using System.Reflection;

namespace SoftwareInstaller.Packaging;

public static class RunnerTemplate
{
    public static string CurrentVersion => typeof(RunnerTemplate).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()!.InformationalVersion;

    public static void Validate(string path)
    {
        if (!File.Exists(path)) throw new FileNotFoundException("未找到安装执行器，请使用包含 Runner 目录的完整程序。", path);
        string? version = FileVersionInfo.GetVersionInfo(path).ProductVersion;
        if (version != CurrentVersion)
            throw new InvalidDataException($"安装执行器版本 {version ?? "未知"} 与当前程序 {CurrentVersion} 不一致，请更新完整程序后重新生成。");
    }
}
