using System.Diagnostics;
using System.Reflection;

namespace SoftwareInstaller.Packaging;

public static class RunnerTemplate
{
    public static string CurrentVersion => typeof(RunnerTemplate).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()!.InformationalVersion.Split('+')[0];
    public static string ExpectedRunnerVersion => CurrentVersion + "+selfcontained-demand-flow";

    public static void Validate(string path)
    {
        if (!File.Exists(path)) throw new FileNotFoundException("未找到安装执行器，请使用包含 Runner 目录的完整程序。", path);
        string? version = FileVersionInfo.GetVersionInfo(path).ProductVersion;
        if (version != ExpectedRunnerVersion)
            throw new InvalidDataException($"安装执行器版本 {version ?? "未知"} 与所需模板 {ExpectedRunnerVersion} 不一致，请更新完整程序后重新生成。");
    }
}
