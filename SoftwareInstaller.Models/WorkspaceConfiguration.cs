namespace SoftwareInstaller.Models;

public sealed class WorkspaceConfiguration
{
    public int FormatVersion { get; set; } = 1;
    public List<SoftwareItem> Software { get; set; } = new();
    public List<string> Categories { get; set; } = new();
    public Dictionary<string, SchemeDefinition> Schemes { get; set; } = new();
}
