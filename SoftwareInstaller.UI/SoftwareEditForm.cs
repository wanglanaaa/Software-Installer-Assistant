using System.Diagnostics;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.UI;

public partial class SoftwareEditForm : Form
{
    public SoftwareItem SoftwareItem { get; private set; }
    private readonly TextBox resources = new(), installDirectory = new(), directoryParameter = new();
    private readonly ComboBox directoryMode = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox category = new() { DropDownStyle = ComboBoxStyle.DropDown };
    private readonly CheckBox network = new() { Text = "此安装包需要联网", AutoSize = true };
    private readonly NumericUpDown order = new() { Minimum = 0, Maximum = 10000 };
    private readonly Label directoryHint = new() { Dock = DockStyle.Fill, AutoSize = false };
    private readonly Label installerHint = new() { Dock = DockStyle.Fill };
    private readonly TableLayoutPanel layout;
    private const int FooterHeight = 52;
    private readonly ToolTip tooltips = new();
    private InstallerProfile profile = InstallerInspector.Unknown;
    private string automaticName = "";
    private static readonly InstallDirectoryMode[] modes = [InstallDirectoryMode.Automatic, InstallDirectoryMode.InnoSetup, InstallDirectoryMode.Nsis, InstallDirectoryMode.MsiProperty, InstallDirectoryMode.Custom];

    public SoftwareEditForm(SoftwareItem? item = null, IEnumerable<SoftwareItem>? available = null, IEnumerable<string>? categories = null)
    {
        InitializeComponent();
        AppBranding.Apply(this);
        SoftwareItem = item ?? new() { Id = Guid.NewGuid().ToString("N"), Name = "", Version = "", Size = "", Description = "", Category = "未分类" };
        Text = item == null ? "新增软件" : "编辑软件";
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new(1080, 520); FormBorderStyle = FormBorderStyle.Sizable;
        Controls.Clear();
        layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new(20, 12, 20, 8), ColumnCount = 4, RowCount = 10 };
        layout.ColumnStyles.Add(new(SizeType.Absolute, 112));
        layout.ColumnStyles.Add(new(SizeType.Absolute, 300));
        layout.ColumnStyles.Add(new(SizeType.Absolute, 78));
        layout.ColumnStyles.Add(new(SizeType.Percent, 100));
        var heights = new[] { 34f, 34, 34, 50, 42, 68, 44, 74, 34, 34 };
        for (int i = 0; i < heights.Length; i++) layout.RowStyles.Add(new(SizeType.Absolute, heights[i]));

        void Row(int row, string title, Control input, string help, Button? browse = null, Control? helpControl = null)
        {
            layout.Controls.Add(new Label { Text = title, Dock = DockStyle.Fill, TextAlign = ContentAlignment.TopLeft, Padding = new(0, 8, 0, 0) }, 0, row);
            input.Dock = DockStyle.Top; input.Margin = new(4, 4, 4, 4); layout.Controls.Add(input, 1, row);
            if (browse != null) { browse.Text = "浏览…"; browse.Dock = DockStyle.Top; browse.Margin = new(3, 3, 8, 4); browse.Height = 27; layout.Controls.Add(browse, 2, row); }
            var notes = helpControl ?? new Label { Text = help, Dock = DockStyle.Fill, Padding = new(14, 5, 0, 0), ForeColor = Color.FromArgb(80, 88, 100) };
            notes.Margin = new(10, 0, 0, 0); layout.Controls.Add(notes, 3, row);
        }

        nameTextBox.Text = SoftwareItem.Name; versionTextBox.Text = SoftwareItem.Version;
        automaticName = Path.GetFileNameWithoutExtension(SoftwareItem.FilePath) ?? "";
        category.Items.AddRange(CategoryNames.Normalize((categories ?? []).Concat((available ?? []).Select(s => s.Category)).Append(SoftwareItem.Category)).Cast<object>().ToArray());
        category.Text = SoftwareItem.Category;
        descriptionTextBox.Multiline = true; descriptionTextBox.Height = 42; descriptionTextBox.Text = SoftwareItem.Description;
        filePathTextBox.Text = SoftwareItem.FilePath;
        installDirectory.Text = SoftwareItem.DefaultInstallDirectory; resources.Text = SoftwareItem.ResourceDirectory;
        silentArgsTextBox.Multiline = true; silentArgsTextBox.Height = 54; silentArgsTextBox.Text = SoftwareItem.SilentInstallArgs;
        silentArgsTextBox.PlaceholderText = "可留空：自动识别，或打开安装界面";
        network.Checked = SoftwareItem.RequiresNetwork; order.Value = Math.Clamp(SoftwareItem.InstallOrder, 0, 10000);
        directoryMode.Items.AddRange(["自动识别（推荐，无需填写类型）", "Inno Setup：使用 /DIR 设置目录", "NSIS：使用末尾的 /D 设置目录", "MSI：使用安装包支持的目录属性", "自定义：使用含 {InstallDir} 的模板"]);
        directoryMode.SelectedIndex = Math.Max(0, Array.IndexOf(modes, SoftwareItem.DirectoryMode));
        directoryParameter.Text = SoftwareItem.DirectoryParameter;

        var directoryNotes = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, Margin = new(12, 4, 0, 2) };
        directoryNotes.RowStyles.Add(new(SizeType.Absolute, 28)); directoryNotes.RowStyles.Add(new(SizeType.AutoSize)); directoryNotes.RowStyles.Add(new(SizeType.Percent, 100));
        directoryMode.Dock = DockStyle.Top; directoryNotes.Controls.Add(directoryMode, 0, 0);
        directoryParameter.Dock = DockStyle.Top; directoryNotes.Controls.Add(directoryParameter, 0, 1);
        directoryNotes.Controls.Add(directoryHint, 0, 2);

        Row(0, "软件名", nameTextBox, "默认填入安装文件名（不含扩展名），可按习惯修改。");
        Row(1, "分类", category, "选择已有分类，也可以直接输入新分类名。");
        Row(2, "版本", versionTextBox, "可留空。选择文件时会尽量读取版本信息。");
        Row(3, "说明", descriptionTextBox, "可选，写下软件用途或安装时需要留意的事。");
        Row(4, "安装文件", filePathTextBox, "", browseButton, installerHint);
        var directoryBrowse = new Button();
        directoryBrowse.Click += (_, _) => { using var dialog = new FolderBrowserDialog(); if (dialog.ShowDialog(this) == DialogResult.OK) installDirectory.Text = dialog.SelectedPath; };
        Row(5, "默认安装目录", installDirectory, "", directoryBrowse, directoryNotes);
        var resourceBrowse = new Button();
        resourceBrowse.Click += (_, _) => { using var dialog = new FolderBrowserDialog(); if (dialog.ShowDialog(this) == DialogResult.OK) resources.Text = dialog.SelectedPath; };
        Row(6, "关联资源目录", resources, "可选。安装包依赖 CAB、配置文件等时，选择包含安装文件的目录；会一同携带子目录。", resourceBrowse);
        Row(7, "静默参数", silentArgsTextBox, "可留空，MSI 无需填写。\n/S：常见 NSIS；/VERYSILENT /NORESTART：常见 Inno。\n参数无效时会打开普通安装界面。");
        Row(8, "联网需求", network, "联网安装包需要目标电脑联网；勾选后仍可打包。");
        Row(9, "安装顺序", order, "数字越小越先安装，默认 0。");

        var footer = new TableLayoutPanel { Dock = DockStyle.Bottom, Height = FooterHeight, ColumnCount = 3, Padding = new(20, 10, 20, 8) };
        footer.ColumnStyles.Add(new(SizeType.Percent, 100)); footer.ColumnStyles.Add(new(SizeType.Absolute, 90)); footer.ColumnStyles.Add(new(SizeType.Absolute, 90));
        footer.Controls.Add(new Label { Text = "未填写静默参数也能保存、打包；打开安装界面后按提示完成即可。", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.DimGray }, 0, 0);
        okButton.DialogResult = DialogResult.None; okButton.Dock = DockStyle.Fill; okButton.Click += Save_Click;
        cancelButton.Dock = DockStyle.Fill; footer.Controls.Add(okButton, 1, 0); footer.Controls.Add(cancelButton, 2, 0);
        Controls.Add(layout); Controls.Add(footer); AcceptButton = okButton; CancelButton = cancelButton;
        filePathTextBox.TextChanged += (_, _) => UpdateInstallerHelp();
        directoryMode.SelectedIndexChanged += (_, _) => UpdateDirectoryHelp();
        UpdateInstallerHelp();
        FormClosed += (_, _) => tooltips.Dispose();
    }

    private void UpdateInstallerHelp()
    {
        var path = filePathTextBox.Text.Trim();
        if (path.Length > 0 && (string.IsNullOrWhiteSpace(nameTextBox.Text) || nameTextBox.Text == automaticName))
        {
            automaticName = Path.GetFileNameWithoutExtension(path);
            nameTextBox.Text = automaticName;
        }
        profile = InstallerInspector.Inspect(path);
        installerHint.Text = $"支持 EXE / MSI。当前识别：{profile.Name}。\n识别时只读取文件，不会运行安装程序。";
        installerHint.Padding = new(14, 2, 0, 0);
        tooltips.SetToolTip(filePathTextBox, filePathTextBox.Text);
        UpdateDirectoryHelp();
    }

    private void UpdateDirectoryHelp()
    {
        var mode = modes[Math.Max(0, directoryMode.SelectedIndex)];
        directoryParameter.Visible = mode is InstallDirectoryMode.MsiProperty or InstallDirectoryMode.Custom;
        directoryParameter.PlaceholderText = mode == InstallDirectoryMode.MsiProperty ? "厂商支持的属性，例如 INSTALLDIR" : "例如 --install-dir=\"{InstallDir}\"";
        directoryHint.Text = mode switch
        {
            InstallDirectoryMode.Automatic => profile.DirectoryMode == InstallDirectoryMode.Automatic
                ? "可留空。未识别到可自动设置目录的方式；填写后会打开安装界面，请在其中选择该目录。D: 按 D:\\ 保存。"
                : $"已识别 {profile.Name}，目录参数会自动生成。可留空使用软件默认位置，例如 D:\\Apps\\软件名。",
            InstallDirectoryMode.InnoSetup => "仅适用于 Inno Setup。通过 /DIR=目录 传递位置；可留空使用软件默认目录。",
            InstallDirectoryMode.Nsis => "仅适用于 NSIS。/D=目录 会自动放在最后且不加引号；可留空。",
            InstallDirectoryMode.MsiProperty => "通常无需选择此项。MSI 不统一使用同一个目录属性，需以软件厂商说明为准。",
            _ => "通常无需选择此项。仅在明确知道厂商参数时使用，模板中的 {InstallDir} 会替换为左侧目录。"
        };
        directoryHint.ForeColor = Color.DimGray;
        ResizeToContent();
    }

    private void ResizeToContent()
    {
        float scale = DeviceDpi / 96f;
        bool hasParameter = modes[Math.Max(0, directoryMode.SelectedIndex)] is InstallDirectoryMode.MsiProperty or InstallDirectoryMode.Custom;
        layout.RowStyles[5].Height = (hasParameter ? 92 : 68) * scale;
        int height = (int)Math.Ceiling(layout.RowStyles.Cast<RowStyle>().Sum(r => r.Height)) + layout.Padding.Vertical + (int)(FooterHeight * scale);
        MinimumSize = SizeFromClientSize(new Size((int)(1060 * scale), height));
        ClientSize = new Size(Math.Max(ClientSize.Width, (int)(1060 * scale)), height);
    }

    private void Save_Click(object? sender, EventArgs e)
    {
        try
        {
            var draft = SoftwareItem.Clone();
            draft.Name = nameTextBox.Text.Trim(); draft.Version = versionTextBox.Text.Trim(); draft.Description = descriptionTextBox.Text.Trim();
            draft.Category = string.IsNullOrWhiteSpace(category.Text) ? "未分类" : category.Text.Trim();
            draft.FilePath = filePathTextBox.Text.Trim(); draft.SilentInstallArgs = silentArgsTextBox.Text.ReplaceLineEndings(" ").Trim();
            draft.DefaultInstallDirectory = InstallDirectoryOptions.Normalize(installDirectory.Text);
            draft.DirectoryMode = modes[directoryMode.SelectedIndex]; draft.DirectoryParameter = directoryParameter.Text.Trim();
            draft.ResourceDirectory = resources.Text.Trim(); draft.RequiresNetwork = network.Checked;
            draft.InstallOrder = (int)order.Value; draft.AllowInteractiveFallback = true;
            if (draft.Name.Length == 0 || !File.Exists(draft.FilePath)) throw new InvalidDataException("请填写软件名并选择存在的安装文件。");
            if (!new[] { ".exe", ".msi" }.Contains(Path.GetExtension(draft.FilePath), StringComparer.OrdinalIgnoreCase)) throw new InvalidDataException("请选择 EXE 或 MSI 安装文件。");
            InstallDirectoryOptions.BuildArguments(draft);
            if (draft.FilePath != SoftwareItem.FilePath || draft.SilentInstallArgs != SoftwareItem.SilentInstallArgs ||
                draft.DefaultInstallDirectory != SoftwareItem.DefaultInstallDirectory || draft.DirectoryMode != SoftwareItem.DirectoryMode || draft.DirectoryParameter != SoftwareItem.DirectoryParameter)
                draft.VerifiedSha256 = null;
            draft.Size = $"{new FileInfo(draft.FilePath).Length / 1048576d:F2} MB";
            foreach (var property in typeof(SoftwareItem).GetProperties().Where(p => p.CanWrite)) property.SetValue(SoftwareItem, property.GetValue(draft));
            DialogResult = DialogResult.OK; Close();
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "无法保存", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
    }

    private void browseButton_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog { Filter = "安装文件 (*.exe;*.msi)|*.exe;*.msi", Title = "选择安装文件" };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        filePathTextBox.Text = dialog.FileName;
        try
        {
            var info = FileVersionInfo.GetVersionInfo(dialog.FileName);
            if (string.IsNullOrWhiteSpace(versionTextBox.Text)) versionTextBox.Text = info.FileVersion ?? "";
            if (string.IsNullOrWhiteSpace(descriptionTextBox.Text)) descriptionTextBox.Text = info.FileDescription ?? "";
        }
        catch (Exception ex) when (ex is IOException or System.ComponentModel.Win32Exception) { if (nameTextBox.Text.Length == 0) nameTextBox.Text = Path.GetFileNameWithoutExtension(dialog.FileName); }
    }
}
