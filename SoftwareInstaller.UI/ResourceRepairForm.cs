using SoftwareInstaller.Core;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.UI;

public sealed class ResourceRepairForm : Form
{
    private readonly WorkspaceConfiguration configuration;
    private readonly TextBox oldRoot = new() { Dock = DockStyle.Fill }, newRoot = new() { Dock = DockStyle.Fill };
    private readonly DataGridView grid = new() { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
        RowHeadersVisible = false, BackgroundColor = SystemColors.Window, AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false };
    private readonly Label status = new() { Dock = DockStyle.Bottom, Height = 34, TextAlign = ContentAlignment.MiddleLeft };
    private readonly Button preview = new() { Text = "预览修复", AutoSize = true }, apply = new() { Text = "应用修复", AutoSize = true, Enabled = false };
    public int AppliedCount { get; private set; }

    public ResourceRepairForm(WorkspaceConfiguration configuration)
    {
        this.configuration = configuration;
        AppBranding.Apply(this);
        Text = "批量修复安装源路径"; ClientSize = new(1000, 540); MinimumSize = new(820, 460); StartPosition = FormStartPosition.CenterParent;
        Padding = new(16);
        var intro = new Label { Text = "选择搬家前后的对应目录，按原有子目录结构修复失效路径。不会搜索同名文件，也不会移动文件。\n例如原目录 F:\\软件库 → 新目录 E:\\软件库；当前列表和已保存方案一同检查，有效路径保持原样。",
            Dock = DockStyle.Top, Height = 48 };
        var folders = new TableLayoutPanel { Dock = DockStyle.Top, Height = 76, ColumnCount = 3, RowCount = 2 };
        folders.ColumnStyles.Add(new(SizeType.Absolute, 108)); folders.ColumnStyles.Add(new(SizeType.Percent, 100)); folders.ColumnStyles.Add(new(SizeType.Absolute, 88));
        void Folder(int row, string title, TextBox input)
        {
            folders.RowStyles.Add(new(SizeType.Percent, 50));
            folders.Controls.Add(new Label { Text = title, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, row);
            input.Margin = new(0, 5, 8, 5); folders.Controls.Add(input, 1, row);
            var browse = new Button { Text = "浏览…", Dock = DockStyle.Fill, Margin = new(0, 2, 0, 6) };
            browse.Click += (_, _) => { using var dialog = new FolderBrowserDialog { Description = title, SelectedPath = Directory.Exists(input.Text) ? input.Text : "" };
                if (dialog.ShowDialog(this) == DialogResult.OK) input.Text = dialog.SelectedPath; };
            folders.Controls.Add(browse, 2, row);
        }
        Folder(0, "原资源根目录", oldRoot); Folder(1, "新资源根目录", newRoot);
        grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "修复", Width = 44 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "软件", Width = 120, ReadOnly = true });
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "所属", Width = 108, ReadOnly = true });
        foreach (string title in new[] { "原目录中的位置", "新目录中的位置" }) grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = title, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 140, ReadOnly = true });
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "检查结果", Width = 150, ReadOnly = true });
        grid.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
        foreach (DataGridViewColumn column in grid.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
        grid.CurrentCellDirtyStateChanged += (_, _) => { if (grid.IsCurrentCellDirty) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
        grid.CellValueChanged += (_, _) => UpdateApplyState();
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, FlowDirection = FlowDirection.RightToLeft, Padding = new(0, 8, 0, 0) };
        var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, AutoSize = true };
        buttons.Controls.Add(cancel); buttons.Controls.Add(apply); buttons.Controls.Add(preview);
        Controls.Add(grid); Controls.Add(status); Controls.Add(buttons); Controls.Add(folders); Controls.Add(intro);
        preview.Click += (_, _) => PreviewChanges(); apply.Click += (_, _) => ApplyChanges();
        oldRoot.Text = ResourceRelocation.SuggestOldRoot(configuration);
        status.Text = $"发现 {ResourceRelocation.MissingCount(configuration)} 条失效的软件记录（含方案快照）。选择新目录后点击预览。";
        void InvalidatePreview() { grid.Rows.Clear(); apply.Enabled = false; status.Text = "目录已修改，请点击预览修复。"; }
        oldRoot.TextChanged += (_, _) => InvalidatePreview(); newRoot.TextChanged += (_, _) => InvalidatePreview();
        CancelButton = cancel;
    }

    private void PreviewChanges()
    {
        grid.Rows.Clear(); apply.Enabled = false;
        try
        {
            static string Locations(string? installer, string? resources, string? relativeTo = null)
            {
                string Display(string? path) => string.IsNullOrWhiteSpace(path) || relativeTo == null ? path ?? "（未设置）" : Path.GetRelativePath(relativeTo, path);
                return "安装：" + Display(installer) + (string.IsNullOrWhiteSpace(resources) ? "" : "\n资源：" + Display(resources));
            }
            foreach (var change in ResourceRelocation.Prepare(configuration, oldRoot.Text.Trim(), newRoot.Text.Trim()))
            {
                int index = grid.Rows.Add(change.CanApply, change.Item.Name, change.Scope, Locations(change.Item.FilePath, change.Item.ResourceDirectory, oldRoot.Text.Trim()), Locations(change.Installer, change.Resources, newRoot.Text.Trim()), change.Status);
                var row = grid.Rows[index]; row.Tag = change; row.Cells[0].ReadOnly = !change.CanApply;
                row.Cells[3].ToolTipText = Locations(change.Item.FilePath, change.Item.ResourceDirectory);
                row.Cells[4].ToolTipText = Locations(change.Installer, change.Resources);
                if (!change.CanApply) row.DefaultCellStyle.ForeColor = Color.Firebrick;
            }
            UpdateApplyState();
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or ArgumentException or UnauthorizedAccessException)
        { status.Text = ex.Message; }
    }

    private List<ResourcePathChange> SelectedChanges() => grid.Rows.Cast<DataGridViewRow>()
        .Where(row => row.Cells[0].Value is true && row.Tag is ResourcePathChange { CanApply: true }).Select(row => (ResourcePathChange)row.Tag!).ToList();

    private void UpdateApplyState()
    {
        int selected = SelectedChanges().Count;
        apply.Enabled = selected > 0;
        status.Text = $"已选 {selected} 条可修复记录，共 {grid.Rows.Count} 条失效记录。应用后请在主界面保存配置。";
    }

    private void ApplyChanges()
    {
        try { AppliedCount = ResourceRelocation.Apply(SelectedChanges()); if (AppliedCount > 0) { DialogResult = DialogResult.OK; Close(); } }
        catch (IOException ex) { apply.Enabled = false; status.Text = ex.Message; }
    }
}
