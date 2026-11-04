using System.ComponentModel;
using SoftwareInstaller.Core;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.UI;

public sealed class BatchImportForm : Form
{
    private readonly DataGridView grid = new() { Dock = DockStyle.Fill, AutoGenerateColumns = false, AllowUserToAddRows = false,
        AllowUserToDeleteRows = false, RowHeadersVisible = false, BackgroundColor = Color.White, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
    private readonly BindingList<SoftwareItem> drafts;
    public List<SoftwareItem> Software => drafts.Where(s => s.IsSelected).Select(s => s.Clone()).ToList();

    public BatchImportForm(ImportPreparation preparation, IEnumerable<string> categories)
    {
        AppBranding.Apply(this);
        Text = "批量添加软件"; ClientSize = new(920, 470); MinimumSize = new(800, 440); StartPosition = FormStartPosition.CenterParent;
        drafts = new(preparation.Items);
        var help = new Label { Dock = DockStyle.Top, Height = 58, Padding = new(16, 12, 16, 4),
            Text = $"已读取 {drafts.Count} 个安装文件，跳过 {preparation.Duplicates} 个重复文件。\n可修改名称、分类和版本；添加后双击软件，可继续设置安装目录、静默参数及关联资源。" };
        var categoryBar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 44, Padding = new(16, 4, 16, 4), WrapContents = false };
        var category = new ComboBox { Width = 200, DropDownStyle = ComboBoxStyle.DropDown };
        category.Items.AddRange(CategoryNames.Normalize(categories.Concat(drafts.Select(s => s.Category))).Cast<object>().ToArray());
        category.Text = drafts[0].Category;
        var apply = new Button { Text = "统一分类", Size = new(96, 28) };
        apply.Click += (_, _) => { grid.EndEdit(); foreach (var item in drafts.Where(s => s.IsSelected)) item.Category = string.IsNullOrWhiteSpace(category.Text) ? "未分类" : category.Text.Trim(); grid.Refresh(); };
        categoryBar.Controls.Add(new Label { Text = "分类", AutoSize = true, Margin = new(0, 7, 8, 0) }); categoryBar.Controls.Add(category); categoryBar.Controls.Add(apply);
        grid.Columns.Add(new DataGridViewCheckBoxColumn { DataPropertyName = nameof(SoftwareItem.IsSelected), HeaderText = "添加", Width = 52 });
        void Column(string property, string title, int width, bool readOnly = false) => grid.Columns.Add(new DataGridViewTextBoxColumn
            { DataPropertyName = property, HeaderText = title, Width = width, ReadOnly = readOnly, SortMode = DataGridViewColumnSortMode.NotSortable });
        Column(nameof(SoftwareItem.Name), "软件名", 190); Column(nameof(SoftwareItem.Category), "分类", 110); Column(nameof(SoftwareItem.Version), "版本", 100);
        Column(nameof(SoftwareItem.Size), "大小", 90, true); Column(nameof(SoftwareItem.FilePath), "安装文件", 320, true);
        grid.DataSource = drafts;
        grid.CurrentCellDirtyStateChanged += (_, _) => { if (grid.IsCurrentCellDirty) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
        var footer = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 58, Padding = new(12), FlowDirection = FlowDirection.RightToLeft };
        var cancel = new Button { Text = "取消", Size = new(96, 32), DialogResult = DialogResult.Cancel };
        var confirm = new Button { Text = "确认添加", Size = new(110, 32) };
        confirm.Click += (_, _) =>
        {
            grid.EndEdit();
            if (!drafts.Any(s => s.IsSelected)) { MessageBox.Show(this, "请至少勾选一款软件。"); return; }
            foreach (var item in drafts.Where(s => s.IsSelected))
            {
                item.Name = item.Name?.Trim() ?? "";
                if (item.Name.Length == 0) { MessageBox.Show(this, "软件名不能为空。"); return; }
                item.Category = string.IsNullOrWhiteSpace(item.Category) ? "未分类" : item.Category.Trim();
            }
            DialogResult = DialogResult.OK; Close();
        };
        footer.Controls.Add(cancel); footer.Controls.Add(confirm);
        Controls.Add(grid); Controls.Add(categoryBar); Controls.Add(help); Controls.Add(footer);
        AcceptButton = confirm; CancelButton = cancel;
    }
}
