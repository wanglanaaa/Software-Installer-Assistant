using SoftwareInstaller.Core;
using SoftwareInstaller.Models;
using SoftwareInstaller.Utils;

namespace SoftwareInstaller.UI;

public partial class MainForm
{
    private readonly TextBox softwareSearch = new() { PlaceholderText = "搜索软件名、版本或说明" };
    private readonly ComboBox softwareCategory = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private bool updatingSoftwareList, preparingImport;
    private int configurationRevision;
    private sealed record CategoryFilter(string? Name)
    {
        public override string ToString() => Name ?? "全部分类";
    }

    private void InitializeSoftwareList()
    {
        InitializeSoftwareIcons();
        Controls.Add(softwareSearch); Controls.Add(softwareCategory);
        softwareListView.CheckBoxes = true; softwareListView.HideSelection = false;
        softwareListView.Columns.Add("分类", 88);
        categoryTreeView.AfterSelect += (_, _) => UpdateImportAvailability();
        softwareSearch.TextChanged += (_, _) => UpdateSoftwareListView();
        softwareCategory.SelectedIndexChanged += (_, _) => UpdateSoftwareListView();
        softwareListView.ItemChecked += (_, e) =>
        {
            if (updatingSoftwareList || e.Item.Tag is not SoftwareItem item) return;
            item.IsSelected = e.Item.Checked;
            bool previous = isUpdatingByCode; isUpdatingByCode = true;
            try
            {
                void Update(TreeNodeCollection nodes)
                {
                    foreach (TreeNode node in nodes)
                    {
                        if (node.Tag is SoftwareItem entry && entry.Id == item.Id) node.Checked = item.IsSelected;
                        Update(node.Nodes);
                        if (node.Tag is not SoftwareItem) UpdateParentNodeCheckState(node);
                    }
                }
                Update(categoryTreeView.Nodes);
            }
            finally { isUpdatingByCode = previous; }
            UpdateSelectionSummary();
        };
        softwareListView.ItemActivate += (_, _) =>
        {
            if (softwareListView.FocusedItem?.Tag is SoftwareItem item) EditSoftware(item);
        };
        foreach (var control in new Control[] { this, softwareListView, categoryTreeView })
        {
            control.AllowDrop = true;
            control.DragEnter += (_, e) => e.Effect = !preparingImport && ImportCategory() != null && e.Data?.GetDataPresent(DataFormats.FileDrop) == true ? DragDropEffects.Copy : DragDropEffects.None;
            control.DragDrop += async (_, e) =>
            {
                if (e.Data?.GetData(DataFormats.FileDrop) is string[] files) await ImportSoftwareAsync(files);
            };
        }
        UpdateImportAvailability();
    }

    private void RefreshSoftwareCategories()
    {
        string? selected = (softwareCategory.SelectedItem as CategoryFilter)?.Name;
        updatingSoftwareList = true;
        try
        {
            softwareCategory.Items.Clear(); softwareCategory.Items.Add(new CategoryFilter(null));
            foreach (var name in CategoryNames.Normalize(_softwareManager!.ActiveCategories.Concat(_softwareManager.ActiveSoftwareItems.Select(s => s.Category))))
                softwareCategory.Items.Add(new CategoryFilter(name));
            softwareCategory.SelectedIndex = Math.Max(0, softwareCategory.Items.Cast<CategoryFilter>().ToList().FindIndex(c => c.Name == selected));
        }
        finally { updatingSoftwareList = false; }
    }

    private void UpdateSelectionSummary()
    {
        var all = _softwareManager?.ActiveSoftwareItems;
        selectionSummary.Text = $"已添加 {all?.Count ?? 0} 款 · 已勾选 {all?.Count(s => s.IsSelected) ?? 0} 款";
    }

    private void EditSoftware(SoftwareItem item)
    {
        using var form = new SoftwareEditForm(item, _softwareManager?.ActiveSoftwareItems, _softwareManager?.ActiveCategories);
        if (form.ShowDialog(this) != DialogResult.OK) return;
        _softwareManager?.AddCategory(item.Category); RebuildTreeView();
    }

    private string? ImportCategory()
    {
        var node = categoryTreeView.SelectedNode;
        string? category = node?.Tag is SoftwareItem item ? item.Category : node?.Parent != null ? node.Text : null;
        return category != null && _softwareManager?.ActiveCategories.Contains(category, StringComparer.OrdinalIgnoreCase) == true ? category : null;
    }

    private void UpdateImportAvailability()
    {
        addSoftwareButton.Enabled = !preparingImport && ImportCategory() != null;
        configurationTips.SetToolTip(addSoftwareButton, _softwareManager?.ActiveCategories.Count > 0
            ? ImportCategory() == null ? "请先在左侧选中分类，再添加或拖入软件" : $"添加软件到「{ImportCategory()}」，支持多选或拖入 EXE / MSI"
            : "请先新建分类，再在分类下添加软件");
    }

    private async Task ImportSoftwareAsync(IEnumerable<string> files)
    {
        if (preparingImport || _softwareManager == null) return;
        string? category = ImportCategory();
        if (category == null) return;
        preparingImport = true; addSoftwareButton.Enabled = false;
        try
        {
            int revision = configurationRevision;
            var paths = files.ToArray(); var existing = _softwareManager.ActiveSoftwareItems.Select(s => s.Clone()).ToList();
            var preparation = await Task.Run(() => SoftwareImport.Prepare(paths, existing, category));
            if (IsDisposed || Disposing) return;
            if (revision != configurationRevision) { MessageBox.Show(this, "配置已切换，请在当前配置中重新添加软件。"); return; }
            if (!_softwareManager.ActiveCategories.Contains(category, StringComparer.OrdinalIgnoreCase))
            { MessageBox.Show(this, "原分类已删除，请先选择分类再添加软件。"); return; }
            List<SoftwareItem> additions = new();
            if (preparation.Items.Count == 1 && paths.Length == 1)
            {
                using var editor = new SoftwareEditForm(preparation.Items[0], existing, _softwareManager.ActiveCategories) { Text = "新增软件" };
                if (editor.ShowDialog(this) == DialogResult.OK) additions.Add(editor.SoftwareItem);
            }
            else if (preparation.Items.Count > 0)
            {
                using var batch = new BatchImportForm(preparation, _softwareManager.ActiveCategories);
                if (batch.ShowDialog(this) == DialogResult.OK) additions = batch.Software;
            }
            if (additions.Any(item => !_softwareManager.ActiveCategories.Contains(item.Category, StringComparer.OrdinalIgnoreCase)))
            { MessageBox.Show(this, "请先在主界面新建分类，再添加软件。"); return; }
            foreach (var item in additions) _softwareManager.AddSoftwareToActiveList(item);
            if (additions.Count > 0) { softwareSearch.Clear(); softwareCategory.SelectedIndex = 0; RebuildTreeView(); }
            if (preparation.Errors.Count > 0 || preparation.Items.Count == 0)
                MessageBox.Show(this, $"添加 {additions.Count} 款，跳过 {preparation.Duplicates} 个重复文件。\n" + string.Join("\n", preparation.Errors), "添加软件", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex) { ErrorHandler.ShowError("添加软件失败：" + ex.Message); }
        finally { preparingImport = false; if (!IsDisposed) UpdateImportAvailability(); }
    }
}
