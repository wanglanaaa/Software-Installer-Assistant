using SoftwareInstaller.Models;

namespace SoftwareInstaller.UI;

public partial class MainForm
{
    private readonly Button treeMoveUp = new() { Text = "上移" }, treeMoveDown = new() { Text = "下移" };
    private readonly Button softwareMoveUp = new() { Text = "上移" }, softwareMoveDown = new() { Text = "下移" };

    private void InitializeOrdering()
    {
        categoryTreeView.HideSelection = false;
        foreach (var button in new[] { treeMoveUp, treeMoveDown, softwareMoveUp, softwareMoveDown })
        {
            StyleButton(button); Controls.Add(button); button.Enabled = false;
        }
        treeMoveUp.Click += (_, _) => MoveTreeSelection(-1);
        treeMoveDown.Click += (_, _) => MoveTreeSelection(1);
        softwareMoveUp.Click += (_, _) => MoveListSelection(-1);
        softwareMoveDown.Click += (_, _) => MoveListSelection(1);
        categoryTreeView.AfterSelect += (_, _) => UpdateOrderingButtons();
        softwareListView.SelectedIndexChanged += (_, _) => UpdateOrderingButtons();
        configurationTips.SetToolTip(treeMoveUp, "上移选中的分类，或分类中的软件；不改变勾选状态");
        configurationTips.SetToolTip(treeMoveDown, "下移选中的分类，或分类中的软件；不改变勾选状态");
        configurationTips.SetToolTip(softwareMoveUp, "上移高亮的软件；筛选时只在可见项目之间移动，可多选");
        configurationTips.SetToolTip(softwareMoveDown, "下移高亮的软件；筛选时只在可见项目之间移动，可多选");
    }

    private void UpdateOrderingButtons()
    {
        var node = categoryTreeView.SelectedNode;
        treeMoveUp.Enabled = node?.Parent != null && node.PrevNode != null;
        treeMoveDown.Enabled = node?.Parent != null && node.NextNode != null;
        var rows = softwareListView.Items.Cast<ListViewItem>().ToList();
        softwareMoveUp.Enabled = rows.Where((row, i) => i > 0 && row.Selected && !rows[i - 1].Selected).Any();
        softwareMoveDown.Enabled = rows.Where((row, i) => i + 1 < rows.Count && row.Selected && !rows[i + 1].Selected).Any();
    }

    // Swap visible slots so filtered-out entries retain their positions and selection.
    private static bool MoveVisible<T>(List<T> all, List<T> visible, HashSet<T> selected, int direction)
    {
        bool moved = false;
        var indices = direction < 0 ? Enumerable.Range(1, Math.Max(0, visible.Count - 1))
            : Enumerable.Range(0, Math.Max(0, visible.Count - 1)).Reverse();
        foreach (int i in indices)
        {
            int adjacent = i + direction;
            if (!selected.Contains(visible[i]) || selected.Contains(visible[adjacent])) continue;
            int currentSlot = all.IndexOf(visible[i]), adjacentSlot = all.IndexOf(visible[adjacent]);
            (all[currentSlot], all[adjacentSlot]) = (all[adjacentSlot], all[currentSlot]);
            (visible[i], visible[adjacent]) = (visible[adjacent], visible[i]);
            moved = true;
        }
        return moved;
    }

    private void MoveTreeSelection(int direction)
    {
        var node = categoryTreeView.SelectedNode;
        if (_softwareManager == null || node?.Parent == null) return;
        if (node.Tag is SoftwareItem item)
        {
            var siblings = node.Parent.Nodes.Cast<TreeNode>().Select(n => (SoftwareItem)n.Tag!).ToList();
            if (!MoveVisible(_softwareManager.ActiveSoftwareItems, siblings, [item], direction)) return;
        }
        else
        {
            var categories = _softwareManager.ActiveCategories;
            if (!MoveVisible(categories, new(categories), [node.Text], direction)) return;
        }
        RebuildTreeView();
    }

    private void MoveListSelection(int direction)
    {
        if (_softwareManager == null) return;
        var visible = softwareListView.Items.Cast<ListViewItem>().Select(row => (SoftwareItem)row.Tag!).ToList();
        var selected = softwareListView.SelectedItems.Cast<ListViewItem>().Select(row => (SoftwareItem)row.Tag!).ToHashSet();
        if (MoveVisible(_softwareManager.ActiveSoftwareItems, visible, selected, direction)) RebuildTreeView();
    }
}
