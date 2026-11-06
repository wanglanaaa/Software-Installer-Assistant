using SoftwareInstaller.Utils;

namespace SoftwareInstaller.UI;

public partial class MainForm
{
    private readonly ImageList softwareIcons = new() { ColorDepth = ColorDepth.Depth32Bit, ImageSize = new(20, 20) };
    private readonly Dictionary<string, int> softwareIconCache = new(StringComparer.OrdinalIgnoreCase);

    private void InitializeSoftwareIcons()
    {
        _ = softwareIcons.Handle;
        using var empty = new Bitmap(20, 20);
        softwareIcons.Images.Add(empty);
        categoryTreeView.ImageList = softwareIcons;
        categoryTreeView.ImageIndex = categoryTreeView.SelectedImageIndex = 0;
        // The name is a subitem: draw its icon here so the number column stays intact.
        softwareListView.OwnerDraw = true;
        softwareListView.DrawColumnHeader += (_, e) => e.DrawDefault = true;
        softwareListView.DrawSubItem += (_, e) =>
        {
            if (e.ColumnIndex != 1 || e.Item == null || e.SubItem == null) { e.DrawDefault = true; return; }
            bool selected = e.Item.Selected && (softwareListView.Focused || !softwareListView.HideSelection);
            Color background = selected ? (softwareListView.Focused ? SystemColors.Highlight : SystemColors.Control) : e.SubItem.BackColor;
            Color foreground = selected && softwareListView.Focused ? SystemColors.HighlightText : e.SubItem.ForeColor;
            using var brush = new SolidBrush(background);
            e.Graphics.FillRectangle(brush, e.Bounds);
            var saved = e.Graphics.Save();
            e.Graphics.SetClip(e.Bounds);
            int size = Math.Min(softwareIcons.ImageSize.Width, e.Bounds.Height - 2);
            if (e.Item.ImageIndex >= 0)
                softwareIcons.Draw(e.Graphics, e.Bounds.Left + 3, e.Bounds.Top + (e.Bounds.Height - size) / 2, size, size, e.Item.ImageIndex);
            var text = new Rectangle(e.Bounds.Left + size + 7, e.Bounds.Top, Math.Max(0, e.Bounds.Width - size - 10), e.Bounds.Height);
            TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.SubItem.Font, text, foreground,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            e.Graphics.Restore(saved);
        };
        softwareListView.Disposed += (_, _) => softwareIcons.Dispose();
    }

    private int GetSoftwareIcon(string? path)
    {
        // Include file metadata so replacing or repairing an installer refreshes its icon.
        string key = "missing";
        try
        {
            var file = string.IsNullOrWhiteSpace(path) ? null : new FileInfo(path);
            if (file?.Exists == true) key = $"{file.FullName}|{file.Length}|{file.LastWriteTimeUtc.Ticks}";
        }
        catch (Exception ex) when (ex is ArgumentException or IOException or UnauthorizedAccessException or NotSupportedException)
        { }
        if (softwareIconCache.TryGetValue(key, out int index)) return index;
        using var image = SoftwareIcons.FromFile(path);
        softwareIcons.Images.Add(image);
        softwareIconCache[key] = index = softwareIcons.Images.Count - 1;
        return index;
    }
}
