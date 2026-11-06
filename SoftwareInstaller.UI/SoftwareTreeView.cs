using SoftwareInstaller.Models;

namespace SoftwareInstaller.UI;

internal sealed class SoftwareTreeView : TreeView
{
    public SoftwareTreeView()
    {
        DrawMode = TreeViewDrawMode.OwnerDrawText;
    }

    protected override void OnDrawNode(DrawTreeNodeEventArgs e)
    {
        if (e.Node == null || e.Node.Tag is SoftwareItem || ImageList == null)
        {
            e.DrawDefault = true;
            base.OnDrawNode(e);
            return;
        }
        // Draw category labels immediately after their checkbox, without an image slot.
        var bounds = e.Bounds;
        bounds.X -= ImageList.ImageSize.Width;
        bool selected = e.Node.IsSelected && (Focused || !HideSelection);
        Color background = selected ? Focused ? SystemColors.Highlight : SystemColors.Control : BackColor;
        Color foreground = selected && Focused ? SystemColors.HighlightText : ForeColor;
        using var clear = new SolidBrush(BackColor);
        using var fill = new SolidBrush(background);
        e.Graphics.FillRectangle(clear, e.Bounds);
        e.Graphics.FillRectangle(fill, bounds);
        TextRenderer.DrawText(e.Graphics, e.Node.Text, e.Node.NodeFont ?? Font, bounds, foreground,
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding);
        if (selected && Focused) ControlPaint.DrawFocusRectangle(e.Graphics, bounds, foreground, background);
        base.OnDrawNode(e);
    }
}
