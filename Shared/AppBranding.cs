using System.Drawing;
using System.Windows.Forms;

namespace SoftwareInstaller;

internal static class AppBranding
{
    private static readonly Lazy<Icon> icon = new(() =>
    {
        using var stream = typeof(AppBranding).Assembly.GetManifestResourceStream("SoftwareInstaller.AppIcon.ico")!;
        using var source = new Icon(stream, new Size(32, 32));
        return (Icon)source.Clone();
    });
    private static readonly Lazy<Bitmap> image = new(() => icon.Value.ToBitmap());

    public static Image Image => image.Value;

    public static void Apply(Form form)
    {
        form.Icon = icon.Value;
        form.ShowIcon = true;
    }
}
