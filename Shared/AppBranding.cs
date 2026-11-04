using System.Drawing;
using System.Drawing.Drawing2D;
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
    private static readonly Lazy<Bitmap> image = new(() =>
    {
        using var stream = typeof(AppBranding).Assembly.GetManifestResourceStream("SoftwareInstaller.AppLogo.png")!;
        using var source = System.Drawing.Image.FromStream(stream);
        return new Bitmap(source);
    });

    public static Image Image => image.Value;

    public static void Apply(Form form)
    {
        form.Icon = icon.Value;
        form.ShowIcon = true;
    }
}

internal sealed class BrandLogo : PictureBox
{
    public BrandLogo()
    {
        Image = AppBranding.Image;
        SizeMode = PictureBoxSizeMode.Zoom;
        TabStop = false;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        e.Graphics.CompositingQuality = CompositingQuality.HighQuality;
        base.OnPaint(e);
    }
}
