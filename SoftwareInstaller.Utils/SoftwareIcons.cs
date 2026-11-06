using System.Drawing.Imaging;

namespace SoftwareInstaller.Utils;

public static class SoftwareIcons
{
    public static Bitmap FromFile(string? path)
    {
        try
        {
            if (File.Exists(path))
            {
                using var icon = Icon.ExtractAssociatedIcon(Path.GetFullPath(path));
                if (icon != null) return icon.ToBitmap();
            }
        }
        catch (Exception ex) when (ex is ArgumentException or IOException or UnauthorizedAccessException or System.ComponentModel.Win32Exception)
        { }
        return SystemIcons.Application.ToBitmap();
    }

    public static byte[] ReadPng(string? path)
    {
        using var icon = FromFile(path);
        using var image = new Bitmap(icon, new Size(32, 32));
        using var stream = new MemoryStream();
        image.Save(stream, ImageFormat.Png);
        return stream.ToArray();
    }

    public static Bitmap FromPng(byte[]? data)
    {
        try
        {
            if (data is { Length: > 0 and <= 65536 })
            {
                using var stream = new MemoryStream(data);
                using var image = Image.FromStream(stream);
                if (image.Width <= 256 && image.Height <= 256) return new Bitmap(image);
            }
        }
        catch (Exception ex) when (ex is ArgumentException or OutOfMemoryException)
        { }
        return SystemIcons.Application.ToBitmap();
    }
}
