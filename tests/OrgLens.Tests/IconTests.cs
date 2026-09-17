using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using OrgLens.Desktop;
using OrgLens.Outlook;

namespace OrgLens.Tests
{
    internal static partial class Program
    {
        [ComImport]
        [Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB")]
        [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        private interface IPictureDimensions
        {
            [DispId(4)] int Width { get; }
            [DispId(5)] int Height { get; }
        }

        private static void CheckIconAssets()
        {
            Check("branded icon resources retain transparent edges and the blue organization mark", () =>
            {
                using (var image = OrgLensImages.CreateRibbonBitmap())
                {
                    Equal(32, image.Width);
                    Equal(32, image.Height);
                    Equal(0, (int)image.GetPixel(0, 0).A);
                    int bluePixels = 0;
                    for (int y = 0; y < image.Height; y++)
                        for (int x = 0; x < image.Width; x++)
                        {
                            var pixel = image.GetPixel(x, y);
                            if (pixel.A == 255 && pixel.B > 200 && pixel.R < 60) bluePixels++;
                        }
                    True(bluePixels > 40);
                }
                using (var icon = OrgLensImages.CreateWindowIcon())
                {
                    foreach (int size in new[] { 16, 32, 48 })
                        using (var sized = new Icon(icon, size, size))
                        using (var image = sized.ToBitmap())
                        {
                            Equal(size, image.Width);
                            Equal(size, image.Height);
                        }
                }
                using (var stream = typeof(OrgLensImages).Assembly.GetManifestResourceStream("OrgLens.Images.Application.ico"))
                using (var reader = new BinaryReader(stream))
                {
                    Equal((ushort)0, reader.ReadUInt16());
                    Equal((ushort)1, reader.ReadUInt16());
                    Equal((ushort)9, reader.ReadUInt16());
                    foreach (int size in new[] { 16, 20, 24, 32, 40, 48, 64, 128, 256 })
                    {
                        int width = reader.ReadByte();
                        int height = reader.ReadByte();
                        Equal(size, width == 0 ? 256 : width);
                        Equal(size, height == 0 ? 256 : height);
                        reader.ReadUInt16();
                        Equal((ushort)1, reader.ReadUInt16());
                        Equal((ushort)32, reader.ReadUInt16());
                        int length = reader.ReadInt32();
                        int offset = reader.ReadInt32();
                        long next = stream.Position;
                        stream.Position = offset;
                        using (var frame = new MemoryStream(reader.ReadBytes(length)))
                        using (var image = new Bitmap(frame))
                        {
                            Equal(size, image.Width);
                            Equal(size, image.Height);
                        }
                        stream.Position = next;
                    }
                }
                True(typeof(OrgLensImages).Assembly.GetManifestResourceNames()
                    .Contains("OrgLens.Images.License.txt"));
            });
            Check("Ribbon picture exposes IPictureDisp and is cached until disconnect", () =>
            {
                var connect = new Connect();
                Array custom = Array.CreateInstance(typeof(object), 0);
                try
                {
                    object image = connect.GetRibbonImage(null);
                    True(Marshal.IsComObject(image));
                    True(ReferenceEquals(image, connect.GetRibbonImage(null)));
                    var picture = (IPictureDimensions)image;
                    True(picture.Width > 0 && picture.Width == picture.Height);
                    connect.OnBeginShutdown(ref custom);
                    object reopened = connect.GetRibbonImage(null);
                    True(!ReferenceEquals(image, reopened));
                    True(((IPictureDimensions)reopened).Width > 0);
                }
                finally
                {
                    connect.OnDisconnection(DisconnectMode.HostShutdown, ref custom);
                    connect.OnBeginShutdown(ref custom);
                }
                var callback = typeof(IRibbonCallbacks).GetMethod("GetRibbonImage");
                Equal(2, ((DispIdAttribute)callback.GetCustomAttributes(typeof(DispIdAttribute), false).Single()).Value);
                Equal(UnmanagedType.IDispatch, ((MarshalAsAttribute)callback.ReturnParameter
                    .GetCustomAttributes(typeof(MarshalAsAttribute), false).Single()).Value);
            });
        }
    }
}
