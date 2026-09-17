using System;
using System.Drawing;
using System.IO;

namespace OrgLens.Desktop
{
    public static class OrgLensImages
    {
        public static Bitmap CreateRibbonBitmap()
        {
            using (var stream = Open("Ribbon.png"))
            using (var image = new Bitmap(stream))
                return new Bitmap(image);
        }

        public static Icon CreateWindowIcon()
        {
            using (var stream = Open("Application.ico"))
            using (var icon = new Icon(stream))
                return (Icon)icon.Clone();
        }

        private static Stream Open(string name)
        {
            return typeof(OrgLensImages).Assembly.GetManifestResourceStream("OrgLens.Images." + name)
                ?? throw new InvalidOperationException("The embedded OrgLens icon is missing: " + name);
        }
    }
}
