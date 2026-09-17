using System.Windows.Forms;
using OrgLens.Desktop;

namespace OrgLens.Outlook
{
    internal sealed class RibbonPicture : AxHost
    {
        private RibbonPicture() : base(string.Empty) { }

        internal static object Create()
        {
            using (var image = OrgLensImages.CreateRibbonBitmap())
                return GetIPictureDispFromPicture(image);
        }
    }
}
