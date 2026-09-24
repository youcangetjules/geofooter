using System.Drawing;
using System.Windows.Forms;

namespace Aliniant.AesRibbonHost
{
    /// <summary>Convert System.Drawing.Image to an OLE IPictureDisp for Ribbon getImage.</summary>
    internal sealed class PictureConverter : AxHost
    {
        private PictureConverter() : base(string.Empty) { }

        public static object ImageToPictureDisp(Image image)
        {
            // Returns the OLE picture object; Office receives it as IPictureDisp.
            return GetIPictureDispFromPicture(image);
        }
    }
}
