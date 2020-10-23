using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WiaDotNet
{
    public class ScanSettings
    {
        /// <summary>
        /// Show scanning progress UI.
        /// </summary>
        public bool ShowUI = false;
        public int Contrast = 0;
        public int Brightness = 0;
        public WiaOrientation? Orientation;
        //public WiaOrientation? Rotation;
        public int? DPI;
        public WiaImageSize ImageSize = null;
        /// <summary>
        /// Guid of supported format.
        /// </summary>
        public string ImageFormat = WiaImageFormat.PNG;
    }

    public enum WiaOrientation
    {
        PORTRAIT,
        LANSCAPE,
        ROT180,
        ROT270
    }

    public class WiaImageSize
    {
        public enum Unit
        {
            Pixel,
            TInch
        }

        public Unit ImageUnit;
        public int ImageWidth, ImageHeight;
    }
}
