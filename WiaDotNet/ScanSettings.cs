using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WIA;

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
      //david 11/01/2023 16:39:11	
      public bool Duplex = false;
      public WiaDocumentHandling DocumentHandling => (((int)WiaDocumentHandling.FEEDER) + (Duplex? WiaDocumentHandling.DUPLEX:0));
   }
   //SetItemIntProperty(ref item, 6146, 2); // greyscale
   //SetItemIntProperty(ref item, 4104, 8); // bit depth
   public enum WiaOrientation
   {
      PORTRAIT,
      LANSCAPE,
      ROT180,
      ROT270
   }
   public enum WiaDocumentHandling : uint
   {//https://github.com/tpn/winddk-8.1/blob/master/Include/um/WiaDef.h
      FEEDER = 0x001,
      FLATBED = 0x002,
      DUPLEX = 0x004,
      FRONT_FIRST = 0x008,
      BACK_FIRST = 0x010,
      FRONT_ONLY = 0x020,
      BACK_ONLY = 0x040,
      NEXT_PAGE = 0x080,
      PREFEED = 0x100,
      AUTO_ADVANCE = 0x200
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
