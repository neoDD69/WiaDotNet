using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WiaDotNet
{
    public enum WiaError : uint
    {
        SUCCESS = 0,
        UNKNOWN = 1,
        DEVICE_NOT_FOUND = 2,
        WIA_ERROR_BUSY = 0x80210006,
        WIA_ERROR_COVER_OPEN = 0x80210016,
        WIA_ERROR_DEVICE_COMMUNICATION = 0x8021000A,
        WIA_ERROR_DEVICE_LOCKED = 0x8021000D,
        WIA_ERROR_EXCEPTION_IN_DRIVER = 0x8021000E,
        WIA_ERROR_GENERAL_ERROR = 0x80210001,
        WIA_ERROR_INCORRECT_HARDWARE_SETTING = 0x8021000C,
        WIA_ERROR_INVALID_COMMAND = 0x8021000B,
        WIA_ERROR_INVALID_DRIVER_RESPONSE = 0x8021000F,
        WIA_ERROR_ITEM_DELETED = 0x80210009,
        WIA_ERROR_LAMP_OFF = 0x80210017,
        WIA_ERROR_MAXIMUM_PRINTER_ENDORSER_COUNTER = 0x80210021,
        WIA_ERROR_MULTI_FEED = 0x80210020,
        WIA_ERROR_OFFLINE = 0x80210005,
        WIA_ERROR_PAPER_EMPTY = 0x80210003,
        WIA_ERROR_PAPER_JAM = 0x80210002,
        WIA_ERROR_PAPER_PROBLEM = 0x80210004,
        WIA_ERROR_WARMING_UP = 0x80210007,
        WIA_ERROR_USER_INTERVENTION = 0x80210008,
        WIA_S_NO_DEVICE_AVAILABLE = 0x80210015
    }

    #region WIA Interop Defines
    internal class WIA_DPS_DOCUMENT_HANDLING_SELECT
    {
        public const uint FEEDER = 0x00000001;
        public const uint FLATBED = 0x00000002;
    }
    internal class WIA_DPS_DOCUMENT_HANDLING_STATUS
    {
        public const uint FEED_READY = 0x00000001;
    }
    internal class WIA_PROPERTIES
    {
        public const uint WIA_RESERVED_FOR_NEW_PROPS = 1024;
        public const uint WIA_DIP_FIRST = 2;
        public const uint WIA_DPA_FIRST = WIA_DIP_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
        public const uint WIA_DPC_FIRST = WIA_DPA_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
        //
        // Scanner only device properties (DPS)
        //
        public const uint WIA_DPS_FIRST = WIA_DPC_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
        public const uint WIA_DPS_DOCUMENT_HANDLING_STATUS = WIA_DPS_FIRST + 13;
        public const uint WIA_DPS_DOCUMENT_HANDLING_SELECT = WIA_DPS_FIRST + 14;
        // Scanner WIA Item Property Constants
        // see: https://github.com/tpn/winddk-8.1/blob/master/Include/um/WiaDef.h
        public const string WIA_IPS_OPTICAL_XRES = "3090";  // Horizontal Optical Resolution
        public const string WIA_IPS_OPTICAL_YRES = "3091";  // Vertical Optical Resolution
        public const string WIA_IPS_PAGE_WIDTH = "3098";  // Width of the current page selected, in thousandths of an inch (.001)
        public const string WIA_IPS_PAGE_HEIGHT = "3099";
        public const string WIA_IPS_XRES = "6147";  // Horizontal Resolution
        public const string WIA_IPS_YRES = "6148";  // Vertical Resolution
        public const string WIA_IPS_XEXTENT = "6151";  // Width, in pixels, of a selected image to acquire.
        public const string WIA_IPS_YEXTENT = "6152";
        public const string WIA_IPS_BRIGHTNESS = "6154";
        public const string WIA_IPS_CONTRAST = "6155";
        public const string WIA_IPS_ORIENTATION = "6156";
        public const string WIA_IPS_ROTATION = "6157";
        public const string WIA_IPS_DESKEW_X = "6162";
        public const string WIA_IPS_DESKEW_Y = "6163";
    }
    #endregion
}
