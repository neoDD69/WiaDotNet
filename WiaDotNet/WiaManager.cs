using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using WIA;

namespace WiaDotNet
{
    public class WiaManager : IDisposable
    {
        private readonly string path;
        private readonly DeviceManager _manager;
        public event EventHandler<WiaEventArgs> DeviceConneted;
        public event EventHandler<WiaEventArgs> DeviceDisonneted;

        public WiaManager()
        {
            _manager = new DeviceManagerClass();
            RegisterEvent();
            path = Path.GetTempPath() + @"WiaDotNet\";
        }

        public void Dispose()
        {
            UnregisterEvent();
        }

        private void RegisterEvent(string deviceId = "*")
        {
            if (string.IsNullOrEmpty(deviceId))
                deviceId = "*";
            _manager.RegisterEvent(EventID.wiaEventDeviceConnected, deviceId);
            _manager.RegisterEvent(EventID.wiaEventDeviceDisconnected, deviceId);
            //_manager.RegisterEvent(EventID.wiaEventItemCreated, deviceId);
            //_manager.RegisterEvent(EventID.wiaEventItemDeleted, deviceId);
            _manager.OnEvent += Manager_OnEvent;
        }

        private void UnregisterEvent(string deviceId = "*")
        {
            _manager.UnregisterEvent(EventID.wiaEventDeviceConnected, deviceId);
            _manager.UnregisterEvent(EventID.wiaEventDeviceDisconnected, deviceId);
        }

        private void Manager_OnEvent(string EventID, string DeviceID, string ItemID)
        {
            Console.WriteLine($"EventID={EventID}, DeviceID={DeviceID}, ItemID={ItemID}.");
            if (EventID == WIA.EventID.wiaEventDeviceConnected)
            {
                WiaEventArgs args = new WiaEventArgs {DeviceID = DeviceID, ItemID = ItemID};
                DeviceConneted?.Invoke(this, args);
            }
            else if (EventID == WIA.EventID.wiaEventDeviceDisconnected)
            {
                WiaEventArgs args = new WiaEventArgs { DeviceID = DeviceID, ItemID = ItemID };
                DeviceDisonneted?.Invoke(this, args);
            }
        }

        public List<WiaScanner> GetDevices()
        {
            List<WiaScanner> devices = new List<WiaScanner>();
            foreach (DeviceInfo info in _manager.DeviceInfos)
            {
                devices.Add(new WiaScanner
                {
                    Id = info.DeviceID,
                    Name = info.Properties["Name"].get_Value().ToString()
                });
            }
            return devices;
        }

        public WiaScanner ShowSelectDevice()
        {
            ICommonDialog dialog = new CommonDialog();
            Device device = dialog.ShowSelectDevice(WiaDeviceType.UnspecifiedDeviceType, true, false);
            if (device != null)
                return new WiaScanner{Id = device.DeviceID, Name = device.Properties["Name"].get_Value().ToString() };
            else
                return null;
        }

        public void ClearTempPath()
        {
            if (Directory.Exists(path))
            {
                try
                {
                    Directory.Delete(path, true);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }

      /// <summary>
      /// Scan images.
      /// </summary>
      /// <param name="images">Result image file path.</param>
      /// <param name="scannerId">Guid of selected scanner.</param>
      /// <param name="settings"></param>
      public WiaResult Scan(out List<string> images, string scannerId, ScanSettings settings, bool color=false)
      {
         // give a default setting if null
         if (settings == null)
            settings = new ScanSettings();
         images = new List<string>();
         var liImages = new List<string>();
         settings.ImageFormat = WiaImageFormat.GIF;
         // parse file extension
         string extension;
         switch (settings.ImageFormat)
         {
            default:
            case WiaImageFormat.PNG:
               extension = ".png";
               break;
            case WiaImageFormat.BMP:
               extension = ".bmp";
               break;
            case WiaImageFormat.GIF:
               extension = ".gif";
               break;
            case WiaImageFormat.JPEG:
               extension = ".jpg";
               break;
            case WiaImageFormat.TIFF:
               extension = ".tif";
               break;
         }
         // prepare result
         WiaResult result = new WiaResult();
         result.Error = WiaError.SUCCESS;

         bool hasMorePages = true;
         while (hasMorePages)
         {
            //david 17/01/2023 12:52:04	bisogna lasciarlo qui se no dopo la prima pagina va in HANG il COM
            //DEVICE
            // select the correct scanner using the provided scannerId parameter
            Device device = null;

            foreach (DeviceInfo info in _manager.DeviceInfos)
            {
               if (info.DeviceID == scannerId)
               {
                  // connect to scanner
                  device = info.Connect();
                  Property nameProp = info.Properties["Name"];
                  Property typeProp = info.Properties["Manufacturer"];
                  Property manufacturerProp = info.Properties["Manufacturer"];
                  Console.WriteLine(nameProp.get_Value().ToString());
                  Console.WriteLine(typeProp.get_Value().ToString());
                  Console.WriteLine(manufacturerProp.get_Value().ToString());
                  //SetDocumentHandling(device, settings.DocumentHandling);
                  //SetDeviceProperty(device, DEVICE_PROPERTY_PAGES_ID, 1); https://stackoverflow.com/questions/7077020/wia-scanning-via-feeder
                  //var   DEVICE_PROPERTY_PAGES_ID = "3096";

                  //SetDevice(device, WIA_PROPERTIES.WIA_DPS_PAGES_ID, 0);
                  		//david 17/01/2023 15:45:04	MUST
                  object val = ((int)settings.DocumentHandling).ToString();

                  SetDevice(device, "Document Handling Select", val);
                  //Property docHand = device.Properties["Document Handling Select"];
                  //docHand.set_Value(ref val);
                  break;
               }
            }
            // device was not found
            if (device == null)
            {
               // enumerate available devices
               string availableDevices = "";
               foreach (DeviceInfo info in _manager.DeviceInfos)
               {
                  availableDevices += info.DeviceID + "\n";
               }
               // show error with available devices
               string message = "The device with provided ID could not be found. Available Devices:\n" + availableDevices;
               WiaResult error = new WiaResult { Error = WiaError.DEVICE_NOT_FOUND, Message = message };
               return error;
            }
            // check paper
            hasMorePages = HasMorePages(device);
            if (!hasMorePages)
               break;
            // scan
            Item item = device.Items[1];
            try
            {
               SetScanSettings(item, settings);
               // COLOR / GREY
               var wiaItent = (WIA_PROPERTIES.WIA_INTENT_MINIMIZE_SIZE + (color ? WIA_PROPERTIES.WIA_INTENT_IMAGE_TYPE_COLOR : WIA_PROPERTIES.WIA_INTENT_IMAGE_TYPE_GRAYSCALE)).ToString();
               SetItem(item, WIA_PROPERTIES.WIA_IPS_CUR_INTENT_str,wiaItent);
               ImageFile image = null;
               var fScanStreamHandling = new Action(() =>
               {
                  // create temp directory

                  if (!Directory.Exists(path))
                     Directory.CreateDirectory(path);
                  // save to temp file
                  string fileName = path + DateTime.Now.ToString("yyyyMMdd-HHmmss-ffffff", System.Globalization.CultureInfo.InvariantCulture) + extension;
                  if (File.Exists(fileName))
                     File.Delete(fileName);
                  image.SaveFile(fileName);
                  
                  // add file to output list
                  liImages.Add(fileName);
               });
               if (settings.ShowUI)
               {
                  // scan image with dialog
                  ICommonDialog wiaCommonDialog = new CommonDialog();
                  image = (ImageFile)wiaCommonDialog.ShowTransfer(item, settings.ImageFormat, false);
                  fScanStreamHandling();
               }
               else
               {
                  image = (ImageFile)item.Transfer(settings.ImageFormat);
                  fScanStreamHandling();
                  if ((settings.DocumentHandling & WiaDocumentHandling.DUPLEX) == WiaDocumentHandling.DUPLEX)
                  {
                     
                     image = (ImageFile)item.Transfer(settings.ImageFormat);
                     fScanStreamHandling();
                  }
               }
            }
            catch (COMException ex)
            {
               if (liImages.Count == 0)
               {//first error ... really no pages else 
                  result = ShowWiaErrorCode(ex);
                  // if there are successfull scans, we juse break without showing error
                  if (result.Error == WiaError.WIA_ERROR_PAPER_EMPTY)
                     result.Error = WiaError.SUCCESS;
               }
               break;
            }
            catch (Exception ex)
            {
               result.Error = WiaError.UNKNOWN;
               result.Message = ex.Message;
               break;
            }
            finally
            {
               hasMorePages = HasMorePages(device);
            }
         }
         images.AddRange(liImages);

         return result;
      }

      private static void SetScanSettings(Item item, ScanSettings settings)
      {
         if (settings == null)
            return;
         SetBrightness(item, settings.Brightness);
         SetContrast(item, settings.Contrast);
         if (settings.DPI.HasValue)
            SetResolution(item, settings.DPI.Value);
         if (settings.Orientation.HasValue)
            SetOrientation(item, settings.Orientation.Value);
         if (settings.ImageSize != null)
         {
            if (settings.ImageSize.ImageUnit == WiaImageSize.Unit.Pixel)
            {
               SetXExtent(item, settings.ImageSize.ImageWidth);
               SetYExtent(item, settings.ImageSize.ImageHeight);
            }
            else  // inch/1000
            {
               SetWidth(item, settings.ImageSize.ImageWidth);
               SetHeight(item, settings.ImageSize.ImageHeight);
            }
         }
      }

      private static void SetDocumentHandling(Device device, WiaDocumentHandling value)
      {		//david 11/01/2023 16:40:42	
         var val = (int)value;
         SetDevice(device, "Document Handling Select", val);
      }
      //private static void SetDocumentHandling(Item item, WiaDocumentHandling value)
      //{		//david 11/01/2023 16:40:42	
      //   var val = (int)value;
      //   SetItem(item, WIA_PROPERTIES.WIA_DPS_DOCUMENT_HANDLING_SELECT, 5);
      //}
      private static void SetResolution(Item item, int dpi)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_XRES, dpi);
            SetItem(item, WIA_PROPERTIES.WIA_IPS_YRES, dpi);
        }

        private static void SetBrightness(Item item, int value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_BRIGHTNESS, value);
        }

        private static void SetContrast(Item item, int value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_CONTRAST, value);
        }

        /// <summary>
        /// Set the current width, in pixels, of a selected image to acquire.
        /// </summary>
        private static void SetXExtent(Item item, int value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_XEXTENT, value);
        }

        /// <summary>
        /// Set the current width, in pixels, of a selected image to acquire.
        /// </summary>
        private static void SetYExtent(Item item, int value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_YEXTENT, value);
        }

        private static void SetWidth(Item item, int value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_PAGE_WIDTH, value);
        }

        private static void SetHeight(Item item, int value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_PAGE_HEIGHT, value);
        }

        private static void SetOrientation(Item item, WiaOrientation value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_ORIENTATION, (int) value);
        }

        /// <summary>
        /// Property name should use string to get.
        /// </summary>
        private static void SetItem(IItem item, object property, object value)
        {
            Property aProperty = item.Properties.get_Item(ref property);
            aProperty?.set_Value(ref value);
        }
      private static void SetDevice(Device device, string property, object value)
      {
         Property aProperty = device.Properties[property];
         aProperty?.set_Value(ref value);

         //Property docHand = device.Properties["Document Handling Select"];
         //object val = ((int)value).ToString();
         //docHand.set_Value(ref val);

      }
      private static bool HasMorePages(Device device)
        {
            // assume there are no more pages
            bool hasMorePages = false;
            //determine if there are any more pages waiting
            Property documentHandlingSelect = null;
            Property documentHandlingStatus = null;
            try
            {
                foreach (Property prop in device.Properties)
                {
                    if (prop.PropertyID == WIA_PROPERTIES.WIA_DPS_DOCUMENT_HANDLING_SELECT)
                        documentHandlingSelect = prop;
                    if (prop.PropertyID == WIA_PROPERTIES.WIA_DPS_DOCUMENT_HANDLING_STATUS)
                        documentHandlingStatus = prop;
                }
                // may not exist on flatbed scanner but required for feeder
                if (documentHandlingSelect != null)
                {
                    // check for document feeder
                    if ((Convert.ToUInt32(documentHandlingSelect.get_Value()) &
                         WIA_DPS_DOCUMENT_HANDLING_SELECT.FEEDER) != 0)
                    {
                        hasMorePages = ((Convert.ToUInt32(documentHandlingStatus.get_Value()) &
                                         WIA_DPS_DOCUMENT_HANDLING_STATUS.FEED_READY) != 0);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return hasMorePages;
        }

        private static WiaResult ShowWiaErrorCode(COMException ex)
        {
            WiaResult error = new WiaResult();
            if (Enum.IsDefined(typeof(WiaError), (uint)ex.ErrorCode))
            {
                error.Error = ((WiaError)ex.ErrorCode);
            }
            else
            {
                error.Error = WiaError.UNKNOWN;
            }
            error.Message = ex.Message;
            return error;
        }
    }
}
