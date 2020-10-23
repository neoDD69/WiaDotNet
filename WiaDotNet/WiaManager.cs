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
        public WiaResult Scan(out List<string> images, string scannerId, ScanSettings settings)
        {
            images = new List<string>();
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
                    WiaResult error = new WiaResult {Error = WiaError.DEVICE_NOT_FOUND, Message = message};
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
                    ImageFile image;
                    if (settings.ShowUI)
                    {
                        // scan image with dialog
                        ICommonDialog wiaCommonDialog = new CommonDialog();
                        image = (ImageFile)wiaCommonDialog.ShowTransfer(item, settings.ImageFormat, false);
                    }
                    else
                    {
                        image = (ImageFile)item.Transfer(settings.ImageFormat);
                    }

                    // create temp directory
                    
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    // save to temp file
                    string fileName = path + DateTime.Now.ToString("yyyyMMdd-HHmmss", System.Globalization.CultureInfo.InvariantCulture) + extension;
                    if (File.Exists(fileName))
                        File.Delete(fileName);
                    image.SaveFile(fileName);
                    // add file to output list
                    images.Add(fileName);
                }
                catch (COMException ex)
                {
                    result = ShowWiaErrorCode(ex);
                    // if there are successfull scans, we juse break without showing error
                    if (result.Error == WiaError.WIA_ERROR_PAPER_EMPTY && images.Count > 0)
                        result.Error = WiaError.SUCCESS;
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
            return result;
        }

        private void SetScanSettings(Item item, ScanSettings settings)
        {
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

        private void SetResolution(Item item, int dpi)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_XRES, dpi);
            SetItem(item, WIA_PROPERTIES.WIA_IPS_YRES, dpi);
        }

        private void SetBrightness(Item item, int value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_BRIGHTNESS, value);
        }

        private void SetContrast(Item item, int value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_CONTRAST, value);
        }

        /// <summary>
        /// Set the current width, in pixels, of a selected image to acquire.
        /// </summary>
        private void SetXExtent(Item item, int value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_XEXTENT, value);
        }

        /// <summary>
        /// Set the current width, in pixels, of a selected image to acquire.
        /// </summary>
        private void SetYExtent(Item item, int value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_YEXTENT, value);
        }

        private void SetWidth(Item item, int value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_PAGE_WIDTH, value);
        }

        private void SetHeight(Item item, int value)
        {
            SetItem(item, WIA_PROPERTIES.WIA_IPS_PAGE_HEIGHT, value);
        }

        private void SetOrientation(Item item, WiaOrientation value)
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
