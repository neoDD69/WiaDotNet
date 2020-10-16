using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media.Imaging;
using WIA;

namespace WIA_Test
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<ScannerModel> scanners;
        private readonly DeviceManager _manager;
        private ScannerModel SelectedScanner => (ScannerModel) scannersCombo.SelectedItem;

        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;
            _manager = new DeviceManagerClass();
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            UnregisterEvent();
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            CheckConnectedScanners();
            RegisterEvent();
        }

        private void CheckConnectedScanners()
        {
            scanners = GetDevices();
            scannersCombo.ItemsSource = scanners;
            if (scanners.Count > 0)
            {
                scannersCombo.SelectedIndex = 0;
                Panel1.IsEnabled = true;
            }
            else
            {
                Panel1.IsEnabled = false;
            }
        }

        private void BtnScan_OnClick(object sender, RoutedEventArgs e)
        {
            if (SelectedScanner == null)
                return;

            bool showDialog = chkShow.IsChecked == true;
            string scannerId = SelectedScanner.Id;
            try
            {
                List<string> images = Scan(scannerId, showDialog);
                if (images.Count > 0)
                {
                    previewImage.Source = new BitmapImage(new Uri(images[0]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        /// <summary>
        /// Use scanner to scan an image (with user selecting the scanner from a dialog).
        /// </summary>
        /// <returns>Scanned images.</returns>
        public List<string> Scan(bool showDialog)
        {
            ICommonDialog dialog = new CommonDialog();
            Device device = dialog.ShowSelectDevice(WIA.WiaDeviceType.UnspecifiedDeviceType, true, false);
            if (device != null)
            {
                return Scan(device.DeviceID, showDialog);
            }
            else
            {
                throw new Exception("You must select a device for scanning.");
            }
        }

        /// <summary>
        /// Use scanner to scan an image (scanner is selected by its unique id).
        /// </summary>
        /// <returns>Scanned images.</returns>
        public List<string> Scan(string scannerId, bool showDialog)
        {
            List<string> images = new List<string>();
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
                    throw new Exception("The device with provided ID could not be found. Available Devices:\n" + availableDevices);
                }
                // check paper
                hasMorePages = HasMorePages(device);
                if (!hasMorePages)
                    break;
                // scan
                Item item = device.Items[1];
                try
                {
                    SetItem(item, WIA_PROPERTIES.WIA_IPS_XRES, 300);
                    SetItem(item, WIA_PROPERTIES.WIA_IPS_YRES, 300);
                    ImageFile image;
                    if (showDialog)
                    {
                        // scan image with dialog
                        ICommonDialog wiaCommonDialog = new CommonDialog();
                        image = (ImageFile)wiaCommonDialog.ShowTransfer(item, FormatID.wiaFormatPNG, false);
                    }
                    else
                    {
                        image = (ImageFile)item.Transfer(FormatID.wiaFormatPNG);
                    }

                    // save to temp file
                    string fileName = Path.GetTempFileName();
                    File.Delete(fileName);
                    image.SaveFile(fileName);
                    image = null;
                    // add file to output list
                    images.Add(fileName);
                }
                catch (COMException ex)
                {
                    if (ShowWiaErrorCode(ex.ErrorCode, out WiaError error))
                    {
                        // if there are successfull scans, we juse break without showing error
                        if (error == WiaError.WIA_ERROR_PAPER_EMPTY && images.Count > 0)
                            break;
                        MessageBox.Show(error.ToString());
                        break;
                    }
                    else
                        throw ex;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    hasMorePages = HasMorePages(device);
                }
            }
            return images;
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

        /// <summary>
        /// Property name should use string to get.
        /// </summary>
        private static void SetItem(IItem item, object property, object value)
        {
            Property aProperty = item.Properties.get_Item(ref property);
            aProperty.set_Value(ref value);
        }

        /// <summary>
        /// Gets the list of available WIA devices.
        /// </summary>
        /// <returns></returns>
        private List<ScannerModel> GetDevices()
        {
            List<ScannerModel> devices = new List<ScannerModel>();
            foreach (DeviceInfo info in _manager.DeviceInfos)
            {
                devices.Add(new ScannerModel
                {
                    Id = info.DeviceID,
                    Name = info.Properties["Name"].get_Value().ToString()
                });
            }
            return devices;
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
                Console.WriteLine("Connected");
                CheckConnectedScanners();
            }
            else if (EventID == WIA.EventID.wiaEventDeviceDisconnected)
            {
                Console.WriteLine("Disconnected");
                CheckConnectedScanners();
            }
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
            public const string WIA_IPS_XRES = "6147";  // Horizontal Resolution
            public const string WIA_IPS_YRES = "6148";  // Vertical Resolution
            public const string WIA_IPS_BRIGHTNESS = "6154";
            public const string WIA_IPS_CONTRAST = "6155";
        }

        internal enum WiaError : uint
        {
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
        #endregion

        private static bool ShowWiaErrorCode(int errorCode, out WiaError error)
        {
            if (Enum.IsDefined(typeof(WiaError), (uint)errorCode))
            {
                error = ((WiaError) errorCode);
                return true;
            }
            else
            {
                error = WiaError.WIA_ERROR_GENERAL_ERROR;
                return false;
            }
        }
    }

    public class ScannerModel
    {
        public string Name { get; set; }
        public string Id { get; set; }
    }
}
