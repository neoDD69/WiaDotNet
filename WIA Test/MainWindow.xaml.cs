using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media.Imaging;
using WiaDotNet;

namespace WIA_Test
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<WiaScanner> scanners;
        private readonly WiaManager _manager;
        private WiaScanner SelectedScanner => (WiaScanner) scannersCombo.SelectedItem;

        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;
            _manager = new WiaManager();
            _manager.DeviceConneted += Manager_DeviceConnetedEvent;
            _manager.DeviceDisonneted += Manager_DeviceDisonnetedEvent;
        }

        private void Manager_DeviceDisonnetedEvent(object sender, WiaEventArgs e)
        {
            CheckConnectedScanners();
        }

        private void Manager_DeviceConnetedEvent(object sender, WiaEventArgs e)
        {
            CheckConnectedScanners();
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _manager.ClearTempPath();
            _manager.Dispose();
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            CheckConnectedScanners();
        }

        private void CheckConnectedScanners()
        {
            scanners = _manager.GetDevices();
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
            ScanSettings settings = new ScanSettings
            {
                ShowUI = showDialog,
                DPI = 200,
                ImageFormat = WiaImageFormat.PNG
            };

            WiaResult result = _manager.Scan(out List<string> images, scannerId, settings);
            if (result.Error == WiaError.SUCCESS)
            {
                if (images.Count > 0)
                    previewImage.Source = new BitmapImage(new Uri(images[0]));
                else
                    MessageBox.Show("No images!");
            }
            else
                MessageBox.Show(result.Error.ToString());
        }
    }
}
