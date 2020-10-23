# WiaDotNet
A wrapper for WIA scanners

## Usage
Working sample code in `WIA Test` WPF project.

### Initialize
Create a new `WiaManager` class, and listen to device insert/delete events.

### Select a scanner
Get a scanner's id by using `WiaManager.GetDevices`, or manually select a scanner in WIA GUI by `WiaManager.ShowSelectDevice`.

### Scan
Call the `WiaManager.Scan` method.
You must provide a scanner id. Additional `ScanSettings` parameter could be optional.
