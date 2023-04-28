# Load the WIA .NET library
[void][reflection.assembly]::LoadWithPartialName("WIA")

# Create a new WIA Device Manager
$deviceManager = New-Object -ComObject WIA.DeviceManager

# List all available scanners (ScannerDeviceType is 1)
$scannerInfos = $deviceManager.DeviceInfos | Where-Object { $_.Type -eq 1 }

foreach ($scannerInfo in $scannerInfos) { 
    Write-Host "Scanner ID: $($scannerInfo.DeviceID)"
    Write-Host "Scanner Name: $($scannerInfo.Properties['Name'].Value)"
    Write-Host "Scanner Description: $($scannerInfo.Properties['Description'].Value)"
    Write-Host "---"
}
