# Load the WIA .NET library
[void][reflection.assembly]::LoadWithPartialName("WIA")

# Create a new WIA Device Manager
$deviceManager = New-Object -ComObject WIA.DeviceManager

# Select the specific scanner by its ID
$deviceInfo = $deviceManager.DeviceInfos |
    Where-Object { $_.Type -eq 1 -and $_.DeviceID -eq '{6BDD1FC6-810F-11D0-BEC7-08002BE2092F}\0000' } |
    Select-Object -First 1

if ($null -eq $deviceInfo) {
    Write-Host "Scanner not found. Please make sure your scanner is connected and try again."
    exit
}

$device = $deviceInfo.Connect()

# Define the output file path
$outputFilePath = "C:\Users\chris\Desktop\scrub\scan.jpg"

# Start the scan and save the image to the output file
$item = $device.Items[1]

if ($null -eq $item) {
    Write-Host "No items found on the scanner. Please make sure your document is correctly placed and try again."
    exit
}

$imageFile = $item.Transfer()
$imageFile.SaveFile($outputFilePath)

# Inform the user
Write-Host "Saved scanned image to $outputFilePath"
