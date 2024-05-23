# Description: This script downloads the HP Universal Printing PCL 6 driver from the specified URL, extracts the driver files, and installs the driver on the local machine. It then adds the driver to the system, creates a printer port using the specified IP address, and adds a printer using the driver and port.

# Background: Printer management is a common task in IT environments, and automating the installation of printer drivers can save time and effort. This script demonstrates how to download, extract, and install a printer driver using PowerShell cmdlets.

# Usage: Run the script in a PowerShell environment with administrative privileges. Modify the variables at the beginning of the script to customize the driver, printer, IP address, and download path.

# Copyright: South Arm Technonology Services Ltd. This script is provided as-is with no warranties or guarantees. Use at your own risk.

# Author: Minghui Yu (myu@southarm.ca)

# GPL-3.0 License

# Download URL for HP universal driver
$url = "https://ftp.hp.com/pub/softlib/software13/printers/UPD/upd-pcl6-x64-7.2.0.25780.exe"
$driverName = "HP Universal Printing PCL 6"
$printerName = "Office HP Printer"
$ipAddr = "10.10.0.168"
# Specified extraction path 
$specifiedPath = "C:\temp"
# Extract the filename (version) from the URL
$fileName = ($url -split '/' | Select-Object -Last 1)
Write-Host "file name is" $fileName
$updVersion = $fileName.Substring($fileName.IndexOf('upd-') + 4).Replace('.exe', '') 
Write-Host "upd version is " $updVersion


# Download path with extracted filename
$downloadPath = "$specifiedPath\$fileName"
Write-Host "download path is" $downloadPath

# Renamed ZIP path
$renamedZipPath = "$specifiedPath\$updVersion.zip"
Write-Host "renamed zip path is" $renamedZipPath

# Destination path for extracted files (using extracted version)
$extractPath = "$specifiedPath\$updVersion"
Write-Host "extract path is" $extractPath

# Download the file
Invoke-WebRequest -Uri $url -OutFile $downloadPath

# Rename the downloaded file to ZIP (for potential compatibility)
Rename-Item -Path $downloadPath -NewName $renamedZipPath

# Expand the archive
try {
    Expand-Archive -Path $renamedZipPath -DestinationPath $extractPath -Force
    Write-Host "Download and extraction completed successfully"
  } catch {
    Write-Error "Error during extraction using Expand-Archive: $_"
  }

# Add INF to Windows store using pnp
try {
    Get-ChildItem -Path $extractPath -Filter *.inf | ForEach-Object {
    pnputil /add-driver $_.FullName /install
    }
} catch {
    Write-Error "Error in loading and adding drivers: $_"
}

# Add the printer driver
try {
    Add-PrinterDriver -Name $driverName
} catch {
    Write-Error "Error in adding $($driverName): $_"
}

$portName = "IP_{0}" -f $ipAddr  # Format the port name with the IP address

# Check if the port already exists
$existingPort = Get-PrinterPort -Name $portName

if (!$existingPort) {
  # Port doesn't exist, create it using Add-PrinterPort
    try{
        Add-PrinterPort -Name $portName -PrinterHostAddress $ipAddr
    }  catch {
        Write-Error "Error in adding printer port: $_"
    }
} else {
  # Port already exists
  Write-Host "Printer port $portName already exists. Skipping creation."
}

# Check if the printer already exists
$existingPrinter = Get-Printer -Name $printerName

if (!$existingPrinter) {
  # Printer doesn't exist, add it
    try {
        Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
    } catch {
        Write-Error "Error in adding printer: $_"
    }
} else {
  # Printer already exists
  Write-Host "Printer $printerName already exists. Skipping creation."
}

# Clean up renamed EXE/ZIP file
try {
    Remove-Item -Path $renamedZipPath -Force
    Write-Host "Downloaded ZIP file removed."
} catch {
    Write-Error "Failed to delete downloaded ZIP file: $_"
}
