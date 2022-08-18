# v2: detects USB vs network connection before install
# v2: detects USB vs network connection before install
# Get the hostname of the computer
$hostname = $env:COMPUTERNAME

# Check if the hostname contains 'FRD' or 'FD'
if ($hostname -match 'FRD|FD') {
    # Define the source and destination paths
    $sourcePath = "\\SYSADMIN01\Tools\Printers\Clubs\Front Desk\Epson TM-T88V Series Receipt Printer\Silent Installer\Receipt Printer.exe"
    $destinationPath = "C:\temp\Receipt Printer.exe"

    # Create the temp directory if it doesn't exist
    $destinationDir = Split-Path -Path $destinationPath -Parent
    if (-not (Test-Path -Path $destinationDir)) {
        New-Item -ItemType Directory -Path $destinationDir
    }

    # Copy the installer from the file share to the local temp directory
    Copy-Item -Path $sourcePath -Destination $destinationPath

    # Define the silent switch for the installer
    # (Replace '/S' with the appropriate silent switch for your installer)
    $silentSwitch = '/S'

    # Run the installer silently
    Start-Process -FilePath $destinationPath -ArgumentList $silentSwitch -Wait -NoNewWindow

    # Output completion message
    Write-Output "Installation completed for $hostname."

    # Optional: Remove the installer file after installation if you no longer need it
    Remove-Item -Path $destinationPath
} else {
    Write-Output "This script does not apply to the hostname $hostname."
}
