# PowerShell Endpoint Management

Managing workstations across 100+ club locations meant constant printer setup, network troubleshooting, and hardware issues. Every new workstation needed the right printers installed silently, and barcode scanners at the front desk needed a reliable way to reboot when they locked up.

## Structure

### printers/club-locations/
Per-club printer deployment scripts. Each numbered script (601.ps1, 602.ps1, etc.) installs the correct printers for that specific club location, including the right IP addresses and driver configurations.

### printers/models/
Printer model setup scripts. Each handles driver installation and port configuration for a specific printer model (HP LaserJet, Canon imageRUNNER, receipt printers, etc.).

### printers/
- **Add-Printers-BO.ps1** - Adds back office printers to a workstation.
- **Add-Printers-FD.ps1** - Adds front desk printers to a workstation.
- **QUICKPRINT.ps1** - Quick printer deployment utility for common configurations.

### network/
- **Reset_Network.bat** - Resets network stack on a workstation (DNS flush, Winsock reset, IP release/renew).
- **Find Printers.ps1** - Discovers printers on the local network segment.
- **Check Latest Windows Updates.ps1** - Reports the latest installed Windows updates.

### scanners/
- **Troubleshoot-Scanner v3.ps1** - Reboots and resets barcode scanners at club front desks when they stop responding.

## Usage

```powershell
# Deploy printers for club 623
.\printers\club-locations\623.ps1

# Install an HP LaserJet Pro 4001n
.\printers\models\HP LaserJet Pro 4001n.ps1

# Reboot a stuck barcode scanner
.\scanners\Troubleshoot-Scanner v3.ps1
```

Most of these are designed to run silently via RMM or as startup scripts through GPO.

## Requirements

- PowerShell 5.1+
- Administrator access on target workstations
- Printer drivers available on the workstation or a network share
- Network access to printer IP addresses

## Blog Post

[Silent Printer Deployment Across 100+ Locations with PowerShell](https://blog.soarsystems.cc/powershell-endpoint-management)
