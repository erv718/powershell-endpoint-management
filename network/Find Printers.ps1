function Test-Port {
    param(
        [string]$IP,
        [int]$Port
    )
    $tcpClient = New-Object System.Net.Sockets.TcpClient
    try {
        $tcpClient.Connect($IP, $Port)
        if ($tcpClient.Connected) {
            $tcpClient.Close()
            return $true
        }
    } catch {
        return $false
    }
}

# Specify the full path to snmpwalk
$snmpwalkPath = "\\SYSADMIN01\Tools\snmpwalk\snmpwalk.exe"

# Discover the local subnet dynamically
$ipConfig = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" -and $_.PrefixOrigin -ne "WellKnown" }
$localIP = $ipConfig.IPAddress
$subnetBase = $localIP -replace '(\d+\.\d+\.\d+\.)\d+', '$1'
$startRange = 1
$endRange = 254

# Common printer ports
$printerPorts = 9100, 515, 631

# Initialize an array to hold active IPs with open printer ports
$activePrinterIPs = @()

Write-Output "Scanning for printers in the subnet: $subnetBase*"
for ($i = $startRange; $i -le $endRange; $i++) {
    $ip = "$subnetBase$i"
    foreach ($port in $printerPorts) {
        if (Test-Port -IP $ip -Port $port) {
            $activePrinterIPs += $ip
            break # Move to the next IP after finding any open printer port
        }
    }
}

# Attempt SNMP queries on devices with open printer ports using snmpwalk
foreach ($ip in $activePrinterIPs) {
    try {
        $community = "public" # SNMP community string
        $oid = ".1.3.6.1.2.1.1.1" # OID for system description, adjust if necessary
        
        # Construct and invoke the snmpwalk command
        $command = "& `"$snmpwalkPath`" -r:$ip -v:2c -c:$community -os:$oid"
        $result = Invoke-Expression $command
        
        # Process the result to extract Printer Model
        if ($result) {
            $model = $result -replace '.*STRING:\s*', '' # Adjust based on actual output format
            Write-Output "Printer Model: $model at IP: $ip"
        }
    } catch {
        Write-Output "Failed to query ${ip} via SNMP with snmpwalk."
    }
}

Write-Output "Script execution complete."