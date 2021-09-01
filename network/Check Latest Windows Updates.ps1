Get-WinEvent -LogName "System" | 
Where-Object { $_.ProviderName -eq "Microsoft-Windows-WindowsUpdateClient" -and $_.Id -eq 19 } | 
Select-Object TimeCreated, Message | 
Sort-Object TimeCreated -Descending | Format-Table -AutoSize
