Get-WmiObject -ComputerName $env:COMPUTERNAME -Class Win32_Product | Sort-Object Name

<#
Other Invocations:

.\Get-Software.ps1 | Format-Table -auto -property Name,Vendor,Version,Caption

.\Get-Software.ps1 | ConvertTo-Csv -NoTypeInformation
(.\Get-Software.ps1 | ConvertTo-Csv -NoTypeInformation) -replace('"','')

(.\Get-Software.ps1 | Export-Csv <file path> -NoTypeInformation) -replace('"','')
#>
