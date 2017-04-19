<#
    .Synopsis
    Tests finance.psm1 with examples

#>
Import-Module $PSScriptRoot\finance.psm1 -Force

Write-Output $("Carnegie Clean Enery {0:c3}" -f $(Get-Price -Code CCE))

Write-Output $("Aussie Dollar to Japanese Yen Exchange rate: `$1.00 = {1}{0:n1}" -f  $(Get-CurrencyRate -fromCurrency "AUD" -toCurrency "JPY"), [char]0x00A5)


Get-Portfolio "D:\local\scripts\powershell\finance\example\short_sunny_day.csv"
Get-Portfolio "D:\local\scripts\powershell\finance\example\sunny_day.csv"
Get-Portfolio "D:\local\scripts\powershell\finance\example\bad_asx_code.csv"

Write-Output "Expect error as it is not a csv file"
Get-Portfolio "D:\local\scripts\powershell\finance\example\sunny_day.htm"
Get-Portfolio "D:\local\scripts\powershell\finance\example\bad_long_day.csv"