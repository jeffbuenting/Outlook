Import-Module 'F:\OneDrive for Business\Scripts\Outlook\outlook.psm1' -force

$Items = Get-OLCalendarItem -start '08/14/2015' -Verbose

$Items