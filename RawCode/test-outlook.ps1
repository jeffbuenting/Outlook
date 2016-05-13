Import-Module 'F:\OneDrive for Business\Scripts\Outlook\outlook.psm1' -force

$Items = Get-OLCalendarItem  -Verbose

$Items | foreach {
    
