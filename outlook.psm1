Function Get-OutlookCalendarItem {

    <#
        .Synopsis
            Returns Outlook Calander Items.

        .Description
            Gets a list of Calendar items within the date range from the current logged in user.

        .Parameter BeginDate
            Beginning of date range.  Defaults to todays date.

        .Parameter EndDate
            End of date range.  Defaults to yesterday (-1)

        .Link
            Outlook Application

            https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application

        .Link
            Outlook Namespace folder types

            https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders

    #>

    [CmdletBinding()]
    Param ( 
        [DateTime]$BeginDate = $(Get-date),

        [DateTime]$EndDate = $((Get-Date).AddDays( -1 ))
    )

    Try {
        Write-Verbose "Connecting to Outlook"

        $Outlook = new-object -comobject outlook.application -ErrorAction Stop
    }
    Catch {
        $EXceptionMessage = $_.Exception.Message
        $ExceptionType = $_.exception.GetType().fullname
        Throw "Get-OutlookCalendarItem : Failed to create Outlook Object.`n     Possibly outlook is not installed.`n`n     $ExceptionMessage`n`n     Exception : $ExceptionType" 
    }
    write-verbose 'one'
    $namespace = $outlook.GetNameSpace("MAPI")
   write-verbose 'two'
    #if ( $StartDate -gt $EndDate ) {

        $Filter = "[Start] < '$($BeginDate.ToString())' AND [End] > '$($EndDate.ToString())'"
    #}

    Write-verbose $Filter

    Write-Output $NameSpace.GetDefaultFolder( 9 ).items.restrict($FIlter)

}