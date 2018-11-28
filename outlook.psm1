Function Get-OutlookCalendarItem {

    <#
        .Synopsis
            Returns Outlook Calander Items.

        .Description
            Gets a list of Calendar items within the date range from the current logged in user.

        .Parameter BeginDate
            Beginning of date range. 

        .Parameter EndDate
            End of date range.  

        .Example
            Retrieve all calendar objects with a categorie of 'Client'

            Get-OutlookCalendarItem -Categories Client

        .Link
            Outlook Application

            https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application

        .Link
            Outlook Namespace folder types

            https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders

        .Notes
            Author : Jeff Buenting
            Date : 2018 NOV 26

    #>

    [CmdletBinding()]
    Param ( 
        [DateTime]$BeginDate,

        [String[]]$Categories
    )

    # ----- EndDate is only mandatory if BeginDate is included
    # ----- Dynamic Parameters : https://www.powershellmagazine.com/2014/05/29/dynamic-parameters-in-powershell/
    DynamicParam {
        if ( $BeginDate ) {
          
            
            $Attrib = New-object System.Management.Automation.ParameterAttribute 
            $Attrib.Mandatory = $True
            $attrib.ParameterSetName = '__AllParameterSets'

            $AttribCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $AttribCollection.Add($Attrib)

            $Param = New-object System.Management.Automation.RuntimeDefinedParameter('EndDate',[DateTime],$AttribCollection)
            
            $ParamList = NEw-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $ParamList.Add( 'EndDate', $Param )

            return $ParamList
        }
    }

    Process {

        # ----- This is how you reference the dynamic parameter with $PSBoundPa
        $EndDate = $PSBoundParameters.EndDate

        
        Try {
            Write-Verbose "Connecting to Outlook"

            $Outlook = new-object -comobject outlook.application -ErrorAction Stop -Verbose:$false
        
        }
        Catch {
            $EXceptionMessage = $_.Exception.Message
            $ExceptionType = $_.exception.GetType().fullname
            Throw "Get-OutlookCalendarItem : Failed to create Outlook Object.`n     Possibly outlook is not installed.`n`n     $ExceptionMessage`n`n     Exception : $ExceptionType" 
        }
        
        $namespace = $outlook.GetNameSpace("MAPI")
       
        # ----- Build FIlter
        # ----- https://docs.microsoft.com/en-us/office/vba/api/Outlook.Items.Restrict
        $Filter = @()

        # ----- Date Range
        # ----- Date format needs to be a specific form.  per above website
        if ( $BeginDate ) {
            write-verbose "begindate = $(get-date $BeginDate -UFormat '%m/%d/%y %I:%M %p') "
            Write-Verbose "enddate = $enddate"
            if ( $Filter -ne $Null ) {
                $Filert = "$Filter AND "
            }

           if ( $BeginDate -gt $EndDate ) {
               $Filter = "[Start] <= '$(Get-Date $BeginDate -Uformat "%m/%d/%y %I:%M %p")' AND [End] >= '$(Get-Date $EndDate -Uformat "%m/%d/%y %I:%M %p")'"
           }
           Else {
                
                $Filter = "[Start] >= '$(Get-Date $BeginDate -Uformat "%m/%d/%y %I:%M %p")' AND [End] <= '$(Get-Date $EndDate -Uformat "%m/%d/%y %I:%M %p")'"
           }
        }

        # ----- categories
        If ( $Categories ) {
            Write-Verbose "Adding Categories to filter"
            if ( $Filter -ne $Null ) {
                $Filter = "$Filter AND "
            }

            Foreach ( $C in $Categories ) {
                $Filter = "$Filter[Categories] = '$C' OR "
            }

            # ----- Remove the last ' OR '
            $Filter = $Filter.Trim( ' OR ' )
        }     
        
    
        if ( $Filter ) {
            Write-verbose "Retrieving Calendar objects that match = $Filter"


            Write-Output $NameSpace.GetDefaultFolder( 9 ).items.restrict($FIlter)
        }
        Else {
            Write-Verbose "Retrieving all Calendar objects"

            Write-Output $NameSpace.GetDefaultFolder( 9 ).items
        }
    }

}