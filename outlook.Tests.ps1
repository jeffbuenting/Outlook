# ----- Get the module name
if ( -Not $PSScriptRoot ) { $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent }

$ModulePath = $PSScriptRoot

$Global:ModuleName = $ModulePath | Split-Path -Leaf

# ----- Remove and then import the module.  This is so any new changes are imported.
Get-Module -Name $ModuleName -All | Remove-Module -Force -Verbose

Import-Module "$ModulePath\$ModuleName.PSD1" -Force -ErrorAction Stop  

InModuleScope $ModuleName {

    #-------------------------------------------------------------------------------------
    # ----- Check if all fucntions in the module have a unit tests

    Describe "$ModuleName : Module Tests" {

        $Module = Get-module -Name $ModuleName -Verbose

        $testFile = Get-ChildItem $module.ModuleBase -Filter '*.Tests.ps1' -File -verbose
    
        $testNames = Select-String -Path $testFile.FullName -Pattern 'describe\s[^\$](.+)?\s+{' | ForEach-Object {
            [System.Management.Automation.PSParser]::Tokenize($_.Matches.Groups[1].Value, [ref]$null).Content
        }

        $moduleCommandNames = (Get-Command -Module $ModuleName | where CommandType -ne Alias)

        it 'should have a test for each function' {
            Compare-Object $moduleCommandNames $testNames | where { $_.SideIndicator -eq '<=' } | select inputobject | should beNullOrEmpty
        }
    }

    #-------------------------------------------------------------------------------------

    Write-Output "`n`n"

    Describe "$ModuleName : Get-OutlookCalendarItem" {
       
        
        # ----- Get Function Help

        # ----- Pester to test Comment based help

        # ----- http://www.lazywinadmin.com/2016/05/using-pester-to-test-your-comment-based.html

        Context "Help" {



            $H = Help Get-OutlookCalendarItem -Full



            # ----- Help Tests

            It "has Synopsis Help Section" {

                { $H.Synopsis }  | Should Not BeNullorEmpty

            }



            It "has Synopsis Help Section that it not start with the command name" {

                 $H.Synopsis | Should Not Match $H.Name

            }



            It "has Description Help Section" {

                 { $H.Description } | Should Not BeNullorEmpty

            }

            It "has Parameters Help Section" {

                 { $H.Parameters.parameter } | Should Not BeNullorEmpty

            }



            # Examples

            it "Example - Count should be greater than 0"{

                 { $H.examples.example } | Measure-Object | Select-Object -ExpandProperty Count | Should BeGreaterthan 0

            }

            

            # Examples - Remarks (small description that comes with the example)

            foreach ($Example in $H.examples.example)

            {

                it "Example - Remarks on $($Example.Title)"{

                     $Example.remarks  | Should not BeNullOrEmpty

                }

            }



            It "has Notes Help Section" {

                { $H.alertSet } | Should Not BeNullorEmpty

            }

        } 

        Mock -CommandName New-Object -ParameterFilter { $comobject } -MockWith {
            $Obj = New-Object -TypeName PSObject
            
            $Obj | Add-Member -MemberType ScriptMethod -Name GetNameSpace -value {
                Param ( $Source )

                $GetDefaultFolder= New-Object -TypeName PSObject
                $GetDefaultFolder | Add-Member -MemberType ScriptMethod -Name GetDefaultFolder -Value {
                    Param ( $FolderType )

                    $Restrict = New-Object -TypeName PSObject
                    $Restrict | Add-Member -MemberType ScriptMethod -Name Restrict -value {
                        param ( $filter )

                        Write-output ( New-Object -TypeName PSObject )
                    }


                    $Items= New-Object -TypeName PSObject
                    $Items | Add-Member -MemberType NoteProperty -Name Items -value $Restrict

                    Write-Output $Items
                }

                Write-Output $GetDefaultFolder
            }            
            
            Return $Obj
        }

        Context Execution {



            It "SHould throw an error if outlook is not installed" {
                Mock -CommandName New-Object -MockWith { Throw "error" }

                { Get-OutlookCalendarItem }  | Should Throw
            }

            It "Should throw an error if end date is included with begin date" {
                { Get-OutlookCalendarItem -BeginDate (Get-Date) }  | Should Throw
            } 

            It "Should accept a date range" {
                { Get-OutlookCalendarItem -BeginDate (Get-Date) -EndDate (get-date).adddays( -2) }  | Should Throw
            } 
        }

        Context Output {
            
            It "Should Return a calendar object" {

                Get-OutlookCalendarItem | Should beoftype PSObject
            } 

            It "Should Return a Calendar object when using a filter" {
                Get-OutlookCalendarItem -Categories 'one'  | Should beoftype PSObject
            } 

            It "should return Calendar object when more than one categorie is included" {
                Get-OutlookCalendarItem -Categories 'one','two'  | Should beoftype PSObject
            } 

            It "Should accept a date range and return a Calendar Object" {
                Get-OutlookCalendarItem -BeginDate (Get-Date) -EndDate (get-date).adddays( -2) | Should beoftype PSObject
            } 

        }
    }

}