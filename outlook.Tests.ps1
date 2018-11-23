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

    Describe "$ModuleName : Get-OutlookCalanderItem" {
       
        Context Output {
            
            It "Should Return a calander object" {
            } -Pending

        }
    }

}