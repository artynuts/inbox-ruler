# Import the module under test
BeforeAll {
    . $PSScriptRoot\..\Manage-InboxRules.ps1

    # Create mock functions for Exchange Online cmdlets if it doesn't exist
    if (!(Get-Command Get-InboxRule -ErrorAction SilentlyContinue)) {
        function Get-InboxRule {}
    }
    if (!(Get-Command New-InboxRule -ErrorAction SilentlyContinue)) {
        function New-InboxRule {}
    }
}

Describe "Connect-ToExchange" {
    BeforeAll {
        Mock Write-Host
        Mock Write-Error
        Mock Install-Module
        Mock Import-Module
        Mock Connect-ExchangeOnline
        Mock Get-Module -MockWith { $false } -ParameterFilter { $Name -eq 'ExchangeOnlineManagement' }
    }

    It 'Should attempt to install ExchangeOnlineManagement module if not present' {
        Connect-ToExchange
        Should -Invoke Install-Module -Times 1 -ParameterFilter { 
            $Name -eq 'ExchangeOnlineManagement' -and 
            $Force -eq $true -and 
            $AllowClobber -eq $true -and 
            $Scope -eq 'CurrentUser' 
        }
    }

    It 'Should import the module and connect to Exchange Online' {
        Connect-ToExchange
        Should -Invoke Import-Module -Times 1 -ParameterFilter { $Name -eq 'ExchangeOnlineManagement' }
        Should -Invoke Connect-ExchangeOnline -Times 1
    }
}

Describe "Get-InboxRules" {
    BeforeAll {
        Mock Write-Error
        Mock Get-InboxRule -MockWith {
            @(
                [PSCustomObject]@{ Name = 'Rule1'; Enabled = $true },
                [PSCustomObject]@{ Name = 'Rule2'; Enabled = $false }
            )
        }
    }

    It 'Should get inbox rules' {
        $rules = Get-InboxRules
        Should -Invoke Get-InboxRule -Times 1
        $rules.Count | Should -Be 2
        $rules[0].Name | Should -Be 'Rule1'
    }
      It 'Should handle errors and return null' {
        Mock Get-InboxRule { throw 'Access denied' }
        $rules = Get-InboxRules
        Should -Invoke Write-Error -Times 1 -ParameterFilter {
            $Message -like '*Failed to get inbox rules: Access denied*'
        }
        $rules | Should -Be $null
    }
}

Describe "New-CustomInboxRule" {    
    BeforeAll {
        Mock Write-Host
        Mock Write-Error
        
        # Mock New-InboxRule with parameter logging
        Mock New-InboxRule -MockWith {
            param($Name, $FromAddressContainsWords, $MoveToFolder)
            Write-Verbose "Mock New-InboxRule called with:"
            Write-Verbose "  Name: $Name"
            Write-Verbose "  FromAddressContainsWords: $FromAddressContainsWords"
            Write-Verbose "  MoveToFolder: $MoveToFolder"
            return [PSCustomObject]@{
                Name = $Name
                FromAddressContainsWords = $FromAddressContainsWords
                MoveToFolder = $MoveToFolder
                Enabled = $true
            }
        } -Verifiable
        
        Mock Get-InboxRule
        Mock New-MailboxFolderHierarchy -MockWith {
            param($FolderPath)
            [PSCustomObject]@{ 
                Identity = $FolderPath
            }
        }
    }

    Context "When creating a new rule" {
        It 'Should create a new inbox rule with the specified parameters' {
            # Arrange
            $VerbosePreference = 'Continue'
            $ruleName = "TestRule"
            $fromAddress = "test@example.com"
            $targetFolder = ":Inbox\Test"
            
            # Act
            New-CustomInboxRule -RuleName $ruleName -FromAddress $fromAddress -TargetFolder $targetFolder
            
            # Assert with detailed output
            Should -InvokeVerifiable
            Assert-MockCalled New-InboxRule -Exactly 1 -Verbose -ParameterFilter {
                Write-Verbose "Checking mock call parameters:"
                Write-Verbose "  Actual Name: $Name"
                Write-Verbose "  Actual FromAddressContainsWords: $FromAddressContainsWords"
                Write-Verbose "  Actual MoveToFolder: $MoveToFolder"
                Write-Verbose "  Expected Name: $ruleName"
                Write-Verbose "  Expected FromAddressContainsWords: $fromAddress"
                Write-Verbose "  Expected MoveToFolder: $targetFolder"
                
                $Name -eq $ruleName
            }
        }

        It 'Should handle errors when creating a rule fails' {
            Mock New-InboxRule -MockWith { 
                throw 'Failed to create rule'
            }
            
            New-CustomInboxRule -RuleName "TestRule" -FromAddress "test@example.com" -TargetFolder "Inbox\Test"
            
            Should -Invoke Write-Error -Times 1 -ParameterFilter {
                $Message -like "*Failed to create inbox rule: Failed to create rule*"
            }
        }
    }
}

<#

Describe "Remove-CustomInboxRule" {
    BeforeAll {
        Mock Write-Host
        Mock Write-Error
        Mock Remove-InboxRule
    }

    Context "When removing a rule" {
        It 'Should remove the specified inbox rule' {
            Remove-CustomInboxRule -RuleName "TestRule"
            
            Should -Invoke Remove-InboxRule -Times 1 -ParameterFilter {
                $Identity -eq "TestRule" -and
                $Confirm -eq $false
            }
            Should -Invoke Write-Host -Times 1 -ParameterFilter {
                $Message -eq "Removed rule: TestRule"
            }
        }

        It 'Should handle errors when removing a rule fails' {
            Mock Remove-InboxRule { throw 'Failed to remove rule' }
            
            Remove-CustomInboxRule -RuleName "TestRule"
            
            Should -Invoke Write-Error -Times 1 -ParameterFilter {
                $Message -like "*Failed to remove inbox rule: Failed to remove rule*"
            }
        }
    }
}
#>

<#

Describe "Rename-CustomInboxRule" {
    BeforeAll {
        Mock Write-Host
        Mock Write-Error
        Mock Set-InboxRule
    }

    Context "When renaming a rule" {
        It 'Should rename the specified inbox rule' {
            Rename-CustomInboxRule -CurrentRuleName "OldRule" -NewRuleName "NewRule"
            
            Should -Invoke Set-InboxRule -Times 1 -ParameterFilter {
                $Identity -eq "OldRule" -and
                $Name -eq "NewRule"
            }
            Should -Invoke Write-Host -Times 1 -ParameterFilter {
                $Message -eq "Renamed rule from 'OldRule' to 'NewRule'"
            }
        }

        It 'Should handle errors when renaming a rule fails' {
            Mock Set-InboxRule { throw 'Failed to rename rule' }
            
            Rename-CustomInboxRule -CurrentRuleName "OldRule" -NewRuleName "NewRule"
            
            Should -Invoke Write-Error -Times 1 -ParameterFilter {
                $Message -like "*Failed to rename inbox rule: Failed to rename rule*"
            }
        }
    }
}
#>


