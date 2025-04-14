# Import the module under test
BeforeAll {
    . $PSScriptRoot\..\Manage-InboxRules.ps1

    # Create mock functions for Exchange Online cmdlets if they don't exist
    if (!(Get-Command Get-AcceptedDomain -ErrorAction SilentlyContinue)) {
        function Get-AcceptedDomain {}
    }
    if (!(Get-Command Get-InboxRule -ErrorAction SilentlyContinue)) {
        function Get-InboxRule {}
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
        Mock Get-AcceptedDomain -MockWith { 
            [PSCustomObject]@{ Default = $true; DomainName = 'contoso.com' }
        }
        Mock Get-CurrentUser -MockWith {
            $mockIdentity = [PSCustomObject]@{ Name = 'DOMAIN\testuser' }
            $mockIdentity | Add-Member -MemberType ScriptMethod -Name "GetType" -Value { 
                return [System.Security.Principal.WindowsIdentity] 
            } -Force
            return $mockIdentity
        }
        Mock Get-InboxRule -MockWith {
            @(
                [PSCustomObject]@{ Name = 'Rule1'; Enabled = $true },
                [PSCustomObject]@{ Name = 'Rule2'; Enabled = $false }
            )
        }
    }

    It 'Should get rules for default mailbox when no mailbox specified' {
        $rules = Get-InboxRules
        Should -Invoke Get-AcceptedDomain -Times 1
        Should -Invoke Get-CurrentUser -Times 1
        Should -Invoke Get-InboxRule -Times 1
        $rules.Count | Should -Be 2
        $rules[0].Name | Should -Be 'Rule1'
    }

    It 'Should get rules for specified mailbox' {
        $rules = Get-InboxRules -Mailbox 'user@example.com'
        Should -Not -Invoke Get-AcceptedDomain
        Should -Not -Invoke Get-CurrentUser
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
