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
        
        # Add mocks for Exchange cmdlets if they don't exist
        if (!(Get-Command Get-MailboxFolder -ErrorAction SilentlyContinue)) {
            function Get-MailboxFolder {}
        }
        if (!(Get-Command New-MailboxFolder -ErrorAction SilentlyContinue)) {
            function New-MailboxFolder {}
        }
    }

    Context "When creating a new rule" {
        BeforeEach {
            # Reset captured parameters for each test
            $script:newInboxRuleCalls = @()
            $script:newMailboxFolderCalls = @()
            $script:writeHostCalls = @()
            
            # Mock with parameter capture
            Mock New-InboxRule -MockWith {
                param($Name, $FromAddressContainsWords, $MoveToFolder)
                $script:newInboxRuleCalls += @{
                    Name = $Name
                    FromAddressContainsWords = $FromAddressContainsWords
                    MoveToFolder = $MoveToFolder
                }
                return [PSCustomObject]@{
                    Name = $Name
                    FromAddressContainsWords = $FromAddressContainsWords
                    MoveToFolder = $MoveToFolder
                    Enabled = $true
                }
            }
            
            Mock Write-Host -MockWith {
                param($Object)
                $script:writeHostCalls += @{ Message = $Object }
            }
            
            Mock Get-InboxRule -MockWith { $null }
            
            Mock New-MailboxFolderHierarchy -MockWith {
                param($FolderPath)
                $script:newMailboxFolderCalls += @{ FolderPath = $FolderPath }
                return [PSCustomObject]@{ 
                    Identity = $FolderPath
                }
            }
        }
        
        It 'Should create a new inbox rule with the specified parameters' {
            # Arrange
            $ruleName = "TestRule"
            $fromAddress = "test@example.com"
            $targetFolder = ":Inbox\Test"
            
            # Act
            New-CustomInboxRule -RuleName $ruleName -FromAddress $fromAddress -TargetFolder $targetFolder
            
            # Assert
            $script:newInboxRuleCalls.Count | Should -Be 1
            $ruleCall = $script:newInboxRuleCalls[0]
            $ruleCall.Name | Should -Be $ruleName
            $ruleCall.FromAddressContainsWords | Should -Be $fromAddress
            $ruleCall.MoveToFolder | Should -Be $targetFolder
            
            $createMessage = $script:writeHostCalls | Where-Object { $_.Message -eq "Creating new inbox rule: $ruleName" }
            $createMessage | Should -Not -BeNull
        }
        
        It 'Should handle errors when creating a rule fails' {
            # Arrange
            Mock New-InboxRule { throw 'Failed to create rule' }
            
            # Act
            New-CustomInboxRule -RuleName "TestRule" -FromAddress "test@example.com" -TargetFolder "Inbox\Test"
              # Assert
            Should -Invoke Write-Error -Times 1 -ParameterFilter {
                Write-Verbose "Error message: $Message"
                $Message -like "*Failed to create inbox rule: Failed to create rule*"
            }
        }
        
        It 'Should create folder hierarchy before creating rule' {
            # Arrange
            $ruleName = "TestRule"
            $fromAddress = "test@example.com"
            $targetFolder = ":Inbox\NewFolder\SubFolder"
            
            # Act
            New-CustomInboxRule -RuleName $ruleName -FromAddress $fromAddress -TargetFolder $targetFolder
            
            # Assert
            $script:newMailboxFolderCalls.Count | Should -Be 1
            $script:newMailboxFolderCalls[0].FolderPath | Should -Be $targetFolder
            
            $script:newInboxRuleCalls.Count | Should -Be 1
            $ruleCall = $script:newInboxRuleCalls[0]
            $ruleCall.Name | Should -Be $ruleName
            $ruleCall.FromAddressContainsWords | Should -Be $fromAddress
            $ruleCall.MoveToFolder | Should -Be $targetFolder
            
            $createMessage = $script:writeHostCalls | Where-Object { $_.Message -eq "Creating new inbox rule: $ruleName" }
            $createMessage | Should -Not -BeNull
        }
    }
}

Describe "New-MailboxFolderHierarchy" {
    BeforeAll {
        Mock Write-Host
        Mock Write-Error
        
        # Add mocks for Exchange cmdlets if they don't exist
        if (!(Get-Command Get-MailboxFolder -ErrorAction SilentlyContinue)) {
            function Get-MailboxFolder {}
        }
        if (!(Get-Command New-MailboxFolder -ErrorAction SilentlyContinue)) {
            function New-MailboxFolder {}
        }
    }
    
    Context "When creating folder hierarchy" {
        BeforeEach {
            # Reset captured parameters for each test
            $script:getMailboxFolderCalls = @()
            $script:newMailboxFolderCalls = @()
            
            # Default mock behavior
            Mock Get-MailboxFolder -MockWith { 
                param($Identity)
                $script:getMailboxFolderCalls += @{ Identity = $Identity }
                throw "Folder not found" 
            }
            
            Mock New-MailboxFolder -MockWith {
                param($Parent, $Name)
                $script:newMailboxFolderCalls += @{
                    Parent = $Parent
                    Name = $Name
                }
                return [PSCustomObject]@{
                    Identity = "$Parent\$Name"
                }
            }
        }
        
        It 'Should normalize folder path to start with :\' {
            # Act
            $result = New-MailboxFolderHierarchy -FolderPath "Inbox\TestFolder"
            
            # Assert
            $result.Identity | Should -Be ":\Inbox\TestFolder"
            $script:newMailboxFolderCalls.Count | Should -BeGreaterThan 0
        }
        
        It 'Should create nested folder structure' {
            # Act
            $result = New-MailboxFolderHierarchy -FolderPath ":\Inbox\Parent\Child\Grandchild"
            
            # Assert
            $script:newMailboxFolderCalls.Count | Should -Be 4
            $script:newMailboxFolderCalls[0].Name | Should -Be "Inbox"
            $script:newMailboxFolderCalls[1].Name | Should -Be "Parent"
            $script:newMailboxFolderCalls[2].Name | Should -Be "Child"
            $script:newMailboxFolderCalls[3].Name | Should -Be "Grandchild"
            $result.Identity | Should -Be ":\Inbox\Parent\Child\Grandchild"
        }
        
        It 'Should skip existing folders' {
            # Arrange
            Mock Get-MailboxFolder -MockWith { 
                param($Identity)
                $script:getMailboxFolderCalls += @{ Identity = $Identity }
                if ($Identity -eq ":\Inbox\Existing") {
                    return [PSCustomObject]@{ Identity = $Identity }
                }
                throw "Folder not found"
            }
            
            # Act
            $result = New-MailboxFolderHierarchy -FolderPath ":\Inbox\Existing\New"
            
            # Assert
            $createdFolder = $script:newMailboxFolderCalls | Where-Object { 
                $_.Parent -eq ":\Inbox\Existing" -and $_.Name -eq "New" 
            }
            $createdFolder | Should -Not -BeNull
            $script:newMailboxFolderCalls.Count | Should -Be 2 # Should only create Inbox and New
            $result.Identity | Should -Be ":\Inbox\Existing\New"
        }
        
        It 'Should handle folder creation errors' {
            # Arrange
            Mock New-MailboxFolder -MockWith { 
                param($Parent, $Name)
                $script:newMailboxFolderCalls += @{
                    Parent = $Parent
                    Name = $Name
                }
                throw "Access denied" 
            }
            
            # Act & Assert
            { New-MailboxFolderHierarchy -FolderPath ":\Inbox\Test" } | 
                Should -Throw "Failed to create folder hierarchy ':\Inbox\Test': Failed to create folder ':\Inbox': Access denied"
        }
    }
}

Describe "Remove-CustomInboxRule" {
    BeforeAll {
        Mock Write-Host
        Mock Write-Warning
        Mock Write-Error
    }

    Context "When removing rules" {
        BeforeEach {
            # Reset captured parameters
            $script:removeInboxRuleCalls = @()
            
            # Mock Get-InboxRule to return test rules
            Mock Get-InboxRule -MockWith {
                @(
                    [PSCustomObject]@{ 
                        Name = "TestRule"
                        Identity = "Rule1"
                        FromAddressContainsWords = "test1@example.com"
                        MoveToFolder = ":\Inbox\Folder1"
                    },
                    [PSCustomObject]@{ 
                        Name = "TestRule"
                        Identity = "Rule2"
                        FromAddressContainsWords = "test2@example.com"
                        MoveToFolder = ":\Inbox\Folder2"
                    }
                )
            }
            
            # Mock Remove-InboxRule with parameter capture
            Mock Remove-InboxRule -MockWith {
                param($Identity, $Confirm)
                $script:removeInboxRuleCalls += @{
                    Identity = $Identity
                    Confirm = $Confirm
                }
            }
        }
        
        It 'Should warn when multiple rules found without RemoveAll' {
            # Act
            Remove-CustomInboxRule -RuleName "TestRule"
            
            # Assert
            Should -Invoke Write-Warning -Times 1 -ParameterFilter {
                $Message -like "*Found 2 rules with name 'TestRule'*"
            }
            $script:removeInboxRuleCalls.Count | Should -Be 0
        }
        
        It 'Should remove all matching rules when RemoveAll specified' {
            # Act
            Remove-CustomInboxRule -RuleName "TestRule" -RemoveAll
            
            # Assert
            $script:removeInboxRuleCalls.Count | Should -Be 2
            $script:removeInboxRuleCalls[0].Identity | Should -Be "Rule1"
            $script:removeInboxRuleCalls[1].Identity | Should -Be "Rule2"
            $script:removeInboxRuleCalls | ForEach-Object {
                $_.Confirm | Should -Be $false
            }
        }
        
        It 'Should handle no matching rules' {
            # Arrange
            Mock Get-InboxRule -MockWith { @() }
            
            # Act
            Remove-CustomInboxRule -RuleName "NonExistentRule"
            
            # Assert
            Should -Invoke Write-Warning -Times 1 -ParameterFilter {
                $Message -eq "No rules found with name: NonExistentRule"
            }
            $script:removeInboxRuleCalls.Count | Should -Be 0
        }
        
        It 'Should handle rule removal errors' {
            # Arrange
            Mock Remove-InboxRule { throw "Failed to remove rule" }
            
            # Act
            Remove-CustomInboxRule -RuleName "TestRule" -RemoveAll
            
            # Assert
            Should -Invoke Write-Error -Times 1 -ParameterFilter {
                $Message -like "*Failed to remove inbox rule(s): Failed to remove rule*"
            }
        }
    }
}


