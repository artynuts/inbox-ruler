# Manage-InboxRules.ps1
# This script helps manage Exchange/Outlook inbox rules programmatically

# Import Exchange Online Management module if not already loaded
function Connect-ToExchange {
    try {
        if (!(Get-Module -Name ExchangeOnlineManagement -ListAvailable -ErrorAction SilentlyContinue)) {
            Write-Host "Exchange Online Management module not found. Installing..."
            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
        }
        
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $_"
        exit 1
    }
}

function Get-InboxRules {
    try {
        return Get-InboxRule
    }
    catch {
        Write-Error "Failed to get inbox rules: $_"
        return $null
    }
}

function New-MailboxFolderHierarchy {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )    
    try {
        
        # Ensure the folder path is in the correct format
        if (-not $FolderPath.StartsWith(':\')) {
            $FolderPath = ":\$($FolderPath.TrimStart('\'))"
        }

        # Split the path into segments and ensure each segment exists
        $segments = $FolderPath.Split('\')
        $currentPath = $segments[0] # Start with ":"        
        
        # Create each segment of the path if it doesn't exist
        $folderObject = $null
        for ($i = 1; $i -lt $segments.Count; $i++) {            
            $parentPath = $currentPath
            $currentPath = "$currentPath\$($segments[$i])"

            try {
                # Try to get existing folder
                $folderObject = Get-MailboxFolder -Identity $currentPath -ErrorAction Stop
                Write-Host "Folder exists: $currentPath"
            }
            catch {
                Write-Host "Creating folder: $currentPath"
                try {
                    # Create new folder and use its return value directly
                    $folderObject = New-MailboxFolder -Parent $parentPath -Name $segments[$i]
                    Write-Host "Successfully created folder: (Identity: $($folderObject.Identity))"
                }
                catch {
                    throw "Failed to create folder '$currentPath': $_"
                }
            }
        }
        return $folderObject
    }
    catch {
        throw "Failed to create folder hierarchy '$FolderPath': $_"
    }
}

function New-CustomInboxRule {
    param(
        [Parameter(Mandatory=$true)]
        [string]$RuleName,
        
        [Parameter(Mandatory=$true)]
        [string]$FromAddress,
        
        [Parameter(Mandatory=$true)]
        [string]$TargetFolder
    )      
    
    try {
        # Check if a rule with this name already exists
        $existingRule = Get-InboxRule -Identity $RuleName -ErrorAction SilentlyContinue
        if ($existingRule) {
            Write-Warning "Rule '$RuleName' already exists. Use Remove-CustomInboxRule first if you want to recreate it."
            return
        }        
        
        # Create the folder hierarchy if it doesn't exist
        $folderObject = New-MailboxFolderHierarchy -FolderPath $TargetFolder
        $targetFolderUpdated = $folderObject.Identity

        # Create the inbox rule
        Write-Host "Creating new inbox rule: $RuleName"
        Write-Host "moving emails from: $FromAddress to folder: $targetFolderUpdated"
        New-InboxRule -Name $RuleName -FromAddressContainsWords $FromAddress -MoveToFolder $targetFolderUpdated
        Write-Host "Created new rule: $RuleName"
    }
    catch {
        Write-Error "Failed to create inbox rule: $_"
    }
}

function Remove-CustomInboxRule {
    param(
        [Parameter(Mandatory=$true)]
        [string]$RuleName,
        
        [Parameter(Mandatory=$false)]
        [switch]$RemoveAll = $false
    )
      
    try {
        # Get all rules with the specified name
        $matchingRules = Get-InboxRule | Where-Object { $_.Name -eq $RuleName }
        
        if (-not $matchingRules) {
            Write-Warning "No rules found with name: $RuleName"
            return
        }
        
        # If multiple rules found and RemoveAll is not specified
        if ($matchingRules.Count -gt 1 -and -not $RemoveAll) {
            Write-Warning "Found $($matchingRules.Count) rules with name '$RuleName'. Use -RemoveAll to remove all matching rules."
            Write-Host "Matching rules:"
            $matchingRules | ForEach-Object {
                Write-Host "- Rule ID: $($_.Identity), From: $($_.FromAddressContainsWords), To Folder: $($_.MoveToFolder)"
            }
            return
        }
        
        # Remove all matching rules
        $matchingRules | ForEach-Object {
            Remove-InboxRule -Identity $_.Identity -Confirm:$false
            Write-Host "Removed rule: $($_.Name) (ID: $($_.Identity))"
        }
    }
    catch {
        Write-Error "Failed to remove inbox rule(s): $_"
    }
}

function Rename-CustomInboxRule {
    param(
        [Parameter(Mandatory=$true)]
        [string]$CurrentRuleName,
        
        [Parameter(Mandatory=$true)]
        [string]$NewRuleName,
        
        [Parameter(Mandatory=$false)]
        [switch]$RenameAll = $false
    )
    try {
        # Get all rules with the current name
        $matchingRules = Get-InboxRule | Where-Object { $_.Name -eq $CurrentRuleName }
        
        if (-not $matchingRules) {
            Write-Warning "No rules found with name: $CurrentRuleName"
            return
        }
        
        # If multiple rules found and RenameAll is not specified
        if ($matchingRules.Count -gt 1 -and -not $RenameAll) {
            Write-Warning "Found $($matchingRules.Count) rules with name '$CurrentRuleName'. Use -RenameAll to rename all matching rules."
            Write-Host "Matching rules:"
            $matchingRules | ForEach-Object {
                Write-Host "- Rule ID: $($_.Identity), From: $($_.FromAddressContainsWords), To Folder: $($_.MoveToFolder)"
            }
            return
        }
        
        # Check if any rule with the new name already exists
        $existingNewName = Get-InboxRule | Where-Object { $_.Name -eq $NewRuleName }
        if ($existingNewName) {
            Write-Warning "A rule with name '$NewRuleName' already exists. Please choose a different name."
            return
        }
        
        # Rename all matching rules
        $matchingRules | ForEach-Object {
            Set-InboxRule -Identity $_.Identity -Name $NewRuleName
            Write-Host "Renamed rule from '$CurrentRuleName' to '$NewRuleName' (ID: $($_.Identity))"
        }
    }
    catch {
        Write-Error "Failed to rename inbox rule(s): $_"
    }
}

function Test-InboxRuleHygiene {
    <#
    .SYNOPSIS
        Verifies the hygiene of inbox rules by checking for various issues.
    
    .DESCRIPTION
        This function analyzes existing inbox rules and reports any hygiene issues found,
        such as duplicate rule names.
    
    .EXAMPLE
        Test-InboxRuleHygiene
        
        Checks all inbox rules for hygiene issues and reports any problems found.
    
    .OUTPUTS
        [PSCustomObject[]] Array of hygiene issues found, with properties:
        - IssueType: The type of hygiene issue (e.g., "DuplicateRuleName")
        - Description: Detailed description of the issue
        - AffectedRules: Array of rule names or IDs affected by the issue
    #>
    
    try {
        $rules = Get-InboxRules
        if (-not $rules) {
            Write-Warning "No inbox rules found to analyze."
            return @()
        }

        $issues = @()
        
        # Check for duplicate rule names
        $ruleGroups = $rules | Group-Object -Property Name
        $duplicateRules = $ruleGroups | Where-Object { $_.Count -gt 1 }
        
        foreach ($duplicate in $duplicateRules) {
            $issues += [PSCustomObject]@{
                IssueType = "DuplicateRuleName"
                Description = "Found $($duplicate.Count) rules with the same name: '$($duplicate.Name)'"
                AffectedRules = $duplicate.Group | ForEach-Object { 
                    [PSCustomObject]@{
                        Name = $_.Name
                        Identity = $_.Identity
                        FromAddress = $_.FromAddressContainsWords
                        TargetFolder = $_.MoveToFolder
                    }
                }
            }
        }
        
        # Return results
        if ($issues.Count -eq 0) {
            Write-Host "No rule hygiene issues found."
        } else {
            Write-Warning "Found $($issues.Count) rule hygiene issue(s)."
            foreach ($issue in $issues) {
                Write-Host "`nIssue: $($issue.Description)"
                Write-Host "Affected Rules:"
                $issue.AffectedRules | ForEach-Object {                    
                    Write-Host "  - Rule ID: $($_.Identity)"
                    Write-Host "    Name: $($_.Name)"
                    $description = $_ | Get-InboxRuleDescription
                    Write-Host "    Description: $description"
                }
            }
        }
        
        return $issues
    }
    catch {
        Write-Error "Failed to analyze inbox rules: $_"
        return $null
    }
}

function Clean-ExchangeAddress {
    param(
        [Parameter(ValueFromPipeline=$true)]
        [string]$Address
    )
    process {
        if (-not $Address) { return $null }
        
        # Extract display name if present in format "Display Name <email@domain.com>"
        if ($Address -match '^"?([^"]+)"?\s*\[.+\]$') {
            return $Matches[1].Trim()
        }
        # If it's just an Exchange URL without display name, return null
        elseif ($Address -match '^\[.+\]$') {
            return $null
        }
        # Return the original address if no special formatting needed
        else {
            return $Address
        }
    }
}

function Get-InboxRuleDescription {
    <#
    .SYNOPSIS
        Generates a friendly description for an inbox rule.
    
    .DESCRIPTION
        This function analyzes an inbox rule's properties and generates a human-readable
        description of what the rule does.
    
    .PARAMETER Rule
        The inbox rule object to generate a description for.
    
    .EXAMPLE
        Get-InboxRule | Get-InboxRuleDescription
        
        Generates descriptions for all inbox rules.
    
    .OUTPUTS
        [string] A friendly description of the rule's function.
    #>
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [PSObject]$Rule
    )
    
    process {
        try {
            $description = ""
            
            # Handle rules based on sender
            $fromAddresses = @()
            if ($Rule.From) {
                $fromAddresses += $Rule.From
            }
            if ($Rule.FromAddressContainsWords) {
                $fromAddresses += $Rule.FromAddressContainsWords
            }
            if ($fromAddresses) {
                $senders = $fromAddresses | 
                    ForEach-Object { Clean-ExchangeAddress $_ } | 
                    Where-Object { $_ } | 
                    Select-Object -Unique | 
                    Join-String -Separator ", "
                if ($senders) {
                    $description += "From $senders "
                }
            }
            
            # Handle rules based on recipient
            $toAddresses = @()
            if ($Rule.SentTo) {
                $toAddresses += $Rule.SentTo
            }
            if ($Rule.SentToAddressContainsWords) {
                $toAddresses += $Rule.SentToAddressContainsWords
            }
            if ($toAddresses) {
                $recipients = $toAddresses | 
                    ForEach-Object { Clean-ExchangeAddress $_ } | 
                    Where-Object { $_ } | 
                    Select-Object -Unique | 
                    Join-String -Separator ", "
                if ($recipients) {
                    $description += "To $recipients "
                }
            }
            
            # Add destination folder if present
            if ($Rule.MoveToFolder) {
                # Clean up folder path for display
                $folderName = $Rule.MoveToFolder -replace '^:\\?', ''
                $description += "> $folderName "
            }
            
            # Handle other common rule actions
            if ($Rule.DeleteMessage) {
                $description += "> Delete "
            }
            if ($Rule.MarkAsRead) {
                $description += "> Mark as Read "
            }
            if ($Rule.MarkImportance -eq "High") {
                $description += "> Mark Important "
            }
            if ($Rule.FlagMessage) {
                $description += "> Flag "
            }
            
            # Add disabled status if applicable
            if (-not $Rule.Enabled) {
                $description += "(Disabled) "
            }
            
            return $description.Trim()
        }
        catch {
            Write-Error "Failed to generate rule description: $_"
            return "Description unavailable"
        }
    }
}

function Get-InboxRuleDetails {
    <#
    .SYNOPSIS
        Gets inbox rules with only non-empty properties.
    
    .DESCRIPTION
        This function retrieves inbox rules and displays only the properties that have values,
        skipping empty, null, or default values for cleaner output.
    
    .EXAMPLE
        Get-InboxRuleDetails
        
        Shows all inbox rules with only their non-empty properties.
    
    .EXAMPLE
        Get-InboxRules | Where-Object { $_.Name -like "*Project*" } | Get-InboxRuleDetails
        
        Shows non-empty properties for rules with "Project" in their name.
    #>    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true)]
        [PSObject[]]$Rules,
        
        [Parameter()]
        [string]$OutputPath,
        
        [Parameter()]
        [ValidateSet('JSON', 'CSV')]
        [string]$OutputFormat = 'JSON'
    )
    
    begin {
        if (-not $Rules) {
            $Rules = Get-InboxRules
        }
        $allRules = @()
    }
    
    process {
        $Rules | Select-Object -Property * | Where-Object {$_.PSObject.Properties} | ForEach-Object { 
            $obj = $_ 
            Write-Host "`nRule: $($obj.Name)" -ForegroundColor Yellow
            
            # Create a cleaned up object with non-empty properties
            $cleanedRule = [ordered]@{
                Name = $obj.Name
                ShortDescription = ($obj | Get-InboxRuleDescription)
            }
            
            $obj.PSObject.Properties | 
                Where-Object { 
                    $null -ne $_.Value -and           # Skip null values
                    $_.Value -ne '' -and              # Skip empty strings
                    $_.Value -ne @() -and             # Skip empty arrays
                    (-not $_.Value -is [Boolean] -or $_.Value -eq $true) -and  # Only show true booleans
                    $_.Name -ne 'ObjectState' -and    # Skip internal properties
                    $_.Name -ne 'Name'                # Skip Name as it's already added
                } | ForEach-Object {
                    $cleanedRule[$_.Name] = $_.Value
                }
            
            # Display the properties
            $cleanedRule.GetEnumerator() | Format-Table -AutoSize -Wrap
            
            # Add to collection for output
            $allRules += [PSCustomObject]$cleanedRule
        }
    }
    
    end {
        if ($OutputPath) {
            try {
                $parentFolder = Split-Path -Path $OutputPath -Parent
                if ($parentFolder -and -not (Test-Path -Path $parentFolder)) {
                    New-Item -ItemType Directory -Path $parentFolder -Force | Out-Null
                }
                
                switch ($OutputFormat) {
                    'JSON' {
                        $allRules | ConvertTo-Json -Depth 10 | Set-Content -Path $OutputPath
                        Write-Host "Rules exported to JSON file: $OutputPath" -ForegroundColor Green
                    }
                    'CSV' {
                        $allRules | Export-Csv -Path $OutputPath -NoTypeInformation
                        Write-Host "Rules exported to CSV file: $OutputPath" -ForegroundColor Green
                    }
                }
            }
            catch {
                Write-Error "Failed to export rules to file: $_"
            }
        }
          return $allRules
    }
}




