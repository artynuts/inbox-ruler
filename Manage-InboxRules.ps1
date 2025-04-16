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
        [string]$NewRuleName
    )
      try {
        Set-InboxRule -Identity $CurrentRuleName -Name $NewRuleName
        Write-Host "Renamed rule from '$CurrentRuleName' to '$NewRuleName'"
    }
    catch {
        Write-Error "Failed to rename inbox rule: $_"
    }
}




