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
    )    try {
        
        # Ensure the folder path is in the correct format
        if (-not $FolderPath.StartsWith(':\')) {
            $FolderPath = ":\$($FolderPath.TrimStart('\'))"
        }

        # Split the path into segments and ensure each segment exists
        $segments = $FolderPath.Split('\')
        $currentPath = $segments[0] # Start with ":"
        
        # Create each segment of the path if it doesn't exist
        for ($i = 1; $i -lt $segments.Count; $i++) {
            $parentPath = $currentPath
            $currentPath = "$currentPath\$($segments[$i])"
            
            try {
                $null = Get-MailboxFolder -Identity $currentPath -ErrorAction Stop
                Write-Host "Folder exists: $currentPath"
            }
            catch {
                Write-Host "Creating folder: $currentPath"                  
                try {
                    New-MailboxFolder -Parent $parentPath -Name $segments[$i]
                    Wait-MailboxFolderCreation -FolderPath $currentPath
                }
                catch {
                    throw "Failed to create folder '$currentPath': $_"
                }
            }
        }
        return $FolderPath
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
        $TargetFolder = New-MailboxFolderHierarchy -FolderPath $TargetFolder

        # Create the inbox rule
        Write-Host "Creating new inbox rule: $RuleName"
        Write-Host "moving emails from: $FromAddress to folder: $TargetFolder"
        New-InboxRule -Name $RuleName -FromAddressContainsWords $FromAddress -MoveToFolder $TargetFolder
        Write-Host "Created new rule: $RuleName"
    }
    catch {
        Write-Error "Failed to create inbox rule: $_"
    }
}

function Remove-CustomInboxRule {
    param(
        [Parameter(Mandatory=$true)]
        [string]$RuleName
    )
      try {
        Remove-InboxRule -Identity $RuleName -Confirm:$false
        Write-Host "Removed rule: $RuleName"
    }
    catch {
        Write-Error "Failed to remove inbox rule: $_"
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

function Wait-MailboxFolderCreation {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath,
        
        [Parameter(Mandatory=$false)]
        [int]$TimeoutSeconds = 30
    )
    
    try {
        $timeout = [DateTime]::Now.AddSeconds($TimeoutSeconds)
        $folderExists = $false
        
        while (-not $folderExists -and [DateTime]::Now -lt $timeout) {
            try {
                $null = Get-MailboxFolder -Identity $FolderPath -ErrorAction Stop
                $folderExists = $true
                Write-Host "Confirmed folder creation: $FolderPath"
            }
            catch {
                Start-Sleep -Seconds 1
            }
        }
        
        if (-not $folderExists) {
            throw "Timeout waiting for folder '$FolderPath' to be created"
        }
        
        return $true
    }
    catch {
        throw "Failed to confirm folder creation '$FolderPath': $_"
    }
}


