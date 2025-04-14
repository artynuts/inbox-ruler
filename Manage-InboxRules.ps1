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
    param(
        [Parameter(Mandatory=$false)]
        [string]$Mailbox = $null
    )
    
    try {
        if ([string]::IsNullOrEmpty($Mailbox)) {
            $Mailbox = (Get-AcceptedDomain | Where-Object {$_.Default -eq $true}).DomainName
            $Mailbox = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -split '\\')[1] + "@" + $Mailbox
        }
        
        $rules = Get-InboxRule -Mailbox $Mailbox
        return $rules
    }
    catch {
        Write-Error "Failed to get inbox rules: $_"
        return $null
    }
}

function New-CustomInboxRule {
    param(
        [Parameter(Mandatory=$true)]
        [string]$RuleName,
        
        [Parameter(Mandatory=$true)]
        [string]$FromAddress,
        
        [Parameter(Mandatory=$true)]
        [string]$TargetFolder,
        
        [Parameter(Mandatory=$false)]
        [string]$Mailbox = $null
    )
    
    try {
        if ([string]::IsNullOrEmpty($Mailbox)) {
            $Mailbox = (Get-AcceptedDomain | Where-Object {$_.Default -eq $true}).DomainName
            $Mailbox = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -split '\\')[1] + "@" + $Mailbox
        }
        
        New-InboxRule -Mailbox $Mailbox -Name $RuleName -FromAddressContainsWords $FromAddress -MoveToFolder $TargetFolder
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
        [string]$Mailbox = $null
    )
    
    try {
        if ([string]::IsNullOrEmpty($Mailbox)) {
            $Mailbox = (Get-AcceptedDomain | Where-Object {$_.Default -eq $true}).DomainName
            $Mailbox = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -split '\\')[1] + "@" + $Mailbox
        }
        
        Remove-InboxRule -Mailbox $Mailbox -Identity $RuleName -Confirm:$false
        Write-Host "Removed rule: $RuleName"
    }
    catch {
        Write-Error "Failed to remove inbox rule: $_"
    }
}
