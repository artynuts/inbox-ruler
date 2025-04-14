# Example usage of Manage-InboxRules.ps1
. .\Manage-InboxRules.ps1

# First, connect to Exchange Online
Connect-ToExchange

# List all current inbox rules
Write-Host "Current inbox rules:"
Get-InboxRules | Format-Table Name, Description, Enabled

# Example: Create a rule to move emails from a specific sender to a folder
New-CustomInboxRule -RuleName "Newsletter Filter" `
    -FromAddress "newsletter@example.com" `
    -TargetFolder "Inbox\Newsletters"

# Example: Remove a rule
# Remove-CustomInboxRule -RuleName "Newsletter Filter"
