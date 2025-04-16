# Example usage of Manage-InboxRules.ps1
. .\Manage-InboxRules.ps1

# First, connect to Exchange Online
#Connect-ToExchange

# List all current inbox rules
#Write-Host "Current inbox rules:"
#Get-InboxRules | Format-Table Name, Description, Enabled

# Example: Create a rule to move emails from a specific sender to a folder
# Note: Folder paths should start with 'Inbox\'
New-CustomInboxRule -RuleName "Test Newsletter Filter" `
    -FromAddress "newsletter@example.com" `
    -TargetFolder ":\Inbox\TestNewsletters"

# Example: Create another rule with a subfolder
New-CustomInboxRule -RuleName "Test Project Updates" `
    -FromAddress "updates@company.com" `
    -TargetFolder "Inbox\TestWork\Projects"

New-CustomInboxRule -RuleName "Test Project Updates 1" `
    -FromAddress "updates@company.com" `
    -TargetFolder "Inbox\TestWork\Projects"

New-CustomInboxRule -RuleName "Test Project Updates 2" `
    -FromAddress "updates@company.com" `
    -TargetFolder "Inbox\TestWork\Project"

New-CustomInboxRule -RuleName "Test Project Updates 7" `
    -FromAddress "updates@company.com" `
    -TargetFolder "Inbox\TestWork\Project 7"

# Example: Rename a rule
Rename-CustomInboxRule -CurrentRuleName "Test Project Updates" -NewRuleName "Test Project Updates Renamed"
Rename-CustomInboxRule -CurrentRuleName "Foobar" -NewRuleName "Foobar Renamed"
Rename-CustomInboxRule -CurrentRuleName "Test Project Updates 3 Renamed" -NewRuleName "Test Project Updates 3 Renamed Again"
Rename-CustomInboxRule -CurrentRuleName "Test Project Updates 2" -NewRuleName "Test Project Updates 1"

# Example: Remove rules (commented out for safety)
Remove-CustomInboxRule -RuleName "Test Newsletter Filter" -RemoveAll
Remove-CustomInboxRule -RuleName "Foobar"
#Remove-CustomInboxRule -RuleName "Important Alerts"
