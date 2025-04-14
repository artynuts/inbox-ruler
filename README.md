# Inbox-Ruler

A PowerShell module to manage Exchange/Outlook inbox rules programmatically. This module allows you to create, retrieve, and remove inbox rules for Exchange Online mailboxes.

## Prerequisites

- PowerShell 5.1 or later
- Exchange Online Management module (automatically installed if not present)

## Installation

```powershell
# Install from PowerShell Gallery (Once published)
Install-Module -Name inbox-ruler -Scope CurrentUser
```

## Usage

```powershell
# Connect to Exchange Online
Connect-ToExchange

# Get all inbox rules for current user
Get-InboxRules

# Create a new rule
New-CustomInboxRule -RuleName "Newsletter" -FromAddress "newsletter@example.com" -TargetFolder "Inbox\Newsletters"

# Remove a rule
Remove-CustomInboxRule -RuleName "Newsletter"
```

## Functions

- `Connect-ToExchange`: Connects to Exchange Online
- `Get-InboxRules`: Lists all inbox rules for a mailbox
- `New-CustomInboxRule`: Creates a new inbox rule
- `Remove-CustomInboxRule`: Removes an existing inbox rule
- `Rename-CustomInboxRule`: Renames an existing inbox rule

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
