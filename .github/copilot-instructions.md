# GitHub Copilot Instructions for inbox-ruler

This document provides context and guidelines for GitHub Copilot when generating suggestions for the inbox-ruler PowerShell module.

## Project Context

inbox-ruler is a PowerShell module for managing Exchange/Outlook inbox rules programmatically. The module handles:

- Creating and managing Exchange Online inbox rules
- Folder hierarchy management in Exchange mailboxes
- Error handling and validation for Exchange operations

## Code Style Guidelines

### Editor Conventions

- When making file edits, preserve the original whitespace and newline formatting, especially around block comments and structural elements.
- For block comments, always include a newline after the opening comment marker and before the code to maintain readability.

### PowerShell Conventions

- Use PascalCase for function names (e.g., `New-CustomInboxRule`)
- Use approved PowerShell verbs (New, Get, Set, Remove)
- Include parameter validation attributes
- Include comment-based help for all public functions
- Use try-catch blocks for error handling
- Write verbose output for debugging

### Function Structure

- Each function should have a single responsibility
- Include mandatory parameter validation
- Follow the PowerShell pipeline pattern where appropriate
- Include proper error messages with Write-Error
- Return appropriate objects rather than raw output

### Testing

- All public functions should have corresponding Pester tests
- Tests should cover both success and failure scenarios
- Mock Exchange Online cmdlets in tests
- Use BeforeAll blocks for test setup

## Project-Specific Patterns

### Error Handling

```powershell
try {
    # Operation code
    Write-Host "Operation succeeded"
}
catch {
    Write-Error "Operation failed: $_"
}
```

### Parameter Declaration

```powershell
param(
    [Parameter(Mandatory=$true)]
    [string]$ParameterName
)
```

### Exchange Online Operations

- Always check for existing objects before creation
- Use proper Exchange Online cmdlets
- Handle folder paths consistently (starting with ':\')
- Include proper pipeline support

## Documentation

- Use comment-based help for all functions
- Include examples in function documentation
- Document all parameters
- Include links to relevant Exchange Online documentation

## Security Considerations

- Follow principle of least privilege
- Validate input parameters
- Use secure string for sensitive data
- Handle credentials securely

## Performance Guidelines

- Use proper filtering at the server level
- Avoid unnecessary API calls
- Implement proper connection management
- Cache results where appropriate
