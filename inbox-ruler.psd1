@{
    ModuleVersion = '0.1.0'
    GUID = [System.Guid]::NewGuid().ToString()
    Author = 'artynuts'
    Description = 'A PowerShell module to manage Exchange/Outlook inbox rules programmatically'
    PowerShellVersion = '5.1'
    RequiredModules = @('ExchangeOnlineManagement')
    FunctionsToExport = @('Connect-ToExchange', 'Get-InboxRules', 'New-CustomInboxRule', 'Remove-CustomInboxRule')
    CmdletsToExport = @()
    VariablesToExport = '*'
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('Exchange', 'Outlook', 'InboxRules', 'Email')
            LicenseUri = 'https://github.com/artynuts/inbox-ruler/blob/main/LICENSE'
            ProjectUri = 'https://github.com/artynuts/inbox-ruler'
            ReleaseNotes = 'Initial release of inbox-ruler module'
        }
    }
}
