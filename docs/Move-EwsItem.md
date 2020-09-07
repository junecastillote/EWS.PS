# Move-EwsItem

This function uses EWS calls to do the following:
- Move all mailbox items from one folder to another
- Move mailbox items between dates from one folder to another

## Syntax

```PowerShell
Move-EwsItem -Token <AuthenticationResult> -MailboxAddress <string> -SourceFolder <Object> -TargetFolder <Object> [-TestMode <bool>] [<CommonParameters>]
```

```PowerShell
Move-EwsItem -Token <AuthenticationResult> -MailboxAddress <string> -SourceFolder <Object> -TargetFolder <Object> -StartDate <datetime> -EndDate <datetime> [-TestMode <bool>] [<CommonParameters>]
```

## Parameters

**`-Token`**

The Token parameter specifies the OAuth Access Token acquired from the Exchange Online API endpoint. You should use the `Get-MsalToken` cmdlet that is part of the MSAL.PS module to generate this token.

|   |   |
|---|---|
| Type: | Microsoft.Identity.Client.AuthenticationResult |
| Position: | Named |
| Default value : | None |
| Accept pipeline input: | False |
| Accept wildcard characters: | False |

**`-MailboxAddress`**

The MailboxAddress parameter specifies the SMTP email address of the mailbox being impersonated. For example: june@poshlab.ga

|   |   |
|---|---|
| Type: | String |
| Position: | Named |
| Default value : | None |
| Accept pipeline input: | False |
| Accept wildcard characters: | False |

**`-SourceFolder`**

The SourceFolder parameter specifies source folder object of the mailbox where the items to be moved are located. Use the [`Get-EwsFolder`](docs/Get-EwsFolder.md) cmdlet to get this object.

|   |   |
|---|---|
| Type: | Object |
| Position: | Named |
| Default value : | None |
| Accept pipeline input: | False |
| Accept wildcard characters: | False |

**`-TargetFolder`**

The TargetFolder parameter specifies target folder object of the mailbox where the items will be moved to. Use the [`Get-EwsFolder`](docs/Get-EwsFolder.md) cmdlet to get this object.

|   |   |
|---|---|
| Type: | Object |
| Position: | Named |
| Default value : | None |
| Accept pipeline input: | False |
| Accept wildcard characters: | False |

**`-StartDate`**

The StartDate parameter specifies date of the oldest items to move. This parameter must be used together with the `-EndDate` parameter.

|   |   |
|---|---|
| Type: | DateTime |
| Position: | Named |
| Default value : | None |
| Accept pipeline input: | False |
| Accept wildcard characters: | False |

**`-EndDate`**

The EndDate parameter specifies date of the most recent items to move. This parameter must be used together with the `-StartDate` parameter.

|   |   |
|---|---|
| Type: | DateTime |
| Position: | Named |
| Default value : | None |
| Accept pipeline input: | False |
| Accept wildcard characters: | False |

**`-TestMode`**

The TestMode parameter specifies whether the function should actually move the items or only simulate the move and display the output on the screen. If not used, the default value is `True` and the command will run in test mode only.

|   |   |
|---|---|
| Type: | DateTime |
| Position: | Named |
| Default value : | True |
| Accept pipeline input: | False |
| Accept wildcard characters: | False |

## Usage Examples

### Access Token Requirement

Make sure to acquire an access token first. Use the `Get-MsalToken` cmdlet.

```PowerShell
# Get MSAL Token using CLIENT ID,  CLIENT SECRET, and TENANT ID
$msalParams = @{
    ClientId = 'CLIENT ID'
    ClientSecret = (ConvertTo-SecureString 'CLIENT SECRET' -AsPlainText -Force)
    TenantId = 'TENANT ID'
    Scopes   = "https://outlook.office.com/.default"
}
$token = Get-MsalToken @msalParams
```

### Example 1: Move All Items from One Folder to Another

This example moves all messages from the primary mailbox Inbox folder to the archive mailbox Inbox folder.

```PowerShell
# Mailbox SMTP Address of the user to impersonate
$mailbox = 'june@poshlab.ga'

# Get the source folder object
$SourceFolder = Get-EwsFolder -Token $token -MailboxAddress $mailbox -MailboxType Primary -FolderName Inbox
$TargetFolder = Get-EwsFolder -Token $token -MailboxAddress $mailbox -MailboxType Archive -FolderName Inbox

# Move messages from $SourceFolder to $TargetFolder
Move-EwsItem -Token $token -MailboxAddress $mailbox -SourceFolder $SourceFolder -TargetFolder $TargetFolder -TestMode $false
```

### Example 2: Move Items Received Between Specified Dates

This example moves items that were received in the last 3 days

```PowerShell
# Mailbox SMTP Address of the user to impersonate
$mailbox = 'june@poshlab.ga'

# Get the source folder object
$SourceFolder = Get-EwsFolder -Token $token -MailboxAddress $mailbox -MailboxType Primary -FolderName Inbox
$TargetFolder = Get-EwsFolder -Token $token -MailboxAddress $mailbox -MailboxType Archive -FolderName Inbox

# Move messages from $SourceFolder to $TargetFolder
Move-EwsItem -Token $token -MailboxAddress $mailbox -SourceFolder $SourceFolder -TargetFolder $TargetFolder -StartDate (Get-Date).AddDays(-3) -EndDate (Get-Date)
```