# Get EXO Mailbox Delegate Permissions

## Description

 This PowerShell script gathers information about mailbox delegate access, inbox folder permissions, and send permissions for a list of mailboxes. The script supports various input methods, including providing mailbox identities, mailbox objects, a list of mailbox identities from a file, or limiting the query to a specific number of mailboxes.

## Parameters

```Text
PARAMETERS
    -MailboxID <String[]>
        Specifies an array of mailbox identities (PrimarySmtpAddress) for which to retrieve delegate permissions and access rights.

        Required?                    true
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -MailboxObject <Object[]>
        Specifies an array of mailbox objects resulting from the Get-Mailbox command. The script will retrieve delegate permissions and access rights for these mailbox objects.

        Required?                    true
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -MailboxListFile <String>
        Specifies the path to a file containing a list of mailbox identities (PrimarySmtpAddress). The script will retrieve delegate permissions and access rights for mailboxes listed in the file.

        Required?                    true
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -Limit <Object>
        Specifies the maximum number of mailboxes to process. This parameter is used when retrieving the top N mailboxes.

        Required?                    false
        Position?                    named
        Default value                10
        Accept pipeline input?       false
        Accept wildcard characters?  false
```

## Syntax

```Text
SYNTAX
    .\Get-MailboxDelegatePermission.ps1 [-Limit <Object>] [<CommonParameters>]

    .\Get-MailboxDelegatePermission.ps1 -MailboxID <String[]> [<CommonParameters>]

    .\Get-MailboxDelegatePermission.ps1 -MailboxObject <Object[]> [<CommonParameters>]

    .\Get-MailboxDelegatePermission.ps1 -MailboxListFile <String> [<CommonParameters>]
```

## Examples

```PowerShell
# This example retrieves the delegate access list from the provided mailbox identities.

$mailboxList = @(
    'mailbox1@domain.tld',
    'mailbox2@domain.tld',
    'mailbox3@domain.tld'
)
.\Get-MailboxDelegatePermission.ps1 -MailboxID $mailboxList
```

```PowerShell
# This example retrieves the delegate access list from the provided mailbox objects resulting from the Get-Mailbox command.

$mailboxList = Get-Mailbox -ResultSize 2 -RecipientTypeDetails SharedMailbox -WarningAction SilentlyContinue
.\Get-MailboxDelegatePermission.ps1 -MailboxObject $mailboxList
```

```PowerShell
# This example retrieves the delegate access list from the provided mailbox list file.

.\Get-MailboxDelegatePermission.ps1 -MailboxListFile .\mailboxList.txt
```

```PowerShell
# This example retrieves the delegate access list from the top N or All mailboxes.

\Get-MailboxDelegatePermission.ps1 -Limit 10
\Get-MailboxDelegatePermission.ps1 -Limit All
```
