
<#PSScriptInfo

.VERSION 1.0

.GUID 558e9982-9a8c-43d6-8b8d-ce00c90a9a4f

.AUTHOR tcastillotej

.COMPANYNAME

.COPYRIGHT

.TAGS

.LICENSEURI

.PROJECTURI
    https://github.com/junecastillote/Get-EXO-Mailbox-Delegate-Permissions

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<#

.SYNOPSIS
    PowerShell script to retrieve mailbox and inbox access rights as well as send permissions for specified mailboxes.

.DESCRIPTION
    This PowerShell script gathers information about mailbox delegate access, inbox folder permissions, and send permissions for a list of mailboxes. The script supports various input methods, including providing mailbox identities, mailbox objects, a list of mailbox identities from a file, or limiting the query to a specific number of mailboxes.

.PARAMETER MailboxID
    Specifies an array of mailbox identities (PrimarySmtpAddress) for which to retrieve delegate permissions and access rights.

.PARAMETER MailboxObject
    Specifies an array of mailbox objects resulting from the Get-Mailbox command. The script will retrieve delegate permissions and access rights for these mailbox objects.

.PARAMETER MailboxListFile
    Specifies the path to a file containing a list of mailbox identities (PrimarySmtpAddress). The script will retrieve delegate permissions and access rights for mailboxes listed in the file.

.PARAMETER Limit
    Specifies the maximum number of mailboxes to process. This parameter is used when retrieving the top N mailboxes.

.EXAMPLE

    # This example retrieves the delegate access list from the provided mailbox identities.

    PS > $mailboxList = @(
        'mailbox1@domain.tld',
        'mailbox2@domain.tld',
        'mailbox3@domain.tld'
    )
    PS > .\Get-MailboxDelegatePermission.ps1 -MailboxID $mailboxList

.EXAMPLE

    # This example retrieves the delegate access list from the provided mailbox objects resulting from the Get-Mailbox command.

    PS > $mailboxList = Get-Mailbox -ResultSize 2 -RecipientTypeDetails SharedMailbox -WarningAction SilentlyContinue
    PS > .\Get-MailboxDelegatePermission.ps1 -MailboxObject $mailboxList

.EXAMPLE

    # This example retrieves the delegate access list from the provided mailbox list file.

    PS > .\Get-MailboxDelegatePermission.ps1 -MailboxListFile .\mailboxList.txt

.EXAMPLE

    # This example retrieves the delegate access list from the top N or All mailboxes.
    PS > .\Get-MailboxDelegatePermission.ps1 -Limit 10
    PS > .\Get-MailboxDelegatePermission.ps1 -Limit All

#>

[CmdletBinding(DefaultParameterSetName = 'byLimit')]
param (
    [Parameter(
        Mandatory,
        ParameterSetName = 'byMailboxId'
    )]
    [String[]]
    $MailboxID,

    [Parameter(
        Mandatory,
        ParameterSetName = 'byMailboxObject'
    )]
    [System.Object[]]
    $MailboxObject,

    [Parameter(
        Mandatory,
        ParameterSetName = 'byMailboxListFile'
    )]
    [string]
    $MailboxListFile,

    [Parameter(
        ParameterSetName = 'byLimit'
    )]
    $Limit = 10
)

if ($PSCmdlet.ParameterSetName -eq 'byMailboxId') {
    $mailbox = @(
        $MailboxID | ForEach-Object {
            Get-Mailbox $_ -ErrorAction SilentlyContinue
        }
    )
}

if ($PSCmdlet.ParameterSetName -eq 'byMailboxListFile') {
    if (!(Test-Path -Path $MailboxListFile)) {
        "The specified file path does not exist." | Out-Default
        return $null
    }

    $mailbox = @(
        Get-Content -Path $MailboxListFile | ForEach-Object {
            Get-Mailbox $_ -ErrorAction SilentlyContinue
        }
    )
}

if ($PSCmdlet.ParameterSetName -eq 'byMailboxObject') {
    $mailbox = @($MailboxObject)
}

if ($PSCmdlet.ParameterSetName -eq 'byLimit') {
    try {
        $mailbox = @(Get-Mailbox -ResultSize $Limit -ErrorAction Stop -WarningAction SilentlyContinue)
    }
    catch {
        $_.Exception.Message | Out-Default
        return $null
    }
}


if (!$mailbox) {
    return $null
}

$result = [System.Collections.ArrayList]@()
for ($i = 0; $i -lt $mailbox.count; $i++) {
    $currentMailbox = $mailbox[$i]
    "Processing mailbox $($i+1)/$($mailbox.count): [$($currentMailbox.PrimarySmtpAddress)]" | Out-Default

    $sendAsPermissions = @(Get-RecipientPermission -Identity $currentMailbox.ExchangeGuid -ResultSize Unlimited | Where-Object { $_.AccessControlType -eq 'Allow' -and $_.Trustee -like "*@*" })
    $mailboxPermissions = @(Get-MailboxPermission -Identity $currentMailbox.ExchangeGuid -ResultSize Unlimited | Where-Object { $_.User -like "*@*" -and !$_.Deny -and !$_.IsInherited })

    try {
        if (!($inboxFolderPermissions = @(Get-MailboxFolderPermission -Identity "$($currentMailbox.ExchangeGuid):\Inbox" -ErrorAction SilentlyContinue))) {
            $inboxFolder = (((Get-MailboxFolderStatistics -Identity $currentMailbox.ExchangeGuid) | Where-Object { $_.FolderType -eq 'Inbox' }).FolderPath -replace '/', '')
            "     -> Non-English Mailbox. Inbox folder name is - [\$inboxFolder]" | Out-Default
            $inboxFolderPermissions = Get-MailboxFolderPermission -Identity "$($currentMailbox.ExchangeGuid):\$($inboxFolder)" -ResultSize Unlimited -ErrorAction Stop | Where-Object { $_.User.DisplayName -ne 'Anonymous' -and $_.User.DisplayName -ne 'Default' -and $_.IsValid -eq $true }
            if ($inboxFolderPermissions.Count -gt 0) {
                $inboxFolderPermissions | Add-Member -Name UserPrincipalName -MemberType NoteProperty -Value ''
                $inboxFolderPermissions | ForEach-Object {
                    $_.UserPrincipalName = (Get-User $_.User -ErrorAction SilentlyContinue).UserPrincipalName
                }
            }
        }
    }
    catch {
        $_.Exception.Message | Out-Default
    }

    $sendOnBehalfPermissions = @(
        if ($currentMailbox.GrantSendOnBehalfTo.Count -gt 0) {
            foreach ($item in $currentMailbox.GrantSendOnBehalfTo) {
                (Get-User -Identity $item).UserPrincipalName
            }
        }
    )

    $delegatesList = [System.Collections.ArrayList]@()
    $delegatesList.AddRange(@(($sendAsPermissions).Trustee))
    $delegatesList.AddRange(@(($mailboxPermissions).User))
    $delegatesList.AddRange(@(($inboxFolderPermissions).UserPrincipalName))
    $delegatesList.AddRange(@($sendOnBehalfPermissions))

    if (!$delegatesList) {
        "     -> Delegates: $($delegatesList.Count)" | Out-Default
    }

    if ($delegatesList) {
        $delegatesList = $delegatesList | Sort-Object | Select-Object -Unique
        "     -> Delegates: $($delegatesList.Count)" | Out-Default
        $delegatesList | ForEach-Object {
            $currentDelegate = $_

            $hasSenderPermission = @()
            if (@($sendAsPermissions | Where-Object { $_.Trustee -eq $currentDelegate })) {
                $hasSenderPermission += "SendAs"
            }

            if (@($sendOnBehalfPermissions | Where-Object { $_ -eq $currentDelegate })) {
                $hasSenderPermission += "SendOnBehalf"
            }

            $hasInboxFolderPermission = @($inboxFolderPermissions | Where-Object { $_.UserPrincipalName -eq $currentDelegate })
            $hasMailboxPermission = @($mailboxPermissions | Where-Object { $_.User -eq $currentDelegate })


            $null = $result.Add(
                $(
                    New-Object psobject -Property $(
                        [ordered]@{
                            MailboxUPN        = $currentMailbox.UserPrincipalName
                            DelegateUPN       = $currentDelegate
                            SenderPermission  = $(
                                if ($hasSenderPermission) {
                                    $hasSenderPermission -join ","
                                }
                            )
                            InboxPermission   = $(
                                if ($hasInboxFolderPermission) {
                                    $hasInboxFolderPermission.AccessRights -join ","
                                }
                            )
                            MailboxPermission = $(
                                if ($hasMailboxPermission) {
                                    $hasMailboxPermission.AccessRights -join ","
                                }
                            )
                        }
                    )
                )
            )
        }
    }
}

return $result