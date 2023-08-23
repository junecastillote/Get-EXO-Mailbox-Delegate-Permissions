# Technical Info

This PowerShell script is designed to retrieve and analyze mailbox delegate permissions, inbox folder permissions, and send permissions for a list of mailboxes. The script supports various methods of input: specifying mailbox identities, mailbox objects, a list of mailbox identities from a file, or limiting the query to a specific number of mailboxes. Here's a breakdown of the script's functionality:

## Parameter Definitions

The script uses the param block to define its input parameters. The parameters are categorized into four parameter sets: '**byMailboxId**', '**byMailboxObject**', '**byMailboxListFile**', and '**byLimit**'. Each parameter set corresponds to a different way of providing input to the script. The script uses parameter sets to ensure that only one set of parameters can be used at a time.

## Parameter Validation and Data Retrieval

Depending on the selected parameter set, the script retrieves mailbox data for further processing. Here's how the data retrieval is done for each parameter set:

'**byMailboxId**': The script iterates through the provided $MailboxID array, retrieves mailbox information using the Get-Mailbox cmdlet for each mailbox identity, and stores the results in the $mailbox array.

'**byMailboxObject**': The script uses the provided $MailboxObject array directly as the mailbox data.

'**byMailboxListFile**': The script reads mailbox identities from the specified file path $MailboxListFile, retrieves mailbox information using the Get-Mailbox cmdlet for each identity, and stores the results in the $mailbox array.

'**byLimit**': The script retrieves mailbox information using the Get-Mailbox cmdlet with the specified $Limit parameter and stores the results in the $mailbox array.

## Processing Mailboxes

After retrieving mailbox data, the script enters a loop to process each mailbox. It performs the following tasks for each mailbox:

Outputs information about the mailbox being processed.

Retrieves permissions using various PowerShell cmdlets such as `Get-RecipientPermission`, `Get-MailboxPermission`, and `Get-MailboxFolderPermission`. These cmdlets are used to gather send permissions, mailbox permissions, and inbox folder permissions, respectively.

Handles exceptions that might occur during permission retrieval.

Collects data about send permissions, inbox folder permissions, and mailbox permissions for each mailbox and its delegates.

## Data Aggregation and Formatting

The script organizes the collected data into a structured format using custom objects. It creates an array list `$result` to store these objects. For each mailbox delegate, the script determines if they have specific permissions (e.g., **SendAs**, **SendOnBehalf**) and inbox folder access rights. The collected data is then added as properties to a custom object, which is then added to the `$result` array list.

## Final Output

Once all mailboxes have been processed, the script returns the populated `$result` array list. This array list contains custom objects for each mailbox delegate, each object representing the mailbox delegate's access and permissions information.
