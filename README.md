# Get-MailboxReport.ps1 Forked
PowerShell script to generate a report of mailboxes, including information such as item count, total size, and other useful attributes. Color code for users who have exceeded quotas.

## SYNOPSIS
Get-MailboxReport.ps1 - Mailbox report generation script.

Generates a report of useful information for the specified server, database, mailbox or list of mailboxes. Use only one parameter at a time depending on the scope of your mailbox report.

Single mailbox reports are output to the console, while all other reports are output to email with -SendEmail parameter

## Parameters
- -All
Generates a report for all mailboxes in the organization.

- -Server
Generates a report for all mailboxes on the specified server.

- -Database
Generates a report for all mailboxes on the specified database.

- -File
Generates a report for mailbox names listed in the specified text file.

- -Mailbox
Generates a report only for the specified mailbox.

- -SendEmail
Specifies that an email report with the CSV file attached should be sent.

- -MailFrom
The SMTP address to send the email from.

- -MailTo
The SMTP address to send the email to.

- -MailServer
The SMTP server to send the email through.

## Usage examples
> .\Get-MailboxReport.ps1 -Database DB01

Returns a report with the mailbox statistics for all mailbox users in database DB01

> .\Get-MailboxReport.ps1 -All -SendEmail -MailFrom exchangereports@domain.local -MailTo it@domain.local -MailServer smtp.domain.local

Returns a report with the mailbox statistics for all mailbox users and send an email report to the specified recipient.


## Source
Written By: Paul Cunningham ( http://exchangeserverpro.com/powershell-script-create-mailbox-size-report-exchange-server-2010 )
