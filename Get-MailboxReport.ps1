<#
.SYNOPSIS
Get-MailboxReport.ps1 - Mailbox report generation script.
.DESCRIPTION 
Generates a report of useful information for
the specified server, database, mailbox or list of mailboxes.
Use only one parameter at a time depending on the scope of
your mailbox report.
.OUTPUTS
Single mailbox reports are output to the console, while all other reports are output to email.
.PARAMETER All
Generates a report for all mailboxes in the organization.
.PARAMETER Server
Generates a report for all mailboxes on the specified server.
.PARAMETER Database
Generates a report for all mailboxes on the specified database.
.PARAMETER File
Generates a report for mailbox names listed in the specified text file.
.PARAMETER Mailbox
Generates a report only for the specified mailbox.
.PARAMETER SendEmail
Specifies that an email report with the CSV file attached should be sent.
.PARAMETER MailFrom
The SMTP address to send the email from.
.PARAMETER MailTo
The SMTP address to send the email to.
-MailServer The SMTP server to send the email through.
.EXAMPLE
.\Get-MailboxReport.ps1 -Database DB01
Returns a report with the mailbox statistics for all mailbox users in
database HO-MB-01
.EXAMPLE
.\Get-MailboxReport.ps1 -All -SendEmail -MailFrom exchangereports@domain.local -MailTo it@domain.local -MailServer smtp.domain.local
Returns a report with the mailbox statistics for all mailbox users and
sends an email report to the specified recipient.
.NOTES
Written by: Geo Holz
Find me on:
* My Blog:	https://blog.jolos.fr
Original script by Paul Cunningham ( http://paulcunningham.me)
License:
The MIT License (MIT)
Copyright (c) 2015 Paul Cunningham
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
Change Log:
V1.00, 9/5/2018 : Original script
#>

#requires -version 2

param(
	[Parameter(ParameterSetName='database')]
    [string]$Database,

	[Parameter(ParameterSetName='file')]
    [string]$File,

	[Parameter(ParameterSetName='server')]
    [string]$Server,

	[Parameter(ParameterSetName='mailbox')]
    [string]$Mailbox,

	[Parameter(ParameterSetName='all')]
    [switch]$All,

    [Parameter( Mandatory=$false)]
	[switch]$SendEmail,

	[Parameter( Mandatory=$false)]
	[string]$MailFrom,

	[Parameter( Mandatory=$false)]
	[string]$MailTo,

	[Parameter( Mandatory=$false)]
	[string]$MailServer,

    [Parameter( Mandatory=$false)]
    [int]$Top = 10

)

#...................................
# Variables
#...................................

$now = Get-Date

$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"

$reportemailsubject = "Exchange Mailbox Size Report - $now"
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$report = @()


#...................................
# Email Settings
#...................................

$smtpsettings = @{
	To =  $MailTo
	From = $MailFrom
    Subject = $reportemailsubject
	SmtpServer = $MailServer
	}

Function Set-CellColor
{   <#
    .SYNOPSIS
        Function that allows you to set individual cell colors in an HTML table
    .DESCRIPTION
        To be used inconjunction with ConvertTo-HTML this simple function allows you
        to set particular colors for cells in an HTML table.  You provide the criteria
        the script uses to make the determination if a cell should be a particular 
        color (property -gt 5, property -like "*Apple*", etc).
        
        You can add the function to your scripts, dot source it to load into your current
        PowerShell session or add it to your $Profile so it is always available.
        
        To dot source:
            .".\Set-CellColor.ps1"
            
    .PARAMETER Property
        Property, or column that you will be keying on.  
    .PARAMETER Color
        Name or 6-digit hex value of the color you want the cell to be
    .PARAMETER InputObject
        HTML you want the script to process.  This can be entered directly into the
        parameter or piped to the function.
    .PARAMETER Filter
        Specifies a query to determine if a cell should have its color changed.  $true
        results will make the color change while $false result will return nothing.
        
        Syntax
        <Property Name> <Operator> <Value>
        
        <Property Name>::= the same as $Property.  This must match exactly
        <Operator>::= "-eq" | "-le" | "-ge" | "-ne" | "-lt" | "-gt"| "-approx" | "-like" | "-notlike" 
            <JoinOperator> ::= "-and" | "-or"
            <NotOperator> ::= "-not"
        
        The script first attempts to convert the cell to a number, and if it fails it will
        cast it as a string.  So 40 will be a number and you can use -lt, -gt, etc.  But 40%
        would be cast as a string so you could only use -eq, -ne, -like, etc.  
    .PARAMETER Row
        Instructs the script to change the entire row to the specified color instead of the individual cell.
    .INPUTS
        HTML with table
    .OUTPUTS
        HTML
    .EXAMPLE
        get-process | convertto-html | set-cellcolor -Propety cpu -Color red -Filter "cpu -gt 1000" | out-file c:\test\get-process.html
        Assuming Set-CellColor has been dot sourced, run Get-Process and convert to HTML.  
        Then change the CPU cell to red only if the CPU field is greater than 1000.
        
    .EXAMPLE
        get-process | convertto-html | set-cellcolor cpu red -filter "cpu -gt 1000 -and cpu -lt 2000" | out-file c:\test\get-process.html
        
        Same as Example 1, but now we will only turn a cell red if CPU is greater than 100 
        but less than 2000.
        
    .EXAMPLE
        $HTML = $Data | sort server | ConvertTo-html -head $header | Set-CellColor cookedvalue red -Filter "cookedvalue -gt 1"
        PS C:\> $HTML = $HTML | Set-CellColor Server green -Filter "server -eq 'dc2'"
        PS C:\> $HTML | Set-CellColor Path Yellow -Filter "Path -like ""*memory*""" | Out-File c:\Test\colortest.html
        
        Takes a collection of objects in $Data, sorts on the property Server and converts to HTML.  From there 
        we set the "CookedValue" property to red if it's greater then 1.  We then send the HTML through Set-CellColor
        again, this time setting the Server cell to green if it's "dc2".  One more time through Set-CellColor
        turns the Path cell to Yellow if it contains the word "memory" in it.
        
    .EXAMPLE
        $HTML = $Data | sort server | ConvertTo-html -head $header | Set-CellColor cookedvalue red -Filter "cookedvalue -gt 1" -Row
        
        Now, if the cookedvalue property is greater than 1 the function will highlight the entire row red.
        
    .NOTES
        Author:             Martin Pugh
        Twitter:            @thesurlyadm1n
        Spiceworks:         Martin9700
        Blog:               www.thesurlyadmin.com
          
        Changelog:
            1.5             Added ability to set row color with -Row switch instead of the individual cell
            1.03            Added error message in case the $Property field cannot be found in the table header
            1.02            Added some additional text to help.  Added some error trapping around $Filter
                            creation.
            1.01            Added verbose output
            1.0             Initial Release
    .LINK
        http://community.spiceworks.com/scripts/show/2450-change-cell-color-in-html-table-with-powershell-set-cellcolor
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory,Position=0)]
        [string]$Property,
        [Parameter(Mandatory,Position=1)]
        [string]$Color,
        [Parameter(Mandatory,ValueFromPipeline)]
        [Object[]]$InputObject,
        [Parameter(Mandatory)]
        [string]$Filter,
        [switch]$Row
    )
    
    Begin {
        Write-Verbose "$(Get-Date): Function Set-CellColor begins"
        If ($Filter)
        {   If ($Filter.ToUpper().IndexOf($Property.ToUpper()) -ge 0)
            {   $Filter = $Filter.ToUpper().Replace($Property.ToUpper(),"`$Value")
                Try {
                    [scriptblock]$Filter = [scriptblock]::Create($Filter)
                }
                Catch {
                    Write-Warning "$(Get-Date): ""$Filter"" caused an error, stopping script!"
                    Write-Warning $Error[0]
                    Exit
                }
            }
            Else
            {   Write-Warning "Could not locate $Property in the Filter, which is required.  Filter: $Filter"
                Exit
            }
        }
    }
    
    Process {
        $InputObject = $InputObject -split "`r`n"
        ForEach ($Line in $InputObject)
        {   If ($Line.IndexOf("<tr><th") -ge 0)
            {   Write-Verbose "$(Get-Date): Processing headers..."
                $Search = $Line | Select-String -Pattern '<th ?[a-z\-:;"=]*>(.*?)<\/th>' -AllMatches
                $Index = 0
                ForEach ($Match in $Search.Matches)
                {   If ($Match.Groups[1].Value -eq $Property)
                    {   Break
                    }
                    $Index ++
                }
                If ($Index -eq $Search.Matches.Count)
                {   Write-Warning "$(Get-Date): Unable to locate property: $Property in table header"
                    Exit
                }
                Write-Verbose "$(Get-Date): $Property column found at index: $Index"
            }
            If ($Line -match "<tr( style=""background-color:.+?"")?><td")
            {   $Search = $Line | Select-String -Pattern '<td ?[a-z\-:;"=]*>(.*?)<\/td>' -AllMatches
                $Value = $Search.Matches[$Index].Groups[1].Value -as [double]
                If (-not $Value)
                {   $Value = $Search.Matches[$Index].Groups[1].Value
                }
                If (Invoke-Command $Filter)
                {   If ($Row)
                    {   Write-Verbose "$(Get-Date): Criteria met!  Changing row to $Color..."
                        If ($Line -match "<tr style=""background-color:(.+?)"">")
                        {   $Line = $Line -replace "<tr style=""background-color:$($Matches[1])","<tr style=""background-color:$Color"
                        }
                        Else
                        {   $Line = $Line.Replace("<tr>","<tr style=""background-color:$Color"">")
                        }
                    }
                    Else
                    {   Write-Verbose "$(Get-Date): Criteria met!  Changing cell to $Color..."
                        $Line = $Line.Replace($Search.Matches[$Index].Value,"<td style=""background-color:$Color"">$Value</td>")
                    }
                }
            }
            Write-Output $Line
        }
    }
    
    End {
        Write-Verbose "$(Get-Date): Function Set-CellColor completed"
    }
}

#...................................
# Script
#...................................

#Add dependencies
Import-Module ActiveDirectory -ErrorAction STOP


#Get the mailbox list

Write-Host -ForegroundColor White "Collecting mailbox list"

if($all) { $mailboxes = @(Get-Mailbox -resultsize unlimited -IgnoreDefaultScope) }

if($server)
{
    $databases = @(Get-MailboxDatabase -Server $server)
    $mailboxes = @($databases | Get-Mailbox -resultsize unlimited -IgnoreDefaultScope)
}

if($database){ $mailboxes = @(Get-Mailbox -database $database -resultsize unlimited -IgnoreDefaultScope) }

if($file) {	$mailboxes = @(Get-Content $file | Get-Mailbox -resultsize unlimited) }

if($mailbox) { $mailboxes = @(Get-Mailbox $mailbox) }

#Get the report

Write-Host -ForegroundColor White "Collecting report data"

$mailboxcount = $mailboxes.count
$i = 0

$mailboxdatabases = @(Get-MailboxDatabase)

#Loop through mailbox list and collect the mailbox statistics
foreach ($mb in $mailboxes)
{
	$i = $i + 1
	$pct = $i/$mailboxcount * 100
	Write-Progress -Activity "Collecting mailbox details" -Status "Processing mailbox $i of $mailboxcount - $mb" -PercentComplete $pct

	$stats = $mb | Get-MailboxStatistics | Select-Object TotalItemSize,TotalDeletedItemSize,ItemCount,LastLogonTime,LastLoggedOnUserAccount
    
    if ($mb.ArchiveDatabase)
    {
        $archivestats = $mb | Get-MailboxStatistics -Archive | Select-Object TotalItemSize,TotalDeletedItemSize,ItemCount
    }
    else
    {
        $archivestats = "n/a"
    }

    $inboxstats = Get-MailboxFolderStatistics $mb -FolderScope Inbox | Where {$_.FolderPath -eq "/Boîte de réception"}
    $sentitemsstats = Get-MailboxFolderStatistics $mb -FolderScope SentItems | Where {$_.FolderPath -eq "/Éléments envoyés"}
    $deleteditemsstats = Get-MailboxFolderStatistics $mb -FolderScope DeletedItems | Where {$_.FolderPath -eq "/Éléments supprimés"}


	$lastlogon = $stats.LastLogonTime

	$user = Get-User $mb
	$aduser = Get-ADUser $mb.samaccountname -Properties Enabled,AccountExpirationDate
    
    $primarydb = $mailboxdatabases | where {$_.Name -eq $mb.Database.Name}
    $archivedb = $mailboxdatabases | where {$_.Name -eq $mb.ArchiveDatabase.Name}

	#Create a custom PS object to aggregate the data we're interested in
	
	$userObj = New-Object PSObject
	$userObj | Add-Member NoteProperty -Name "DisplayName" -Value $mb.DisplayName
	$userObj | Add-Member NoteProperty -Name "Mailbox Type" -Value $mb.RecipientTypeDetails
	# $userObj | Add-Member NoteProperty -Name "Title" -Value $user.Title
    # $userObj | Add-Member NoteProperty -Name "Department" -Value $user.Department
    # $userObj | Add-Member NoteProperty -Name "Office" -Value $user.Office

    # $userObj | Add-Member NoteProperty -Name "TotalMailboxSize" -Value ($stats.TotalItemSize.Value.ToMB() + $stats.TotalDeletedItemSize.Value.ToMB())
	$userObj | Add-Member NoteProperty -Name "MailboxSize" -Value $stats.TotalItemSize.Value.ToMB()
	# $userObj | Add-Member NoteProperty -Name "Mailbox Recoverable Item Size (Mb)" -Value $stats.TotalDeletedItemSize.Value.ToMB()
	$userObj | Add-Member NoteProperty -Name "Mailbox Items" -Value $stats.ItemCount

    $userObj | Add-Member NoteProperty -Name "Inbox Size" -Value $inboxstats.FolderandSubFolderSize.ToMB()
    $userObj | Add-Member NoteProperty -Name "Sent Folder Size" -Value $sentitemsstats.FolderandSubFolderSize.ToMB()
    $userObj | Add-Member NoteProperty -Name "Deleted Items Size" -Value $deleteditemsstats.FolderandSubFolderSize.ToMB()

    # if ($archivestats -eq "n/a")
    # {
        # $userObj | Add-Member NoteProperty -Name "Total Archive Size (Mb)" -Value "n/a"
	    # $userObj | Add-Member NoteProperty -Name "Archive Size (Mb)" -Value "n/a"
	    # $userObj | Add-Member NoteProperty -Name "Archive Deleted Item Size (Mb)" -Value "n/a"
	    # $userObj | Add-Member NoteProperty -Name "Archive Items" -Value "n/a"
    # }
    # else
    # {
        # $userObj | Add-Member NoteProperty -Name "Total Archive Size (Mb)" -Value ($archivestats.TotalItemSize.Value.ToMB() + $archivestats.TotalDeletedItemSize.Value.ToMB())
	    # $userObj | Add-Member NoteProperty -Name "Archive Size (Mb)" -Value $archivestats.TotalItemSize.Value.ToMB()
	    # $userObj | Add-Member NoteProperty -Name "Archive Deleted Item Size (Mb)" -Value $archivestats.TotalDeletedItemSize.Value.ToMB()
	    # $userObj | Add-Member NoteProperty -Name "Archive Items" -Value $archivestats.ItemCount
    # }

    # $userObj | Add-Member NoteProperty -Name "Audit Enabled" -Value $mb.AuditEnabled
    # $userObj | Add-Member NoteProperty -Name "Email Address Policy Enabled" -Value $mb.EmailAddressPolicyEnabled
    # $userObj | Add-Member NoteProperty -Name "Hidden From Address Lists" -Value $mb.HiddenFromAddressListsEnabled
    # $userObj | Add-Member NoteProperty -Name "Use Database Quota Defaults" -Value $mb.UseDatabaseQuotaDefaults
    
    if ($mb.UseDatabaseQuotaDefaults -eq $true)
    {
        $userObj | Add-Member NoteProperty -Name "IssueWarningQuota" -Value $primarydb.IssueWarningQuota.Value.ToMB()	
        $userObj | Add-Member NoteProperty -Name "ProhibitSendQuota" -Value $primarydb.ProhibitSendQuota.Value.ToMB()	
        $userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota" -Value $primarydb.ProhibitSendReceiveQuota.Value.ToMB()	
    }
    elseif ($mb.UseDatabaseQuotaDefaults -eq $false)
    {    
        $userObj | Add-Member NoteProperty -Name "IssueWarningQuota" -Value $mb.IssueWarningQuota.Value.ToMB()	
        $userObj | Add-Member NoteProperty -Name "ProhibitSendQuota" -Value $mb.ProhibitSendQuota.Value.ToMB()	
        $userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota" -Value $mb.ProhibitSendReceiveQuota.Value.ToMB()	
    }


    if( $userObj.MailboxSize -lt $userObj.IssueWarningQuota)
    {
        $userObj | Add-Member NoteProperty -Name "Statut" -Value "0"
    }
    elseif ($userObj.MailboxSize -gt $userObj.IssueWarningQuota -AND $userObj.MailboxSize -lt $userObj.ProhibitSendQuota )
    {
        $userObj | Add-Member NoteProperty -Name "Statut" -Value "1"
    }
    elseif ($userObj.MailboxSize -gt $userObj.ProhibitSendQuota -AND $userObj.MailboxSize -lt $userObj.ProhibitSendReceiveQuota )
    {
        $userObj | Add-Member NoteProperty -Name "Statut" -Value "2"
    }
        elseif ($userObj.MailboxSize -gt $userObj.ProhibitSendReceiveQuota)
    {
        $userObj | Add-Member NoteProperty -Name "Statut" -Value "3"
    }

	# $userObj | Add-Member NoteProperty -Name "Account Enabled" -Value $aduser.Enabled
	# $userObj | Add-Member NoteProperty -Name "Account Expires" -Value $aduser.AccountExpirationDate
	 $userObj | Add-Member NoteProperty -Name "Last Mailbox Logon" -Value $lastlogon
	# $userObj | Add-Member NoteProperty -Name "Last Logon By" -Value $stats.LastLoggedOnUserAccount
    

	$userObj | Add-Member NoteProperty -Name "Primary Mailbox Database" -Value $mb.Database
	# $userObj | Add-Member NoteProperty -Name "Primary Server/DAG" -Value $primarydb.MasterServerOrAvailabilityGroup

	# $userObj | Add-Member NoteProperty -Name "Archive Mailbox Database" -Value $mb.ArchiveDatabase
	# $userObj | Add-Member NoteProperty -Name "Archive Server/DAG" -Value $archivedb.MasterServerOrAvailabilityGroup

    # $userObj | Add-Member NoteProperty -Name "Primary Email Address" -Value $mb.PrimarySMTPAddress
    # $userObj | Add-Member NoteProperty -Name "Organizational Unit" -Value $user.OrganizationalUnit

	
	#Add the object to the report
	$report = $report += $userObj
}

#Catch zero item results
$reportcount = $report.count

if ($reportcount -eq 0)
{
	Write-Host -ForegroundColor Yellow "No mailboxes were found matching that criteria."
}
else
{
	#Output single mailbox report to console, otherwise only output to email with -SendEmail parameter
	 if ($mailbox) 
	 {
		 $report | Sort "MailboxSize" -Desc | Format-Table
	 }
}


if ($SendEmail)
{

    $topmailboxeshtml = $report | Sort "MailboxSize" -Desc  | ConvertTo-Html -Fragment | set-cellcolor Statut yellow -Filter "Statut -eq 1"
    $topmailboxeshtml = $topmailboxeshtml | set-cellcolor Statut orange -Filter "Statut -eq 2"
    $topmailboxeshtml = $topmailboxeshtml | set-cellcolor Statut red -Filter "Statut -eq 3"



	$htmlhead="<html>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
				TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
				TD{border: 1px solid #969595; padding: 5px; }
				</style>
				<body>
                <h1 align=""center"">Exchange Server Mailbox Report</h1>
                <h3 align=""center"">Generated: $now</h3>
                <p>Key : <span style='background-color: #ffff00;'>IssueWarningQuota</span> <span style='background-color: #ff9900;'>ProhibitSendQuota</span>&nbsp;<span style='background-color: #ff0000;'>ProhibitSendReceiveQuota</span> </p>
                <p>Report of Exchange mailboxes.</p>"    

    $spacer = "<br />"

	$htmltail = "</body></html>"

	$htmlreport = $htmlhead + $topmailboxeshtml + $htmltail

	try
    {
        Write-Host "Sending email report..."
        Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -ErrorAction STOP
        Write-Host "Finished."
    }
    catch
    {
        Write-Warning "An SMTP error has occurred, refer to log file for more details."
        $_.Exception.Message | Out-File "$myDir\get-mailboxreport-error.log"
        EXIT
    }
}
