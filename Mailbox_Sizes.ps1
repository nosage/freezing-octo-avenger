# Fill these in:
$mailboxserver = "exch1"
$smtpserver = "exch1"
$From = "Exch1@domain.local"
$To = "it@domain.local"

Add-PSSnapin *exchange* -erroraction SilentlyContinue

$userlist = Get-Mailbox -server $mailboxserver -resultSize unlimited
$masterlist = @()
ForEach($user in $userlist)
{
  $mailbox = New-Object PSObject -Property @{
    DisplayName = $null
    TotalItemSize = $null
    ProhibitSendQuota = $null
    ProhibitSendReceiveQuota = $null
    IssueWarningQuota = $null
    DBProhibitSendQuota = $null
    DBProhibitSendReceiveQuota = $null
    PercentUsed = $null
    PercentUsedFmt = $null
    }
    $mailbox.DisplayName = ($user).DisplayName
    $mailbox.TotalItemSize = if ( (Get-MailboxStatistics -Identity $user).TotalItemSize.Value) {
        (Get-MailboxStatistics -Identity $user).TotalItemSize.Value.ToMB()
        } else {"-"}

    $mailbox.ProhibitSendQuota = if ( (Get-Mailbox $user).ProhibitSendQuota.Value) {
        (Get-Mailbox $user).ProhibitSendQuota.Value.ToMB()
        } else {"-"}

    $mailbox.ProhibitSendReceiveQuota = if ( (Get-Mailbox $user).ProhibitSendReceiveQuota.Value) {
        (Get-Mailbox $user).ProhibitSendReceiveQuota.Value.ToMB()
        } else {"-"}

    $mailbox.IssueWarningQuota = if ( (Get-Mailbox $user).IssueWarningQuota.value) {
        (Get-Mailbox $user).IssueWarningQuota.value.ToMB()
        } else {"-"}

    $mailbox.DBProhibitSendQuota = if ( (Get-MailboxDatabase -Identity $user.Database).ProhibitSendQuota.Value) {
        (Get-MailboxDatabase -Identity $user.Database).ProhibitSendQuota.Value.ToMB()
        } else {"-"}

    $mailbox.DBProhibitSendReceiveQuota = if ( (Get-MailboxDatabase -Identity $user.Database).ProhibitSendReceiveQuota.Value) {
        (Get-MailboxDatabase -Identity $user.Database).ProhibitSendReceiveQuota.Value.ToMB()
        } else {"-"}

    $mailbox.PercentUsed = if ( (Get-Mailbox $user).ProhibitSendReceiveQuota.Value) {
        (Get-MailboxStatistics -Identity $user).TotalItemSize.Value.ToMB() / (Get-Mailbox $user).ProhibitSendReceiveQuota.Value.ToMB()
        } else {
        (Get-MailboxStatistics -Identity $user).TotalItemSize.Value.ToMB() / (Get-MailboxDatabase -Identity $user.Database).ProhibitSendReceiveQuota.Value.ToMB()
        }
    $mailbox.PercentUsedFmt = "{0:P0}" -f $mailbox.PercentUsed
    $masterlist += $mailbox
}

$date = Get-Date -Format M-d-yy
$Header = @"
<style>
TABLE {border:#000 1px solid; border-collapse: collapse;}
TH {border:#000 1px solid; padding: 5px; background-color: #06F; font-family:sans-serif; font-size:8pt; font-weight: bold; color:#FFF;}
TD {border:#000 1px solid; padding: 5px; font-family:sans-serif;font-size:10pt;color:black;}
</style>
<title>
Mailbox Size Report
</title>
"@
$masterlist | select-object DisplayName, PercentUsedFmt, TotalItemSize, ProhibitSendQuota, ProhibitSendReceiveQuota, IssueWarningQuota, DBProhibitSendQuota, DBProhibitSendReceiveQuota | sort-object { [INT]($_.PercentUsedFmt -replace '%')  } -Descending | ConvertTo-Html -Head $Header | Set-Content c:\scripts\mailbox_sizes_$date.htm
$body = Get-Content c:\scripts\mailbox_sizes_$date.htm | Out-String
Send-MailMessage -From $From -To $To -Subject "Mailbox Sizes $date" -SmtpServer $smtpserver -Body $body -BodyAsHtml
