# Prerequisites:
# Audit logging must be enabled in the Microsoft 365 organization and you must be assigned the View-Only AuditLogs or Audit Logs role in Exchange Online. https://learn.microsoft.com/en-us/purview/audit-log-search-script#before-you-run-the-script

# - Connect to Exchange Online using Connect-ExchangeOnline 

$Logs = Search-UnifiedAuditLog `
    -StartDate (Get-Date).AddDays(-90) `
    -EndDate (Get-Date) `
    -Operations "MeetingParticipantDetail" `
    -ResultSize 100

$Logs.Count

$BotPattern = '(?i)otter|fireflies|read\.ai|fathom|notetaker|recorder|bot|assistant|tl;dv|gong'

$Results = foreach ($Log in $Logs) {

    $Audit = $Log.AuditData | ConvertFrom-Json

    foreach ($Attendee in @($Audit.Attendees)) {

        if (
            $Attendee.DisplayName -match $BotPattern -or
            $Attendee.UPN -match $BotPattern
        ) {

            [PSCustomObject]@{
                CreationTime    = $Audit.CreationTime
                DisplayName     = $Attendee.DisplayName
                UPN             = $Attendee.UPN
                RecipientType   = $Attendee.RecipientType
                UserObjectId    = $Attendee.UserObjectId
                OrganizerId     = $Audit.OrganizerId
                MeetingDetailId = $Audit.MeetingDetailId
                JoinTime        = $Audit.JoinTime
                LeaveTime       = $Audit.LeaveTime
                JoinedFromLobby = $Audit.IsJoinedFromLobby
            }
        }
    }
}

$Results | Format-Table -AutoSize