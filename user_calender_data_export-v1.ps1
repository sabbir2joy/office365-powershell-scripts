# Required variables
$tenantId = "23c2b29f-23d7-488d-8a9f-6a486301094a"
$clientId = "f802cb03-e1b7-4aa3-a36d-65bd33a82bbd"
$clientSecret = "ftm8Q~x1yN4OI7dCKvWGLEnBmE7P05b369-ITdmU"

# Scope for Graph API
$scopes = "https://graph.microsoft.com/.default"

# Authenticate to Microsoft Graph
$body = @{
    client_id     = $clientId
    scope         = $scopes
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -Body $body
$accessToken = $tokenResponse.access_token

# Prompt for the user's email address
$userPrincipalName = Read-Host "Enter the user's email address"
$monthsBack = Read-Host "Enter the number of months to go back (e.g. 9)"

# Date range
$startDate = (Get-Date).AddMonths(-[int]$monthsBack).ToString("yyyy-MM-ddTHH:mm:ssZ")
$endDate   = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")

Write-Host "Fetching events from $startDate to $endDate for user $userPrincipalName"

# Headers
$headers = @{
    Authorization = "Bearer $accessToken"
    'Content-Type' = "application/json"
    Prefer = "outlook.timezone=""Bangladesh Standard Time"""
}

# Collect all events
$allEvents = @()
$eventsUri = "https://graph.microsoft.com/v1.0/users/$userPrincipalName/calendarView?startDateTime=$startDate&endDateTime=$endDate&`$top=1000&`$orderby=start/dateTime"

do {
    $eventsResponse = Invoke-RestMethod -Uri $eventsUri -Method Get -Headers $headers
    $allEvents += $eventsResponse.value
    $eventsUri = $eventsResponse.'@odata.nextLink'
} while ($eventsUri)

# Prepare CSV data
$eventList = foreach ($event in $allEvents) {
    $startTime = [datetime]$event.start.dateTime
    $endTime = [datetime]$event.end.dateTime
    $duration = $endTime - $startTime

    [PSCustomObject]@{
        Subject   = $event.subject
        Start     = $startTime.ToString("yyyy-MM-dd HH:mm")
        End       = $endTime.ToString("yyyy-MM-dd HH:mm")
        DurationMin = [math]::Round($duration.TotalMinutes, 0)
        Organizer = $event.organizer.emailAddress.address
        Location  = $event.location.displayName
    }
}

# Export
$username = $userPrincipalName.Split('@')[0]
$csvFilePath = "C:\Users\sabbir.hossain\Documents\${username}_calendar_events_${monthsBack}_months.csv"
$eventList | Export-Csv -Path $csvFilePath -NoTypeInformation

Write-Host "✅ Calendar events exported to $csvFilePath"
