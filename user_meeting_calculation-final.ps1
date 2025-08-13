<<<<<<< HEAD
﻿# ==============================
# Config
# ==============================
$tenantId = "23c2b29f-23d7-488d-8a9f-6a486301094a"
$clientId = "f802cb03-e1b7-4aa3-a36d-65bd33a82bbd"
$clientSecret = "ftm8Q~x1yN4OI7dCKvWGLEnBmE7P05b369-ITdmU"

$scopes = "https://graph.microsoft.com/.default"

$body = @{
    client_id     = $clientId
    scope         = $scopes
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -Body $body
$accessToken = $tokenResponse.access_token

# ==============================
# User list
# ==============================

$userEmails = @(
    "mdsofiqul.islam@bracits.com",
    "mdmorsalin.anwar@bracits.com",
    "farzana.meher@bracits.com",
    "prodip.saha@bracits.com",
    "mohammad.hossen@bracits.com",
    "arif.muhammad@bracits.com",
    "sadia.badhan@bracits.com",
    "tahmid.lodi@bracits.com",
    "sadia.chowdhury@bracits.com",
    "hossain.ahmed@bracits.com",
    "sayed.labeed@bracits.com",
    "mesbah.uddin@bracits.com",
    "shakib.chowdhury@bracits.com",
    "iftekhar.rahman@bracits.com",
    "lubaba.salsabil@bracits.com",
    "tabassum.ashrafi@bracits.com",
    "morsalin.mithun@bracits.com",
    "sohel.adnan@bracits.com",
    "salman.azad@bracits.com",
    "mashroora.nadi@bracits.com",
    "adnan.hossain@bracits.com",
    "faisal.rhythm@bracits.com",
    "sayan.manoranjan@bracits.com",
    "raisa.tabassum@bracits.com",
    "salman.masud@bracits.com",
    "mostahid.ahmed@bracits.com",
    "sheikh.salman@bracits.com",
    "afsha.haque@bracits.com",
    "mithila.hoq@bracits.com",
    "shahnaz.sarker@bracits.com",
    "tausir.khan@bracits.com",
    "farah.ahmed@bracits.com",
    "mahedi.hasan@bracits.com",
    "shadman.sakib@bracits.com",
    "abdul.mannan@bracits.com",
    "mohammad.faisal@bracits.com"
)


# ==============================
# Inputs
# ==============================
$monthsBack = [int](Read-Host "Enter number of months to go back (e.g. 3)")

$headers = @{
    Authorization = "Bearer $accessToken"
    'Content-Type' = "application/json"
    Prefer = 'outlook.timezone="Bangladesh Standard Time"'
}

# ==============================
# Main Loop
# ==============================
$currentDate = Get-Date
$pivotData = @{}

foreach ($userEmail in $userEmails) {
    $pivotData[$userEmail] = @{}

    for ($i = $monthsBack; $i -ge 1; $i--) {
        $monthStart = (Get-Date -Year $currentDate.Year -Month $currentDate.Month -Day 1).AddMonths(-$i)
        $monthEnd = $monthStart.AddMonths(1).AddSeconds(-1)
        $monthKey = $monthStart.ToString("yyyy-MM")

        $allEvents = @()
        $startDateTime = $monthStart.ToString("yyyy-MM-ddTHH:mm:ssZ")
        $endDateTime = $monthEnd.ToString("yyyy-MM-ddTHH:mm:ssZ")
        $eventsUri = "https://graph.microsoft.com/v1.0/users/$userEmail/calendarView?startDateTime=$startDateTime&endDateTime=$endDateTime&`$top=1000&`$orderby=start/dateTime"

        do {
            try {
                $eventsResponse = Invoke-RestMethod -Uri $eventsUri -Method Get -Headers $headers
                $allEvents += $eventsResponse.value
                $eventsUri = $eventsResponse.'@odata.nextLink'
            } catch {
                Write-Warning "Failed to fetch events for $userEmail in $monthKey : $_"
                break
            }
        } while ($eventsUri)

        $totalMinutes = 0
        foreach ($event in $allEvents) {
            $startTime = [datetime]$event.start.dateTime
            $endTime = [datetime]$event.end.dateTime
            $duration = ($endTime - $startTime).TotalMinutes
            $totalMinutes += $duration
        }

        $pivotData[$userEmail][$monthKey] = [math]::Round($totalMinutes / 60, 2)
    }
}

# ==============================
# Convert to Pivot Table
# ==============================
$allMonths = @()
for ($i = $monthsBack; $i -ge 1; $i--) {
    $monthStart = (Get-Date -Year $currentDate.Year -Month $currentDate.Month -Day 1).AddMonths(-$i)
    $allMonths += $monthStart.ToString("yyyy-MM")
}

$pivotTable = @()
foreach ($userEmail in $pivotData.Keys) {
    $row = [ordered]@{ UserEmail = $userEmail }
    foreach ($month in $allMonths) {
        $row[$month] = $pivotData[$userEmail][$month]
    }
    $pivotTable += New-Object PSObject -Property $row
}

# ==============================
# Export to CSV
# ==============================
$outputFile = "C:\Users\sabbir.hossain\Documents\MeetingHours_Pivoted_${monthsBack}months.csv"
$pivotTable | Export-Csv -Path $outputFile -NoTypeInformation

Write-Host "✅ Pivoted output saved to $outputFile"
=======
﻿# ==============================
# Config
# ==============================
$tenantId = "23c2b29f-23d7-488d-8a9f-6a486301094a"
$clientId = "f802cb03-e1b7-4aa3-a36d-65bd33a82bbd"
$clientSecret = "ftm8Q~x1yN4OI7dCKvWGLEnBmE7P05b369-ITdmU"

$scopes = "https://graph.microsoft.com/.default"

$body = @{
    client_id     = $clientId
    scope         = $scopes
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -Body $body
$accessToken = $tokenResponse.access_token

# ==============================
# User list
# ==============================
$userEmails = @(
    "mdsofiqul.islam@bracits.com",
    "mdmorsalin.anwar@bracits.com",
    "farzana.meher@bracits.com",
    "prodip.saha@bracits.com",
    "mohammad.hossen@bracits.com",
    "arif.muhammad@bracits.com",
    "sadia.badhan@bracits.com",
    "tahmid.lodi@bracits.com",
    "sadia.chowdhury@bracits.com",
    "hossain.ahmed@bracits.com",
    "sayed.labeed@bracits.com",
    "mesbah.uddin@bracits.com",
    "shakib.chowdhury@bracits.com",
    "iftekhar.rahman@bracits.com",
    "lubaba.salsabil@bracits.com",
    "tabassum.ashrafi@bracits.com",
    "morsalin.mithun@bracits.com",
    "sohel.adnan@bracits.com",
    "salman.azad@bracits.com",
    "mashroora.nadi@bracits.com",
    "adnan.hossain@bracits.com",
    "faisal.rhythm@bracits.com",
    "sayan.manoranjan@bracits.com",
    "raisa.tabassum@bracits.com",
    "salman.masud@bracits.com",
    "mostahid.ahmed@bracits.com",
    "sheikh.salman@bracits.com",
    "afsha.haque@bracits.com",
    "mithila.hoq@bracits.com",
    "shahnaz.sarker@bracits.com",
    "tausir.khan@bracits.com",
    "farah.ahmed@bracits.com",
    "mahedi.hasan@bracits.com",
    "shadman.sakib@bracits.com",
    "abdul.mannan@bracits.com",
    "mohammad.faisal@bracits.com"
)

# ==============================
# Inputs
# ==============================
$monthsBack = [int](Read-Host "Enter number of months to go back (e.g. 3)")

$headers = @{
    Authorization = "Bearer $accessToken"
    'Content-Type' = "application/json"
    Prefer = 'outlook.timezone="Bangladesh Standard Time"'
}

# ==============================
# Main Loop
# ==============================
$currentDate = Get-Date
$pivotData = @{}

foreach ($userEmail in $userEmails) {
    $pivotData[$userEmail] = @{}

    for ($i = $monthsBack; $i -ge 1; $i--) {
        Write-Progress -Activity "Processing $userEmail" -Status "$i months back"

        $monthStart = (Get-Date -Year $currentDate.Year -Month $currentDate.Month -Day 1).AddMonths(-$i)
        $monthEnd = $monthStart.AddMonths(1).AddSeconds(-1)
        $monthKey = $monthStart.ToString("MMM-yyyy")  # Changed format

        $allEvents = @()
        $startDateTime = $monthStart.ToString("yyyy-MM-ddTHH:mm:ssZ")
        $endDateTime = $monthEnd.ToString("yyyy-MM-ddTHH:mm:ssZ")
        $eventsUri = "https://graph.microsoft.com/v1.0/users/$userEmail/calendarView?startDateTime=$startDateTime&endDateTime=$endDateTime&`$top=1000&`$orderby=start/dateTime"

        do {
            try {
                $eventsResponse = Invoke-RestMethod -Uri $eventsUri -Method Get -Headers $headers
                $allEvents += $eventsResponse.value
                $eventsUri = $eventsResponse.'@odata.nextLink'
            } catch {
                Write-Warning "Failed to fetch events for $userEmail in $monthKey : $_"
                break
            }
        } while ($eventsUri)

        $totalMinutes = 0
        foreach ($event in $allEvents) {
            # Skip cancelled or free meetings
            if ($event.status -eq "cancelled" -or $event.showAs -eq "free") { continue }

            $startTime = [datetime]$event.start.dateTime
            $endTime = [datetime]$event.end.dateTime
            $duration = ($endTime - $startTime).TotalMinutes
            $totalMinutes += $duration
        }

        $pivotData[$userEmail][$monthKey] = [math]::Round($totalMinutes / 60, 2)
    }
}

# ==============================
# Convert to Pivot Table
# ==============================
$allMonths = @()
for ($i = $monthsBack; $i -ge 1; $i--) {
    $monthStart = (Get-Date -Year $currentDate.Year -Month $currentDate.Month -Day 1).AddMonths(-$i)
    $allMonths += $monthStart.ToString("MMM-yyyy") # Changed format
}

$pivotTable = @()
foreach ($userEmail in $pivotData.Keys) {
    $row = [ordered]@{ UserEmail = $userEmail }
    foreach ($month in $allMonths) {
        $row[$month] = $pivotData[$userEmail][$month]
    }
    $pivotTable += New-Object PSObject -Property $row
}

# ==============================
# Export to CSV
# ==============================
$outputFile = "C:\Users\sabbir.hossain\Documents\MeetingHours_Pivoted_${monthsBack}months_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
$pivotTable | Export-Csv -Path $outputFile -NoTypeInformation

Write-Host "✅ Pivoted output saved to $outputFile"
>>>>>>> bf72cca (Initial commit or file update)
