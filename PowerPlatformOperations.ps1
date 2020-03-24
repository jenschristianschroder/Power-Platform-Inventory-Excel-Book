$installModules = $false
$days = 21

$dayspan = New-TimeSpan -Days $days
$startDate = (Get-Date) - $dayspan
$endDate = (Get-Date)

if($installModules -eq $true) {
    Install-Module PowershellGet -Force
    Install-Module -Name ExchangeOnlineManagement
}

Write-Host "Collecting Power Apps Operations from $startDate to $endDate"

$Array = @()
ForEach($PowerAppEvent in Search-UnifiedAuditLog -StartDate $startDate.ToString("yyyy-MM-dd") -EndDate $endDate.ToString("yyyy-MM-dd") -RecordType PowerAppsApp -ResultSize 250) {
        $PowerAppEventObject = New-Object PSObject -Property @{
        CreationTime = $PowerAppEvent | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty CreationTime
        Operation = $PowerAppEvent.Operations
        EnvironemtnName = $PowerAppEvent | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty AdditionalInfo | ConvertFrom-Json | Select-Object -ExpandProperty environmentName
        AppName = $PowerAppEvent | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty AppName
        UserId = $PowerAppEvent | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty UserId
    }
    $Array += $PowerAppEventObject
}
$Array | Export-Csv $PSScriptRoot\Data\PowerAppOperations.csv
Write-Host "Power Apps Operation collection complete." $Array.Count "Operations collected"-ForegroundColor Green
Write-Host "Last sample:"
$PowerAppEventObject

if($Array.Count -eq 250) {
    WriteHost "There where 250 events in the returned result. There might be more events in your select period." -ForegroundColor Yellow
}

Write-Host "Collecting Power Automate Operations from $startDate to $endDate"

$Array = @()
ForEach($PowerAutomateEvent in Search-UnifiedAuditLog -StartDate $startDate.ToString("yyyy-MM-dd") -EndDate $endDate.ToString("yyyy-MM-dd") -RecordType MicrosoftFlow -ResultSize 250) {
        $FlowDetailsUrl = $PowerAutomateEvent | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty FlowDetailsUrl
        $FlowDetailsUrlArray = $FlowDetailsUrl -split "/"
        $PowerAutomateEventObject = New-Object PSObject -Property @{
        CreationTime = $PowerAutomateEvent | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty CreationTime
        Operation = $PowerAutomateEvent.Operations
        EnvironemtnName = $FlowDetailsUrlArray[4]
        AppName = $PowerAutomateEvent | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty ObjectId
        UserId = $PowerAutomateEvent | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty UserId
    }
    $Array += $PowerAutomateEventObject
}
$Array | Export-Csv $PSScriptRoot\Data\PowerAutomateOperations.csv
Write-Host "Power Automate Operation collection complete." $Array.Count "Operations collected"-ForegroundColor Green
Write-Host "Last sample:"
$PowerAutomateEventObject

if($Array.Count -eq 250) {
    WriteHost "There where 250 events in the returned result. There might be more events in your select period." -ForegroundColor Yellow
}
