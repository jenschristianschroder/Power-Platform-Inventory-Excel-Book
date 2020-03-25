$installModules = $fals
$days = 365

$dayspan = New-TimeSpan -Days $days
$startDate = (Get-Date) - $dayspan
$endDate = Get-Date

if($installModules -eq $true) {
    Install-Module PowershellGet -Force
    Install-Module -Name ExchangeOnlineManagement
}

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

Write-Host "Collecting Power Apps Operations from $startDate to $endDate"


$eventCount = 0
$Array = @()
DO {
    $PowerAppEventCollection = Search-UnifiedAuditLog -StartDate $startDate.ToString("yyyy-MM-dd") -EndDate $endDate.ToString("yyyy-MM-dd") -RecordType PowerAppsApp -ResultSize 5000
    ForEach($PowerAppEvent in $PowerAppEventCollection) {
            $PowerAppEventObject = New-Object PSObject -Property @{
            CreationTime = $PowerAppEvent | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty CreationTime
            Operation = $PowerAppEvent.Operations
            EnvironemtnName = $PowerAppEvent | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty AdditionalInfo | ConvertFrom-Json | Select-Object -ExpandProperty environmentName
            AppName = $PowerAppEvent | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty AppName
            UserId = $PowerAppEvent | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty UserId
        }
        $Array += $PowerAppEventObject
    }
    $eventCount = $PowerAppEventCollection.Length
    $endDate = [datetime]::parseexact($Array[-1].CreationTime, 'yyyy-MM-ddTHH:mm:ss', $null)
    "$endDate - $startDate : $eventCount operations"
} While ($eventCount -ne 0)

$Array | Export-Csv $PSScriptRoot\Data\PowerAppOperations.csv

Write-Host "Power Apps Operation collection complete." $Array.Count "Operations collected"-ForegroundColor Green
Write-Host "Last sample:"
$PowerAppEventObject

$endDate = Get-Date

Write-Host "Collecting Power Automate Operations from $startDate to $endDate"

$eventCount = 0
$Array = @()
DO {
    $PowerAutomateEventCollection = Search-UnifiedAuditLog -StartDate $startDate.ToString("yyyy-MM-dd") -EndDate $endDate.ToString("yyyy-MM-dd") -RecordType MicrosoftFlow -ResultSize 5000
    ForEach($PowerAutomateEvent in $PowerAutomateEventCollection) {
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
    $eventCount = $PowerAutomateEventCollection.Length
    $endDate = [datetime]::parseexact($Array[-1].CreationTime, 'yyyy-MM-ddTHH:mm:ss', $null)
    "$endDate - $startDate : $eventCount operations"
} While ($eventCount -ne 0)

$Array | Export-Csv $PSScriptRoot\Data\PowerAutomateOperations.csv

Write-Host "Power Automate Operation collection complete." $Array.Count "Operations collected"-ForegroundColor Green
Write-Host "Last sample:"
$PowerAutomateEventObject

Get-PSSession | Remove-PSSession