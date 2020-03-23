Write-Host "Updating required PowerShell modules"

Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -AllowClobber -Force -ErrorAction Stop
$AdminModuleVersion = Get-Module -Name Microsoft.PowerApps.Administration.PowerShell | Select-Object Version
Write-Host "Microsoft.PowerApps.Administration.PowerShell is version" $AdminModuleVersion.Version  -ForegroundColor Green

Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber -Force -ErrorAction Stop
$PowerAppsModuleVersion = Get-Module -Name Microsoft.PowerApps.PowerShell | Select-Object Version
Write-Host "Microsoft.PowerApps.PowerShell is version" $PowerAppsModuleVersion.Version  -ForegroundColor Green

Add-PowerAppsAccount
Write-Host "Connected to Power Platform" -ForegroundColor Green

Connect-AzureAD
Write-Host "Connected to Azure AD" -ForegroundColor Green

#Get Environment Inventory
Write-Host "`n"
Write-Host "Collecting Environment Inventory"
$Array = @()
ForEach($Environment in Get-PowerAppEnvironment) {
    $UserObject = Get-AzureADUser -All $true | Where-Object {$_.objectId -eq $Environment.CreatedBy.Id } | Select-Object UserPrincipalName, DisplayName
    $EnvironmentObject = New-Object PSObject -Property @{
        Environment = $Environment.Displayname
        EnvironmentOwner = $Environment.CreatedBy
        EnvironmentCreatedTime = $Environment.CreatedTime
        Location = $Environment.Location
        IsDefault = $Environment.IsDefault
    }
    $Array += $EnvironmentObject
}
$Array | Export-Csv $PSScriptRoot\Data\Environments.csv
Write-Host "Environment Inventory collection complete." $Array.Count "Environments collected"-ForegroundColor Green

#Get Power Apps Inventory
Write-Host "`n"
Write-Host "Collecting Power Apps Inventory"
$Array = @()
ForEach($Environment in Get-PowerAppEnvironment) {
    ForEach($PowerApp in Get-AdminPowerApp -EnvironmentName $Environment.EnvironmentName){
        $UserObject = Get-AzureADUser -All $true | Where-Object {$_.objectId -eq $PowerApp.Owner.Id } | Select-Object UserPrincipalName, DisplayName
        $PowerAppObject = New-Object PSObject -Property @{
            AppName = $PowerApp.AppName
            DisplayName = $PowerApp.DisplayName
            Environment = $Environment.Displayname
            EnvironmentCreatedTime = $Environment.CreatedTime
            Location = $Environment.Location
            IsDefault = $Environment.IsDefault
            Owner = $UserObject.DisplayName
            OwnerEmail = $UserObject.UserPrincipalName
            OwnerUPN = $UserObject.UserPrincipalName
        }
        $Array += $PowerAppObject
    }
}
$Array | Export-Csv $PSScriptRoot\Data\PowerAppInventory.csv
Write-Host "Power Apps Inventory collection complete." $Array.Count "Power Apps collected"-ForegroundColor Green

#Get Power Automate Inventory
Write-Host "`n"
Write-Host "Collecting Power Automate Inventory"
$Array = @()
ForEach($Environment in Get-FlowEnvironment) {
    ForEach($Flow in Get-AdminFlow -EnvironmentName $Environment.EnvironmentName){
        
        $UserObject = Get-AzureADUser -All $true | Where-Object {$_.objectId -eq $Flow.CreatedBy.objectId } | Select-Object UserPrincipalName, DisplayName
        $FlowObject = New-Object PSObject -Property @{
            AppName = $Flow.FlowName
            DisplayName = $Flow.DisplayName
            Environment = $Environment.Displayname
            EnvironmentCreatedTime = $Environment.CreatedTime
            Location = $Environment.Location
            IsDefault = $Environment.IsDefault
            Owner = $UserObject | Select-Object -ExpandProperty DisplayName
            OwnerEmail = $UserObject | Select-Object -ExpandProperty UserPrincipalName 
            OwnerUPN = $UserObject | Select-Object -ExpandProperty UserPrincipalName 
        }
        $Array += $FlowObject
    }
}
$Array | Export-Csv $PSScriptRoot\Data\FlowInventory.csv
Write-Host "Power Automate Inventory collection complete." $Array.Count "Flows collected"-ForegroundColor Green
Write-Host "`n"
