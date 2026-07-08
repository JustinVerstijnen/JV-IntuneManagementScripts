<#
.SYNOPSIS
Sets the Intune Management name of all managed devices to match the Device name.

.DESCRIPTION
This script retrieves all Microsoft Intune managed devices and compares:
- deviceName
- managedDeviceName

If managedDeviceName does not match deviceName, the script updates managedDeviceName
so it becomes equal to deviceName.

By default, the script applies the change IMMEDIATELY.
Use -DryRun if you only want to check what would be changed without making changes.

.REQUIREMENTS
- PowerShell 5.1 or PowerShell 7+
- Microsoft.Graph.Authentication module
- Graph permission: DeviceManagementManagedDevices.ReadWrite.All
- Account with sufficient Intune/Graph permissions

.EXAMPLES
Apply changes directly:
.\Set-IntuneManagementNameToDeviceName.ps1

Check only:
.\Set-IntuneManagementNameToDeviceName.ps1 -DryRun

Specific tenant:
.\Set-IntuneManagementNameToDeviceName.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [string]$ReportPath = ".\Intune-ManagementNameToDeviceName-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
)

$ErrorActionPreference = "Stop"

function Get-SafePropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName,

        [Parameter(Mandatory = $false)]
        [object]$DefaultValue = $null
    )

    if ($null -eq $InputObject) {
        return $DefaultValue
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($PropertyName)) {
            return $InputObject[$PropertyName]
        }

        return $DefaultValue
    }

    $property = $InputObject.PSObject.Properties |
        Where-Object { $_.Name -eq $PropertyName } |
        Select-Object -First 1

    if ($null -ne $property) {
        return $property.Value
    }

    return $DefaultValue
}

function Test-GraphAuthenticationModule {
    $moduleName = "Microsoft.Graph.Authentication"

    $loadedModule = Get-Module -Name $moduleName -ErrorAction SilentlyContinue

    if ($loadedModule) {
        Write-Host "Graph Authentication module is already loaded: version $($loadedModule.Version)" -ForegroundColor Cyan
        return
    }

    $availableModules = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue |
        Sort-Object Version -Descending

    if (-not $availableModules) {
        throw @"
Module '$moduleName' is not installed.

Run the following command first:

Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber

Or as administrator:

Install-Module Microsoft.Graph.Authentication -Scope AllUsers -Force -AllowClobber
"@
    }

    $latestModule = $availableModules | Select-Object -First 1

    Write-Host "Loading Graph Authentication module: version $($latestModule.Version)" -ForegroundColor Cyan

    Import-Module $moduleName -RequiredVersion $latestModule.Version -ErrorAction Stop -WarningAction SilentlyContinue
}

function Connect-ToMicrosoftGraph {
    param(
        [Parameter(Mandatory = $false)]
        [string]$TenantId
    )

    $scopes = @(
        "DeviceManagementManagedDevices.ReadWrite.All"
    )

    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

    if ([string]::IsNullOrWhiteSpace($TenantId)) {
        Connect-MgGraph -Scopes $scopes -NoWelcome
    }
    else {
        Connect-MgGraph -TenantId $TenantId -Scopes $scopes -NoWelcome
    }

    $context = Get-MgContext

    if (-not $context) {
        throw "No Graph context found. Sign-in failed."
    }

    Write-Host "Connected to tenant: $($context.TenantId)" -ForegroundColor Green
    Write-Host "Signed in as: $($context.Account)" -ForegroundColor Green
}

function Get-AllGraphPages {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $items = @()
    $nextUri = $Uri

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject

        $pageItems = Get-SafePropertyValue -InputObject $response -PropertyName "value" -DefaultValue @()

        if ($null -ne $pageItems) {
            $items += @($pageItems)
        }

        $nextLink = Get-SafePropertyValue -InputObject $response -PropertyName "@odata.nextLink" -DefaultValue $null

        if ([string]::IsNullOrWhiteSpace($nextLink)) {
            $nextUri = $null
        }
        else {
            $nextUri = $nextLink
        }
    }

    return $items
}

function Get-IntuneManagedDevices {
    $selectProperties = @(
        "id",
        "deviceName",
        "managedDeviceName",
        "operatingSystem",
        "serialNumber",
        "userPrincipalName",
        "model",
        "manufacturer",
        "lastSyncDateTime"
    ) -join ","

    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=$selectProperties&`$top=100"

    Write-Host "Retrieving all Intune managed devices..." -ForegroundColor Cyan

    $devices = Get-AllGraphPages -Uri $uri

    return @($devices)
}

function Set-IntuneDeviceManagementName {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Device,

        [Parameter(Mandatory = $true)]
        [string]$NewManagementName
    )

    $deviceIdRaw = Get-SafePropertyValue -InputObject $Device -PropertyName "id" -DefaultValue ""

    if ([string]::IsNullOrWhiteSpace($deviceIdRaw)) {
        throw "Device has no Intune device ID. Cannot update management name."
    }

    if ([string]::IsNullOrWhiteSpace($NewManagementName)) {
        throw "New management name is empty. Cannot update management name."
    }

    $deviceId = [uri]::EscapeDataString($deviceIdRaw)
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$deviceId"

    $bodyObject = @{
        managedDeviceName = $NewManagementName
    }

    $bodyJson = $bodyObject | ConvertTo-Json -Depth 5

    Invoke-MgGraphRequest `
        -Method PATCH `
        -Uri $uri `
        -Body $bodyJson `
        -ContentType "application/json" `
        -OutputType PSObject
}

function New-ResultObject {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Device,

        [Parameter(Mandatory = $true)]
        [string]$OldManagementName,

        [Parameter(Mandatory = $true)]
        [string]$NewManagementName,

        [Parameter(Mandatory = $true)]
        [bool]$Changed,

        [Parameter(Mandatory = $true)]
        [string]$Result,

        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = ""
    )

    return [PSCustomObject]@{
        DeviceName        = Get-SafePropertyValue -InputObject $Device -PropertyName "deviceName" -DefaultValue ""
        OldManagementName = $OldManagementName
        NewManagementName = $NewManagementName
        Changed           = $Changed
        Result            = $Result
        Error             = $ErrorMessage
        OperatingSystem   = Get-SafePropertyValue -InputObject $Device -PropertyName "operatingSystem" -DefaultValue ""
        SerialNumber      = Get-SafePropertyValue -InputObject $Device -PropertyName "serialNumber" -DefaultValue ""
        UserPrincipalName = Get-SafePropertyValue -InputObject $Device -PropertyName "userPrincipalName" -DefaultValue ""
        Model             = Get-SafePropertyValue -InputObject $Device -PropertyName "model" -DefaultValue ""
        Manufacturer      = Get-SafePropertyValue -InputObject $Device -PropertyName "manufacturer" -DefaultValue ""
        LastSyncDateTime  = Get-SafePropertyValue -InputObject $Device -PropertyName "lastSyncDateTime" -DefaultValue ""
        IntuneDeviceId    = Get-SafePropertyValue -InputObject $Device -PropertyName "id" -DefaultValue ""
    }
}

try {
    Test-GraphAuthenticationModule
    Connect-ToMicrosoftGraph -TenantId $TenantId

    $allDevices = Get-IntuneManagedDevices

    Write-Host ""
    Write-Host "Number of Intune managed devices found: $($allDevices.Count)" -ForegroundColor Cyan

    $devicesToChange = @(
        $allDevices | Where-Object {
            $deviceName = Get-SafePropertyValue -InputObject $_ -PropertyName "deviceName" -DefaultValue ""
            $managementName = Get-SafePropertyValue -InputObject $_ -PropertyName "managedDeviceName" -DefaultValue ""

            -not [string]::IsNullOrWhiteSpace($deviceName) -and $managementName -ne $deviceName
        }
    )

    $devicesSkippedBecauseDeviceNameIsEmpty = @(
        $allDevices | Where-Object {
            $deviceName = Get-SafePropertyValue -InputObject $_ -PropertyName "deviceName" -DefaultValue ""
            [string]::IsNullOrWhiteSpace($deviceName)
        }
    )

    $devicesAlreadyCorrect = @(
        $allDevices | Where-Object {
            $deviceName = Get-SafePropertyValue -InputObject $_ -PropertyName "deviceName" -DefaultValue ""
            $managementName = Get-SafePropertyValue -InputObject $_ -PropertyName "managedDeviceName" -DefaultValue ""

            -not [string]::IsNullOrWhiteSpace($deviceName) -and $managementName -eq $deviceName
        }
    )

    Write-Host "Number of devices that will be changed: $($devicesToChange.Count)" -ForegroundColor Yellow
    Write-Host "Number of devices already correct: $($devicesAlreadyCorrect.Count)" -ForegroundColor Green
    Write-Host "Number of devices skipped because Device name is empty: $($devicesSkippedBecauseDeviceNameIsEmpty.Count)" -ForegroundColor Yellow
    Write-Host ""

    if ($DryRun) {
        Write-Host "DRY-RUN mode is active. No changes will be made." -ForegroundColor Yellow
        Write-Host ""
    }
    else {
        Write-Host "DIRECT EXECUTION mode is active. Management names will now be changed to match Device names." -ForegroundColor Red
        Write-Host ""
    }

    $results = @()

    foreach ($device in $devicesToChange) {
        $deviceName = Get-SafePropertyValue -InputObject $device -PropertyName "deviceName" -DefaultValue ""
        $oldManagementName = Get-SafePropertyValue -InputObject $device -PropertyName "managedDeviceName" -DefaultValue ""
        $serialNumber = Get-SafePropertyValue -InputObject $device -PropertyName "serialNumber" -DefaultValue ""
        $operatingSystem = Get-SafePropertyValue -InputObject $device -PropertyName "operatingSystem" -DefaultValue ""

        Write-Host "Device: $deviceName | OS: $operatingSystem | Serial: $serialNumber | Old Management name: $oldManagementName" -ForegroundColor Cyan

        if ($DryRun) {
            $results += New-ResultObject `
                -Device $device `
                -OldManagementName $oldManagementName `
                -NewManagementName $deviceName `
                -Changed $false `
                -Result "Dry-run; no change made"

            continue
        }

        try {
            Set-IntuneDeviceManagementName -Device $device -NewManagementName $deviceName | Out-Null

            Write-Host "Management name changed to: $deviceName" -ForegroundColor Green

            $results += New-ResultObject `
                -Device $device `
                -OldManagementName $oldManagementName `
                -NewManagementName $deviceName `
                -Changed $true `
                -Result "Changed"
        }
        catch {
            Write-Host "ERROR while changing device: $($_.Exception.Message)" -ForegroundColor Red

            $results += New-ResultObject `
                -Device $device `
                -OldManagementName $oldManagementName `
                -NewManagementName $deviceName `
                -Changed $false `
                -Result "Error" `
                -ErrorMessage $_.Exception.Message
        }
    }

    foreach ($device in $devicesAlreadyCorrect) {
        $deviceName = Get-SafePropertyValue -InputObject $device -PropertyName "deviceName" -DefaultValue ""
        $managementName = Get-SafePropertyValue -InputObject $device -PropertyName "managedDeviceName" -DefaultValue ""

        $results += New-ResultObject `
            -Device $device `
            -OldManagementName $managementName `
            -NewManagementName $deviceName `
            -Changed $false `
            -Result "Already correct"
    }

    foreach ($device in $devicesSkippedBecauseDeviceNameIsEmpty) {
        $managementName = Get-SafePropertyValue -InputObject $device -PropertyName "managedDeviceName" -DefaultValue ""

        $results += New-ResultObject `
            -Device $device `
            -OldManagementName $managementName `
            -NewManagementName "" `
            -Changed $false `
            -Result "Skipped; Device name is empty"
    }

    $results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Done." -ForegroundColor Green
    Write-Host "Report saved to: $ReportPath" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Script stopped because of an error:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
finally {
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
        # No action needed
    }
}
