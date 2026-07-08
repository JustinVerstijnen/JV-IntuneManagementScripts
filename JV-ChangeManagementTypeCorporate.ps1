<#
.SYNOPSIS
Sets all Microsoft Intune managed devices directly to Corporate ownership.

.DESCRIPTION
This script finds all Intune managed devices where managedDeviceOwnerType
is not equal to company.

By default, the script applies the change IMMEDIATELY.
Use -DryRun if you only want to check what would be changed without making changes.

.REQUIREMENTS
- PowerShell 5.1 or PowerShell 7+
- Microsoft.Graph.Authentication module
- Graph permission: DeviceManagementManagedDevices.ReadWrite.All
- Account with sufficient Intune/Graph permissions

.EXAMPLES
Apply changes directly:
.\Set-IntuneAllDevicesCorporate.ps1

Check only:
.\Set-IntuneAllDevicesCorporate.ps1 -DryRun

Specific tenant:
.\Set-IntuneAllDevicesCorporate.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

Only change Personal devices:
.\Set-IntuneAllDevicesCorporate.ps1 -Mode PersonalOnly

Only change Unknown/empty owner type devices:
.\Set-IntuneAllDevicesCorporate.ps1 -Mode UnknownOnly
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [ValidateSet("AllNonCorporate", "PersonalOnly", "UnknownOnly")]
    [string]$Mode = "AllNonCorporate",

    [Parameter(Mandatory = $false)]
    [string]$ReportPath = ".\Intune-AllDevicesOwnershipChange-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
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
        "operatingSystem",
        "managedDeviceOwnerType",
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

function Get-DevicesToChange {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Devices,

        [Parameter(Mandatory = $true)]
        [string]$Mode
    )

    switch ($Mode) {
        "AllNonCorporate" {
            return @(
                $Devices | Where-Object {
                    $ownerType = Get-SafePropertyValue -InputObject $_ -PropertyName "managedDeviceOwnerType" -DefaultValue ""
                    $ownerType -ne "company"
                }
            )
        }

        "PersonalOnly" {
            return @(
                $Devices | Where-Object {
                    $ownerType = Get-SafePropertyValue -InputObject $_ -PropertyName "managedDeviceOwnerType" -DefaultValue ""
                    $ownerType -eq "personal"
                }
            )
        }

        "UnknownOnly" {
            return @(
                $Devices | Where-Object {
                    $ownerType = Get-SafePropertyValue -InputObject $_ -PropertyName "managedDeviceOwnerType" -DefaultValue ""
                    $ownerType -eq "unknown" -or [string]::IsNullOrWhiteSpace($ownerType)
                }
            )
        }

        default {
            throw "Unknown Mode: $Mode"
        }
    }
}

function Set-IntuneDeviceCorporateOwnership {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Device
    )

    $deviceIdRaw = Get-SafePropertyValue -InputObject $Device -PropertyName "id" -DefaultValue ""

    if ([string]::IsNullOrWhiteSpace($deviceIdRaw)) {
        throw "Device has no Intune device ID. Cannot change ownership."
    }

    $deviceId = [uri]::EscapeDataString($deviceIdRaw)
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$deviceId"

    $bodyObject = @{
        managedDeviceOwnerType = "company"
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
        [string]$OldOwnerType,

        [Parameter(Mandatory = $true)]
        [bool]$Changed,

        [Parameter(Mandatory = $true)]
        [string]$Result,

        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = ""
    )

    return [PSCustomObject]@{
        DeviceName        = Get-SafePropertyValue -InputObject $Device -PropertyName "deviceName" -DefaultValue ""
        SerialNumber      = Get-SafePropertyValue -InputObject $Device -PropertyName "serialNumber" -DefaultValue ""
        UserPrincipalName = Get-SafePropertyValue -InputObject $Device -PropertyName "userPrincipalName" -DefaultValue ""
        OperatingSystem   = Get-SafePropertyValue -InputObject $Device -PropertyName "operatingSystem" -DefaultValue ""
        Model             = Get-SafePropertyValue -InputObject $Device -PropertyName "model" -DefaultValue ""
        Manufacturer      = Get-SafePropertyValue -InputObject $Device -PropertyName "manufacturer" -DefaultValue ""
        OldOwnerType      = $OldOwnerType
        NewOwnerType      = "company"
        Changed           = $Changed
        Result            = $Result
        Error             = $ErrorMessage
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

    $devicesToChange = Get-DevicesToChange -Devices $allDevices -Mode $Mode

    Write-Host "Number of devices that will be changed: $($devicesToChange.Count)" -ForegroundColor Yellow
    Write-Host ""

    if ($DryRun) {
        Write-Host "DRY-RUN mode is active. No changes will be made." -ForegroundColor Yellow
        Write-Host ""
    }
    else {
        Write-Host "DIRECT EXECUTION mode is active. Devices will now be changed to Corporate." -ForegroundColor Red
        Write-Host ""
    }

    $results = @()

    foreach ($device in $devicesToChange) {
        $deviceName = Get-SafePropertyValue -InputObject $device -PropertyName "deviceName" -DefaultValue ""
        $serialNumber = Get-SafePropertyValue -InputObject $device -PropertyName "serialNumber" -DefaultValue ""
        $operatingSystem = Get-SafePropertyValue -InputObject $device -PropertyName "operatingSystem" -DefaultValue ""
        $oldOwnerType = Get-SafePropertyValue -InputObject $device -PropertyName "managedDeviceOwnerType" -DefaultValue ""

        if ([string]::IsNullOrWhiteSpace($oldOwnerType)) {
            $oldOwnerType = "empty"
        }

        Write-Host "Device: $deviceName | OS: $operatingSystem | Serial: $serialNumber | OwnerType: $oldOwnerType" -ForegroundColor Cyan

        if ($DryRun) {
            $results += New-ResultObject `
                -Device $device `
                -OldOwnerType $oldOwnerType `
                -Changed $false `
                -Result "Dry-run; no change made"

            continue
        }

        try {
            Set-IntuneDeviceCorporateOwnership -Device $device | Out-Null

            Write-Host "Changed to Corporate." -ForegroundColor Green

            $results += New-ResultObject `
                -Device $device `
                -OldOwnerType $oldOwnerType `
                -Changed $true `
                -Result "Changed to Corporate"
        }
        catch {
            Write-Host "ERROR while changing device: $($_.Exception.Message)" -ForegroundColor Red

            $results += New-ResultObject `
                -Device $device `
                -OldOwnerType $oldOwnerType `
                -Changed $false `
                -Result "Error" `
                -ErrorMessage $_.Exception.Message
        }
    }

    $alreadyCorporateDevices = $allDevices | Where-Object {
        $ownerType = Get-SafePropertyValue -InputObject $_ -PropertyName "managedDeviceOwnerType" -DefaultValue ""
        $ownerType -eq "company"
    }

    foreach ($device in $alreadyCorporateDevices) {
        $results += New-ResultObject `
            -Device $device `
            -OldOwnerType "company" `
            -Changed $false `
            -Result "Already Corporate"
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
