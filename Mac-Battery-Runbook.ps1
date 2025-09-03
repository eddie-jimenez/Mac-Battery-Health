<#
.SYNOPSIS
    macOS Battery Health Report - Azure Automation Runbook
    
.DESCRIPTION
    Generates comprehensive battery health reports for macOS devices in Intune using 
    Custom Attribute script data. Creates HTML report with professional styling 
    and email delivery with CSV attachment.
    
.PARAMETER EmailRecipient
    Comma-separated list of email addresses to receive battery health reports.
    Default: 'test-email@contoso.com'
    
.PARAMETER CustomAttributeId
    GUID of the Custom Attribute script in Intune for battery health data.
    Default: 'YOUR CUSTOM ATTRIBUTE GUID'
    
.PARAMETER MinHealthThreshold
    Minimum battery health percentage to highlight devices needing attention.
    Default: 80
    
.NOTES
    Prerequisites:
    - Uses Azure Managed Identity
    - Required permissions:
      * DeviceManagementConfiguration.Read.All
      * DeviceManagementManagedDevices.Read.All  
      * Mail.Send
    - Custom Attribute script deployed to macOS devices (see https://github.com/eddie-jimenez/Mac-Battery-Health)
    
.AUTHOR
    Eddie Jimenez @eddie-jimenez https://github.com/eddie-jimenez
    
.VERSION
    1.1 - Fixed all parsing issues
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$EmailRecipient = 'test-email@contoso.com',
    
    [Parameter(Mandatory = $false)]
    [string]$CustomAttributeId = '00000000-0000-0000-0000-000000000000', #Your custom attribute ID
    
    [Parameter(Mandatory = $false)]
    [int]$MinHealthThreshold = 80
)

# Store Start Time
$scriptStart = Get-Date

# Environment check
if (-not $PSPrivateMetadata.JobId.Guid) {
    Write-Error "This script requires Azure Automation"
    exit 1
}

Write-Output "Running Battery Health Report in Azure Automation"

# Import required modules
Write-Output "=== Module Setup ==="
try {
    Import-Module Microsoft.Graph.Authentication -Force
    Import-Module Microsoft.Graph.Mail -Force
    Write-Output "Modules imported successfully"
} catch {
    Write-Error "Module import failed: $_"
    exit 1
}

# Get authentication token
Write-Output "=== Authentication ==="
try {
    $resourceURI = "https://graph.microsoft.com"
    $tokenAuthURI = $env:IDENTITY_ENDPOINT + "?resource=$resourceURI&api-version=2019-08-01"
    $tokenResponse = Invoke-RestMethod -Method Get -Headers @{"X-IDENTITY-HEADER" = $env:IDENTITY_HEADER} -Uri $tokenAuthURI
    $accessToken = $tokenResponse.access_token
    
    $headers = @{
        "Authorization" = "Bearer $accessToken"
        "Content-Type" = "application/json"
    }
    
    Write-Output "Authentication successful"
    
} catch {
    Write-Error "Authentication failed: $_"
    exit 1
}

# Function to get all pages
function Get-AllPages {
    param([string]$Uri)
    
    $results = New-Object System.Collections.ArrayList
    $nextUri = $Uri
    
    do {
        try {
            $response = Invoke-RestMethod -Uri $nextUri -Headers $headers -Method GET
            
            if ($response.value) {
                foreach ($item in $response.value) {
                    $null = $results.Add($item)
                }
            }
            
            $nextUri = $response.'@odata.nextLink'
            
            if ($nextUri) {
                Start-Sleep -Milliseconds 200
            }
            
        } catch {
            Write-Warning "Error getting $nextUri : $_"
            break
        }
    } while ($nextUri)
    
    return $results
}

# Function to convert battery data to CSV format
function ConvertTo-BatteryCSV {
    param($BatteryData)
    
    $csv = "Device,User,Health%,Cycles,Full_mAh,Design_mAh,Current_mAh,Charging,Ext_Power,Min_Left,mV,Condition,Updated`n"
    
    foreach ($row in $BatteryData) {
        # Escape fields that might contain commas or quotes
        $device = ($row.DeviceName -replace '[",]', '_')
        $user = if ($row.UserPrincipalName) { ($row.UserPrincipalName -replace '[",]', '_') } else { "" }
        $health = if ($row.HealthPercent) { $row.HealthPercent } else { "" }
        $cycles = if ($row.CycleCount) { $row.CycleCount } else { "" }
        $full = if ($row.FullCharge_mAh) { $row.FullCharge_mAh } else { "" }
        $design = if ($row.Design_mAh) { $row.Design_mAh } else { "" }
        $current = if ($row.Current_mAh) { $row.Current_mAh } else { "" }
        $charging = if ($row.IsCharging -eq $true) { "Yes" } elseif ($row.IsCharging -eq $false) { "No" } else { "" }
        $extPower = if ($row.ExternalPower -eq $true) { "Yes" } elseif ($row.ExternalPower -eq $false) { "No" } else { "" }
        $minLeft = if ($row.TimeRemaining) { $row.TimeRemaining } else { "" }
        $voltage = if ($row.Voltage_mV) { $row.Voltage_mV } else { "" }
        $condition = if ($row.Condition) { ($row.Condition -replace '[",]', '_') } else { "" }
        $updated = if ($row.LastUpdate) { [datetime]::Parse($row.LastUpdate).ToString('yyyy-MM-dd HH:mm:ss') } else { "" }
        
        $csv += "$device,$user,$health,$cycles,$full,$design,$current,$charging,$extPower,$minLeft,$voltage,$condition,$updated`n"
    }
    
    return $csv
}

# Main execution
try {
    Write-Output "=== Fetching Battery Health Data ==="
    
    # Get Custom Attribute run states
    $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceCustomAttributeShellScripts/$CustomAttributeId/deviceRunStates?`$select=id,lastStateUpdateDateTime,resultMessage,runState,errorCode,errorDescription&`$expand=managedDevice(`$select=id,deviceName,userPrincipalName,osVersion,userId,model)&`$top=200"
    
    $runStates = Get-AllPages -Uri $uri
    Write-Output "Found $($runStates.Count) device reports"
    
    # Parse battery data
    $batteryData = @()
    
    # Known desktop Mac model identifiers
    $desktopModels = @(
        'Mac13,1', 'Mac13,2', 'Mac14,3', 'Mac14,12', 'Mac14,13', 'Mac14,14',
        'Mac15,4', 'Mac15,5', 'Mac14,8', 'iMac21,1', 'iMac21,2', 'Macmini9,1'
    )
    
    foreach ($state in $runStates) {
        if ($state.managedDevice -and $state.resultMessage) {
            # Check if device model is a desktop Mac
            $deviceModel = $state.managedDevice.model
            if ($deviceModel -and $deviceModel -in $desktopModels) {
                Write-Output "Skipping desktop Mac: $($state.managedDevice.deviceName) (Model: $deviceModel)"
                continue
            }
            
            $csv = $state.resultMessage.Trim()
            
            # Skip desktop Macs (no battery)
            if ($csv -eq "None") {
                continue
            }
            
            $parts = $csv -split ',' | ForEach-Object { $_.Trim() }
            
            # Parse CSV fields (11 fields expected)
            if ($parts.Count -ge 11) {
                # Skip if all battery values are None
                if ($parts[0] -eq 'None' -and $parts[1] -eq 'None' -and $parts[2] -eq 'None') {
                    continue
                }
                
                $battery = [PSCustomObject]@{
                    DeviceName = $state.managedDevice.deviceName
                    UserPrincipalName = $state.managedDevice.userPrincipalName
                    DeviceModel = $deviceModel
                    HealthPercent = if ($parts[0] -ne 'None' -and $parts[0] -match '^\d+$') { [int]$parts[0] } else { $null }
                    CycleCount = if ($parts[1] -ne 'None' -and $parts[1] -match '^\d+$') { [int]$parts[1] } else { $null }
                    FullCharge_mAh = if ($parts[2] -ne 'None' -and $parts[2] -match '^\d+$') { [int]$parts[2] } else { $null }
                    Design_mAh = if ($parts[3] -ne 'None' -and $parts[3] -match '^\d+$') { [int]$parts[3] } else { $null }
                    Current_mAh = if ($parts[4] -ne 'None' -and $parts[4] -match '^\d+$') { [int]$parts[4] } else { $null }
                    IsCharging = $parts[5] -eq 'True'
                    ExternalPower = $parts[6] -eq 'True'
                    TimeRemaining = if ($parts[7] -ne 'None' -and $parts[7] -match '^\d+$') { [int]$parts[7] } else { $null }
                    Voltage_mV = if ($parts[8] -ne 'None' -and $parts[8] -match '^\d+$') { [int]$parts[8] } else { $null }
                    Condition = if ($parts[9] -ne 'None') { $parts[9] } else { 'Unknown' }
                    OverThreshold = $parts[10] -eq 'True'
                    LastUpdate = $state.lastStateUpdateDateTime
                    RunState = $state.runState
                }
                
                # Only add devices that actually have battery data
                if ($battery.HealthPercent -or $battery.CycleCount -or $battery.FullCharge_mAh) {
                    $batteryData += $battery
                }
            }
        }
    }
    
    Write-Output "Parsed $($batteryData.Count) battery reports (MacBooks only)"
    
    # Calculate statistics
    $validHealthData = $batteryData | Where-Object { $null -ne $_.HealthPercent }
    $avgHealth = if ($validHealthData) { 
        [math]::Round(($validHealthData.HealthPercent | Measure-Object -Average).Average, 1) 
    } else { 0 }
    
    $over1000Cycles = ($batteryData | Where-Object { $_.CycleCount -ge 1000 }).Count
    $onExternalPower = ($batteryData | Where-Object { $_.ExternalPower -eq $true }).Count
    $belowThreshold = ($validHealthData | Where-Object { $_.HealthPercent -lt $MinHealthThreshold }).Count
    $chargingNow = ($batteryData | Where-Object { $_.IsCharging -eq $true }).Count
    $needsService = ($batteryData | Where-Object { $_.Condition -in @("Replace Soon", "Replace Now", "Service Battery") }).Count
    $below70 = ($validHealthData | Where-Object { $_.HealthPercent -lt 70 }).Count
    
    # Sort devices by health (worst first)
    $sortedDevices = $batteryData | Sort-Object { 
        if ($null -eq $_.HealthPercent) { 999 } else { $_.HealthPercent }
    }
    
    # Calculate duration
    $scriptEnd = Get-Date
    $duration = $scriptEnd - $scriptStart
    $durationFormatted = "{0:hh\:mm\:ss}" -f $duration
    
    # Build HTML report
    $htmlReport = New-Object System.Text.StringBuilder
    
    # HTML Header
    [void]$htmlReport.AppendLine('<!DOCTYPE html>')
    [void]$htmlReport.AppendLine('<html lang="en">')
    [void]$htmlReport.AppendLine('<head>')
    [void]$htmlReport.AppendLine('<meta charset="UTF-8">')
    [void]$htmlReport.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1.0">')
    [void]$htmlReport.AppendLine('<title>macOS Battery Health Report</title>')
    [void]$htmlReport.AppendLine('<style>')
    [void]$htmlReport.AppendLine('* { margin: 0; padding: 0; box-sizing: border-box; }')
    [void]$htmlReport.AppendLine('body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Oxygen, Ubuntu, sans-serif; background: linear-gradient(135deg, #1e1e2e 0%, #2d2d44 100%); color: #e0e0e0; padding: 20px; min-height: 100vh; }')
    [void]$htmlReport.AppendLine('.container { max-width: 1400px; margin: 0 auto; background: rgba(30, 30, 46, 0.95); border-radius: 20px; padding: 30px; box-shadow: 0 20px 60px rgba(0, 0, 0, 0.5); }')
    [void]$htmlReport.AppendLine('.header { display: flex; align-items: center; gap: 15px; margin-bottom: 30px; padding-bottom: 20px; border-bottom: 1px solid rgba(255, 255, 255, 0.1); }')
    [void]$htmlReport.AppendLine('.header-icon { width: 40px; height: 40px; background: linear-gradient(135deg, #4ade80, #22c55e); border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 24px; }')
    [void]$htmlReport.AppendLine('h1 { font-size: 28px; font-weight: 600; background: linear-gradient(135deg, #4ade80, #22c55e); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }')
    [void]$htmlReport.AppendLine('.export-info { margin-left: auto; font-size: 14px; color: #9ca3af; }')
    [void]$htmlReport.AppendLine('.kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }')
    [void]$htmlReport.AppendLine('.kpi-card { background: linear-gradient(135deg, rgba(74, 222, 128, 0.15), transparent); border-radius: 14px; padding: 20px; border: 1px solid rgba(74, 222, 128, 0.2); }')
    [void]$htmlReport.AppendLine('.kpi-title { font-size: 12px; color: #9ca3af; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px; }')
    [void]$htmlReport.AppendLine('.kpi-value { font-size: 32px; font-weight: 700; color: #4ade80; }')
    [void]$htmlReport.AppendLine('.stats-section { background: rgba(255, 255, 255, 0.05); border-radius: 14px; padding: 20px; margin-bottom: 30px; }')
    [void]$htmlReport.AppendLine('.stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; }')
    [void]$htmlReport.AppendLine('.stat-item { padding: 10px; background: rgba(255, 255, 255, 0.03); border-radius: 8px; border-left: 3px solid #4ade80; }')
    [void]$htmlReport.AppendLine('.stat-label { font-size: 11px; color: #9ca3af; text-transform: uppercase; letter-spacing: 0.5px; }')
    [void]$htmlReport.AppendLine('.stat-value { font-size: 20px; font-weight: 600; color: #e0e0e0; margin-top: 4px; }')
    [void]$htmlReport.AppendLine('.alert-section { background: rgba(239, 68, 68, 0.1); border: 1px solid rgba(239, 68, 68, 0.3); border-radius: 8px; padding: 15px; margin-bottom: 20px; }')
    [void]$htmlReport.AppendLine('.alert-title { color: #ef4444; font-weight: 600; margin-bottom: 10px; }')
    [void]$htmlReport.AppendLine('.alert-list { list-style: none; padding: 0; }')
    [void]$htmlReport.AppendLine('.alert-item { padding: 5px 0; color: #fca5a5; }')
    [void]$htmlReport.AppendLine('.data-table { background: rgba(255, 255, 255, 0.03); border-radius: 14px; overflow: hidden; max-height: 600px; overflow-y: auto; }')
    [void]$htmlReport.AppendLine('table { width: 100%; border-collapse: collapse; }')
    [void]$htmlReport.AppendLine('th { background: rgba(30, 30, 46, 0.95); padding: 12px; text-align: left; font-size: 12px; font-weight: 600; color: #9ca3af; text-transform: uppercase; letter-spacing: 1px; border-bottom: 2px solid rgba(74, 222, 128, 0.3); position: sticky; top: 0; }')
    [void]$htmlReport.AppendLine('td { padding: 12px; font-size: 14px; border-bottom: 1px solid rgba(255, 255, 255, 0.05); }')
    [void]$htmlReport.AppendLine('tbody tr:hover { background: rgba(255, 255, 255, 0.05); }')
    [void]$htmlReport.AppendLine('.pill { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 500; }')
    [void]$htmlReport.AppendLine('.pill-yes { background: rgba(74, 222, 128, 0.2); color: #4ade80; border: 1px solid rgba(74, 222, 128, 0.3); }')
    [void]$htmlReport.AppendLine('.pill-no { background: rgba(239, 68, 68, 0.2); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.3); }')
    [void]$htmlReport.AppendLine('.health-badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-weight: 600; }')
    [void]$htmlReport.AppendLine('.health-good { background: rgba(74, 222, 128, 0.2); color: #4ade80; }')
    [void]$htmlReport.AppendLine('.health-warning { background: rgba(251, 191, 36, 0.2); color: #fbbf24; }')
    [void]$htmlReport.AppendLine('.health-danger { background: rgba(239, 68, 68, 0.2); color: #ef4444; }')
    [void]$htmlReport.AppendLine('.text-right { text-align: right; }')
    [void]$htmlReport.AppendLine('.text-center { text-align: center; }')
    [void]$htmlReport.AppendLine('.summary { margin-top: 30px; padding: 20px; background: rgba(255, 255, 255, 0.05); border-radius: 12px; font-size: 14px; color: #9ca3af; text-align: center; }')
    [void]$htmlReport.AppendLine('.attachment-note { margin-top: 20px; padding: 15px; background: rgba(74, 222, 128, 0.1); border: 1px solid rgba(74, 222, 128, 0.3); border-radius: 8px; text-align: center; }')
    [void]$htmlReport.AppendLine('.attachment-note strong { color: #4ade80; }')
    [void]$htmlReport.AppendLine('.docs-section { margin-top: 20px; padding: 15px; background: rgba(255, 255, 255, 0.03); border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 8px; text-align: center; }')
    [void]$htmlReport.AppendLine('.docs-section a { color: #4ade80; text-decoration: none; font-weight: 500; }')
    [void]$htmlReport.AppendLine('.docs-section a:hover { text-decoration: underline; }')
    [void]$htmlReport.AppendLine('</style>')
    [void]$htmlReport.AppendLine('</head>')
    [void]$htmlReport.AppendLine('<body>')
    [void]$htmlReport.AppendLine('<div class="container">')
    
    # Header
    [void]$htmlReport.AppendLine('<div class="header">')
    [void]$htmlReport.AppendLine('<div class="header-icon">⚡</div>')
    [void]$htmlReport.AppendLine('<h1>macOS Battery Health Report</h1>')
    [void]$htmlReport.AppendLine('<div class="export-info">')
    [void]$htmlReport.AppendLine("<div>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>")
    [void]$htmlReport.AppendLine("<div style='font-size: 12px; margin-top: 4px;'>MacBooks with Batteries: $($batteryData.Count)</div>")
    [void]$htmlReport.AppendLine('</div>')
    [void]$htmlReport.AppendLine('</div>')
    
    # Attachment note
    [void]$htmlReport.AppendLine('<div class="attachment-note">')
    [void]$htmlReport.AppendLine('📎 <strong>CSV data attached</strong> - Check email attachment for full dataset export')
    [void]$htmlReport.AppendLine('</div>')
    
    # KPI Cards
    [void]$htmlReport.AppendLine('<div class="kpi-row">')
    [void]$htmlReport.AppendLine('<div class="kpi-card"><div class="kpi-title">MacBooks</div><div class="kpi-value">' + $batteryData.Count + '</div></div>')
    [void]$htmlReport.AppendLine('<div class="kpi-card"><div class="kpi-title">Average Health</div><div class="kpi-value">' + $avgHealth + '%</div></div>')
    [void]$htmlReport.AppendLine('<div class="kpi-card"><div class="kpi-title">≥1000 Cycles</div><div class="kpi-value">' + $over1000Cycles + '</div></div>')
    [void]$htmlReport.AppendLine('<div class="kpi-card"><div class="kpi-title">On External Power</div><div class="kpi-value">' + $onExternalPower + '</div></div>')
    [void]$htmlReport.AppendLine('</div>')
    
    # Alert section if needed
    if ($belowThreshold -gt 0) {
        [void]$htmlReport.AppendLine('<div class="alert-section">')
        [void]$htmlReport.AppendLine("<div class='alert-title'>⚠️ Warning: $belowThreshold MacBooks below $MinHealthThreshold% health threshold</div>")
        [void]$htmlReport.AppendLine('<ul class="alert-list">')
        
        $alertDevices = $validHealthData | Where-Object { $_.HealthPercent -lt $MinHealthThreshold } | Sort-Object HealthPercent
        foreach ($device in ($alertDevices | Select-Object -First 5)) {
            [void]$htmlReport.AppendLine("<li class='alert-item'>$($device.DeviceName) - $($device.HealthPercent)% health, $($device.CycleCount) cycles</li>")
        }
        if ($belowThreshold -gt 5) {
            [void]$htmlReport.AppendLine("<li class='alert-item'>... and $($belowThreshold - 5) more</li>")
        }
        [void]$htmlReport.AppendLine('</ul>')
        [void]$htmlReport.AppendLine('</div>')
    }
    
    # Statistics section
    [void]$htmlReport.AppendLine('<div class="stats-section">')
    [void]$htmlReport.AppendLine('<div style="font-size: 14px; color: #9ca3af; margin-bottom: 15px; text-transform: uppercase; letter-spacing: 1px;">MACBOOK BATTERY STATISTICS</div>')
    [void]$htmlReport.AppendLine('<div class="stats-grid">')
    [void]$htmlReport.AppendLine('<div class="stat-item"><div class="stat-label">MacBooks with Batteries</div><div class="stat-value">' + $batteryData.Count + '</div></div>')
    [void]$htmlReport.AppendLine('<div class="stat-item"><div class="stat-label">Below 70% Health</div><div class="stat-value" style="color: ' + $(if ($below70 -gt 0) { '#ef4444' } else { '#4ade80' }) + ';">' + $below70 + '</div></div>')
    [void]$htmlReport.AppendLine('<div class="stat-item"><div class="stat-label">Need Service</div><div class="stat-value" style="color: ' + $(if ($needsService -gt 0) { '#fbbf24' } else { '#4ade80' }) + ';">' + $needsService + '</div></div>')
    [void]$htmlReport.AppendLine('<div class="stat-item"><div class="stat-label">Currently Charging</div><div class="stat-value">' + $chargingNow + '</div></div>')
    [void]$htmlReport.AppendLine('</div>')
    [void]$htmlReport.AppendLine('</div>')
    
    # Data table
    [void]$htmlReport.AppendLine('<div class="data-table">')
    [void]$htmlReport.AppendLine('<table>')
    [void]$htmlReport.AppendLine('<thead>')
    [void]$htmlReport.AppendLine('<tr>')
    [void]$htmlReport.AppendLine('<th>Device</th>')
    [void]$htmlReport.AppendLine('<th>User</th>')
    [void]$htmlReport.AppendLine('<th class="text-right">Health%</th>')
    [void]$htmlReport.AppendLine('<th class="text-right">Cycles</th>')
    [void]$htmlReport.AppendLine('<th class="text-right">Full mAh</th>')
    [void]$htmlReport.AppendLine('<th class="text-right">Design mAh</th>')
    [void]$htmlReport.AppendLine('<th class="text-center">Charging</th>')
    [void]$htmlReport.AppendLine('<th class="text-center">Ext Pwr</th>')
    [void]$htmlReport.AppendLine('<th>Condition</th>')
    [void]$htmlReport.AppendLine('<th>Updated</th>')
    [void]$htmlReport.AppendLine('</tr>')
    [void]$htmlReport.AppendLine('</thead>')
    [void]$htmlReport.AppendLine('<tbody>')
    
    foreach ($device in ($sortedDevices | Select-Object -First 50)) {
        $healthClass = if ($device.HealthPercent) {
            if ($device.HealthPercent -ge 90) { "health-good" }
            elseif ($device.HealthPercent -ge 70) { "health-warning" }
            else { "health-danger" }
        } else { "" }
        
        $healthDisplay = if ($device.HealthPercent) { 
            "<span class='health-badge $healthClass'>$($device.HealthPercent)%</span>" 
        } else { "-" }
        
        $chargingDisplay = if ($device.IsCharging -eq $true) { 
            "<span class='pill pill-yes'>Yes</span>" 
        } elseif ($device.IsCharging -eq $false) { 
            "<span class='pill pill-no'>No</span>" 
        } else { "-" }
        
        $extPowerDisplay = if ($device.ExternalPower -eq $true) { 
            "<span class='pill pill-yes'>Yes</span>" 
        } elseif ($device.ExternalPower -eq $false) { 
            "<span class='pill pill-no'>No</span>" 
        } else { "-" }
        
        $updated = if ($device.LastUpdate) {
            [datetime]::Parse($device.LastUpdate).ToString('yyyy-MM-dd HH:mm')
        } else { "-" }
        
        [void]$htmlReport.AppendLine('<tr>')
        [void]$htmlReport.AppendLine("<td>$($device.DeviceName)</td>")
        [void]$htmlReport.AppendLine("<td>$($device.UserPrincipalName)</td>")
        [void]$htmlReport.AppendLine("<td class='text-right'>$healthDisplay</td>")
        [void]$htmlReport.AppendLine("<td class='text-right'>$(if ($device.CycleCount) { $device.CycleCount } else { '-' })</td>")
        [void]$htmlReport.AppendLine("<td class='text-right'>$(if ($device.FullCharge_mAh) { $device.FullCharge_mAh } else { '-' })</td>")
        [void]$htmlReport.AppendLine("<td class='text-right'>$(if ($device.Design_mAh) { $device.Design_mAh } else { '-' })</td>")
        [void]$htmlReport.AppendLine("<td class='text-center'>$chargingDisplay</td>")
        [void]$htmlReport.AppendLine("<td class='text-center'>$extPowerDisplay</td>")
        [void]$htmlReport.AppendLine("<td>$($device.Condition)</td>")
        [void]$htmlReport.AppendLine("<td>$updated</td>")
        [void]$htmlReport.AppendLine('</tr>')
    }
    
    [void]$htmlReport.AppendLine('</tbody>')
    [void]$htmlReport.AppendLine('</table>')
    [void]$htmlReport.AppendLine('</div>')
    
    # Apple documentation reference
    [void]$htmlReport.AppendLine('<div class="docs-section">')
    [void]$htmlReport.AppendLine('📖 For official battery health information, see <a href="https://support.apple.com/en-us/102888" target="_blank">Apple Support: Determine battery cycle count for Mac laptops</a>')
    [void]$htmlReport.AppendLine('</div>')
    
    # Summary footer
    [void]$htmlReport.AppendLine('<div class="summary">')
    [void]$htmlReport.AppendLine("Generated by Azure Automation - $($batteryData.Count) MacBooks analyzed - Runtime: $durationFormatted")
    [void]$htmlReport.AppendLine('</div>')
    
    [void]$htmlReport.AppendLine('</div>')
    [void]$htmlReport.AppendLine('</body>')
    [void]$htmlReport.AppendLine('</html>')
    
    # Convert StringBuilder to string
    $htmlContent = $htmlReport.ToString()
    
    # Generate CSV data
    $csvContent = ConvertTo-BatteryCSV -BatteryData $sortedDevices
    $csvBytes = [System.Text.Encoding]::UTF8.GetBytes($csvContent)
    $csvBase64 = [Convert]::ToBase64String($csvBytes)
    $csvFileName = "battery_health_$(Get-Date -Format 'MM-dd-yyyy').csv"
    
    Write-Output "CSV file size: $([math]::Round($csvBytes.Length / 1KB, 2)) KB"
    
    # Send email report with CSV attachment
    Write-Output "=== Sending Email Report with CSV ==="
    Connect-MgGraph -AccessToken (ConvertTo-SecureString -String $accessToken -AsPlainText -Force) -NoWelcome
    
    $emailRecipients = $EmailRecipient -split ',' | ForEach-Object { $_.Trim() }
    
    foreach ($recipient in $emailRecipients) {
        $message = @{
            subject = "🍎⚡ macOS Battery Health Report - $(Get-Date -Format 'MM-dd-yyyy')"
            body = @{
                contentType = "HTML"
                content = $htmlContent
            }
            toRecipients = @(
                @{ emailAddress = @{ address = $recipient } }
            )
            attachments = @(
                @{
                    "@odata.type" = "#microsoft.graph.fileAttachment"
                    name = $csvFileName
                    contentType = "text/csv"
                    contentBytes = $csvBase64
                }
            )
        }
        
        # Check if CSV is too large
        if ($csvBytes.Length -gt 3145728) {
            Write-Warning "CSV file too large for direct attachment ($([math]::Round($csvBytes.Length / 1MB, 2)) MB). Sending without attachment."
            $message.Remove("attachments")
        }
        
        $requestBody = @{ 
            message = $message
            saveToSentItems = $true
        } | ConvertTo-Json -Depth 10
        
        $emailUri = "https://graph.microsoft.com/v1.0/users/YOUR-AUTOMATION-EMAIL/sendMail" # Update with your automation email account
        Invoke-MgGraphRequest -Uri $emailUri -Method POST -Body $requestBody -ContentType "application/json"
    }
    
    Write-Output "Battery health report sent successfully with CSV attachment"
    
    # Final summary
    Write-Output ""
    Write-Output "BATTERY HEALTH REPORT COMPLETE"
    Write-Output "=============================="
    Write-Output "Total MacBooks: $($batteryData.Count)"
    Write-Output "Average Health: $avgHealth%"
    Write-Output "Below Threshold: $belowThreshold devices"
    Write-Output "CSV Size: $([math]::Round($csvBytes.Length / 1KB, 2)) KB"
    Write-Output "Duration: $durationFormatted"
    Write-Output ""
    
} catch {
    Write-Error "Battery health report failed: $_"
    
    # Send failure notification
    try {
        Connect-MgGraph -AccessToken (ConvertTo-SecureString -String $accessToken -AsPlainText -Force) -NoWelcome
        
        $failureBody = "<html><body style='font-family: Segoe UI, sans-serif;'>"
        $failureBody += "<h2 style='color: #dc3545;'>Battery Health Report Failed</h2>"
        $failureBody += "<p><strong>Error:</strong> $_</p>"
        $failureBody += "<p><strong>Time:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>"
        $failureBody += "<p>Please check the Azure Automation logs for details.</p>"
        $failureBody += "</body></html>"
        
        $message = @{
            subject = "❌ Battery Health Report Failed - $(Get-Date -Format 'yyyy-MM-dd')"
            body = @{
                contentType = "HTML"
                content = $failureBody
            }
            toRecipients = @(
                @{ emailAddress = @{ address = $EmailRecipient.Split(',')[0].Trim() } }
            )
        }
        
        $requestBody = @{ message = $message } | ConvertTo-Json -Depth 10
        $emailUri = "https://graph.microsoft.com/v1.0/users/YOUR-AUTOMATION-EMAIL/sendMail" # Update with your automation email account
        Invoke-MgGraphRequest -Uri $emailUri -Method POST -Body $requestBody -ContentType "application/json"
        
    } catch {
        Write-Error "Failed to send failure notification: $_"
    }
    
    exit 1
}