#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Fetches Microsoft 365 endpoints and creates firewall list files.

.DESCRIPTION
    This script fetches Microsoft 365 endpoints from the official Microsoft API
    and generates list files that can be used in firewall configurations.
    
    Two types of files are generated:
    
    1. Category-based files: ms365_{{serviceArea}}_{{addrType}}_{{category}}.txt
       where:
       - serviceArea: common, exchange, sharepoint, teams, etc.
       - addrType: url, ipv4, ipv6
       - category: opt, allow, default
    
    2. Port-based files: ms365_{{serviceArea}}_{{addrType}}_port{{ports}}.txt
       where:
       - serviceArea: common, exchange, sharepoint, teams, etc.
       - addrType: url, ipv4, ipv6
       - ports: port numbers separated by hyphens (e.g., 25, 80-443, 143-587-993-995)
       
       These files contain the same IPs or URLs but are organized by the TCP ports
       they use, making it easier to configure port-specific firewall rules.

.PARAMETER OutputDirectory
    Directory where the list files will be saved. Default is './lists'

.PARAMETER ClientRequestId
    Optional client request ID for API tracking. A random GUID is generated if not provided.

.EXAMPLE
    .\Get-MSEndpoints.ps1
    .\Get-MSEndpoints.ps1 -OutputDirectory "./output"
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$OutputDirectory = "./lists",
    
    [Parameter(Mandatory = $false)]
    [string]$ClientRequestId = [guid]::NewGuid().ToString()
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Create output directory if it doesn't exist
if (-not (Test-Path -Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    Write-Host "Created output directory: $OutputDirectory"
}

# Clean up output directory - remove all existing files before generating new ones
if (Test-Path -Path $OutputDirectory) {
    $existingFiles = Get-ChildItem -Path $OutputDirectory -Filter "*.txt"
    
    if ($existingFiles.Count -gt 0) {
        Write-Host "Removing $($existingFiles.Count) existing file(s) from output directory..."
        $existingFiles | Remove-Item -Force
    }
}

# Fetch endpoints from Microsoft API
$apiUrl = "https://endpoints.office.com/endpoints/worldwide?clientrequestid=$ClientRequestId"
Write-Host "Fetching endpoints from: $apiUrl"

try {
    $endpoints = Invoke-RestMethod -Uri $apiUrl -Method Get
    Write-Host "Successfully fetched $($endpoints.Count) endpoint entries"
}
catch {
    Write-Error "Failed to fetch endpoints: $_"
    exit 1
}

# Initialize hashtables to group data
$groupedData = @{}
$groupedDataByPort = @{}

# Process each endpoint
foreach ($endpoint in $endpoints) {
    # Determine category (Optimize=opt, Allow=allow, Default=default)
    $category = switch ($endpoint.category) {
        "Optimize" { "opt" }
        "Allow" { "allow" }
        "Default" { "default" }
        default { "default" }
    }
    
    # Get service area (normalize to lowercase)
    $serviceArea = if ($endpoint.serviceArea) { 
        $endpoint.serviceArea.ToLower() 
    } else { 
        "common" 
    }
    
    # Get port information for port-specific lists
    $tcpPorts = if ($endpoint.tcpPorts) { $endpoint.tcpPorts } else { $null }
    $udpPorts = if ($endpoint.udpPorts) { $endpoint.udpPorts } else { $null }
    
    # Normalize port format for filename: remove spaces, replace commas with hyphens
    $normalizedPorts = if ($tcpPorts) { 
        $tcpPorts -replace '\s+', '' -replace ',', '-'
    } else { 
        $null 
    }
    
    # Process URLs
    if ($endpoint.urls) {
        $key = "${serviceArea}_url_${category}"
        if (-not $groupedData.ContainsKey($key)) {
            $groupedData[$key] = @()
        }
        foreach ($url in $endpoint.urls) {
            if ($url -and $url.Trim() -ne "") {
                $groupedData[$key] += $url
                
                # Also add to port-specific lists if port info exists
                if ($normalizedPorts) {
                    $portKey = "${serviceArea}_url_port${normalizedPorts}"
                    if (-not $groupedDataByPort.ContainsKey($portKey)) {
                        $groupedDataByPort[$portKey] = @()
                    }
                    $groupedDataByPort[$portKey] += $url
                }
            }
        }
    }
    
    # Process IPv4 addresses
    if ($endpoint.ips) {
        foreach ($ip in $endpoint.ips) {
            # Check if it's IPv4 (contains dots but not colons)
            if ($ip -match '^\d+\.\d+\.\d+\.\d+' -and $ip -notmatch ':') {
                $key = "${serviceArea}_ipv4_${category}"
                if (-not $groupedData.ContainsKey($key)) {
                    $groupedData[$key] = @()
                }
                if ($ip -and $ip.Trim() -ne "") {
                    $groupedData[$key] += $ip
                    
                    # Also add to port-specific lists if port info exists
                    if ($normalizedPorts) {
                        $portKey = "${serviceArea}_ipv4_port${normalizedPorts}"
                        if (-not $groupedDataByPort.ContainsKey($portKey)) {
                            $groupedDataByPort[$portKey] = @()
                        }
                        $groupedDataByPort[$portKey] += $ip
                    }
                }
            }
            # Check if it's IPv6 (contains colons)
            elseif ($ip -match ':') {
                $key = "${serviceArea}_ipv6_${category}"
                if (-not $groupedData.ContainsKey($key)) {
                    $groupedData[$key] = @()
                }
                if ($ip -and $ip.Trim() -ne "") {
                    $groupedData[$key] += $ip
                    
                    # Also add to port-specific lists if port info exists
                    if ($normalizedPorts) {
                        $portKey = "${serviceArea}_ipv6_port${normalizedPorts}"
                        if (-not $groupedDataByPort.ContainsKey($portKey)) {
                            $groupedDataByPort[$portKey] = @()
                        }
                        $groupedDataByPort[$portKey] += $ip
                    }
                }
            }
        }
    }
}

# Write data to files (original format by category)
$fileCount = 0
foreach ($key in $groupedData.Keys | Sort-Object) {
    # Remove duplicates and sort
    $uniqueData = $groupedData[$key] | Select-Object -Unique | Sort-Object
    
    if ($uniqueData.Count -gt 0) {
        $fileName = "ms365_$key.txt"
        $filePath = Join-Path -Path $OutputDirectory -ChildPath $fileName
        
        # Write to file
        $uniqueData | Out-File -FilePath $filePath -Encoding UTF8 -Force
        
        Write-Host "Created: $fileName ($($uniqueData.Count) entries)"
        $fileCount++
    }
}

# Write data to files (port-specific format)
foreach ($key in $groupedDataByPort.Keys | Sort-Object) {
    # Remove duplicates and sort
    $uniqueData = $groupedDataByPort[$key] | Select-Object -Unique | Sort-Object
    
    if ($uniqueData.Count -gt 0) {
        $fileName = "ms365_$key.txt"
        $filePath = Join-Path -Path $OutputDirectory -ChildPath $fileName
        
        # Write to file
        $uniqueData | Out-File -FilePath $filePath -Encoding UTF8 -Force
        
        Write-Host "Created: $fileName ($($uniqueData.Count) entries)"
        $fileCount++
    }
}

Write-Host "`nTotal files created: $fileCount"
Write-Host "Output directory: $OutputDirectory"
