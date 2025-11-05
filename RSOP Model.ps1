# RSOP Modeling + Full GPO Configuration Report
# Requires: RSAT/GPMC and GroupPolicy module
# Written by Nick Pylarinos
# Script v1.0 contains advanced GPO modeling with full GPO configurations using my own lab domain details change to suit your environmnent
# Script is provided "as-is" without any warranties.

# Scenario variables
$domain                 = "skynet.local"
$userSamAccountName     = "nickosp"
$computerSamAccountName = "SCCMTARGET$"
$userOU                 = "OU=TestUsers,DC=SKYNET,DC=local"
$computerOU             = "OU=Desktop,OU=Citrix,DC=SKYNET,DC=local"
$dcName                 = "NPADC02.skynet.local"
$outputFile             = "$env:TEMP\RSOP_Full_Report.html"

Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host "GROUP POLICY MODELING WITH FULL GPO CONFIGURATIONS" -ForegroundColor Cyan
Write-Host "=" * 80 -ForegroundColor Cyan

# Load GroupPolicy module
try {
    Import-Module GroupPolicy -ErrorAction Stop
    Write-Host "[OK] GroupPolicy module loaded" -ForegroundColor Green
} catch {
    Write-Warning "Could not import GroupPolicy module. Report will be limited."
}

# Load GPMC COM
try {
    $gpm = New-Object -ComObject GPMgmt.GPM
    Write-Host "[OK] GPMC COM object loaded" -ForegroundColor Green
} catch {
    Write-Error "Failed to load GPMC COM. Ensure RSAT/GPMC is installed."
    exit 1
}

$constants = $gpm.GetConstants()

# Connect to domain
try {
    $gpmDomain = $gpm.GetDomain($domain, "", $constants.UseAnyDC)
    Write-Host "[OK] Connected to domain: $domain" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to domain: $domain"
    exit 1
}

Write-Host "`nModeling scenario:" -ForegroundColor Yellow
Write-Host "  User:        $userSamAccountName" -ForegroundColor White
Write-Host "  Computer:    $($computerSamAccountName -replace '\$','')" -ForegroundColor White
Write-Host "  User OU:     $userOU" -ForegroundColor White
Write-Host "  Computer OU: $computerOU" -ForegroundColor White

# Get GPOs from both OUs
Write-Host "`nRetrieving applied GPOs..." -ForegroundColor Yellow

$allGpos = @()

# User GPOs
try {
    $userSOM = $gpmDomain.GetSOM($userOU)
    $userLinks = $userSOM.GetGPOLinks()
    $userInherited = $userSOM.GetInheritedGPOLinks()
    
    foreach ($link in $userLinks) {
        $gpoId = $link.GPOID
        if ($gpoId) {
            try {
                $gpo = $gpmDomain.GetGPO($gpoId)
                $allGpos += [PSCustomObject]@{
                    DisplayName = $gpo.DisplayName
                    Guid = $gpoId
                    Scope = 'User'
                    Source = 'Direct Link'
                    LinkOrder = $link.SOMLinkOrder
                    Enforced = $link.Enforced
                }
            } catch {
                Write-Host "  Warning: Could not retrieve GPO $gpoId" -ForegroundColor DarkYellow
            }
        }
    }
    
    foreach ($link in $userInherited) {
        $gpoId = $link.GPOID
        if ($gpoId) {
            try {
                $gpo = $gpmDomain.GetGPO($gpoId)
                $allGpos += [PSCustomObject]@{
                    DisplayName = $gpo.DisplayName
                    Guid = $gpoId
                    Scope = 'User'
                    Source = "Inherited: $($link.SOM.Path)"
                    LinkOrder = 999
                    Enforced = $link.Enforced
                }
            } catch {
                Write-Host "  Warning: Could not retrieve inherited GPO $gpoId" -ForegroundColor DarkYellow
            }
        }
    }
    
    Write-Host "[OK] Retrieved $($userLinks.Count) direct + $($userInherited.Count) inherited User GPOs" -ForegroundColor Green
} catch {
    Write-Warning "Could not retrieve user GPOs: $_"
}

# Computer GPOs
try {
    $computerSOM = $gpmDomain.GetSOM($computerOU)
    $computerLinks = $computerSOM.GetGPOLinks()
    $computerInherited = $computerSOM.GetInheritedGPOLinks()
    
    foreach ($link in $computerLinks) {
        $gpoId = $link.GPOID
        if ($gpoId) {
            try {
                $gpo = $gpmDomain.GetGPO($gpoId)
                $allGpos += [PSCustomObject]@{
                    DisplayName = $gpo.DisplayName
                    Guid = $gpoId
                    Scope = 'Computer'
                    Source = 'Direct Link'
                    LinkOrder = $link.SOMLinkOrder
                    Enforced = $link.Enforced
                }
            } catch {
                Write-Host "  Warning: Could not retrieve GPO $gpoId" -ForegroundColor DarkYellow
            }
        }
    }
    
    foreach ($link in $computerInherited) {
        $gpoId = $link.GPOID
        if ($gpoId) {
            try {
                $gpo = $gpmDomain.GetGPO($gpoId)
                $allGpos += [PSCustomObject]@{
                    DisplayName = $gpo.DisplayName
                    Guid = $gpoId
                    Scope = 'Computer'
                    Source = "Inherited: $($link.SOM.Path)"
                    LinkOrder = 999
                    Enforced = $link.Enforced
                }
            } catch {
                Write-Host "  Warning: Could not retrieve inherited GPO $gpoId" -ForegroundColor DarkYellow
            }
        }
    }
    
    Write-Host "[OK] Retrieved $($computerLinks.Count) direct + $($computerInherited.Count) inherited Computer GPOs" -ForegroundColor Green
} catch {
    Write-Warning "Could not retrieve computer GPOs: $_"
}

# Remove duplicates (same GPO might apply to both user and computer)
$uniqueGpos = $allGpos | Sort-Object Guid -Unique

Write-Host "`n[INFO] Total unique GPOs to process: $($uniqueGpos.Count)" -ForegroundColor Cyan

if ($uniqueGpos.Count -eq 0) {
    Write-Warning "No GPOs found. Check OUs and permissions."
    exit 1
}

# Generate consolidated HTML report
Write-Host "`n" + ("=" * 80) -ForegroundColor Green
Write-Host "GENERATING FULL GPO CONFIGURATION REPORT" -ForegroundColor Green
Write-Host ("=" * 80) -ForegroundColor Green

$sb = New-Object System.Text.StringBuilder

[void]$sb.AppendLine('<!DOCTYPE html>')
[void]$sb.AppendLine('<html><head><meta charset="utf-8">')
[void]$sb.AppendLine('<title>GPO Modeling Report - Full Configuration</title>')
[void]$sb.AppendLine('<style>')
[void]$sb.AppendLine('body{font-family:Segoe UI,Arial,sans-serif;margin:20px;background:#f5f5f5}')
[void]$sb.AppendLine('.header{background:#0066cc;color:white;padding:20px;margin-bottom:20px}')
[void]$sb.AppendLine('.info-box{background:#e7f3ff;border-left:4px solid #0066cc;padding:15px;margin:15px 0}')
[void]$sb.AppendLine('table{border-collapse:collapse;width:100%;margin:15px 0;background:white}')
[void]$sb.AppendLine('th{background:#0066cc;color:white;padding:12px;text-align:left}')
[void]$sb.AppendLine('td{padding:10px;border-bottom:1px solid #ddd}')
[void]$sb.AppendLine('tr:hover{background:#f0f8ff}')
[void]$sb.AppendLine('.enforced{color:red;font-weight:bold}')
[void]$sb.AppendLine('.gpo-section{background:white;padding:20px;margin:20px 0;box-shadow:0 2px 4px rgba(0,0,0,0.1)}')
[void]$sb.AppendLine('h1,h2,h3{margin-top:0}')
[void]$sb.AppendLine('code{background:#f6f8fa;padding:2px 6px;border-radius:3px}')
[void]$sb.AppendLine('</style></head><body>')

[void]$sb.AppendLine('<div class="header">')
[void]$sb.AppendLine('<h1>Group Policy Modeling Report</h1>')
[void]$sb.AppendLine("<p>Full Configuration of Applied GPOs</p>")
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="info-box">')
[void]$sb.AppendLine("<strong>Report Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')<br>")
[void]$sb.AppendLine("<strong>Domain:</strong> $domain<br><br>")
[void]$sb.AppendLine("<strong>User:</strong> $userSamAccountName<br>")
[void]$sb.AppendLine("<strong>User OU:</strong> $userOU<br><br>")
[void]$sb.AppendLine("<strong>Computer:</strong> $($computerSamAccountName -replace '\$','')<br>")
[void]$sb.AppendLine("<strong>Computer OU:</strong> $computerOU")
[void]$sb.AppendLine('</div>')

# Table of Contents
[void]$sb.AppendLine('<div class="gpo-section">')
[void]$sb.AppendLine('<h2>Applied GPOs Summary</h2>')
[void]$sb.AppendLine('<table>')
[void]$sb.AppendLine('<tr><th>#</th><th>GPO Name</th><th>Scope</th><th>Source</th><th>Link Order</th><th>GUID</th></tr>')

$idx = 0
foreach ($gpo in ($allGpos | Sort-Object Scope, LinkOrder)) {
    $idx++
    $enforcedMark = if ($gpo.Enforced) { ' <span class="enforced">(Enforced)</span>' } else { '' }
    [void]$sb.AppendLine("<tr>")
    [void]$sb.AppendLine("<td>$idx</td>")
    [void]$sb.AppendLine("<td><a href='#gpo-$($gpo.Guid)'>$($gpo.DisplayName)</a>$enforcedMark</td>")
    [void]$sb.AppendLine("<td>$($gpo.Scope)</td>")
    [void]$sb.AppendLine("<td>$($gpo.Source)</td>")
    [void]$sb.AppendLine("<td>$($gpo.LinkOrder)</td>")
    [void]$sb.AppendLine("<td><code>$($gpo.Guid)</code></td>")
    [void]$sb.AppendLine("</tr>")
}

[void]$sb.AppendLine('</table>')
[void]$sb.AppendLine('</div>')

# Full GPO configurations
$processedCount = 0
foreach ($gpo in $uniqueGpos) {
    $processedCount++
    Write-Host "  [$processedCount/$($uniqueGpos.Count)] Processing: $($gpo.DisplayName)" -ForegroundColor Cyan
    
    [void]$sb.AppendLine("<div class='gpo-section' id='gpo-$($gpo.Guid)'>")
    [void]$sb.AppendLine("<h2>$($gpo.DisplayName)</h2>")
    [void]$sb.AppendLine("<p><strong>GUID:</strong> <code>$($gpo.Guid)</code></p>")
    
    try {
        # Get full GPO report using Get-GPOReport
        $gpoReportPath = Join-Path $env:TEMP "gpo_$($gpo.Guid).html"
        Get-GPOReport -Guid $gpo.Guid -Domain $domain -ReportType Html -Path $gpoReportPath -ErrorAction Stop
        
        # Read and embed the report
        $gpoHtml = Get-Content -Path $gpoReportPath -Raw
        
        # Extract just the body content (strip html/head tags)
        if ($gpoHtml -match '<body[^>]*>(.*)</body>') {
            [void]$sb.AppendLine($matches[1])
        } else {
            [void]$sb.AppendLine($gpoHtml)
        }
        
        # Cleanup temp file
        Remove-Item -Path $gpoReportPath -Force -ErrorAction SilentlyContinue
        
    } catch {
        [void]$sb.AppendLine("<p style='color:red'>Failed to retrieve GPO configuration: $($_.Exception.Message)</p>")
    }
    
    [void]$sb.AppendLine('</div>')
}

[void]$sb.AppendLine('</body></html>')

# Save report
try {
    Set-Content -Path $outputFile -Value $sb.ToString() -Encoding UTF8
    Write-Host "`n[SUCCESS] Full report generated: $outputFile" -ForegroundColor Green
    Write-Host "Opening report..." -ForegroundColor Yellow
    Start-Process $outputFile
} catch {
    Write-Error "Failed to save report: $_"
}

Write-Host "`n" + ("=" * 80) -ForegroundColor Cyan
Write-Host "COMPLETE" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
