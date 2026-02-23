#Requires -Version 5.1
<#
.SYNOPSIS
    Network Test Tool - Comprehensive network diagnostics for L1/L2/L3 support
.DESCRIPTION
    Collects network details, runs connectivity tests, and generates a structured report
    for helpdesk and network engineering troubleshooting.
.NOTES
    Run as: .\NetworkTest.ps1
    Output: $env:USERPROFILE\<hostname>_NetTest_YYYYmmDD_HHMM.txt
#>

#region ── Configuration ──────────────────────────────────────────────────────

# Load external config file from same directory as this script
$configFile = Join-Path $PSScriptRoot "NetworkTest.config.ps1"
if (-not (Test-Path $configFile)) {
    Write-Host ""
    Write-Host "  [ERROR] Config file not found: $configFile" -ForegroundColor Red
    Write-Host "  Place NetworkTest.config.ps1 in the same folder as this script." -ForegroundColor Red
    Write-Host ""
    exit 1
}
. $configFile   # dot-source to load all $cfg_* variables into current scope

# Assemble into a single Config object for use throughout the script
$Config = @{
    InternalHosts        = $cfg_InternalHosts
    ExternalHosts        = $cfg_ExternalHosts
    InternalWebURLs      = $cfg_InternalWebURLs
    ExternalWebURLs      = $cfg_ExternalWebURLs
    CloudPlatforms       = $cfg_CloudPlatforms
    InternetSummaryNames = $cfg_InternetSummaryNames
    NslookupTestDomain   = $cfg_NslookupTestDomain
    PingCount            = $cfg_PingCount
    PingTimeoutMs        = $cfg_PingTimeoutMs
    WebTimeoutSec        = $cfg_WebTimeoutSec
    # Thresholds
    PingLatencyWarnMs    = $cfg_PingLatencyWarnMs
    PingLatencyFailMs    = $cfg_PingLatencyFailMs
    PingJitterWarnMs     = $cfg_PingJitterWarnMs
    PingJitterFailMs     = $cfg_PingJitterFailMs
    PingLossWarnPct      = $cfg_PingLossWarnPct
    PingLossFailPct      = $cfg_PingLossFailPct
    WebLatencyWarnMs     = $cfg_WebLatencyWarnMs
    WebLatencyFailMs     = $cfg_WebLatencyFailMs
}

#endregion

#region ── Helper Functions ───────────────────────────────────────────────────

function Write-Progress-Step {
    param([int]$Step, [int]$Total, [string]$Activity, [string]$Status)
    $pct = [int](($Step / $Total) * 100)
    Write-Progress -Activity $Activity -Status "[$Step/$Total] $Status" -PercentComplete $pct
    Write-Host "  >> $Status" -ForegroundColor Cyan
}

function Get-PingStats {
    param([string]$Target, [int]$Count = 10, [int]$TimeoutMs = 2000)
    $results = @{ Success = 0; Failed = 0; Latencies = @() }
    for ($i = 0; $i -lt $Count; $i++) {
        try {
            $ping = New-Object System.Net.NetworkInformation.Ping
            $reply = $ping.Send($Target, $TimeoutMs)
            if ($reply.Status -eq 'Success') {
                $results.Success++
                $results.Latencies += $reply.RoundtripTime
            } else {
                $results.Failed++
            }
        } catch {
            $results.Failed++
        }
    }
    $lats = $results.Latencies
    if ($lats.Count -gt 0) {
        $avg  = [math]::Round(($lats | Measure-Object -Average).Average, 1)
        $min  = ($lats | Measure-Object -Minimum).Minimum
        $max  = ($lats | Measure-Object -Maximum).Maximum
        # Jitter = mean absolute deviation of consecutive differences
        $diffs = @()
        for ($i = 1; $i -lt $lats.Count; $i++) { $diffs += [math]::Abs($lats[$i] - $lats[$i-1]) }
        $jitter = if ($diffs.Count -gt 0) { [math]::Round(($diffs | Measure-Object -Average).Average, 1) } else { 0 }
        $loss   = [math]::Round(($results.Failed / $Count) * 100, 0)
        return @{ Reachable=$true; Avg=$avg; Min=$min; Max=$max; Jitter=$jitter; Loss=$loss; Sent=$Count; Recv=$results.Success }
    } else {
        return @{ Reachable=$false; Avg="N/A"; Min="N/A"; Max="N/A"; Jitter="N/A"; Loss=100; Sent=$Count; Recv=0 }
    }
}

function Get-DNSLookup {
    param([string]$Target)
    try {
        $result = [System.Net.Dns]::GetHostEntry($Target)
        $ips = ($result.AddressList | ForEach-Object { $_.ToString() }) -join ", "
        return @{ Success=$true; Hostname=$result.HostName; IPs=$ips; Error="" }
    } catch {
        return @{ Success=$false; Hostname=""; IPs=""; Error=$_.Exception.Message }
    }
}

function Test-WebURL {
    param([string]$URL, [int]$TimeoutSec = 15)
    try {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $req = [System.Net.HttpWebRequest]::Create($URL)
        $req.Timeout = $TimeoutSec * 1000
        $req.AllowAutoRedirect = $true
        $req.UserAgent = "Mozilla/5.0 NetworkTestTool/1.0"
        $resp = $req.GetResponse()
        $sw.Stop()
        $code = [int]$resp.StatusCode
        $resp.Close()
        return @{ Success=$true; StatusCode=$code; LatencyMs=$sw.ElapsedMilliseconds; Error="" }
    } catch [System.Net.WebException] {
        $sw.Stop()
        $code = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
        return @{ Success=($code -gt 0 -and $code -lt 500); StatusCode=$code; LatencyMs=$sw.ElapsedMilliseconds; Error=$_.Exception.Message }
    } catch {
        return @{ Success=$false; StatusCode=0; LatencyMs=0; Error=$_.Exception.Message }
    }
}

function Get-NsLookup {
    param([string]$Domain)
    try {
        $out = & nslookup $Domain 2>&1
        return @{ Success=$true; Output=($out -join "`n") }
    } catch {
        return @{ Success=$false; Output=$_.Exception.Message }
    }
}

function Get-LLDPInfo {
    # Attempt to retrieve CDP/LLDP neighbour info from Windows LLDP agent
    $switch = "N/A"; $port = "N/A"; $vlan = "N/A"
    try {
        $lldp = Get-NetNeighbor -ErrorAction SilentlyContinue | Where-Object { $_.State -ne 'Unreachable' } | Select-Object -First 1
        if ($lldp) { $switch = $lldp.IPAddress }
        # Try WMI-based LLDP (requires Win32_NetworkAdapterConfiguration or vendor tool)
        # Most Windows environments need WinPcap/LLDP service; surface what we can
        $lldpSvc = Get-Service -Name "lldpsvc" -ErrorAction SilentlyContinue
        if (-not $lldpSvc) {
            # Try reading from registry if Cisco or HP LLDP agent installed
            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\lltdio"
            if (Test-Path $regPath) { $switch = "LLTD capable (use vendor tool for details)" }
        }
    } catch { }
    return @{ Switch=$switch; Port=$port; VLAN=$vlan }
}

function Get-ADUserFullName {
    param([string]$SamAccount)
    try {
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.Filter = "(&(objectClass=user)(sAMAccountName=$SamAccount))"
        $searcher.PropertiesToLoad.Add("displayName") | Out-Null
        $result = $searcher.FindOne()
        if ($result) { return $result.Properties["displayName"][0] }
    } catch { }
    return "N/A (AD lookup failed)"
}

function Pad-Right { param([string]$s, [int]$n) $s.PadRight($n) }
function HR { param([string]$c='=', [int]$n=80) return $c * $n }
function SectionHeader {
    param([string]$Title, [int]$Num)
    $bar = HR '='
    return @"
$bar
  SECTION $Num : $($Title.ToUpper())
$bar
"@
}

function StatusBadge {
    param([bool]$ok, [string]$errMsg="")
    if ($ok) { return "[ OK    ]" }
    else      { return "[ ERROR ] $errMsg" }
}

function Get-PingThresholdTag {
    param([hashtable]$r)
    if (-not $r.Reachable) { return "[FAIL]" }
    $tag = "[OK]"
    if ($Config.PingLatencyFailMs -gt 0 -and $r.Avg -ne "N/A" -and [double]$r.Avg -ge $Config.PingLatencyFailMs) { $tag = "[FAIL] Latency ${$r.Avg}ms exceeds ${$Config.PingLatencyFailMs}ms limit" }
    elseif ($Config.PingLossFailPct -gt 0 -and $r.Loss -ge $Config.PingLossFailPct)    { $tag = "[FAIL] Loss $($r.Loss)% exceeds $($Config.PingLossFailPct)% limit" }
    elseif ($Config.PingJitterFailMs -gt 0 -and $r.Jitter -ne "N/A" -and [double]$r.Jitter -ge $Config.PingJitterFailMs) { $tag = "[FAIL] Jitter $($r.Jitter)ms exceeds $($Config.PingJitterFailMs)ms limit" }
    elseif ($Config.PingLatencyWarnMs -gt 0 -and $r.Avg -ne "N/A" -and [double]$r.Avg -ge $Config.PingLatencyWarnMs) { $tag = "[WARN] Latency $($r.Avg)ms exceeds $($Config.PingLatencyWarnMs)ms threshold" }
    elseif ($Config.PingLossWarnPct -gt 0 -and $r.Loss -ge $Config.PingLossWarnPct)    { $tag = "[WARN] Loss $($r.Loss)% exceeds $($Config.PingLossWarnPct)% threshold" }
    elseif ($Config.PingJitterWarnMs -gt 0 -and $r.Jitter -ne "N/A" -and [double]$r.Jitter -ge $Config.PingJitterWarnMs) { $tag = "[WARN] Jitter $($r.Jitter)ms exceeds $($Config.PingJitterWarnMs)ms threshold" }
    return $tag
}

function Get-WebThresholdTag {
    param([hashtable]$r)
    if (-not $r -or -not $r.Success) { return "[FAIL]" }
    if ($Config.WebLatencyFailMs -gt 0 -and $r.LatencyMs -ge $Config.WebLatencyFailMs) { return "[FAIL] $($r.LatencyMs)ms exceeds $($Config.WebLatencyFailMs)ms limit" }
    if ($Config.WebLatencyWarnMs -gt 0 -and $r.LatencyMs -ge $Config.WebLatencyWarnMs) { return "[WARN] $($r.LatencyMs)ms exceeds $($Config.WebLatencyWarnMs)ms threshold" }
    return "[OK]"
}

#endregion

#region ── Main Script ────────────────────────────────────────────────────────

Clear-Host
Write-Host ""
Write-Host "+================================================================================+" -ForegroundColor Yellow
Write-Host "|              NETWORK TEST TOOL  -  Multi-Level Support Report               |" -ForegroundColor Yellow
Write-Host "+================================================================================+" -ForegroundColor Yellow
Write-Host ""

$totalSteps = 18
$step = 0
$activity = "Network Test Tool"

#── Step 1: Gather Client Info ─────────────────────────────────────────────────
$step++; Write-Progress-Step $step $totalSteps $activity "Gathering client information..."

$genTime      = Get-Date
$computerName = $env:COMPUTERNAME
$username     = $env:USERNAME
$domainName   = (Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue).Domain
if (-not $domainName) { $domainName = $env:USERDOMAIN }
$adFullName   = Get-ADUserFullName $username

# Network adapter info
$step++; Write-Progress-Step $step $totalSteps $activity "Collecting network adapter details..."

$allAdapters  = Get-NetAdapter -ErrorAction SilentlyContinue
$ipConfigs    = Get-NetIPConfiguration -ErrorAction SilentlyContinue
$activeConfig = $ipConfigs | Where-Object { $_.IPv4DefaultGateway -ne $null } | Select-Object -First 1

$primaryIPv4  = if ($activeConfig) { ($activeConfig.IPv4Address | Select-Object -First 1).IPAddress } else { "N/A" }
$primaryIPv6  = if ($activeConfig) { ($activeConfig.IPv6Address | Where-Object { $_.PrefixOrigin -ne 'WellKnown' } | Select-Object -First 1).IPAddress } else { $null }
$gateway      = if ($activeConfig) { ($activeConfig.IPv4DefaultGateway | Select-Object -First 1).NextHop } else { "N/A" }
$dnsServers   = if ($activeConfig) { ($activeConfig.DNSServer | Where-Object { $_.AddressFamily -eq 2 } | Select-Object -ExpandProperty ServerAddresses) -join ", " } else { "N/A" }
$primaryMAC   = if ($activeConfig) { ($allAdapters | Where-Object { $_.Name -eq $activeConfig.InterfaceAlias } | Select-Object -First 1).MacAddress } else { "N/A" }
if (-not $primaryMAC) { $primaryMAC = "N/A" }

# Inject dynamic targets
if ($gateway -ne "N/A") {
    $Config.InternalHosts[0].Host = $gateway
    $Config.InternalHosts[0].Name = "Default Gateway ($gateway)"
}
$firstDNS = if ($activeConfig) { ($activeConfig.DNSServer | Where-Object { $_.AddressFamily -eq 2 } | Select-Object -First 1 -ExpandProperty ServerAddresses) } else { $null }
if ($firstDNS) {
    $Config.InternalHosts[1].Host = $firstDNS
    $Config.InternalHosts[1].Name = "Primary DNS ($firstDNS)"
}

# LLDP / switch info
$step++; Write-Progress-Step $step $totalSteps $activity "Querying switch/LLDP information..."
$lldp = Get-LLDPInfo

#── Section 2 Tests ────────────────────────────────────────────────────────────
$step++; Write-Progress-Step $step $totalSteps $activity "Testing local network access (gateway & DNS ping)..."

$gwPing  = if ($gateway -ne "N/A") { Get-PingStats -Target $gateway -Count 4 } else { @{ Reachable=$false; Loss=100 } }
$dnsPing = if ($firstDNS)          { Get-PingStats -Target $firstDNS -Count 4 } else { @{ Reachable=$false; Loss=100 } }

$step++; Write-Progress-Step $step $totalSteps $activity "Running nslookup test..."
$nsTest = Get-NsLookup -Domain $Config.NslookupTestDomain

$step++; Write-Progress-Step $step $totalSteps $activity "Testing internal web services..."
$internalWebResults = @{}
foreach ($url in $Config.InternalWebURLs) {
    $internalWebResults[$url.Name] = Test-WebURL -URL $url.URL -TimeoutSec $Config.WebTimeoutSec
}
# DNS lookup for each internal web target
$internalDNSResults = @{}
foreach ($url in $Config.InternalWebURLs) {
    $host = ([System.Uri]$url.URL).Host
    $internalDNSResults[$url.Name] = Get-DNSLookup $host
}

$step++; Write-Progress-Step $step $totalSteps $activity "Testing internet services..."
$extURLTests = @{}
foreach ($url in $Config.ExternalWebURLs) {
    $extURLTests[$url.Name] = Test-WebURL -URL $url.URL -TimeoutSec $Config.WebTimeoutSec
}

# Section 2 overall statuses
$localNetOK    = $gwPing.Reachable -and $dnsPing.Reachable -and $nsTest.Success
$internalSvcOK = ($internalDNSResults.Values | Where-Object { $_.Success }).Count -gt 0 -and
                 ($internalWebResults.Values  | Where-Object { $_.Success }).Count -gt 0
$internetOK    = ($Config.InternetSummaryNames | Where-Object { $extURLTests[$_] -and $extURLTests[$_].Success }).Count -gt 0

#── Section 3 Tests ────────────────────────────────────────────────────────────
$step++; Write-Progress-Step $step $totalSteps $activity "Collecting all network interface details..."
# (allAdapters already collected)

$step++; Write-Progress-Step $step $totalSteps $activity "Collecting route table..."
$routeTable = Get-NetRoute -ErrorAction SilentlyContinue | Sort-Object RouteMetric, DestinationPrefix
$pingResults = @{}
$allTargets  = @()
$allTargets += $Config.InternalHosts | Where-Object { $_.Host }
$allTargets += $Config.ExternalHosts

foreach ($t in $allTargets) {
    $pingResults[$t.Name] = Get-PingStats -Target $t.Host -Count $Config.PingCount -TimeoutMs $Config.PingTimeoutMs
}

$step++; Write-Progress-Step $step $totalSteps $activity "Running detailed DNS lookups..."
$dnsResults = @{}
foreach ($t in $allTargets) {
    if ($t.Host -notmatch '^\d+\.\d+\.\d+\.\d+$') {  # skip pure IPs
        $dnsResults[$t.Name] = Get-DNSLookup $t.Host
    }
}

$step++; Write-Progress-Step $step $totalSteps $activity "Running detailed web access tests..."
# extURLTests already populated above; re-use

$step++; Write-Progress-Step $step $totalSteps $activity "Testing cloud platform connectivity endpoints..."
$cloudConnTests = @{}
foreach ($platform in $Config.CloudPlatforms) {
    if ($platform.ConnectURL -ne "") {
        $cloudConnTests[$platform.Name] = Test-WebURL -URL $platform.ConnectURL -TimeoutSec $Config.WebTimeoutSec
    }
}

#── Build Report ───────────────────────────────────────────────────────────────
$step++; Write-Progress-Step $step $totalSteps $activity "Building report..."

$reportLines = [System.Collections.Generic.List[string]]::new()

function Add { param([string]$line="") $reportLines.Add($line) }

# ─── Header ────────────────────────────────────────────────────────────────────
Add "$(HR '=')"
Add "  NETWORK TEST TOOL REPORT"
Add "$(HR '=')"
Add ""

# ─── SECTION 1 ─────────────────────────────────────────────────────────────────
Add (SectionHeader "Client Information" 1)
Add ""
Add "  Generated          : $($genTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Add "  Computer Name      : $computerName"
Add "  Domain Name        : $domainName"
Add "  Logged-in User     : $username"
Add "  AD Display Name    : $adFullName"
Add ""
Add "  Primary IPv4       : $primaryIPv4"
$ipv6Line = if ($primaryIPv6) { $primaryIPv6 } else { "Not configured" }
Add "  Primary IPv6       : $ipv6Line"
Add "  MAC Address        : $primaryMAC"
Add "  Default Gateway    : $gateway"
Add "  DNS Servers        : $dnsServers"
Add ""
Add "  Connected Switch   : $($lldp.Switch)"
Add "  Switch Port        : $($lldp.Port)"
Add "  VLAN Info          : $($lldp.VLAN)"
Add ""
Add "  NOTE: Switch/port/VLAN data requires an LLDP agent (e.g., Cisco CDP, HP LLDP)"
Add "        or a network management platform. Install vendor LLDP tool for full data."
Add ""

# ─── SECTION 2 ─────────────────────────────────────────────────────────────────
Add (SectionHeader "Helpdesk Troubleshooting Information" 2)
Add ""
Add "  Legend: [ OK    ] = Test passed    [ ERROR ] = Test failed"
Add ""
Add "$(HR '-')"
Add "  LOCAL NETWORK ACCESS"
Add "$(HR '-')"

$gwErr  = if (-not $gwPing.Reachable)  { "Gateway unreachable" } else { "" }
$dnsErr = if (-not $dnsPing.Reachable) { "DNS server unreachable" } else { "" }
$nsErr  = if (-not $nsTest.Success)    { "nslookup failed" } else { "" }
$localOK = $gwPing.Reachable -and $dnsPing.Reachable -and $nsTest.Success

$localErrMsg   = ($gwErr,$dnsErr,$nsErr | Where-Object {$_}) -join '; '
$localBadge    = StatusBadge $localOK $localErrMsg
$gwPingStr     = if ($gwPing.Reachable)  { "OK ($($gwPing.Avg) ms avg)"  } else { "FAIL - gateway not responding" }
$dnsPingStr    = if ($dnsPing.Reachable) { "OK ($($dnsPing.Avg) ms avg)" } else { "FAIL - DNS server not responding" }
$nsTestStr     = if ($nsTest.Success)    { "OK - $($Config.NslookupTestDomain) resolved" } else { "FAIL - $($Config.NslookupTestDomain) not resolved" }
Add "  $localBadge  Local Network Access"
Add "    Gateway Ping  : $gwPingStr"
Add "    DNS Ping      : $dnsPingStr"
Add "    nslookup test : $nsTestStr"
Add ""

Add "$(HR '-')"
Add "  INTERNAL SERVICES"
Add "$(HR '-')"

$intBadge = StatusBadge $internalSvcOK
Add "  $intBadge  Internal Services"
foreach ($url in $Config.InternalWebURLs) {
    $dnsR = $internalDNSResults[$url.Name]
    $webR = $internalWebResults[$url.Name]
    $dnsStr = if ($dnsR -and $dnsR.Success) { "OK -> $($dnsR.IPs)" } else { "FAIL - $(if($dnsR){$dnsR.Error}else{'no result'})" }
    $webStr = if ($webR -and $webR.Success) { "OK (HTTP $($webR.StatusCode), $($webR.LatencyMs) ms)" } else { "FAIL (HTTP $(if($webR){$webR.StatusCode}else{0})) - $(if($webR){$webR.Error}else{'no result'})" }
    $label = $url.Name.PadRight(20)
    Add "    $label DNS : $dnsStr"
    Add "    $label Web : $webStr"
}
Add ""

Add "$(HR '-')"
Add "  INTERNET SERVICES"
Add "$(HR '-')"

$internetSvcOK = ($Config.InternetSummaryNames | Where-Object { $extURLTests[$_] -and $extURLTests[$_].Success }).Count -gt 0
foreach ($name in $Config.InternetSummaryNames) {
    $r = $extURLTests[$name]
    if ($r) {
        $badge   = StatusBadge $r.Success
        $errTrunc = if ($r.Error) { $r.Error.Substring(0,[Math]::Min(60,$r.Error.Length)) } else { "" }
        $str     = if ($r.Success) { "OK (HTTP $($r.StatusCode), $($r.LatencyMs) ms)" } else { "FAIL (HTTP $($r.StatusCode)) - $errTrunc" }
        $namePad = $name.PadRight(20)
        Add "  $badge  $namePad $str"
    }
}
Add ""

Add "$(HR '-')"
Add "  CLOUD PLATFORM AVAILABILITY"
Add "$(HR '-')"

foreach ($platform in $Config.CloudPlatforms) {
    $connR  = $cloudConnTests[$platform.Name]
    $webR   = $extURLTests[$platform.WebURLName]
    $connOK = if ($connR) { $connR.Success } elseif ($webR) { $webR.Success } else { $false }
    $badge  = StatusBadge $connOK
    $detail = if ($connR) {
        "Connectivity check: HTTP $($connR.StatusCode), $($connR.LatencyMs) ms"
    } elseif ($webR) {
        "Web access: HTTP $($webR.StatusCode), $($webR.LatencyMs) ms"
    } else { "No test result" }
    $platPad = $platform.Name.PadRight(22)
    Add "  $badge  $platPad $detail"
}
Add ""
Add "  TIP: Check each platform status page (listed in Section 3.6) for authoritative health info."
Add ""

# ─── SECTION 3 ─────────────────────────────────────────────────────────────────
Add (SectionHeader "Network Engineer Troubleshooting Information" 3)
Add ""

# 3.1 All Interfaces
Add "$(HR '-')"
Add "  3.1  NETWORK INTERFACES (ALL - ACTIVE AND INACTIVE)"
Add "$(HR '-')"
Add ""

foreach ($adapter in $allAdapters | Sort-Object Status, Name) {
    $cfg = $ipConfigs | Where-Object { $_.InterfaceAlias -eq $adapter.Name }
    $status = $adapter.Status
    $statusColor = if ($status -eq 'Up') { "UP  " } else { "DOWN" }
    Add "  [$statusColor] $($adapter.Name)"
    Add "         Description  : $($adapter.InterfaceDescription)"
    Add "         MAC Address  : $($adapter.MacAddress)"
    Add "         Link Speed   : $($adapter.LinkSpeed)"
    Add "         Media Type   : $($adapter.MediaType)"

    if ($cfg) {
        $v4 = $cfg.IPv4Address | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" }
        $v6 = $cfg.IPv6Address | Where-Object { $_.PrefixOrigin -ne 'WellKnown' } | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" }
        $gw = $cfg.IPv4DefaultGateway | Select-Object -ExpandProperty NextHop
        $dns= ($cfg.DNSServer | Where-Object { $_.AddressFamily -eq 2 } | Select-Object -ExpandProperty ServerAddresses)
        if ($v4)  { Add "         IPv4 Address : $($v4 -join ', ')" }
        if ($v6)  { Add "         IPv6 Address : $($v6 -join ', ')" }
        if ($gw)  { Add "         Gateway      : $($gw -join ', ')" }
        if ($dns) { Add "         DNS Servers  : $($dns -join ', ')" }
    }
    Add ""
}

# 3.2 Route Table
Add "$(HR '-')"
Add "  3.2  WINDOWS ROUTE TABLE"
Add "$(HR '-')"
Add ""
$rtHdr = "  {0} {1} {2} {3} {4}" -f "Destination".PadRight(25), "Gateway".PadRight(18), "Interface".PadRight(22), "Metric".PadLeft(7), "Protocol".PadLeft(12)
Add $rtHdr
Add "  $(HR '-' 90)"
foreach ($route in $routeTable) {
    $dest    = if ($route.DestinationPrefix) { $route.DestinationPrefix } else { "N/A" }
    $gw      = if ($route.NextHop -and $route.NextHop -ne '0.0.0.0' -and $route.NextHop -ne '::') { $route.NextHop } else { "on-link" }
    $iface   = if ($route.InterfaceAlias) { $route.InterfaceAlias } else { "idx:$($route.InterfaceIndex)" }
    $metric  = [string]$route.RouteMetric
    $proto   = if ($route.Protocol) { $route.Protocol } else { "-" }
    $rline = "  {0} {1} {2} {3} {4}" -f $dest.PadRight(25), $gw.PadRight(18), $iface.PadRight(22), $metric.PadLeft(7), ([string]$proto).PadLeft(12)
    Add $rline
}
Add ""

# 3.3 Ping Results (was 3.2)
Add "$(HR '-')"
Add "  3.3  DETAILED PING RESULTS  ($($Config.PingCount) packets per target)"
Add "$(HR '-')"
Add ""
$pingHdr = "  {0} {1} {2} {3} {4} {5} {6} {7}  {8}" -f "Target".PadRight(35),"Sent".PadLeft(4),"Recv".PadLeft(4),"Loss%".PadLeft(6),"Min ms".PadLeft(7),"Avg ms".PadLeft(7),"Max ms".PadLeft(7),"Jitter".PadLeft(7),"Status"
Add $pingHdr
Add "  $(HR '-' 100)"

foreach ($t in $allTargets) {
    $name = $t.Name
    $r = $pingResults[$name]
    if ($r) {
        $tag  = Get-PingThresholdTag $r
        $line = "  {0} {1} {2} {3} {4} {5} {6} {7}  {8}" -f `
            $name.PadRight(35), `
            ([string]$r.Sent).PadLeft(4), `
            ([string]$r.Recv).PadLeft(4), `
            ("$($r.Loss)%").PadLeft(6), `
            ([string]$r.Min).PadLeft(7), `
            ([string]$r.Avg).PadLeft(7), `
            ([string]$r.Max).PadLeft(7), `
            ([string]$r.Jitter).PadLeft(7), `
            $tag
        Add $line
    }
}
Add ""
Add "  NOTE: All values in milliseconds. Jitter = mean absolute deviation of"
Add "        consecutive round-trip times. Loss% = packet loss percentage."
Add ""

# 3.4 DNS Lookup
Add "$(HR '-')"
Add "  3.4  DETAILED DNS LOOKUPS"
Add "$(HR '-')"
Add ""

foreach ($t in $allTargets) {
    if ($dnsResults.ContainsKey($t.Name)) {
        $r = $dnsResults[$t.Name]
        Add "  Target    : $($t.Host)"
        Add "  Friendly  : $($t.Name)"
        if ($r.Success) {
            Add "  Result    : OK"
            Add "  Hostname  : $($r.Hostname)"
            Add "  IPs       : $($r.IPs)"
        } else {
            Add "  Result    : FAIL"
            Add "  Error     : $($r.Error)"
        }
        Add ""
    }
}

# 3.5 Web Access Tests
Add "$(HR '-')"
Add "  3.5  WEB ACCESS TESTS"
Add "$(HR '-')"
Add ""
$webHdr = "  {0} {1} {2} {3} {4}  {5}" -f "Name".PadRight(24),"URL".PadRight(45),"Status".PadLeft(7),"HTTP".PadLeft(5),"Latency".PadLeft(9),"Threshold"
Add $webHdr
Add "  $(HR '-' 105)"

foreach ($url in ($Config.InternalWebURLs + $Config.ExternalWebURLs)) {
    $r = $internalWebResults[$url.Name]
    if (-not $r) { $r = $extURLTests[$url.Name] }
    if ($r) {
        $wStatus = if ($r.Success) { "OK" } else { "FAIL" }
        $tag     = Get-WebThresholdTag $r
        $line = "  {0} {1} {2} {3} {4}  {5}" -f `
            $url.Name.PadRight(24), `
            $url.URL.PadRight(45), `
            $wStatus.PadLeft(7), `
            ([string]$r.StatusCode).PadLeft(5), `
            ("$($r.LatencyMs) ms").PadLeft(9), `
            $tag
        Add $line
        if (-not $r.Success -and $r.Error) {
            $errTrunc2 = $r.Error.Substring(0,[Math]::Min(80,$r.Error.Length))
            Add "  $(' ' * 24) ERROR: $errTrunc2"
        }
    }
}
Add ""

# 3.6 Cloud Platform Detailed Tests
Add "$(HR '-')"
Add "  3.6  CLOUD PLATFORM AVAILABILITY - DETAILED"
Add "$(HR '-')"
Add ""

Add "  -- Microsoft Teams ------------------------------------------------------------"
$tConnStr  = if ($cloudConnTests['Microsoft Teams']) { if ($cloudConnTests['Microsoft Teams'].Success) {'PASS'} else {'FAIL'} } else { 'N/A' }
$tConnHTTP = if ($cloudConnTests['Microsoft Teams']) { $cloudConnTests['Microsoft Teams'].StatusCode } else { '-' }
$tConnMs   = if ($cloudConnTests['Microsoft Teams']) { $cloudConnTests['Microsoft Teams'].LatencyMs } else { '-' }
$tWebR     = $extURLTests['MS Teams Web']
$tWebStr   = if ($tWebR) { if ($tWebR.Success) {'PASS'} else {'FAIL'} } else { 'N/A' }
$tDNSEntry = $dnsResults['MS Teams']
$tDNSStr   = if ($tDNSEntry -and $tDNSEntry.Success) { "PASS - $($tDNSEntry.IPs)" } else { "FAIL - $(if($tDNSEntry){$tDNSEntry.Error}else{'not tested'})" }
$tPingEntry= $pingResults['MS Teams']
$tPingStr  = if ($tPingEntry) { "$($tPingEntry.Avg) ms avg  /  $($tPingEntry.Loss)% loss" } else { "not tested" }
Add "     Connectivity Check : $tConnStr  HTTP $tConnHTTP  Latency: $tConnMs ms"
Add "     Web Access         : $tWebStr  HTTP $(if($tWebR){$tWebR.StatusCode}else{'-'})  Latency: $(if($tWebR){$tWebR.LatencyMs}else{'-'}) ms"
Add "     DNS Resolution     : $tDNSStr"
Add "     Ping (avg/loss)    : $tPingStr"
Add "     Status Page        : https://admin.microsoft.com  (Health -> Service health)"
Add ""

Add "  -- Cisco Webex ---------------------------------------------------------------"
$wxConnR    = $cloudConnTests['Cisco Webex']
$wxConnStr  = if ($wxConnR) { if ($wxConnR.Success) {'PASS'} else {'FAIL'} } else { 'N/A' }
$wxWebR     = $extURLTests['Webex']
$wxWebStr   = if ($wxWebR) { if ($wxWebR.Success) {'PASS'} else {'FAIL'} } else { 'N/A' }
$wxDNSEntry = $dnsResults['Webex']
$wxDNSStr   = if ($wxDNSEntry -and $wxDNSEntry.Success) { "PASS - $($wxDNSEntry.IPs)" } else { "FAIL - $(if($wxDNSEntry){$wxDNSEntry.Error}else{'not tested'})" }
$wxPingEntry= $pingResults['Webex']
$wxPingStr  = if ($wxPingEntry) { "$($wxPingEntry.Avg) ms avg  /  $($wxPingEntry.Loss)% loss" } else { "not tested" }
Add "     Connectivity Check : $wxConnStr$(if($wxConnR){" HTTP $($wxConnR.StatusCode)  Latency: $($wxConnR.LatencyMs) ms"}else{""})"
Add "     Web Access         : $wxWebStr$(if($wxWebR){" HTTP $($wxWebR.StatusCode)  Latency: $($wxWebR.LatencyMs) ms"}else{""})"
Add "     DNS Resolution     : $wxDNSStr"
Add "     Ping (avg/loss)    : $wxPingStr"
Add "     Status Page        : https://status.webex.com"
Add ""

Add "  -- Microsoft 365 -------------------------------------------------------------"
$m3WebR     = $extURLTests['Office 365']
$m3SpR      = $extURLTests['SharePoint']
$m3WebStr   = if ($m3WebR) { if ($m3WebR.Success) {'PASS'} else {'FAIL'} } else { 'N/A' }
$m3SpStr    = if ($m3SpR)  { if ($m3SpR.Success)  {'PASS'} else {'FAIL'} } else { 'N/A' }
$m3DNSEntry = $dnsResults['Office 365']
$m3DNSStr   = if ($m3DNSEntry -and $m3DNSEntry.Success) { "PASS - $($m3DNSEntry.IPs)" } else { "FAIL - $(if($m3DNSEntry){$m3DNSEntry.Error}else{'not tested'})" }
$m3PingEntry= $pingResults['Office 365']
$m3PingStr  = if ($m3PingEntry) { "$($m3PingEntry.Avg) ms avg  /  $($m3PingEntry.Loss)% loss" } else { "not tested" }
Add "     Portal Access      : $m3WebStr$(if($m3WebR){" HTTP $($m3WebR.StatusCode)  Latency: $($m3WebR.LatencyMs) ms"}else{""})"
Add "     SharePoint Online  : $m3SpStr$(if($m3SpR){" HTTP $($m3SpR.StatusCode)  Latency: $($m3SpR.LatencyMs) ms"}else{""})"
Add "     DNS (office.com)   : $m3DNSStr"
Add "     Ping (avg/loss)    : $m3PingStr"
Add "     Status Page        : https://status.office365.com"
Add ""

Add "$(HR '=')"
Add "  END OF REPORT"
Add "$(HR '=')"

#── Save Report ────────────────────────────────────────────────────────────────
$step++; Write-Progress-Step $step $totalSteps $activity "Saving report file..."

$timestamp    = $genTime.ToString("yyyyMMdd_HHmm")
$baseName     = "${computerName}_NetTest_${timestamp}"
$reportPath   = Join-Path $env:USERPROFILE "${baseName}.txt"

# If the file already exists, append an incrementing counter
if (Test-Path $reportPath) {
    $counter = 2
    do {
        $reportPath = Join-Path $env:USERPROFILE "${baseName}_${counter}.txt"
        $counter++
    } while (Test-Path $reportPath)
}

$reportLines | Set-Content -Path $reportPath -Encoding UTF8

#── Done ───────────────────────────────────────────────────────────────────────
Write-Progress -Activity $activity -Completed

Write-Host ""
Write-Host "$(HR '=' 80)" -ForegroundColor Green
Write-Host "  Report saved to:" -ForegroundColor Green
Write-Host "  $reportPath" -ForegroundColor White
Write-Host "$(HR '=' 80)" -ForegroundColor Green
Write-Host ""

# Quick summary to console
$sumLocal = if ($localNetOK)    { "OK" } else { "ISSUE DETECTED" }
$sumInt   = if ($internalSvcOK) { "OK" } else { "ISSUE DETECTED" }
$sumInet  = if ($internetSvcOK) { "OK" } else { "ISSUE DETECTED" }
$colLocal = if ($localNetOK)    { "Green" } else { "Red" }
$colInt   = if ($internalSvcOK) { "Green" } else { "Red" }
$colInet  = if ($internetSvcOK) { "Green" } else { "Red" }

Write-Host "  SUMMARY" -ForegroundColor Yellow
Write-Host "  Local Network  : $sumLocal" -ForegroundColor $colLocal
Write-Host "  Internal Svcs  : $sumInt"   -ForegroundColor $colInt
Write-Host "  Internet Svcs  : $sumInet"  -ForegroundColor $colInet

foreach ($platform in $Config.CloudPlatforms) {
    $connR  = $cloudConnTests[$platform.Name]
    $webR   = $extURLTests[$platform.WebURLName]
    $platOK = if ($connR) { $connR.Success } elseif ($webR) { $webR.Success } else { $false }
    $platStr = if ($platOK) { "OK" } else { "ISSUE DETECTED" }
    $platCol = if ($platOK) { "Green" } else { "Red" }
    $label   = $platform.Name.PadRight(14)
    Write-Host "  $label : $platStr" -ForegroundColor $platCol
}
Write-Host ""

#endregion
