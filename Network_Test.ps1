#Requires -Version 5.1
<#
.SYNOPSIS
    Network Test Tool - Comprehensive network diagnostics for L1/L2/L3 support
.DESCRIPTION
    Collects network details, runs connectivity tests, and generates a structured report
    for helpdesk and network engineering troubleshooting.
    Created by: Ehsan to use by :)
.NOTES
    Run as: .\NetworkTest.ps1
    Output: $env:USERPROFILE\<hostname>_NetTest_YYYYmmDD_HHMM.txt
#>

#region ── Configuration ──────────────────────────────────────────────────────

$Config = @{
    # Internal targets (DNS name + IP)
    InternalHosts = @(
        @{ Name = "Default Gateway";   Host = $null }          # resolved at runtime
        @{ Name = "DC / DNS Server";   Host = $null }          # resolved at runtime
        @{ Name = "test1.local";       Host = "test1.local" }
        @{ Name = "test2.local";       Host = "test2.local" }
    )
    # External targets
    ExternalHosts = @(
        @{ Name = "Google DNS";        Host = "8.8.8.8" }
        @{ Name = "Google";            Host = "google.com" }
        @{ Name = "Office 365";        Host = "office.com" }
        @{ Name = "SharePoint";        Host = "sharepoint.com" }
        @{ Name = "MS Teams";          Host = "teams.microsoft.com" }
        @{ Name = "Webex";             Host = "webex.com" }
    )
    # Web URLs to test
    WebURLs = @(
        @{ Name = "Google";            URL = "https://www.google.com" }
        @{ Name = "Office 365";        URL = "https://www.office.com" }
        @{ Name = "SharePoint";        URL = "https://www.sharepoint.com" }
        @{ Name = "MS Teams Web";      URL = "https://teams.microsoft.com" }
        @{ Name = "Webex";             URL = "https://www.webex.com" }
        @{ Name = "test1.local";       URL = "http://test1.local" }
        @{ Name = "test2.local";       URL = "http://test2.local" }
    )
    # Internal web services for Section 2
    InternalWebURLs = @(
        @{ Name = "test1.local";       URL = "http://test1.local" }
        @{ Name = "test2.local";       URL = "http://test2.local" }
    )
    # nslookup test domain for Section 2
    NslookupTestDomain = "test.domain"
    PingCount          = 10
    PingTimeoutMs      = 2000
    WebTimeoutSec      = 15
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
function HR { param([char]$c='─', [int]$n=80) return ([string]$c * $n) }
function SectionHeader {
    param([string]$Title, [int]$Num)
    $bar = HR '═'
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

#endregion

#region ── Main Script ────────────────────────────────────────────────────────

Clear-Host
Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
Write-Host "║              NETWORK TEST TOOL  —  Multi-Level Support Report               ║" -ForegroundColor Yellow
Write-Host "╚══════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
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
$intDNS1 = Get-DNSLookup "test1.local"
$intDNS2 = Get-DNSLookup "test2.local"

$step++; Write-Progress-Step $step $totalSteps $activity "Testing internet services (Teams, SharePoint, Google)..."
$extURLTests = @{}
foreach ($url in $Config.WebURLs) {
    $extURLTests[$url.Name] = Test-WebURL -URL $url.URL -TimeoutSec $Config.WebTimeoutSec
}

# Section 2 overall statuses
$localNetOK    = $gwPing.Reachable -and $dnsPing.Reachable -and $nsTest.Success
$internalSvcOK = $intDNS1.Success -and $intDNS2.Success -and
                 ($internalWebResults["test1.local"].Success -or $internalWebResults["test2.local"].Success)
$internetOK    = $extURLTests["Google"].Success -or $extURLTests["Office 365"].Success

#── Section 3 Tests ────────────────────────────────────────────────────────────
$step++; Write-Progress-Step $step $totalSteps $activity "Collecting all network interface details..."
# (allAdapters already collected)

$step++; Write-Progress-Step $step $totalSteps $activity "Running detailed ping tests to all targets..."
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

$step++; Write-Progress-Step $step $totalSteps $activity "Testing MS Teams connectivity..."
# Teams has a native connectivity test endpoint
$teamsConnTest = Test-WebURL -URL "https://connectivity.teams.microsoft.com/api/check" -TimeoutSec $Config.WebTimeoutSec

$step++; Write-Progress-Step $step $totalSteps $activity "Testing Webex connectivity..."
$webexConnTest = Test-WebURL -URL "https://api.ciscospark.com/v1/ping" -TimeoutSec $Config.WebTimeoutSec

$step++; Write-Progress-Step $step $totalSteps $activity "Testing Microsoft 365 endpoints..."
$m365Test = Test-WebURL -URL "https://www.office.com" -TimeoutSec $Config.WebTimeoutSec

#── Build Report ───────────────────────────────────────────────────────────────
$step++; Write-Progress-Step $step $totalSteps $activity "Building report..."

$reportLines = [System.Collections.Generic.List[string]]::new()

function Add { param([string]$line="") $reportLines.Add($line) }

# ─── Header ────────────────────────────────────────────────────────────────────
Add "$(HR '═')"
Add "  NETWORK TEST TOOL REPORT"
Add "$(HR '═')"
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
if ($primaryIPv6) {
    Add "  Primary IPv6       : $primaryIPv6"
}
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
Add "$(HR '─')"
Add "  LOCAL NETWORK ACCESS"
Add "$(HR '─')"

$gwErr  = if (-not $gwPing.Reachable)  { "Gateway unreachable" } else { "" }
$dnsErr = if (-not $dnsPing.Reachable) { "DNS server unreachable" } else { "" }
$nsErr  = if (-not $nsTest.Success)    { "nslookup failed" } else { "" }
$localOK = $gwPing.Reachable -and $dnsPing.Reachable -and $nsTest.Success

Add "  $(StatusBadge $localOK (($gwErr,$dnsErr,$nsErr | Where-Object {$_}) -join '; '))  Local Network Access"
Add "    Gateway Ping  : $(if($gwPing.Reachable){"OK ($($gwPing.Avg) ms avg)"}else{"FAIL — gateway not responding"})"
Add "    DNS Ping      : $(if($dnsPing.Reachable){"OK ($($dnsPing.Avg) ms avg)"}else{"FAIL — DNS server not responding"})"
Add "    nslookup test : $(if($nsTest.Success){"OK — $($Config.NslookupTestDomain) resolved"}else{"FAIL — $($Config.NslookupTestDomain) not resolved"})"
Add ""

Add "$(HR '─')"
Add "  INTERNAL SERVICES"
Add "$(HR '─')"

$int1DNS = if ($intDNS1.Success) { "OK → $($intDNS1.IPs)" } else { "FAIL — $($intDNS1.Error)" }
$int2DNS = if ($intDNS2.Success) { "OK → $($intDNS2.IPs)" } else { "FAIL — $($intDNS2.Error)" }
$int1Web = $internalWebResults["test1.local"]
$int2Web = $internalWebResults["test2.local"]
$int1WebStr = if ($int1Web.Success) { "OK (HTTP $($int1Web.StatusCode), $($int1Web.LatencyMs) ms)" } else { "FAIL (HTTP $($int1Web.StatusCode)) — $($int1Web.Error)" }
$int2WebStr = if ($int2Web.Success) { "OK (HTTP $($int2Web.StatusCode), $($int2Web.LatencyMs) ms)" } else { "FAIL (HTTP $($int2Web.StatusCode)) — $($int2Web.Error)" }

Add "  $(StatusBadge $internalSvcOK)  Internal Services"
Add "    test1.local DNS  : $int1DNS"
Add "    test1.local Web  : $int1WebStr"
Add "    test2.local DNS  : $int2DNS"
Add "    test2.local Web  : $int2WebStr"
Add ""

Add "$(HR '─')"
Add "  INTERNET SERVICES"
Add "$(HR '─')"

$googleOK    = $extURLTests["Google"].Success
$sharepointOK= $extURLTests["SharePoint"].Success
$teamsOK     = $extURLTests["MS Teams Web"].Success
$internetSvcOK = $googleOK -or $sharepointOK -or $teamsOK

foreach ($url in @("Google","Office 365","SharePoint","MS Teams Web","Webex")) {
    $r = $extURLTests[$url]
    if ($r) {
        $str = if ($r.Success) { "OK (HTTP $($r.StatusCode), $($r.LatencyMs) ms)" } else { "FAIL (HTTP $($r.StatusCode)) — $($r.Error.Substring(0,[Math]::Min(60,$r.Error.Length)))" }
        Add "  $(StatusBadge $r.Success)  $($url.PadRight(20)) $str"
    }
}
Add ""

Add "$(HR '─')"
Add "  CLOUD PLATFORM AVAILABILITY"
Add "$(HR '─')"

$teamsConnOK = $teamsConnTest.Success
$webexConnOK = $webexConnTest.Success
$m365OK      = $m365Test.Success

Add "  $(StatusBadge $teamsConnOK)  MS Teams   — Connectivity endpoint: HTTP $($teamsConnTest.StatusCode), $($teamsConnTest.LatencyMs) ms"
Add "  $(StatusBadge $webexConnOK)  Cisco Webex— API ping:              HTTP $($webexConnTest.StatusCode), $($webexConnTest.LatencyMs) ms"
Add "  $(StatusBadge $m365OK)  M365       — Portal access:          HTTP $($m365Test.StatusCode), $($m365Test.LatencyMs) ms"
Add ""
Add "  TIP: For authoritative service health, check:"
Add "       Teams/M365 : https://admin.microsoft.com  →  Health → Service health"
Add "       Webex       : https://status.webex.com"
Add ""

# ─── SECTION 3 ─────────────────────────────────────────────────────────────────
Add (SectionHeader "Network Engineer Troubleshooting Information" 3)
Add ""

# 3.1 All Interfaces
Add "$(HR '─')"
Add "  3.1  NETWORK INTERFACES (ALL — ACTIVE AND INACTIVE)"
Add "$(HR '─')"
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

# 3.2 Ping Results
Add "$(HR '─')"
Add "  3.2  DETAILED PING RESULTS  ($($Config.PingCount) packets per target)"
Add "$(HR '─')"
Add ""
Add "  $("Target".PadRight(35)) $("Sent".PadLeft(4)) $("Recv".PadLeft(4)) $("Loss%".PadLeft(6)) $("Min ms".PadLeft(7)) $("Avg ms".PadLeft(7)) $("Max ms".PadLeft(7)) $("Jitter".PadLeft(7))"
Add "  $(HR '-' 80)"

foreach ($t in $allTargets) {
    $name = $t.Name
    $r = $pingResults[$name]
    if ($r) {
        $line = "  $($name.PadRight(35)) $([string]$r.Sent.PadLeft(4)) $([string]$r.Recv.PadLeft(4)) $([string]("$($r.Loss)%").PadLeft(6)) $([string]"$($r.Min)".PadLeft(7)) $([string]"$($r.Avg)".PadLeft(7)) $([string]"$($r.Max)".PadLeft(7)) $([string]"$($r.Jitter)".PadLeft(7))"
        Add $line
    }
}
Add ""
Add "  NOTE: All values in milliseconds. Jitter = mean absolute deviation of"
Add "        consecutive round-trip times. Loss% = packet loss percentage."
Add ""

# 3.3 DNS Lookup
Add "$(HR '─')"
Add "  3.3  DETAILED DNS LOOKUPS"
Add "$(HR '─')"
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

# 3.4 Web Access Tests
Add "$(HR '─')"
Add "  3.4  WEB ACCESS TESTS"
Add "$(HR '─')"
Add ""
Add "  $("Name".PadRight(24)) $("URL".PadRight(45)) $("Status".PadLeft(7)) $("HTTP".PadLeft(5)) $("Latency".PadLeft(9))"
Add "  $(HR '-' 95)"

foreach ($url in $Config.WebURLs) {
    $r = $extURLTests[$url.Name]
    if (-not $r) { $r = $internalWebResults[$url.Name] }
    if ($r) {
        $status = if ($r.Success) { "OK" } else { "FAIL" }
        $line = "  $($url.Name.PadRight(24)) $($url.URL.PadRight(45)) $($status.PadLeft(7)) $([string]$r.StatusCode.PadLeft(5)) $("$($r.LatencyMs) ms".PadLeft(9))"
        Add $line
        if (-not $r.Success -and $r.Error) {
            Add "  $((' ' * 24)) ERROR: $($r.Error.Substring(0,[Math]::Min(80,$r.Error.Length)))"
        }
    }
}
Add ""

# 3.5 Cloud Platform Detailed Tests
Add "$(HR '─')"
Add "  3.5  CLOUD PLATFORM AVAILABILITY — DETAILED"
Add "$(HR '─')"
Add ""

Add "  ── Microsoft Teams ────────────────────────────────────────────────────────────"
Add "     Connectivity Check : $(if($teamsConnTest.Success){'PASS'}else{'FAIL'})  HTTP $($teamsConnTest.StatusCode)  Latency: $($teamsConnTest.LatencyMs) ms"
Add "     Web Client         : $(if($extURLTests['MS Teams Web'].Success){'PASS'}else{'FAIL'})  HTTP $($extURLTests['MS Teams Web'].StatusCode)  Latency: $($extURLTests['MS Teams Web'].LatencyMs) ms"
Add "     DNS Resolution     : $(if($dnsResults['MS Teams'].Success){'PASS — '+$dnsResults['MS Teams'].IPs}else{'FAIL — '+$dnsResults['MS Teams'].Error})"
Add "     Ping (avg/loss)    : $($pingResults['MS Teams'].Avg) ms  /  $($pingResults['MS Teams'].Loss)% loss"
Add "     Admin Health URL   : https://admin.microsoft.com  → Health → Service health"
Add ""

Add "  ── Cisco Webex ────────────────────────────────────────────────────────────────"
Add "     API Ping           : $(if($webexConnTest.Success){'PASS'}else{'FAIL'})  HTTP $($webexConnTest.StatusCode)  Latency: $($webexConnTest.LatencyMs) ms"
Add "     Web Client         : $(if($extURLTests['Webex'].Success){'PASS'}else{'FAIL'})  HTTP $($extURLTests['Webex'].StatusCode)  Latency: $($extURLTests['Webex'].LatencyMs) ms"
Add "     DNS Resolution     : $(if($dnsResults['Webex'].Success){'PASS — '+$dnsResults['Webex'].IPs}else{'FAIL — '+$dnsResults['Webex'].Error})"
Add "     Ping (avg/loss)    : $($pingResults['Webex'].Avg) ms  /  $($pingResults['Webex'].Loss)% loss"
Add "     Status Page        : https://status.webex.com"
Add ""

Add "  ── Microsoft 365 ──────────────────────────────────────────────────────────────"
Add "     Portal Access      : $(if($m365Test.Success){'PASS'}else{'FAIL'})  HTTP $($m365Test.StatusCode)  Latency: $($m365Test.LatencyMs) ms"
Add "     SharePoint         : $(if($extURLTests['SharePoint'].Success){'PASS'}else{'FAIL'})  HTTP $($extURLTests['SharePoint'].StatusCode)  Latency: $($extURLTests['SharePoint'].LatencyMs) ms"
Add "     Office 365         : $(if($extURLTests['Office 365'].Success){'PASS'}else{'FAIL'})  HTTP $($extURLTests['Office 365'].StatusCode)  Latency: $($extURLTests['Office 365'].LatencyMs) ms"
Add "     DNS (office.com)   : $(if($dnsResults['Office 365'].Success){'PASS — '+$dnsResults['Office 365'].IPs}else{'FAIL — '+$dnsResults['Office 365'].Error})"
Add "     Status Page        : https://status.office365.com"
Add ""

Add "$(HR '═')"
Add "  END OF REPORT"
Add "$(HR '═')"

#── Save Report ────────────────────────────────────────────────────────────────
$step++; Write-Progress-Step $step $totalSteps $activity "Saving report file..."

$timestamp  = $genTime.ToString("yyyyMMdd_HHmm")
$reportName = "${computerName}_NetTest_${timestamp}.txt"
$reportPath = Join-Path $env:USERPROFILE $reportName

$reportLines | Set-Content -Path $reportPath -Encoding UTF8

#── Done ───────────────────────────────────────────────────────────────────────
Write-Progress -Activity $activity -Completed

Write-Host ""
Write-Host "$(HR '═' 80)" -ForegroundColor Green
Write-Host "  Report saved to:" -ForegroundColor Green
Write-Host "  $reportPath" -ForegroundColor White
Write-Host "$(HR '═' 80)" -ForegroundColor Green
Write-Host ""

# Quick summary to console
Write-Host "  SUMMARY" -ForegroundColor Yellow
Write-Host "  Local Network  : $(if($localNetOK){'OK'}else{'ISSUE DETECTED'})"     -ForegroundColor $(if($localNetOK){'Green'}else{'Red'})
Write-Host "  Internal Svcs  : $(if($internalSvcOK){'OK'}else{'ISSUE DETECTED'})"  -ForegroundColor $(if($internalSvcOK){'Green'}else{'Red'})
Write-Host "  Internet Svcs  : $(if($internetSvcOK){'OK'}else{'ISSUE DETECTED'})"  -ForegroundColor $(if($internetSvcOK){'Green'}else{'Red'})
Write-Host "  MS Teams       : $(if($teamsConnOK){'OK'}else{'ISSUE DETECTED'})"    -ForegroundColor $(if($teamsConnOK){'Green'}else{'Red'})
Write-Host "  Cisco Webex    : $(if($webexConnOK){'OK'}else{'ISSUE DETECTED'})"    -ForegroundColor $(if($webexConnOK){'Green'}else{'Red'})
Write-Host "  M365           : $(if($m365OK){'OK'}else{'ISSUE DETECTED'})"         -ForegroundColor $(if($m365OK){'Green'}else{'Red'})
Write-Host ""

#endregion