#Requires -Version 5.1
<#
.SYNOPSIS
    Network Test Tool  -  Multi-level network diagnostics for L1/L2/L3 support.
.DESCRIPTION
    Reads NetworkTest.config.txt from the same folder, runs connectivity tests,
    and writes a structured report to the current user profile folder.
    Filename: <hostname>_NetTest_YYYYmmDD_HHMM.txt
    If that file already exists a counter suffix is appended (_2, _3, ...).
.NOTES
    No admin rights required for most tests.
    LLDP switch/port data requires a vendor LLDP agent on the machine.
#>

# ==============================================================================
#  REGION 1 - CONFIG LOADER
#  Reads NetworkTest.config.txt and builds typed arrays from plain-text entries.
# ==============================================================================
#region Config Loader

$ConfigFile = Join-Path $PSScriptRoot "NetworkTest.config.txt"

if (-not (Test-Path $ConfigFile)) {
    Write-Host ""
    Write-Host "  [ERROR] Config file not found:" -ForegroundColor Red
    Write-Host "  $ConfigFile" -ForegroundColor Red
    Write-Host "  Place NetworkTest.config.txt in the same folder as this script." -ForegroundColor Red
    Write-Host ""
    exit 1
}

# Defaults applied when a setting is absent from the config file
$Cfg = @{
    PingCount        = 10
    PingTimeoutMs    = 2000
    PingWarnMs       = 50
    PingFailMs       = 150
    PingLossWarnPct  = 2
    PingLossFailPct  = 10
    WebTimeoutSec    = 15
    WebWarnMs        = 800
    WebFailMs        = 3000
    NslookupDomain   = "test.domain"
    InternalPing     = [System.Collections.Generic.List[hashtable]]::new()
    ExternalPing     = [System.Collections.Generic.List[hashtable]]::new()
    InternalWeb      = [System.Collections.Generic.List[hashtable]]::new()
    ExternalWeb      = [System.Collections.Generic.List[hashtable]]::new()
}

$currentSection = ""
foreach ($rawLine in (Get-Content $ConfigFile)) {
    $line = $rawLine.Trim()
    if ($line -eq "" -or $line.StartsWith("#")) { continue }

    # Section header  [SectionName]
    if ($line -match '^\[(.+)\]$') {
        $currentSection = $Matches[1].Trim()
        continue
    }

    # Settings  KEY = VALUE  (strip inline comments)
    if ($currentSection -eq "Settings" -and $line -match '^([^=]+)=(.+)$') {
        $key = $Matches[1].Trim()
        $val = ($Matches[2] -replace '#.*$', '').Trim()
        switch ($key) {
            "PingCount"        { $Cfg.PingCount        = [int]$val }
            "PingTimeoutMs"    { $Cfg.PingTimeoutMs    = [int]$val }
            "PingWarnMs"       { $Cfg.PingWarnMs       = [int]$val }
            "PingFailMs"       { $Cfg.PingFailMs       = [int]$val }
            "PingLossWarnPct"  { $Cfg.PingLossWarnPct  = [int]$val }
            "PingLossFailPct"  { $Cfg.PingLossFailPct  = [int]$val }
            "WebTimeoutSec"    { $Cfg.WebTimeoutSec    = [int]$val }
            "WebWarnMs"        { $Cfg.WebWarnMs        = [int]$val }
            "WebFailMs"        { $Cfg.WebFailMs        = [int]$val }
            "NslookupDomain"   { $Cfg.NslookupDomain   = $val }
        }
        continue
    }

    # Data rows  ping | Name | Host   or   web | Name | URL
    if ($line -match '^\s*(ping|web)\s*\|\s*(.+?)\s*\|\s*(.+?)\s*$') {
        $type  = $Matches[1].Trim().ToLower()
        $name  = $Matches[2].Trim()
        $value = $Matches[3].Trim()
        $entry = @{ Name = $name; Value = $value }
        switch ($currentSection) {
            "InternalPing" { if ($type -eq "ping") { $Cfg.InternalPing.Add($entry) } }
            "ExternalPing" { if ($type -eq "ping") { $Cfg.ExternalPing.Add($entry) } }
            "InternalWeb"  { if ($type -eq "web")  { $Cfg.InternalWeb.Add($entry)  } }
            "ExternalWeb"  { if ($type -eq "web")  { $Cfg.ExternalWeb.Add($entry)  } }
        }
    }
}

#endregion

# ==============================================================================
#  REGION 2 - HELPERS
# ==============================================================================
#region Helpers

# Report line accumulator
$ReportLines = [System.Collections.Generic.List[string]]::new()
function RL  { param([string]$L = "") $ReportLines.Add($L) }
function HR  { param([string]$C = "=", [int]$W = 80) return $C * $W }

function Show-Step {
    param([int]$Step, [int]$Total, [string]$Act, [string]$Msg)
    Write-Progress -Activity $Act -Status "[$Step/$Total] $Msg" -PercentComplete ([int](($Step / $Total) * 100))
    Write-Host "  >> $Msg" -ForegroundColor Cyan
}

function Write-SectionHeader {
    param([string]$Title, [int]$Num)
    RL (HR "=")
    RL "  SECTION $Num : $($Title.ToUpper())"
    RL (HR "=")
}

# Ping: returns hashtable {OK, Avg, Min, Max, Jitter, Loss, Sent, Recv}
function Invoke-PingTest {
    param([string]$Target, [int]$Count, [int]$TimeoutMs)
    $sent = 0; $recv = 0; $lats = @()
    $p = New-Object System.Net.NetworkInformation.Ping
    for ($i = 0; $i -lt $Count; $i++) {
        $sent++
        try {
            $r = $p.Send($Target, $TimeoutMs)
            if ($r.Status -eq "Success") { $recv++; $lats += $r.RoundtripTime }
        } catch { }
    }
    if ($recv -gt 0) {
        $avg  = [math]::Round(($lats | Measure-Object -Average).Average, 1)
        $min  = ($lats | Measure-Object -Minimum).Minimum
        $max  = ($lats | Measure-Object -Maximum).Maximum
        $diffs = @()
        for ($i = 1; $i -lt $lats.Count; $i++) { $diffs += [math]::Abs($lats[$i] - $lats[$i-1]) }
        $jitter = if ($diffs.Count -gt 0) { [math]::Round(($diffs | Measure-Object -Average).Average, 1) } else { 0 }
        $loss   = [math]::Round((($sent - $recv) / $sent) * 100, 0)
        return @{ OK=$true;  Avg=$avg; Min=$min; Max=$max; Jitter=$jitter; Loss=$loss; Sent=$sent; Recv=$recv }
    }
    return     @{ OK=$false; Avg="-";  Min="-";  Max="-";  Jitter="-";    Loss=100;   Sent=$sent; Recv=0 }
}

# DNS: returns hashtable {OK, IPs, Error}
function Invoke-DNSTest {
    param([string]$Target)
    try {
        $r   = [System.Net.Dns]::GetHostEntry($Target)
        $ips = ($r.AddressList | ForEach-Object { $_.ToString() }) -join ", "
        return @{ OK=$true;  IPs=$ips; Error="" }
    } catch {
        return @{ OK=$false; IPs="";   Error=$_.Exception.Message }
    }
}

# Web: returns hashtable {OK, Code, Ms, Error}
function Invoke-WebTest {
    param([string]$URL, [int]$TimeoutSec)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $req = [System.Net.HttpWebRequest]::Create($URL)
        $req.Timeout = $TimeoutSec * 1000
        $req.AllowAutoRedirect = $true
        $req.UserAgent = "Mozilla/5.0 NetworkTestTool/1.0"
        $resp = $req.GetResponse()
        $sw.Stop()
        $code = [int]$resp.StatusCode
        $resp.Close()
        return @{ OK=$true;  Code=$code; Ms=$sw.ElapsedMilliseconds; Error="" }
    } catch [System.Net.WebException] {
        $sw.Stop()
        $code = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
        return @{ OK=($code -gt 0 -and $code -lt 500); Code=$code; Ms=$sw.ElapsedMilliseconds; Error=$_.Exception.Message }
    } catch {
        $sw.Stop()
        return @{ OK=$false; Code=0; Ms=0; Error=$_.Exception.Message }
    }
}

# nslookup check: returns hashtable {OK, Output}
function Invoke-NslookupTest {
    param([string]$Domain)
    try   { $o = & nslookup $Domain 2>&1; return @{ OK=$true;  Output=($o -join "`n") } }
    catch { return @{ OK=$false; Output=$_.Exception.Message } }
}

# Ping threshold tag
function Get-PingTag {
    param([hashtable]$R)
    if (-not $R.OK) { return "[FAIL] unreachable" }
    if ($Cfg.PingFailMs      -gt 0 -and $R.Avg  -ne "-" -and [double]$R.Avg  -ge $Cfg.PingFailMs)     { return "[FAIL] avg $($R.Avg)ms >= $($Cfg.PingFailMs)ms" }
    if ($Cfg.PingLossFailPct -gt 0 -and $R.Loss -ge $Cfg.PingLossFailPct)                              { return "[FAIL] loss $($R.Loss)% >= $($Cfg.PingLossFailPct)%" }
    if ($Cfg.PingWarnMs      -gt 0 -and $R.Avg  -ne "-" -and [double]$R.Avg  -ge $Cfg.PingWarnMs)     { return "[WARN] avg $($R.Avg)ms >= $($Cfg.PingWarnMs)ms" }
    if ($Cfg.PingLossWarnPct -gt 0 -and $R.Loss -ge $Cfg.PingLossWarnPct)                              { return "[WARN] loss $($R.Loss)% >= $($Cfg.PingLossWarnPct)%" }
    return "[OK]"
}

# Web threshold tag
function Get-WebTag {
    param([hashtable]$R)
    if (-not $R -or -not $R.OK)                                                   { return "[FAIL]" }
    if ($Cfg.WebFailMs -gt 0 -and $R.Ms -ge $Cfg.WebFailMs)                      { return "[FAIL] $($R.Ms)ms >= $($Cfg.WebFailMs)ms" }
    if ($Cfg.WebWarnMs -gt 0 -and $R.Ms -ge $Cfg.WebWarnMs)                      { return "[WARN] $($R.Ms)ms >= $($Cfg.WebWarnMs)ms" }
    return "[OK]"
}

# AD display name (no RSAT required)
function Get-ADDisplayName {
    param([string]$Sam)
    try {
        $s = New-Object System.DirectoryServices.DirectorySearcher
        $s.Filter = "(&(objectClass=user)(sAMAccountName=$Sam))"
        $s.PropertiesToLoad.Add("displayName") | Out-Null
        $r = $s.FindOne()
        if ($r) { return $r.Properties["displayName"][0] }
    } catch { }
    return "N/A"
}

# LLDP best-effort
function Get-SwitchInfo {
    try {
        $n = Get-NetNeighbor -ErrorAction SilentlyContinue |
             Where-Object { $_.State -ne "Unreachable" } |
             Select-Object -First 1
        if ($n) { return @{ Switch=$n.IPAddress; Port="N/A"; VLAN="N/A" } }
    } catch { }
    return @{ Switch="N/A (requires vendor LLDP agent)"; Port="N/A"; VLAN="N/A" }
}

# Inline ping row writer (used in Section 3.3)
function Write-PingRow {
    param([string]$Name, [hashtable]$R)
    $tag = Get-PingTag $R
    RL ("  {0} {1} {2} {3} {4} {5} {6} {7}  {8}" -f `
        $Name.PadRight(30), `
        ([string]$R.Sent).PadLeft(4), `
        ([string]$R.Recv).PadLeft(4), `
        ("$($R.Loss)%").PadLeft(6), `
        ([string]$R.Min).PadLeft(6), `
        ([string]$R.Avg).PadLeft(6), `
        ([string]$R.Max).PadLeft(6), `
        ([string]$R.Jitter).PadLeft(7), `
        $tag)
}

# DNS result block writer (used in Section 3.4)
function Write-DNSBlock {
    param([string]$Name, [string]$TargetHost, [hashtable]$R)
    RL "  $Name  ($TargetHost)"
    if ($R -and $R.OK) {
        RL "    Result : OK"
        RL "    IPs    : $($R.IPs)"
    } else {
        RL "    Result : FAIL"
        RL "    Error  : $(if ($R) { $R.Error } else { 'not tested' })"
    }
    RL ""
}

# Web row writer (used in Section 3.5)
function Write-WebRow {
    param([string]$Name, [string]$URL, [hashtable]$R)
    $res = if ($R -and $R.OK) { "OK" } else { "FAIL" }
    $tag = Get-WebTag $R
    RL ("  {0} {1} {2} {3} {4}  {5}" -f `
        $Name.PadRight(22), `
        $URL.PadRight(42), `
        $res.PadRight(6), `
        ([string]$(if ($R) { $R.Code } else { "-" })).PadLeft(5), `
        ($(if ($R) { "$($R.Ms)ms" } else { "-" })).PadLeft(8), `
        $tag)
    if ($R -and -not $R.OK -and $R.Error) {
        $e = $R.Error.Substring(0, [Math]::Min(90, $R.Error.Length))
        RL "  $(' ' * 22) ERROR: $e"
    }
}

# Console summary line writer
function Write-SummaryLine {
    param([string]$Label, [bool]$OK)
    $status = if ($OK) { "OK" } else { "ISSUE DETECTED" }
    $color  = if ($OK) { "Green" } else { "Red" }
    Write-Host ("  {0} : {1}" -f $Label.PadRight(22), $status) -ForegroundColor $color
}

#endregion

# ==============================================================================
#  REGION 3 - DATA COLLECTION
# ==============================================================================
#region Data Collection

Clear-Host
Write-Host ""
Write-Host "+================================================================================+" -ForegroundColor Yellow
Write-Host "|         NETWORK TEST TOOL  -  Multi-Level Support Report                      |" -ForegroundColor Yellow
Write-Host "+================================================================================+" -ForegroundColor Yellow
Write-Host ""

# Resolve AUTO placeholders before counting steps, so the count is accurate
foreach ($entry in $Cfg.InternalPing) {
    if ($entry.Value -eq "AUTO") {
        if   ($entry.Name -like "*Gateway*" -and $gateway  -ne "N/A") { $entry.Value = $gateway }
        elseif ($entry.Name -like "*DNS*"   -and $firstDNS)            { $entry.Value = $firstDNS }
        else  { $entry.Value = $null }
    }
}

# Fixed steps: 1 sys info, 2 adapters, 3 LLDP, 4 local net, 5 route table, 6 save
$fixedSteps     = 6
$resolvedIntPin = ($Cfg.InternalPing | Where-Object { $_.Value }).Count
$totalSteps     = $fixedSteps + $resolvedIntPin + $Cfg.ExternalPing.Count +
                  $Cfg.InternalWeb.Count + $Cfg.ExternalWeb.Count
$step = 0
$act  = "Network Test Tool"

# Step 1 - System info
$step++; Show-Step $step $totalSteps $act "Gathering system information..."
$genTime      = Get-Date
$computerName = $env:COMPUTERNAME
$username     = $env:USERNAME
$domainName   = (Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue).Domain
if (-not $domainName) { $domainName = $env:USERDOMAIN }
$adName       = Get-ADDisplayName $username

# Step 2 - Adapter info
$step++; Show-Step $step $totalSteps $act "Reading network adapter configuration..."
$allAdapters  = Get-NetAdapter -ErrorAction SilentlyContinue
$allIPConfigs = Get-NetIPConfiguration -ErrorAction SilentlyContinue
$activeConfig = $allIPConfigs | Where-Object { $_.IPv4DefaultGateway -ne $null } | Select-Object -First 1

$primaryIPv4 = if ($activeConfig) { ($activeConfig.IPv4Address | Select-Object -First 1).IPAddress } else { "N/A" }
$primaryIPv6 = if ($activeConfig) { ($activeConfig.IPv6Address | Where-Object { $_.PrefixOrigin -ne "WellKnown" } | Select-Object -First 1).IPAddress } else { $null }
$gateway     = if ($activeConfig) { ($activeConfig.IPv4DefaultGateway | Select-Object -First 1).NextHop } else { "N/A" }
$dnsServers  = if ($activeConfig) { ($activeConfig.DNSServer | Where-Object { $_.AddressFamily -eq 2 } | Select-Object -ExpandProperty ServerAddresses) -join ", " } else { "N/A" }
$primaryMAC  = if ($activeConfig) { ($allAdapters | Where-Object { $_.Name -eq $activeConfig.InterfaceAlias } | Select-Object -First 1).MacAddress } else { "N/A" }
$firstDNS    = if ($activeConfig) { ($activeConfig.DNSServer | Where-Object { $_.AddressFamily -eq 2 } | Select-Object -First 1 -ExpandProperty ServerAddresses) } else { $null }


# Step 3 - LLDP
$step++; Show-Step $step $totalSteps $act "Querying switch/port information..."
$switchInfo = Get-SwitchInfo

# Step 4 - Local network quick test
$step++; Show-Step $step $totalSteps $act "Testing local network access..."
$gwPing  = if ($gateway -ne "N/A") { Invoke-PingTest $gateway   4 $Cfg.PingTimeoutMs } else { @{ OK=$false; Loss=100; Avg="-" } }
$dnsPing = if ($firstDNS)          { Invoke-PingTest $firstDNS  4 $Cfg.PingTimeoutMs } else { @{ OK=$false; Loss=100; Avg="-" } }
$nsTest  = Invoke-NslookupTest $Cfg.NslookupDomain

# Internal web tests (one step each)
$intWebResults = @{}
$intDNSResults = @{}
foreach ($entry in $Cfg.InternalWeb) {
    $step++; Show-Step $step $totalSteps $act "Internal web: $($entry.Name)..."
    $intWebResults[$entry.Name] = Invoke-WebTest $entry.Value $Cfg.WebTimeoutSec
    $uriHost = try { ([System.Uri]$entry.Value).Host } catch { $entry.Value }
    $intDNSResults[$entry.Name] = Invoke-DNSTest $uriHost
}

# External web tests (one step each)
$extWebResults = @{}
foreach ($entry in $Cfg.ExternalWeb) {
    $step++; Show-Step $step $totalSteps $act "External web: $($entry.Name)..."
    $extWebResults[$entry.Name] = Invoke-WebTest $entry.Value $Cfg.WebTimeoutSec
}

# Internal ping + DNS tests (one step each)
$intPingResults = @{}
$intDNS2Results = @{}
foreach ($entry in $Cfg.InternalPing) {
    if (-not $entry.Value) { continue }
    $step++; Show-Step $step $totalSteps $act "Internal ping: $($entry.Name)..."
    $intPingResults[$entry.Name] = Invoke-PingTest $entry.Value $Cfg.PingCount $Cfg.PingTimeoutMs
    if ($entry.Value -notmatch '^\d+\.\d+\.\d+\.\d+$') {
        $intDNS2Results[$entry.Name] = Invoke-DNSTest $entry.Value
    }
}

# External ping + DNS tests (one step each)
$extPingResults = @{}
$extDNSResults  = @{}
foreach ($entry in $Cfg.ExternalPing) {
    $step++; Show-Step $step $totalSteps $act "External ping: $($entry.Name)..."
    $extPingResults[$entry.Name] = Invoke-PingTest $entry.Value $Cfg.PingCount $Cfg.PingTimeoutMs
    if ($entry.Value -notmatch '^\d+\.\d+\.\d+\.\d+$') {
        $extDNSResults[$entry.Name] = Invoke-DNSTest $entry.Value
    }
}

# Step 5 - Route table
$step++; Show-Step $step $totalSteps $act "Collecting route table..."
$routeTable = Get-NetRoute -ErrorAction SilentlyContinue | Sort-Object RouteMetric, DestinationPrefix

#endregion

# ==============================================================================
#  REGION 4 - REPORT BUILDER
# ==============================================================================
#region Report Builder

$localNetOK    = $gwPing.OK -and $dnsPing.OK -and $nsTest.OK
$internalSvcOK = (($intWebResults.Values  | Where-Object { $_.OK }).Count -gt 0) -or
                 (($intPingResults.Values  | Where-Object { $_.OK }).Count -gt 0)
$internetSvcOK = ($extWebResults.Values | Where-Object { $_.OK }).Count -gt 0

# ============================================================  SECTION 1
Write-SectionHeader "Client Information" 1
RL ""
RL "  Generated       : $($genTime.ToString('yyyy-MM-dd HH:mm:ss'))"
RL "  Computer Name   : $computerName"
RL "  Domain          : $domainName"
RL "  Logged-in User  : $username"
RL "  AD Display Name : $adName"
RL ""
RL "  Primary IPv4    : $primaryIPv4"
RL "  Primary IPv6    : $(if ($primaryIPv6) { $primaryIPv6 } else { 'Not configured' })"
RL "  MAC Address     : $primaryMAC"
RL "  Default Gateway : $gateway"
RL "  DNS Servers     : $dnsServers"
RL ""
RL "  Switch          : $($switchInfo.Switch)"
RL "  Switch Port     : $($switchInfo.Port)"
RL "  VLAN            : $($switchInfo.VLAN)"
RL ""

# ============================================================  SECTION 2
Write-SectionHeader "Helpdesk Troubleshooting Information" 2
RL ""
RL "  Legend:  [ OK    ] = passed    [ ERROR ] = failed"
RL ""

# Local network
RL (HR "-")
RL "  LOCAL NETWORK ACCESS"
RL (HR "-")
$lBadge = if ($localNetOK) { "[ OK    ]" } else { "[ ERROR ]" }
RL "  $lBadge  Local Network Access"
RL "    Gateway ping  : $(if ($gwPing.OK)  { "OK ($($gwPing.Avg) ms avg)"  } else { "FAIL - gateway unreachable" })"
RL "    DNS ping      : $(if ($dnsPing.OK) { "OK ($($dnsPing.Avg) ms avg)" } else { "FAIL - DNS server unreachable" })"
RL "    nslookup      : $(if ($nsTest.OK)  { "OK - $($Cfg.NslookupDomain) resolved" } else { "FAIL - $($Cfg.NslookupDomain) not resolved" })"
RL ""

# Internal services
RL (HR "-")
RL "  INTERNAL SERVICES"
RL (HR "-")
$iBadge = if ($internalSvcOK) { "[ OK    ]" } else { "[ ERROR ]" }
RL "  $iBadge  Internal Services"
RL ""
if ($Cfg.InternalPing.Count -gt 0) {
    RL "    Ping targets:"
    foreach ($entry in $Cfg.InternalPing) {
        if (-not $entry.Value) { continue }
        $pR    = $intPingResults[$entry.Name]
        $label = $entry.Name.PadRight(22)
        $pTag  = Get-PingTag $pR
        $pStr  = if ($pR -and $pR.OK) { "OK   avg $($pR.Avg)ms  loss $($pR.Loss)%" } else { "FAIL - unreachable" }
        RL "      $label  $pStr  $pTag"
    }
    RL ""
}
if ($Cfg.InternalWeb.Count -gt 0) {
    RL "    Web targets:"
    foreach ($entry in $Cfg.InternalWeb) {
        $dR    = $intDNSResults[$entry.Name]
        $wR    = $intWebResults[$entry.Name]
        $label = $entry.Name.PadRight(22)
        $dStr  = if ($dR -and $dR.OK) { "DNS OK" } else { "DNS FAIL" }
        $wStr  = if ($wR -and $wR.OK) { "Web OK  HTTP $($wR.Code)  $($wR.Ms)ms" } else { "Web FAIL HTTP $(if ($wR) { $wR.Code } else { '?' })" }
        $wTag  = Get-WebTag $wR
        RL "      $label  $dStr  |  $wStr  $wTag"
    }
}
RL ""

# Internet services
RL (HR "-")
RL "  INTERNET SERVICES"
RL (HR "-")
$nBadge = if ($internetSvcOK) { "[ OK    ]" } else { "[ ERROR ]" }
RL "  $nBadge  Internet Services"
foreach ($entry in $Cfg.ExternalWeb) {
    $wR    = $extWebResults[$entry.Name]
    $label = $entry.Name.PadRight(18)
    $wStr  = if ($wR -and $wR.OK) { "OK   HTTP $($wR.Code)  $($wR.Ms)ms" } else { "FAIL HTTP $(if ($wR) { $wR.Code } else { '?' })" }
    $tag   = Get-WebTag $wR
    RL "    $label  $wStr  $tag"
}
RL ""

# ============================================================  SECTION 3
Write-SectionHeader "Network Engineer Troubleshooting Information" 3
RL ""

# 3.1 All interfaces
RL (HR "-")
RL "  3.1  NETWORK INTERFACES"
RL (HR "-")
RL ""
foreach ($adp in ($allAdapters | Sort-Object Status, Name)) {
    $ipc = $allIPConfigs | Where-Object { $_.InterfaceAlias -eq $adp.Name }
    $st  = if ($adp.Status -eq "Up") { "UP  " } else { "DOWN" }
    RL "  [$st] $($adp.Name)"
    RL "         Description : $($adp.InterfaceDescription)"
    RL "         MAC         : $($adp.MacAddress)"
    RL "         Speed       : $($adp.LinkSpeed)"
    RL "         Media       : $($adp.MediaType)"
    if ($ipc) {
        $v4  = $ipc.IPv4Address | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" }
        $v6  = $ipc.IPv6Address | Where-Object { $_.PrefixOrigin -ne "WellKnown" } | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" }
        $gw  = $ipc.IPv4DefaultGateway | Select-Object -ExpandProperty NextHop
        $dns = $ipc.DNSServer | Where-Object { $_.AddressFamily -eq 2 } | Select-Object -ExpandProperty ServerAddresses
        if ($v4)  { RL "         IPv4        : $($v4  -join ', ')" }
        if ($v6)  { RL "         IPv6        : $($v6  -join ', ')" }
        if ($gw)  { RL "         Gateway     : $($gw  -join ', ')" }
        if ($dns) { RL "         DNS         : $($dns -join ', ')" }
    }
    RL ""
}

# 3.2 Route table
RL (HR "-")
RL "  3.2  ROUTE TABLE"
RL (HR "-")
RL ""
$rtHdr = "  {0} {1} {2} {3} {4}" -f "Destination".PadRight(25), "Gateway".PadRight(20), "Interface".PadRight(28), "Metric".PadLeft(7), "Protocol"
RL $rtHdr
RL "  $(HR '-' 95)"
foreach ($rt in $routeTable) {
    $dest  = if ($rt.DestinationPrefix) { $rt.DestinationPrefix } else { "N/A" }
    $gw    = if ($rt.NextHop -and $rt.NextHop -ne "0.0.0.0" -and $rt.NextHop -ne "::") { $rt.NextHop } else { "on-link" }
    $iface = if ($rt.InterfaceAlias) { $rt.InterfaceAlias } else { "idx:$($rt.InterfaceIndex)" }
    $met   = [string]$rt.RouteMetric
    $proto = if ($rt.Protocol) { [string]$rt.Protocol } else { "-" }
    RL ("  {0} {1} {2} {3} {4}" -f $dest.PadRight(25), $gw.PadRight(20), $iface.PadRight(28), $met.PadLeft(7), $proto)
}
RL ""

# 3.3 Ping results
RL (HR "-")
RL "  3.3  PING TEST RESULTS  ($($Cfg.PingCount) packets per target)"
RL (HR "-")
RL "       Thresholds:  WARN avg>=$($Cfg.PingWarnMs)ms or loss>=$($Cfg.PingLossWarnPct)%  |  FAIL avg>=$($Cfg.PingFailMs)ms or loss>=$($Cfg.PingLossFailPct)%"
RL ""
$ph = "  {0} {1} {2} {3} {4} {5} {6} {7}  {8}" -f `
    "Target".PadRight(30), "Sent".PadLeft(4), "Recv".PadLeft(4), "Loss%".PadLeft(6), `
    "Min".PadLeft(6), "Avg".PadLeft(6), "Max".PadLeft(6), "Jitter".PadLeft(7), "Status"
RL $ph
RL "  $(HR '-' 105)"
RL "  -- Internal --"
foreach ($entry in $Cfg.InternalPing) {
    if (-not $entry.Value) { continue }
    $r = $intPingResults[$entry.Name]
    if ($r) { Write-PingRow $entry.Name $r }
}
RL ""
RL "  -- External --"
foreach ($entry in $Cfg.ExternalPing) {
    $r = $extPingResults[$entry.Name]
    if ($r) { Write-PingRow $entry.Name $r }
}
RL ""
RL "  Latency in ms. Jitter = mean absolute deviation of consecutive round-trip times."
RL ""

# 3.4 DNS results
RL (HR "-")
RL "  3.4  DNS LOOKUP RESULTS"
RL (HR "-")
RL ""
RL "  -- Internal Ping Targets --"
foreach ($entry in $Cfg.InternalPing) {
    if (-not $entry.Value -or $entry.Value -match '^\d+\.\d+\.\d+\.\d+$') { continue }
    $r = $intDNS2Results[$entry.Name]
    if ($r) { Write-DNSBlock $entry.Name $entry.Value $r }
}
RL "  -- Internal Web Targets --"
foreach ($entry in $Cfg.InternalWeb) {
    $uriHost2 = try { ([System.Uri]$entry.Value).Host } catch { $entry.Value }
    $r = $intDNSResults[$entry.Name]
    if ($r) { Write-DNSBlock $entry.Name $uriHost2 $r }
}
RL "  -- External Ping Targets --"
foreach ($entry in $Cfg.ExternalPing) {
    if ($entry.Value -match '^\d+\.\d+\.\d+\.\d+$') { continue }
    $r = $extDNSResults[$entry.Name]
    if ($r) { Write-DNSBlock $entry.Name $entry.Value $r }
}

# 3.5 Web results
RL (HR "-")
RL "  3.5  WEB ACCESS RESULTS"
RL (HR "-")
RL "       Thresholds:  WARN >=$($Cfg.WebWarnMs)ms  |  FAIL >=$($Cfg.WebFailMs)ms"
RL ""
$wh = "  {0} {1} {2} {3} {4}  {5}" -f `
    "Name".PadRight(22), "URL".PadRight(42), "Result".PadRight(6), "HTTP".PadLeft(5), "Latency".PadLeft(8), "Threshold"
RL $wh
RL "  $(HR '-' 105)"
RL "  -- Internal --"
foreach ($entry in $Cfg.InternalWeb) {
    Write-WebRow $entry.Name $entry.Value $intWebResults[$entry.Name]
}
RL ""
RL "  -- External --"
foreach ($entry in $Cfg.ExternalWeb) {
    Write-WebRow $entry.Name $entry.Value $extWebResults[$entry.Name]
}
RL ""

RL (HR "=")
RL "  END OF REPORT"
RL (HR "=")

#endregion

# ==============================================================================
#  REGION 5 - SAVE FILE
# ==============================================================================
#region Save File

$step++; Show-Step $step $totalSteps $act "Saving report..."

$stamp   = $genTime.ToString("yyyyMMdd_HHmm")
$base    = Join-Path $env:USERPROFILE "${computerName}_NetTest_${stamp}"
$outPath = "${base}.txt"
if (Test-Path $outPath) {
    $n = 2
    do { $outPath = "${base}_${n}.txt"; $n++ } while (Test-Path $outPath)
}
$ReportLines | Set-Content -Path $outPath -Encoding UTF8

#endregion

# ==============================================================================
#  REGION 6 - CONSOLE SUMMARY
# ==============================================================================
#region Console Summary

Write-Progress -Activity $act -Completed
Write-Host ""
Write-Host (HR "=" 80) -ForegroundColor Green
Write-Host "  Report saved: $outPath" -ForegroundColor White
Write-Host (HR "=" 80) -ForegroundColor Green
Write-Host ""
Write-Host "  SUMMARY" -ForegroundColor Yellow
Write-Host "  $(HR '-' 40)" -ForegroundColor DarkGray

Write-SummaryLine "Local Network"     $localNetOK
Write-SummaryLine "Internal Services" $internalSvcOK
Write-SummaryLine "Internet Services" $internetSvcOK
foreach ($entry in $Cfg.ExternalWeb) {
    $r = $extWebResults[$entry.Name]
    Write-SummaryLine $entry.Name ($r -and $r.OK)
}
Write-Host ""

#endregion