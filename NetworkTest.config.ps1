# =============================================================================
#  NetworkTest.config.ps1  -  Configuration for NetworkTest.ps1
#  Edit this file to customise targets without touching the main script.
#
#  RULES:
#    - Name   : Friendly label shown in the report. Must be unique per section.
#    - Host   : Hostname or IP used for ping and DNS tests.
#    - URL    : Full URL used for HTTP tests.
#    - Each section is a plain array - add or remove entries freely.
#    - Two entries in InternalHosts have Host = $null on purpose:
#      "Default Gateway" and "Primary DNS" are filled in automatically
#      at runtime from the active network adapter. Do not remove them.
# =============================================================================


# -----------------------------------------------------------------------------
#  TIMING & BEHAVIOUR
# -----------------------------------------------------------------------------

$cfg_PingCount       = 10     # Number of ICMP packets per target
$cfg_PingTimeoutMs   = 2000   # Ping timeout in milliseconds
$cfg_WebTimeoutSec   = 15     # HTTP test timeout in seconds


# -----------------------------------------------------------------------------
#  TEST METRIC THRESHOLDS
#  Results exceeding these limits are flagged as WARNING or FAIL in the report.
#  Set any value to 0 to disable that threshold check.
# -----------------------------------------------------------------------------

# Ping latency thresholds (milliseconds)
$cfg_PingLatencyWarnMs   = 50    # Avg RTT above this is flagged as WARNING
$cfg_PingLatencyFailMs   = 150   # Avg RTT above this is flagged as FAIL

# Ping jitter thresholds (milliseconds)
$cfg_PingJitterWarnMs    = 10    # Jitter above this is flagged as WARNING
$cfg_PingJitterFailMs    = 30    # Jitter above this is flagged as FAIL

# Ping packet loss thresholds (percentage 0-100)
$cfg_PingLossWarnPct     = 2     # Loss % above this is flagged as WARNING
$cfg_PingLossFailPct     = 10    # Loss % above this is flagged as FAIL

# Web / HTTP response latency thresholds (milliseconds)
$cfg_WebLatencyWarnMs    = 800   # Response time above this is flagged as WARNING
$cfg_WebLatencyFailMs    = 3000  # Response time above this is flagged as FAIL


# -----------------------------------------------------------------------------
#  INTERNAL PING / DNS TARGETS
#  These are tested in both Section 2 (summary) and Section 3 (detail).
#  The first two entries are resolved automatically - leave Host as $null.
# -----------------------------------------------------------------------------

$cfg_InternalHosts = @(

    # Auto-resolved at runtime - do not set Host manually
    @{ Name = "Default Gateway";        Host = $null }
    @{ Name = "Primary DNS Server";     Host = $null }

    # Add your internal servers below
    @{ Name = "test1.local";            Host = "test1.local"   }
    @{ Name = "test2.local";            Host = "test2.local"   }
    @{ Name = "File Server";            Host = "fileserver.local" }
    @{ Name = "Domain Controller";      Host = "dc01.local"    }

)


# -----------------------------------------------------------------------------
#  EXTERNAL PING / DNS TARGETS
#  Tested in Section 3 only (ping + DNS).
# -----------------------------------------------------------------------------

$cfg_ExternalHosts = @(

    @{ Name = "Google Public DNS";      Host = "8.8.8.8"                }
    @{ Name = "Cloudflare DNS";         Host = "1.1.1.1"                }
    @{ Name = "Google";                 Host = "google.com"             }
    @{ Name = "Office 365";             Host = "office.com"             }
    @{ Name = "SharePoint";             Host = "sharepoint.com"         }
    @{ Name = "MS Teams";               Host = "teams.microsoft.com"    }
    @{ Name = "Webex";                  Host = "webex.com"              }

)


# -----------------------------------------------------------------------------
#  INTERNAL WEB TESTS  (Section 2 - Helpdesk view)
#  Short list of internal URLs shown in the helpdesk summary with OK/ERROR.
# -----------------------------------------------------------------------------

$cfg_InternalWebURLs = @(

    @{ Name = "test1.local";            URL = "http://test1.local"      }
    @{ Name = "test2.local";            URL = "http://test2.local"      }
    @{ Name = "Intranet";               URL = "http://intranet.local"   }

)


# -----------------------------------------------------------------------------
#  EXTERNAL WEB TESTS  (Sections 2 & 3)
#  Tested with HTTP GET. StatusCode and latency recorded.
#  Section 2 shows a summary; Section 3 shows full detail.
# -----------------------------------------------------------------------------

$cfg_ExternalWebURLs = @(

    @{ Name = "Google";                 URL = "https://www.google.com"           }
    @{ Name = "Office 365";             URL = "https://www.office.com"           }
    @{ Name = "SharePoint";             URL = "https://www.sharepoint.com"       }
    @{ Name = "MS Teams Web";           URL = "https://teams.microsoft.com"      }
    @{ Name = "Webex";                  URL = "https://www.webex.com"            }
    @{ Name = "Azure Portal";           URL = "https://portal.azure.com"         }

)


# -----------------------------------------------------------------------------
#  CLOUD PLATFORM CHECKS  (Sections 2 & 3)
#  Each platform gets a dedicated block in Section 3.6.
#  Fields:
#    Name          - Display name used in headings
#    ConnectURL    - Native connectivity/health API endpoint (leave "" to skip)
#    WebURLName    - Must match a Name in $cfg_ExternalWebURLs (for web result)
#    DNSHostName   - Must match a Name in $cfg_ExternalHosts  (for DNS result)
#    PingHostName  - Must match a Name in $cfg_ExternalHosts  (for ping result)
#    StatusPage    - URL printed as a reference for manual checks
# -----------------------------------------------------------------------------

$cfg_CloudPlatforms = @(

    @{
        Name         = "Microsoft Teams"
        ConnectURL   = "https://connectivity.teams.microsoft.com/api/check"
        WebURLName   = "MS Teams Web"
        DNSHostName  = "MS Teams"
        PingHostName = "MS Teams"
        StatusPage   = "https://admin.microsoft.com  (Health -> Service health)"
    }

    @{
        Name         = "Cisco Webex"
        ConnectURL   = "https://api.ciscospark.com/v1/ping"
        WebURLName   = "Webex"
        DNSHostName  = "Webex"
        PingHostName = "Webex"
        StatusPage   = "https://status.webex.com"
    }

    @{
        Name         = "Microsoft 365"
        ConnectURL   = ""
        WebURLName   = "Office 365"
        DNSHostName  = "Office 365"
        PingHostName = "Office 365"
        StatusPage   = "https://status.office365.com"
    }

    @{
        Name         = "SharePoint Online"
        ConnectURL   = ""
        WebURLName   = "SharePoint"
        DNSHostName  = "SharePoint"
        PingHostName = "SharePoint"
        StatusPage   = "https://status.office365.com"
    }

)


# -----------------------------------------------------------------------------
#  SECTION 2 - INTERNET SERVICES SUMMARY
#  Names here must match entries in $cfg_ExternalWebURLs.
#  Only these entries appear in the helpdesk summary table.
# -----------------------------------------------------------------------------

$cfg_InternetSummaryNames = @(
    "Google"
    "Office 365"
    "SharePoint"
    "MS Teams Web"
    "Webex"
)


# -----------------------------------------------------------------------------
#  NSLOOKUP TEST DOMAIN  (Section 2 - local network access check)
# -----------------------------------------------------------------------------

$cfg_NslookupTestDomain = "test.domain"
