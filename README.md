This is a PowerShell script for a network test tool that cana analyze network details of Windows listed below and generate reports for different audience. The purpose is to get netowrk details and use for troubleshooting from level1 to level 3 support.
The script show the current activity in the prompt with a progress bar.


Report Structure:
A text file stored in the user folder. Filename format hostname_NetTest_YYYYmmDD_HHMM.txt . contents as below:

# Section 1 - Client Information
- date and time of generation
- computer name
- domain name
- logged in username
- AD full name of the user
- Primary IPv4
- Primary IPv6 if exist
- Default gateway
- DNS servers in use
- Connected switch name
- Connected switch port name
- Connected Switch VLAN Info

# Section 2 - Helpdesk Troubleshooting Information
Show the following details with result of OK, ERROR and a short error message if exist
- Local network access status (Check default gateway ping, DNS servsers ping, nslookup to test.domain)
- Intrnal services status (Result of dns lookup, and web testing webservers of test1.local and test2.local destinations) 
- Internet service status (Result of dns lookup, and web testing web servers for MS Teams, Sharepoing and Google)
- Availability result to MS Teams, Webex, M365 services if they have native tests with minimum details


# Section 3 - Network Engineer Troubleshooting Information
Show follwoing informations and verbose result of tests as per below in a readable structure with a managabe internal and external hostname, IP addresses and URL target lists (example: 8.8.8.8, google.com, office.com, sharepoint.com, test1.local, https://www.google.com):
- All network interfaces listed including active and deactive interfaces with all details
- Detailed ping results including avg, min, max latency, jitter and packet loss rate to targets above. 
- detailed DNS lookup to targets above.
- Web access test to URL targets above with detial of latency and HTTP accesss results
- Availability result to MS Teams, Webex, M365 services if they have native tests with details
