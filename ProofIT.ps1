#  _________                __                                     __    ___________           .__                  .__                .__               
#  \_   ___ \  ____________/  |________ __ __   ____  ____   _____/  |_  \__    ___/___   ____ |  |__   ____   ____ |  |   ____   ____ |__| ____   ______
#  /    \  \/ /  _ \_  __ \   __\_  __ \  |  \_/ ___\/ __ \ /    \   __\   |    |_/ __ \_/ ___\|  |  \ /    \ /  _ \|  |  /  _ \ / ___\|  |/ __ \ /  ___/
#  \     \___(  <_> )  | \/|  |  |  | \/  |  /\  \__\  ___/|   |  \  |     |    |\  ___/\  \___|   Y  \   |  (  <_> )  |_(  <_> ) /_/  >  \  ___/ \___ \ 
#   \______  /\____/|__|   |__|  |__|  |____/  \___  >___  >___|  /__|     |____| \___  >\___  >___|  /___|  /\____/|____/\____/\___  /|__|\___  >____  >
#          \/                                      \/    \/     \/                    \/     \/     \/     \/                  /_____/         \/     \/ 

### Check if ExchangeOnline v2 is installed; if not, install it
if (-not (Get-InstalledModule ExchangeOnlineManagement)){
    Install-Module ExchangeOnlineManagement
}

# Import EXO module
Import-Module ExchangeOnlineManagement

### Connect to Exchange Online
Connect-ExchangeOnline

### Begin Proofpoint configuration
# Globally used declarations
$proofpointCIDRs = "67.231.152.0/24","67.231.153.0/24","67.231.154.0/24",
    "67.231.155.0/24","67.231.156.0/24","67.231.144.0/24","67.231.145.0/24",
    "67.231.146.0/24","67.231.147.0/24","67.231.148.0/24","67.231.149.0/24",
    "148.163.128.0/24","148.163.129.0/24","148.163.130.0/24","148.163.131.0/24",
    "148.163.132.0/24","148.163.133.0/24","148.163.134.0/24","148.163.135.0/24",
    "148.163.136.0/24","148.163.137.0/24","148.163.138.0/24","148.163.139.0/24",
    "148.163.140.0/24","148.163.141.0/24","148.163.142.0/24","148.163.143.0/24",
    "148.163.144.0/24","148.163.145.0/24","148.163.146.0/24","148.163.147.0/24",
    "148.163.148.0/24","148.163.149.0/24","148.163.150.0/24","148.163.151.0/24",
    "148.163.152.0/24","148.163.153.0/24","148.163.154.0/24","148.163.155.0/24",
    "148.163.156.0/24","148.163.157.0/24","148.163.158.0/24","148.163.159.0/24"

# Create mail flow rule to bypass built-in Office 365 spam filtering
$transportRuleName = 'Proofpoint Essentials - Bypass O365 Spam Filtering'
New-TransportRule -Name $transportRuleName -SenderIpRanges $proofpointCIDRs -SetSCL -1

# Create inbound connector
$inboundConnectorName = 'Proofpoint Essentials - Inbound Connector'
New-InboundConnector -Name $inboundConnectorName -SenderDomains '*' -SenderIPAddresses $proofpointCIDRs -RequireTLS $true -ConnectorType Partner -RestrictDomainsToIPAddresses $true

# Create outbound connector
$outboundConnectorName = 'Proofpoint Essentials - Outbound Connector'
$smartHost = 'outbound-us1.ppe-hosted.com'
New-OutboundConnector -Name $outboundConnectorName -RecipientDomains '*' -SmartHosts $smartHost -TlsSettings EncryptionOnly -UseMXRecord $false