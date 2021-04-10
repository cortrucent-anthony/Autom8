#   ______     ______     ______     ______   ______     __  __     ______     ______     __   __     ______  
#  /\  ___\   /\  __ \   /\  == \   /\__  _\ /\  == \   /\ \/\ \   /\  ___\   /\  ___\   /\ "-.\ \   /\__  _\ 
#  \ \ \____  \ \ \/\ \  \ \  __<   \/_/\ \/ \ \  __<   \ \ \_\ \  \ \ \____  \ \  __\   \ \ \-.  \  \/_/\ \/ 
#   \ \_____\  \ \_____\  \ \_\ \_\    \ \_\  \ \_\ \_\  \ \_____\  \ \_____\  \ \_____\  \ \_\\"\_\    \ \_\ 
#    \/_____/   \/_____/   \/_/ /_/     \/_/   \/_/ /_/   \/_____/   \/_____/   \/_____/   \/_/ \/_/     \/_/ 

<#
    .SYNOPSIS
        Enable all Proofpoint features.

    .DESCRIPTION
        Enables the Proofpoint inbound connector, outbound connector, and creates mail flow rule bypassing the O365 spam filter.

    .EXAMPLE
        Enable-Proofpoint

    .NOTES
        TBD
#>
function Enable-Proofpoint {
     if (EXOConnected){
        Disable-EXOSpamfilter
        Enable-ProofpointInbound
        Enable-ProofpointOutbound
     } else {
         Write-Error -Message 'Not connected to Exchange Online. Please run Connect-ExchangeOnline and try again.'
     }
}

<#
    .SYNOPSIS
        Disable all Proofpoint features.

    .DESCRIPTION
        Disable the Proofpoint inbound connector, outbound connector, and creates mail flow rule bypassing the O365 spam filter.

    .EXAMPLE
        Disable-Proofpoint

    .NOTES
        TBD
#>
function Disable-Proofpoint {
    if (EXOConnected){
       Enable-EXOSpamfilter
       Disable-ProofpointInbound
       Disable-ProofpointOutbound
    } else {
        Write-Error -Message 'Not connected to Exchange Online. Please run Connect-ExchangeOnline and try again.'
    }
}

<#
    .SYNOPSIS
        Enable O365 spam filter.

    .DESCRIPTION
        Remove mail flow rule to bypass built-in O365 spam filter for email coming in from Proofpoint.

    .EXAMPLE
        Enable-EXOSpamFilter.

    .NOTES
        TBD.
#>
function Enable-EXOSpamfilter {
    if (EXOConnected){
        # TODO: Identify rule by action instead of name for a better universal-use success rate.
        Get-TransportRule | Where-Object { $_.Name -match "Bypass O365" } | Remove-TransportRule -Confirm:$false
    } else {
        Write-Error -Message 'Not connected to Exchange Online. Please run Connect-ExchangeOnline and try again.'
    }
}

<#
    .SYNOPSIS
        Bypass O365 spam filter.

    .DESCRIPTION
        Create mail flow rule to bypass built-in O365 spam filter for email coming in from Proofpoint.

    .EXAMPLE
        Disable-EXOSpamFilter.

    .NOTES
        TBD.
#>
function Disable-EXOSpamfilter {
    if (EXOConnected){
        $transportRuleName = 'Proofpoint Essentials - Bypass O365 Spam Filtering'
        New-TransportRule -Name $transportRuleName -SenderIpRanges $ProofpointIPs -SetSCL -1 
    } else {
        Write-Error -Message 'Not connected to Exchange Online. Please run Connect-ExchangeOnline and try again.'
    }
}

<#
    .SYNOPSIS
    Enable Proofpoint inbound.

    .DESCRIPTION
    'Enables the Proofpoint receive connector.'

    .EXAMPLE
    An example

    .NOTES
    General notes
#>
function Enable-ProofpointInbound {
    if (EXOConnected){
        $inboundConnectorName = 'Proofpoint Essentials - Inbound Connector'
        New-InboundConnector -Name $inboundConnectorName -SenderDomains '*' -SenderIPAddresses $ProofpointIPs -RequireTLS $true -ConnectorType Partner -RestrictDomainsToIPAddresses $true     
    } else {
        Write-Error -Message 'Not connected to Exchange Online. Please run Connect-ExchangeOnline and try again.'
    }
}

<#
    .SYNOPSIS
        Disables Proofpoint inbound.

    .DESCRIPTION
        Removes the Proofpoint receive connector.

    .EXAMPLE
        Disable-ProofpointInbound

    .NOTES
        TBD.
#>
function Disable-ProofpointInbound {
    if (EXOConnected){
        Get-InboundConnector | Where-Object { $_.Name -match "Proofpoint" } | Remove-InboundConnector -Confirm:$false
    } else {
        Write-Error -Message 'Not connected to Exchange Online. Please run Connect-ExchangeOnline and try again.'
    }
}

<#
    .SYNOPSIS
        Enable Proofpoint inbound.

    .DESCRIPTION
        Configures & enables the Proofpoint receive connector.

    .EXAMPLE
        Enable-ProofpointOutbound

    .NOTES
        TBD.
#>
function Enable-ProofpointOutbound {
    if (EXOConnected){
        $outboundConnectorName = 'Proofpoint Essentials - Outbound Connector'
        New-OutboundConnector -Name $outboundConnectorName -RecipientDomains '*' -SmartHosts $SmartHost -TlsSettings EncryptionOnly -UseMXRecord $false -Confirm:$false
    } else {
        Write-Error -Message 'Not connected to Exchange Online. Please run Connect-ExchangeOnline and try again.'
    }
}

<#
    .SYNOPSIS
        Disable Proofpoint outbound.

    .DESCRIPTION
        Removes the Proofpoint send connector.

    .EXAMPLE
        Disable-ProofpointOutbound

    .NOTES
        TBD.
#>
function Disable-ProofpointOutbound {
    if (EXOConnected){
        Get-OutboundConnector | Where-Object { $_.Name -match "Proofpoint" } | Remove-OutboundConnector -Confirm:$false
    } else {
        Write-Error -Message 'Not connected to Exchange Online. Please run Connect-ExchangeOnline and try again.'
    }
}

<#
    === Utility ===
#>

# GLOBAL DECLARATIONS

# Smart host name record
$SmartHost = 'outbound-us1.ppe-hosted.com'

# Proofpoint static blocks
$ProofpointIPs = "67.231.152.0/24","67.231.153.0/24","67.231.154.0/24",
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

# UTILITY FUNCTIONS
<#
    .SYNOPSIS
        Check if connected to Exchange Online.

    .DESCRIPTION
        Checks existing PSSessions to determine if we're connected to Exchange Online or not.

    .EXAMPLE
        if (EXOConnected){
            # Connected logic
        } else {
            # Not connected logic
        }

    .NOTES
        TBD.
#>
function EXOConnected {
    return ($null -ne (Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' }))
}