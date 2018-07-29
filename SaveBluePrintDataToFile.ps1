<#
TridionBluePrintProject
Step 1
Save Blueprint data to disk

#>
 
using namespace Tridion.ContentManager.CoreService.Client

$Modules = @(
"Tridion-CoreService"
)

foreach($Module in $Modules)
{
    if(-not (Get-Module -Name $Module)){ Install-Module $Module  -Force }
    Import-Module $Module
}

$user = ""
if(-not $credential){
    $credential = Get-Credential -UserName $user -Message "Remote Server Tridion Administrator Account"
}

$tcsconnection = @{
        hostname       = "cms.poc2" 
        version        = "Web-8.5"   # 2011-SP1, 2013, 2013-SP1, Web-8.1, Web-8.5
        ConnectionType = "netTcp"    # Default, SSL, LDAP, LDAP-SSL, netTcp, Basic, Basic-SSL
        CredentialType = "Windows"
}
Set-TridionCoreServiceSettings @tcsconnection 

try
{
    $client = Get-TridionCoreServiceClient
    $client.ChannelFactory.Credentials.Windows.ClientCredential = $credential
    $client.GetApiVersion()

# Get Publications 
    $filter = New-Object PublicationsFilterData      
    $SystemWideItems = $client.GetSystemWideList($Filter)

# Get Publications Data
    $SystemWideItemsArray = @()
    $counter = 0
    $total = $SystemWideItems.count

    foreach ($SystemWideItem in $SystemWideItems)
    {
        $counter += 1
        $progress = @{
			Activity = "Get Publication $($SystemWideItem.Title)"
			PercentComplete = ($counter / $total) * 100
			Status = "Processing $counter of $total" 
		}
        Write-Progress @progress
        $SystemWideItemsArray += $client.Read($SystemWideItem.Id,$null)
    }

    Write-Progress -Activity "Saving Data to file"

    $client.Dispose()

# Save Publication Data to file
    $jfile = Join-path $PSScriptRoot -ChildPath "TridionBlueprint.json"
    $SystemWideItemsArray | ConvertTo-Json -Depth 9 | Out-File -FilePath $jfile
    Write-Progress -Completed -Activity "Saving Data to file"
}

catch 
{
    Write-Output "Failed to get data with TridionCoreService on $($tcsconnection.hostname)"
    exit
}


