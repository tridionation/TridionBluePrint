<#
TridionBluePrintProject
Step 1
Save Blueprint data to disk

#>
 
using namespace Tridion.ContentManager.CoreService.Client

# For Demo or Testing 
# set to true to restrict to first 4 publications and 
# slim down json to speed up saving (REDUCED DATA set)
$DemoOrDebug = $false 

$Modules = @(
"Tridion-CoreService"
)
foreach($Module in $Modules)
{
    if(-not (Get-Module -ListAvailable -Name $Module)){ Install-Module $Module  -Force }
    Import-Module $Module
}

$user = ""
if(-not $credential){
    $credential = Get-Credential -UserName $user -Message "Remote Server Tridion Administrator Account"
}

$tcsconnection = @{
        hostname       = "cms.poc2.ucles.internal" 
        version        = "Web-8.5"   # 2011-SP1, 2013, 2013-SP1, Web-8.1, Web-8.5
        ConnectionType = "netTcp"    # Default, SSL, LDAP, LDAP-SSL, netTcp, Basic, Basic-SSL
        CredentialType = "Windows"
}
Set-TridionCoreServiceSettings @tcsconnection 

try
{
    Write-Progress -Activity "Tridion Core Service" -Status "Connecting ..."
    $client = Get-TridionCoreServiceClient
    $client.ChannelFactory.Credentials.Windows.ClientCredential = $credential
    $client.GetApiVersion() | Out-Null
    Write-Progress -Activity "Tridion Core Service" -Status "Connected to $($tcsconnection.hostname)"
}
catch 
{
    Write-Error "Failed to get data with TridionCoreService on $($tcsconnection.hostname)"
    break
}

# Get Publications 
    $filter = New-Object PublicationsFilterData      
    $publications = $client.GetSystemWideList($Filter)

# Get Publications Data
    $DataArray = @()
    $counter = 0
    $total = $SystemWideItems.count

    if($DemoOrDebug){
        $publications = $publications | Select-Object -First 4
    }

    foreach ($publication in $publications)
    {
        $counter += 1
        $progress = @{
			Activity = "Tridion Publication $($SystemWideItem.Title)"
			PercentComplete = ($counter / $total) * 100
			Status = "Processing $counter of $total" 
		}
        Write-Progress @progress

        $hostname = $tcsconnection.hostname
        $EncodedId = [System.Web.HttpUtility]::UrlEncode($publication.Id)      
        $link = [string]::Format("http://{0}/SDL/#app=wcm&entry=cme&url=%23locationId%3D{1}",$hostname,$EncodedId )
        $ts = Get-Date -Format u 

        $ht = $client.Read($publication.Id,$null)
        $ht | Add-Member -MemberType NoteProperty -Name "Date" -Value $ts
        $ht | Add-Member -MemberType NoteProperty -Name "Link" -Value $link
        
        if($DemoOrDebug){
            $ht.CategoriesXsd = $null
            $ht.AccessControlList = $null
        }
        
        $DataArray += $ht
    }


 


# Save Publication Data to file

$jfile = Join-path $PSScriptRoot -ChildPath "TridionBlueprint.json"
Write-Progress -Activity "Tridion Data " -Status "Saving data to $jfile" 
$DataArray | ConvertTo-Json -Depth 9 | Out-File -FilePath $jfile
Write-Progress -Completed -Activity "Tridion Data" -Status "Saved data to $jfile" 





