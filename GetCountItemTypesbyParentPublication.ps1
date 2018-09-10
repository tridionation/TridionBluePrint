#
#
#
    #[Enum]::GetNames( [Tridion.ContentManager.CoreService.Client.ItemType] )
    #[Enum]::GetNames( [Tridion.ContentManager.CoreService.Client.LoadFlags] )
#
#
using namespace Tridion.ContentManager.CoreService.Client
Using namespace System.Uri

# For Demo or Testing 
# set to true to restrict to first 4 publications and 
# slim down json to speed up saving (REDUCED DATA set)
$DemoOrDebug = $false 

$Modules = @("ImportExcel","Tridion-CoreService")
foreach($Module in $Modules)
{
    if(-not (Get-Module -ListAvailable -Name $Module)){ Install-Module $Module  -Force }
    Import-Module $Module
}

$user = "UCLES\millsdsup"
if(-not $credential){
    $credential = Get-Credential -UserName $user -Message "Remote Server Tridion Administrator Account"
}


$tcsconnection = @{
        hostname       = "cms.poc2.ucles.internal" # Server name not FQDN
        version        = "Web-8.5"   # 2011-SP1, 2013, 2013-SP1, Web-8.1, Web-8.5
        ConnectionType = "Default"    # Default, SSL, LDAP, LDAP-SSL, netTcp, Basic, Basic-SSL
        ConnectionSendTimeout = "00:10:00"
}
Set-TridionCoreServiceSettings @tcsconnection 

try
{
    $client = Get-TridionCoreServiceClient
    $client.ChannelFactory.Credentials.Windows.ClientCredential = $credential
    $client.GetApiVersion() | Out-Null
    Write-Output "Connected to TridionCoreService on $($tcsconnection.hostname)"
}

catch 
{
    Write-Error "Failed to connect TridionCoreService on $($tcsconnection.hostname)"
    break
}

$ItemTypes = @(
    "BusinessProcessType",
    "Schema",
    "TemplateBuildingBlock",
    "ComponentTemplate",
    "PageTemplate",
    "Category",
    "Keyword",
    "Page",
    "Component"
)

$filter = New-Object PublicationsFilterData         
$publications = $client.GetSystemWideList($Filter)

$pcount = 0
$ptotal = $publications.Count
$itotal = $ItemTypes.Count
$zeropadding = "0" * $ptotal.ToString().length
$ItemTypeCounts = @()

if($DemoOrDebug){
    $publications = $publications | Select-Object -First 4
}

foreach ($publication in $publications  )
{  

$link = [string]::Format("http://{0}/SDL/#app=wcm&entry=cme&url=%23locationId%3D{1}",$tcsconnection.hostname,[System.Web.HttpUtility]::UrlEncode($publication.Id) )
$ts = Get-Date -Format u 
    $ht = @{}    
    $ht.Add("Id",$publication.Id)
    $ht.Add("Title",$publication.Title)
    $ht.Add("Link",$link)
    $ht.Add("Date",$ts)

    $pcount +=1
	$progress = @{
		Activity = "Tridion Publications"
		PercentComplete = ($pcount / $ptotal) * 100
		Status = "Processing publication $($pcount.ToString("$zeropadding")) of $ptotal - $($publication.Title) "
        Id = 1
	}
    Write-Progress @progress 

    $icount = 0

    foreach($ItemType in $ItemTypes)
    {

        $icount +=1
	    $progress = @{
		    Activity = "Tridion ItemTypes"
		    PercentComplete = ($icount / $itotal) * 100
		    Status = "Counting Tridion Items of Type [$ItemType] "
            Id = 2
	    }
        Write-Progress @progress 

        $filter = New-Object RepositoryItemsFilterData
        $filter.ItemTypes += [Tridion.ContentManager.CoreService.Client.ItemType]::$ItemType
        $filter.Recursive = $true
        $list = $client.GetList($publication.id,$filter)

        $NumberOfItems = $list | 
            Where-Object {$_.BluePrintInfo.OwningRepository.IdRef -eq $publication.Id } | 
            Measure-Object  | 
            Select-object -ExpandProperty Count

        $ht.Add("$ItemType",$NumberOfItems)
    }
    $ItemTypeCounts += New-Object PSObject -Property $ht

    $ht.ForEach({[PSCustomObject]$_}) | Format-Table -Property Date,Id,Title -HideTableHeaders 

}
    Write-Progress -Completed -Activity "Completed" -Id 2
    Write-Progress -Completed -Activity "Completed" -Id 1

$client.Dispose()


$gridtitle = "Publication Owning Item Type Counts  | SDL* Tridion BluePrint Data Visualization with Microsoft Excel and Microsoft Visio Professional 2016   | © Tridionation $((Get-Date).Year)"
$headers = @("Id","Title") + $ItemTypes + @("Date","Link")

$results = $ItemTypeCounts | Select-Object -Property $headers | Out-GridView -Title $gridtitle -PassThru 
if ($Results){
    foreach($result in $results){
    Start-Process -FilePath $result.Link
    }
}



$param = @{
    Path = [string]::Format("c:\temp\BlueCounts{0}.xlsx", $(Get-Date -Format "yyyyMMdd") )
    TableName =  "Tridionation"
    TableStyle = "Medium6"
    WorkSheetname = "Tridionation"
    AutoSize = $true
    FreezeTopRowFirstColumn = $true
    BoldTopRow = $true
    NoNumberConversion = $true
    ClearSheet = $true
}


Remove-Item $param.Path -Force -ErrorAction SilentlyContinue
$ItemTypeCounts  | 
    Select-Object $headers | 
    Export-Excel @param 


