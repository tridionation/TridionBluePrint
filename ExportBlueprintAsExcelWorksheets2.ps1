<#
TridionBluePrintProject
Step 2 export Blueprint as Excel Worksheets
#>

# For Demo or Testing 
# set to true to restrict to first 10 publications 
# 
$DemoOrDebug = $false 

$Modules = @("ImportExcel")
foreach($Module in $Modules)
{
    if(-not (Get-Module -ListAvailable -Name $Module)){ Install-Module $Module  -Force }
    Import-Module $Module
}

<#
.SYNOPSIS
    QuotedList
.DESCRIPTION
    Creates a quoted list

.PARAMETER list of items

#>
function Get-QuotedList()
{
    param(
        [array]$Items
    )
    # create comma seperated list of quoted items
    $quoteditems = @()
    foreach($item in $Items)
    {   
        $quoteditems += [string]::Format("`"{0}`"",$item) 
    }        
    return ($quoteditems -join ",")   
}

function Get-Ancestors()
{
    param(
        [array]$Id,
        [array]$Ancestors,
        [array]$publications
    )
    <#
        Required to build individual heirachy diagrams for business groups 
        and individual website publications
    #>

    $publication = $publications | Where-Object {$_.Id -eq $Id}
    $Ancestors += $publication.Id
    foreach($parent in $publication.Parents.IdRef)
    {                              
        $Ancestors += Get-Ancestors -Id $parent -Ancestors $Ancestors -Publications $publications         
    }

    return $Ancestors | Select-Object -Unique

}

function Get-Phase()
{
    param(
        [string]$Title
    )
    <#
    Insert your Business Rules to create Cross Functional Workflow "Phases"

    This is just a quick mash of the publication titles to group them by Business
    A real solution should probably use a Publication Metadata value.
    #>

        switch -Wildcard ($Title)
        {
            "*ATST*"    {$Phase = "ATST"}
            "*CA*"      {$Phase = "CA"}
            "*CAM*"     {$Phase = "CAM"}
            "*CE *"     {$Phase = "CE"}
            "*CIE*"     {$Phase = "CIE"}
            "*ESOL*"    {$Phase = "ESOL"}
            "*OCR*"     {$Phase = "OCR"}
            "*DXA*"     {$Phase = "DXA"}
            "*Example*" {$Phase = "DXA"}

            #Private sites
            "*Teacher Support*" {$Phase = "ODTS"}
            "*CIE Tactic*" {$Phase = "ASFA"}
            "*TSL Web*" {$Phase = "SSH"}
            "*CIE CPM*" {$Phase = "CPM"}
     
            "*Schema*" {$Phase = "a"}  
                                                     
            Default  {$Phase = "a"}
        }
    return $Phase
}

function Get-Function(){
    param(
        [string]$Title
    )

    <#
        Insert your Business Rules to create Cross Functional Workflow "Function"
        for ordering the publications in the heirarchy
        Here we extract the order number from the publication title.
        Your need to devise your own rules here if the titles do not have this ordering information
    #>
        $result = [string]::Format("`"{0}`"",$Title.split(" ")[0])
        if ($Title.split(" ")[0] -contains "Accelerated")
        {
            $result = [string]::Format("`"{0}`"",$Title.split(" ")[3])
        }
        if($Title.split(" ")[0] -contains "Decomm_800")
        {
            $result = [string]::Format("`"{0}`"",800)
        }

    return $result

}

function Get-Decendants()
{
    param(
        [array]$Id,
        [array]$Ancestors
    )
}

function Get-ProcessStepDescription()
{
    param(
        $publication
    )
    $result =  $($publication.Title.ToString() -replace '^\d{3}\s' , "" ).Trim()
    return $result
}
function Get-ProcessStepId()
{
    param(
        $publication
    )
    $result = [string]::Format("`"{0}`"",$publication.Id.ToString())
    return $result
}


        





function Get-WSData()
{
    param(
        [array]$publications
    )

    $RootPublicationId =  $publications | 
        Where-Object {$_.parents.count -eq 0} | 
        Select-Object -ExpandProperty Id

    $BPdata = @()
    $counter = 0
    $total = $publications.count
    $zeropadding = "0" * $total.ToString().length



    if($DemoOrDebug){
        $publications = $publications | Select-Object -First 10
    }

    foreach($publication in $publications)
    {
 
        $counter +=1
        $progress = @{
		    Activity = "SDL* Tridion Publication: $($publication.Title)"
		    PercentComplete = ($counter / $total) * 100
		    Status = "Processing $($counter.ToString("$zeropadding")) of $total" 
	    }
        Write-Progress @progress

        # Create "Parents"
        $plist = Get-QuotedList -Items $publication.Parents.IdRef

        # Create "Children"
        $children = $publications | Where-Object {$_.Parents.Idref -eq $publication.Id} 
        $clist = Get-QuotedList -Items $children.Id

        # Create Ancestors
        $ancestors = Get-Ancestors -Id $publication.id -Ancestors @() -Publications $publications
        $alist = Get-QuotedList -Items $ancestors

        # Create "Phase"
        $Phase = Get-Phase -Title $publication.Title

        # Create "Function"
        $Function = Get-Function -Title $publication.Title

        # Create "Process Step Description"
        $psd = Get-ProcessStepDescription -publication $publication

        $psid = Get-ProcessStepId -publication $publication 


        # Create Tridion Publication Link
        $hostname = $tcsconnection.hostname
        $EncodedId = [System.Web.HttpUtility]::UrlEncode($publication.Id)      
        $link = [string]::Format("http://{0}/SDL/#app=wcm&entry=cme&url=%23locationId%3D{1}",$hostname,$EncodedId )
    
        #Create Timestamp   
        $ts = Get-Date -Format u


        # Add New Cross-Functional Flow Chart Process
        $BPPubliction = @{}
        $BPPubliction.Add('Process Step ID'          , $psid )
        $BPPubliction.Add('Process Step Description' , $psd )
        $BPPubliction.Add('Next Step ID'             , "$clist" )
        $BPPubliction.Add('Function'                 , $Function )
        $BPPubliction.Add('Phase'                    , $Phase )

        $BPPubliction.Add('Parents'    , "$plist" )
        $BPPubliction.Add('Children'   , "$clist" )
        $BPPubliction.Add('Ancestors'  , "$alist" )
        $BPPubliction.Add('Date'       , $ts )
        $BPPubliction.Add('Link'       , $link ) 

        $BPPubliction.Add('Shape Type'      , "Process" )
        $BPPubliction.Add('Connector Label' , "" )
        $BPPubliction.Add('Alt Description' , "" )
          
        $BPdata += New-Object PsObject -Property $BPPubliction 
    }
    return $BPdata
}  
 

 function Save-Worksheet()
{
    param(
        [hashtable] $WSparam,
        [array] $Wsdata
         
    )

     $headers = @(
         "Process Step ID", 
         "Process Step Description",
         "Next Step ID",
         "Function", "Phase",
         "Shape Type",

         "Parents","Ancestors","Children",
         "Date", "Link",

         "Connector Label",
         "Alt Description"
     )

     $sortOrder = @(
        "Function","Phase"
     )
     # was Phase, Process Step Description

     $Subject = $WSparam.WorkSheetname
     $gridtitle = "/ $Subject  \\\_____    | SDL* Tridion BluePrint Data Visualization with Microsoft Excel and Microsoft Visio Professional 2016   |  © Tridionation $((Get-Date).Year)"
     $Wsdata  | 
        Select-Object -Property $headers| 
        Sort-Object Function, Phase | 
        Out-GridView -Title $gridtitle

    $Wsdata  | 
        Select-Object $headers | 
        Sort-Object $sortOrder | 
        Export-Excel @WSparam 
}


$WSparam = @{
    Path ="c:\temp\Blue.xlsx"
    TableName =  "Tridionation"
    TableStyle = "Medium6"
    WorkSheetname = "Tridionation"
    AutoSize = $true
    FreezeTopRowFirstColumn = $true
    BoldTopRow = $true
    NoNumberConversion = $true
    ClearSheet = $false
}

try{
    # Delete Excel Spreadsheet
    Remove-Item -Path $WSparam.Path -Force -ErrorAction SilentlyContinue
}
catch{
    Write-Warning " Please close the Excel Output file $($WSparam.Path)"
    break
}

# Rehydrate Publications object from JSON file
$jfile = Join-path $PSScriptRoot -ChildPath "TridionBlueprint.json"
$publications  = Get-Content $jfile -Raw | ConvertFrom-Json

# Add WorkSheet 
$MasterWsdata = Get-WSData -publications $publications
$WSparam
Save-Worksheet -WSparam $WSparam -Wsdata $MasterWsdata 

$publications.Count

$businesseNames = $MasterWsdata | Select-Object -ExpandProperty Phase -Unique
$businesseNames.Count

$Sites = $MasterWsdata  | Where-Object {(-not $_.Children)  }
$Sites.count


foreach($Site in $Sites)
{
    $name = $Site.'Process Step Description'

    $WSparam.WorkSheetname = $name
    $WSparam.TableName = $name -replace " ", ""
    $WSparam.ClearSheet = $false

    $AncestorIds = $Site.Ancestors -replace """", "" -split ","
    $AncestorPublications = $publications | Where-Object {$_.Id -in $AncestorIds}
    $AncestorPublications.Count

    $Wsdata = Get-WSData -publications $AncestorPublications
    $WSparam
    Save-Worksheet -WSparam $WSparam -Wsdata $Wsdata 
    
}






foreach($businessName in $businesseNames)
{
if($businessName -eq "a") {continue}
    $name = $businessName
    $WSparam.WorkSheetname = $name
    $WSparam.TableName = $name -replace " ", ""
    $WSparam.ClearSheet = $false

    $Sites = $MasterWsdata  | Where-Object { ($_.Phase -eq $businessName) -or ($_.Phase -eq "a" )}
    $x = $Sites.Ancestors -join "," -replace """" , ""
    $AncestorIds =  $x -split "," | Select-Object -Unique

    $AncestorPublications = $publications | Where-Object {$_.Id -in $AncestorIds}

    $Wsdata = Get-WSData -publications $AncestorPublications
    $WSparam
    Save-Worksheet -WSparam $WSparam -Wsdata $Wsdata 
}



