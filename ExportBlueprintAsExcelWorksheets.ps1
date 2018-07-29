<#
TridionBluePrintProject
Step 2 export Blueprint as Excel Worksheets
#>

$Modules = @("ImportExcel")
foreach($Module in $Modules)
{
    if(-not (Get-Module -Name $Module)){ Install-Module $Module  -Force }
    Import-Module $Module
}


    function QuotedList()
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
        function GetAncestors()
    {
        param(
        [array]$Id,
        [array]$Ancestors
        )


            $publication = $SystemWideItemsArray | Where-Object {$_.Id -eq $Id}
            $Ancestors += $publication.Id
            foreach($parent in $publication.Parents.IdRef)
            {                              
                $Ancestors += GetAncestors -Id $parent -Ancestors $Ancestors          
            }

        return $Ancestors | Select-Object -Unique

    }


        function GetDecendants()
    {
        param(
        [array]$Id,
        [array]$Ancestors
        )
    }






    #Rehydrate object from JSON file
    $jfile = Join-path $PSScriptRoot -ChildPath "TridionBlueprint.json"
    $SystemWideItemsArray  = Get-Content $jfile -Raw | ConvertFrom-Json

    $RootPublicationId =  $SystemWideItems | Where-Object {$_.parents.count -eq 0} | Select-Object -ExpandProperty Id

    $BPdata = @()
    $counter = 0
    $total = $SystemWideItemsArray.count

    foreach($publication in $SystemWideItemsArray)
    {
    
        $counter +=1
        $progress = @{
			Activity = "Processing $($publication.Title)"
			PercentComplete = ($counter / $total) * 100
			Status = "Processing $counter of $total" 
		}
        Write-Progress @progress

        # Create "Parents"
        $plist = QuotedList -Items $publication.Parents.IdRef

        # Create "Children"
        $children = $SystemWideItemsArray | Where-Object {$_.Parents.Idref -eq $publication.Id} 
        $clist = QuotedList -Items $children.Id

        # Create Ancestors
        $ancestors = GetAncestors -Id $publication.id -Ancestors @()
        $alist = QuotedList -Items $ancestors

        # Create "Phase"
        switch -Wildcard ($publication.Title)
        {
            "*ATST*" {$Phase = "ATST"}
            "*CA*"   {$Phase = "CA"}
            "*CAM*"  {$Phase = "CAM"}
            "*CE*"   {$Phase = "CE"}
            "*CIE*"  {$Phase = "CIE"}
            "*DXA*"  {$Phase = "a"}
            "*ESOL*" {$Phase = "ESOL"}
            "*OCR*"  {$Phase = "OCR"}
                        
            Default  {$Phase = "a"}
        }

        # Create "Function"
        $Function = [string]::Format("`"{0}`"",$publication.Title.split(" ")[0])
        if ($publication.Title.split(" ")[0] -contains "Accelerated")
        {
            $Function = [string]::Format("`"{0}`"",$publication.Title.split(" ")[3])
        }
        if($publication.Title.split(" ")[0] -contains "Decomm_800")
        {
            $Function = [string]::Format("`"{0}`"",800)
        }

        # Create "ProcessStep_Description"
        $psd =  $publication.Title.ToString() -replace '^\d{3}\s' , ""

        # Add New "Process"
        $BPPubliction = @{
            'Process Step ID' = [string]::Format("`"{0}`"",$publication.Id.ToString())
            'Process Step Description' = $psd 
            'Parents' = "$plist"
            'Children' = "$clist"
            'Ancestors' = "$alist"
            'Shape Type' = "Process"
            'Function' = $Function
            'Phase' = $Phase
        }
        $BPdata += New-Object PsObject -Property $BPPubliction 
    }

  
 


 $headers = @(
     "Process Step ID", 
     "Process Step Description",
     "Parents",
     "Next Step ID",
     "Shape Type",
     "Function",
     "Phase",
     "Ancestors",
     "Children"
 )
 $sortOrder = @("Phase","Process_Step_Description")

 #  $BPdata | Where-Object {($_.Phase -contains "a") -or ($_.Phase -contains "ESOL") } | Select $headers | Sort-Object  $sortOrder| Out-GridView 
 # $BPdata  | Select Process_Step_ID,Process_Step_Description,Parents,Children,Ancestors,Shape_Type,Function,Phase | Sort-Object Function, Phase | Out-GridView 
 #
 $BPdata  | Select  $headers| Sort-Object Function, Phase | Out-GridView 
 
 break

 # Also "Connector Label" could use Parent

<#
$BPdata | 
Select Process_Step_ID,Process_Step_Description,Parents,Children,Shape_Type,Function,Phase |  
Sort-Object Phase,Process_Step_Description | 
Export-XLSX -Path $filename -Table -TableStyle "Medium6" -Header $header -WorksheetName "Tridionation" -AutoFit -ReplaceSheet 
#>


$param = @{
    Path ="c:\temp\Blue.xlsx"
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
$BPdata  | 
    Select-Object $headers | 
    Sort-Object $sortOrder | 
    Export-Excel @param 

#ToDo
# filter children to next step id removing unrelated items 
$businesses = $BPdata | Select-Object -ExpandProperty Phase -Unique
foreach($business in $businesses)
{
# Multiple Worksheets
$param.WorkSheetname = $business
$param.TableName = $business
$BPdata | 
    Where-Object {($_.Phase -contains "a") -or ($_.Phase -contains $business) } | 
    Select-Object $headers | 
    Sort-Object $sortOrder | 
    Export-Excel @param 
}

#Websites
#ToDo
# filter children to next step id removing unrelated items 

$Sites = $BPdata | Where-Object {(-not $_.Children)  }

foreach($Site in $Sites)
{
    if($Site.Children){continue}
    # Multiple Worksheets
    $name = $Site.'Process Step Description'
    $id = $Site.'Process Step ID'
    $ancestors = $Site.Ancestors
    $param.WorkSheetname = $name
    $param.TableName = $name -replace " ", ""
    $BPdata | 
        Where-Object { 
        $ancestors.Contains( $_.'Process Step ID')  } | 
        Select-Object $headers | Sort-Object $sortOrder | 
        Export-Excel @param 
}