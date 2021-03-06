# ![SDL](/Images/sdl_logo.png) Tridion BluePrint Data Visualisation in Visio Professional 2016




I have been working on a concept for a [**TDS 2018**](http://2018.tridiondevelopersummit.com/register-tds-2018/)  presentation which is about Tridion  Data Visualization.

There are a lot of places where Tridion structure and configuration is difficult to surface that can be resolved by integrating with MS Office.

The presentation shows automated techniques using PowerShell to visualise Tridion information in MS Excel and MS Visio.
This automatically produces Tridion Blue Prints as Visio Diagrams, creating individual blueprint diagrams targeting 
+ the whole organisation
+ individual Business Units 
+ individual Publications

## Step 1
Export the Tridion Blueprint data to json so that we can work on the data without a connection to the live CMS.
[**SaveBluePrintDataToFile.ps1**](SaveBluePrintDataToFile.ps1)

[Saving Tridion Blueprint to Json](SaveTridionBluePrintDataAsJson.mp4)


## Step 2 
Process the captured data into Microsoft Excel Tables and Worksheets
[**ExportBlueprintAsExcelWorksheets2.ps1**](ExportBlueprintAsExcelWorksheets2.ps1)


## Step 3 
Use Visio Data Visualization templates to connect to the Excel data to automatically draw the diagrams

## Step 4 
Collect Item Type Data for each Publication 
[**GetCountItemTypesbyParentPublication.ps1**](GetCountItemTypesbyParentPublication.ps1)

## Step 5 
Add external data to excel diagram

