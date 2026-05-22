#Export serch service managed property mappings.
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$ssa = Get-SPEnterpriseSearchServiceApplication "Search Service Application"

$outputPath = "C:\Temp\SearchManagedProperties_Backup_WithMappings.csv"

$result = @()

$managedProperties = Get-SPEnterpriseSearchMetadataManagedProperty -SearchApplication $ssa |
    Sort-Object Name

foreach ($mp in $managedProperties) {

    Write-Host "Reading: $($mp.Name)" -ForegroundColor Cyan

    $mappings = @(
        Get-SPEnterpriseSearchMetadataMapping `
            -SearchApplication $ssa `
            -ManagedProperty $mp `
            -ErrorAction SilentlyContinue
    )

    $crawledProperties = @()

    foreach ($mapping in $mappings) {
        if ($mapping.CrawledPropertyName) {
            $crawledProperties += $mapping.CrawledPropertyName
        }
        elseif ($mapping.CrawledProperty -and $mapping.CrawledProperty.Name) {
            $crawledProperties += $mapping.CrawledProperty.Name
        }
        else {
            $crawledProperties += $mapping.ToString()
        }
    }

    $result += New-Object PSObject -Property @{
        Name                = $mp.Name
        Type                = $mp.Type
        Description         = $mp.Description
        Queryable           = $mp.Queryable
        Retrievable         = $mp.Retrievable
        FullTextQueriable   = $mp.FullTextQueriable
        EnabledForScoping   = $mp.EnabledForScoping
        Refinable           = $mp.Refinable
        Sortable            = $mp.Sortable
        HasMultipleValues   = $mp.HasMultipleValues
        SafeForAnonymous    = $mp.SafeForAnonymous
        DeleteDisallowed    = $mp.DeleteDisallowed
        MappingDisallowed   = $mp.MappingDisallowed
        Aliases             = ($mp.Aliases -join ";")
        CrawledProperties   = ($crawledProperties -join ";")
    }
}

$result |
    Sort-Object Name |
    Export-Csv $outputPath -NoTypeInformation -Encoding UTF8

Write-Host "Managed properties and mappings exported to: $outputPath" -ForegroundColor Green

    ----


  #import serch service managed property mappings.
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$ssaName = "Search Service Application"
$csvPath = "C:\Temp\SearchManagedProperties_Backup_WithMappings.csv"

$ssa = Get-SPEnterpriseSearchServiceApplication $ssaName

function Convert-ToBoolOrNull($value) {
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $null
    }

    return [System.Convert]::ToBoolean($value)
}

function Convert-ManagedPropertyType($typeValue) {
    if ([string]::IsNullOrWhiteSpace($typeValue)) {
        throw "Managed property Type is empty."
    }

    switch ($typeValue.ToString().Trim().ToLower()) {
        "1"        { return 1 }
        "text"     { return 1 }

        "2"        { return 2 }
        "integer"  { return 2 }

        "3"        { return 3 }
        "decimal"  { return 3 }

        "4"        { return 4 }
        "datetime" { return 4 }
        "date time" { return 4 }

        "5"        { return 5 }
        "yesno"    { return 5 }
        "yes no"   { return 5 }
        "boolean"  { return 5 }

        "6"        { return 6 }
        "binary"   { return 6 }

        "7"        { return 7 }
        "double"   { return 7 }

        default {
            throw "Unknown managed property type: $typeValue"
        }
    }
}

function Add-ParamIfNotNull {
    param(
        [hashtable]$Hash,
        [string]$Name,
        $Value
    )

    if ($null -ne $Value -and $Value -ne "") {
        $Hash[$Name] = $Value
    }
}

$rows = Import-Csv $csvPath

foreach ($row in $rows) {

    # IMPORTANT:
    # Do not restore every built-in SharePoint property.
    # Adjust this filter according to your naming convention.
    #
    # Examples:
    #   HVL*
    #   YourCompany*
    #   RefinableString00
    #   RefinableDate00
    #   RefinableInt00

    $shouldRestore =
        ($row.Name -like "HVL*") -or
        ($row.Name -like "YourCompany*") -or
        ($row.Name -like "RefinableString*") -or
        ($row.Name -like "RefinableDate*") -or
        ($row.Name -like "RefinableInt*") -or
        ($row.Name -like "RefinableDecimal*") -or
        ($row.Name -like "RefinableDouble*")

    if (-not $shouldRestore) {
        Write-Host "Skipping built-in or unrelated property: $($row.Name)" -ForegroundColor DarkGray
        continue
    }

    Write-Host "Processing managed property: $($row.Name)" -ForegroundColor Cyan

    $existingMp = Get-SPEnterpriseSearchMetadataManagedProperty `
        -SearchApplication $ssa `
        -Identity $row.Name `
        -ErrorAction SilentlyContinue

    if ($null -eq $existingMp) {

        $typeId = Convert-ManagedPropertyType $row.Type

        $newParams = @{
            SearchApplication = $ssa
            Name              = $row.Name
            Type              = $typeId
        }

        Add-ParamIfNotNull $newParams "Description"       $row.Description
        Add-ParamIfNotNull $newParams "Queryable"         (Convert-ToBoolOrNull $row.Queryable)
        Add-ParamIfNotNull $newParams "Retrievable"       (Convert-ToBoolOrNull $row.Retrievable)
        Add-ParamIfNotNull $newParams "FullTextQueriable" (Convert-ToBoolOrNull $row.FullTextQueriable)
        Add-ParamIfNotNull $newParams "EnabledForScoping" (Convert-ToBoolOrNull $row.EnabledForScoping)
        Add-ParamIfNotNull $newParams "SafeForAnonymous"  (Convert-ToBoolOrNull $row.SafeForAnonymous)

        $mp = New-SPEnterpriseSearchMetadataManagedProperty @newParams

        Write-Host "Created: $($row.Name)" -ForegroundColor Green
    }
    else {

        # Existing properties, especially RefinableStringXX, should be updated, not recreated.
        # Set-SPEnterpriseSearchMetadataManagedProperty does not support every field from New.
        $setParams = @{
            SearchApplication = $ssa
            Identity          = $existingMp
        }

        Add-ParamIfNotNull $setParams "Description"       $row.Description
        Add-ParamIfNotNull $setParams "Retrievable"       (Convert-ToBoolOrNull $row.Retrievable)
        Add-ParamIfNotNull $setParams "FullTextQueriable" (Convert-ToBoolOrNull $row.FullTextQueriable)
        Add-ParamIfNotNull $setParams "EnabledForScoping" (Convert-ToBoolOrNull $row.EnabledForScoping)
        Add-ParamIfNotNull $setParams "SafeForAnonymous"  (Convert-ToBoolOrNull $row.SafeForAnonymous)

        Set-SPEnterpriseSearchMetadataManagedProperty @setParams

        $mp = Get-SPEnterpriseSearchMetadataManagedProperty `
            -SearchApplication $ssa `
            -Identity $row.Name

        Write-Host "Updated existing: $($row.Name)" -ForegroundColor Yellow
    }

    if (-not [string]::IsNullOrWhiteSpace($row.CrawledProperties)) {

        $crawledPropertyNames = $row.CrawledProperties -split ";" |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -ne "" }

        foreach ($cpName in $crawledPropertyNames) {

            $crawledProperties = @(
                Get-SPEnterpriseSearchMetadataCrawledProperty `
                    -SearchApplication $ssa `
                    -Name $cpName `
                    -Limit All `
                    -ErrorAction SilentlyContinue
            )

            if ($crawledProperties.Count -eq 0) {
                Write-Warning "Crawled property not found: $cpName. Run a full crawl first, then rerun this script."
                continue
            }

            $cp = $crawledProperties |
                Where-Object { $_.CategoryName -eq "SharePoint" } |
                Select-Object -First 1

            if ($null -eq $cp) {
                $cp = $crawledProperties | Select-Object -First 1
            }

            try {
                New-SPEnterpriseSearchMetadataMapping `
                    -SearchApplication $ssa `
                    -ManagedProperty $mp `
                    -CrawledProperty $cp `
                    -ErrorAction Stop

                Write-Host "Mapped $cpName -> $($row.Name)" -ForegroundColor Green
            }
            catch {
                Write-Warning "Could not map $cpName -> $($row.Name). It may already exist. Error: $($_.Exception.Message)"
            }
        }
    }
}

The order should be:
1. Export CSV before deleting old Search Service Application
2. Delete/recreate Search Service Application
3. Configure content sources
4. Run full crawl once
5. Run the import script
6. Run full crawl again
7. Test search/refiners/display templates



  -----

 #   Clean up search index. Remove orphaned crawled and managed properties
 #https://sharepoint.stackexchange.com/questions/129352/clean-up-search-index-remove-orphaned-crawled-and-managed-properties
 if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) 
{
    Add-PSSnapin "Microsoft.SharePoint.PowerShell"
}

$categoryName = "SharePoint"  

function RemoveCrawledProperty($crawledPropertyName)  
{  
    Write-Host "If a CrawledProperty has option include in index enabled it won be deleted" -foregroundcolor green  

    $searchapp = Get-SPEnterpriseSearchServiceApplication
    #$category = Get-SPEnterpriseSearchMetadataCategory -Identity $categoryName -SearchApplication $searchAppName  
    $category = Get-SPEnterpriseSearchMetadataCategory -Identity $categoryName -SearchApplication $searchapp  
    $crawledProperty = Get-SPEnterpriseSearchMetadataCrawledProperty -Name $crawledPropertyName -SearchApplication $searchAppName -Category $category  
    if ($crawledProperty)  
    {  
        Write-Host "CrawledProperty found for '$crawledPropertyName'." -foregroundcolor yellow  
        $mappings = Get-SPEnterpriseSearchMetadataMapping -SearchApplication $searchAppName -CrawledProperty $crawledProperty  
        if ($mappings)
        {
            Write-Host "Mappings found for '$crawledPropertyName'." -foregroundcolor yellow  
            continue
        }  
        else  
        {  
            Write-Host "No mappings found for '$crawledPropertyName'." -foregroundcolor yellow  
        }  
        $category.DeleteUnmappedProperties()  
        $category.Update()
    }  
    else  
    {  
        Write-Host "Crawled property '$crawledPropertyName' not found." -foregroundcolor yellow  
    }  
}  

RemoveCrawledProperty "<name of property>"
