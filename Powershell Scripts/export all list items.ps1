
# ============================================================
# CONFIGURATION
# Edit these values before running the script in PowerShell ISE
# ============================================================

$WebUrl   = "http://sharepoint/sites/YourSite"
$ListName = "Your List Name"
$OutputCsv = "C:\Temp\YourList.csv"

# Number of items retrieved from SharePoint in each batch
$PageSize = 2000


# ============================================================
# SCRIPT
# ============================================================

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

# Load SharePoint PowerShell snap-in
if (-not (Get-PSSnapin -Name "Microsoft.SharePoint.PowerShell" `
                       -ErrorAction SilentlyContinue))
{
    Add-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction Stop
}

function Convert-SPFieldValueToText
{
    param
    (
        $Item,
        $Field
    )

    try
    {
        $rawValue = $Item[$Field.InternalName]
    }
    catch
    {
        return "[Unable to read field: $($_.Exception.Message)]"
    }

    if ($null -eq $rawValue)
    {
        return ""
    }

    try
    {
        $textValue = $Field.GetFieldValueAsText($rawValue)

        if ($null -eq $textValue)
        {
            return ""
        }

        return [string]$textValue
    }
    catch
    {
        try
        {
            if (($rawValue -is [System.Collections.IEnumerable]) -and
                -not ($rawValue -is [string]))
            {
                return (($rawValue | ForEach-Object {
                    [string]$_
                }) -join "; ")
            }

            return [string]$rawValue
        }
        catch
        {
            return "[Unable to convert field value: $($_.Exception.Message)]"
        }
    }
}

function Get-SPItemAttachmentData
{
    param
    (
        $Web,
        $List,
        $Item
    )

    $attachmentNames =
        New-Object "System.Collections.Generic.List[string]"

    $attachmentUrls =
        New-Object "System.Collections.Generic.List[string]"

    if ($List.EnableAttachments)
    {
        try
        {
            foreach ($attachmentNameValue in $Item.Attachments)
            {
                $attachmentName = [string]$attachmentNameValue

                [void]$attachmentNames.Add($attachmentName)

                $encodedFileName =
                    [System.Uri]::EscapeDataString($attachmentName)

                $serverRelativeUrl =
                    $List.RootFolder.ServerRelativeUrl.TrimEnd("/") +
                    "/Attachments/" +
                    $Item.ID +
                    "/" +
                    $encodedFileName

                $absoluteUrl =
                    $Web.Site.MakeFullUrl($serverRelativeUrl)

                [void]$attachmentUrls.Add($absoluteUrl)
            }
        }
        catch
        {
            Write-Warning (
                "Could not read attachments for item ID {0}: {1}" -f
                $Item.ID,
                $_.Exception.Message
            )
        }
    }

    return [PSCustomObject]@{
        Count = $attachmentNames.Count
        Names = $attachmentNames.ToArray()
        URLs  = $attachmentUrls.ToArray()
    }
}

$web = $null

try
{
    $fullOutputPath =
        [System.IO.Path]::GetFullPath($OutputCsv)

    $outputDirectory =
        Split-Path -Path $fullOutputPath -Parent

    if (-not [string]::IsNullOrWhiteSpace($outputDirectory))
    {
        if (-not (Test-Path -LiteralPath $outputDirectory))
        {
            New-Item `
                -Path $outputDirectory `
                -ItemType Directory `
                -Force | Out-Null
        }
    }

    if (Test-Path -LiteralPath $fullOutputPath)
    {
        Remove-Item `
            -LiteralPath $fullOutputPath `
            -Force
    }

    Write-Host "Opening SharePoint site: $WebUrl"

    $web = Get-SPWeb `
        -Identity $WebUrl `
        -ErrorAction Stop

    $list = $web.Lists.TryGetList($ListName)

    if ($null -eq $list)
    {
        throw "The list '$ListName' was not found in '$WebUrl'."
    }

    Write-Host "List found: $($list.Title)"
    Write-Host "List item count: $($list.ItemCount)"

    # Include all custom, default, hidden and system fields
    $fieldMappings = @(
        foreach ($field in $list.Fields)
        {
            $displayName = $field.Title

            if ([string]::IsNullOrWhiteSpace($displayName))
            {
                $displayName = $field.InternalName
            }

            [PSCustomObject]@{
                Field  = $field
                Header = "{0} [{1}]" -f `
                    $displayName,
                    $field.InternalName
            }
        }
    )

    Write-Host "Fields to export: $($fieldMappings.Count)"

    $query =
        New-Object Microsoft.SharePoint.SPQuery

    $query.RowLimit = [uint32]$PageSize

    # Include items located inside folders
    $query.ViewAttributes = "Scope='RecursiveAll'"

    $query.IncludeMandatoryColumns = $true

    $position = $null
    $totalExported = 0
    $isFirstBatch = $true

    do
    {
        $query.ListItemCollectionPosition = $position

        $items = $list.GetItems($query)

        $exportRows =
            New-Object "System.Collections.Generic.List[object]"

        foreach ($item in $items)
        {
            $row = [ordered]@{}

            foreach ($fieldMapping in $fieldMappings)
            {
                $row[$fieldMapping.Header] =
                    Convert-SPFieldValueToText `
                        -Item $item `
                        -Field $fieldMapping.Field
            }

            $attachmentData =
                Get-SPItemAttachmentData `
                    -Web $web `
                    -List $list `
                    -Item $item

            $row["__AttachmentCount"] =
                $attachmentData.Count

            $row["__AttachmentNames"] =
                $attachmentData.Names -join "`r`n"

            $row["__AttachmentUrls"] =
                $attachmentData.URLs -join "`r`n"

            [void]$exportRows.Add(
                [PSCustomObject]$row
            )
        }

        if ($exportRows.Count -gt 0)
        {
            if ($isFirstBatch)
            {
                $exportRows |
                    Export-Csv `
                        -LiteralPath $fullOutputPath `
                        -NoTypeInformation `
                        -Encoding UTF8

                $isFirstBatch = $false
            }
            else
            {
                $exportRows |
                    Export-Csv `
                        -LiteralPath $fullOutputPath `
                        -NoTypeInformation `
                        -Encoding UTF8 `
                        -Append
            }
        }

        $totalExported += $exportRows.Count

        $position =
            $items.ListItemCollectionPosition

        if ($list.ItemCount -gt 0)
        {
            $percentComplete = [Math]::Min(
                100,
                [Math]::Round(
                    ($totalExported / $list.ItemCount) * 100,
                    0
                )
            )
        }
        else
        {
            $percentComplete = 100
        }

        Write-Progress `
            -Activity "Exporting list '$($list.Title)'" `
            -Status "$totalExported item(s) exported" `
            -PercentComplete $percentComplete
    }
    while ($null -ne $position)

    Write-Progress `
        -Activity "Exporting list '$($list.Title)'" `
        -Completed

    # Create a CSV containing only headers if the list is empty
    if ($isFirstBatch)
    {
        $emptyRow = [ordered]@{}

        foreach ($fieldMapping in $fieldMappings)
        {
            $emptyRow[$fieldMapping.Header] = ""
        }

        $emptyRow["__AttachmentCount"] = ""
        $emptyRow["__AttachmentNames"] = ""
        $emptyRow["__AttachmentUrls"] = ""

        $csvLines =
            [PSCustomObject]$emptyRow |
            ConvertTo-Csv -NoTypeInformation

        Set-Content `
            -LiteralPath $fullOutputPath `
            -Value $csvLines[0] `
            -Encoding UTF8
    }

    Write-Host ""
    Write-Host "Export completed successfully." `
        -ForegroundColor Green

    Write-Host "Items exported : $totalExported"
    Write-Host "CSV file       : $fullOutputPath"
}
catch
{
    Write-Host ""
    Write-Host "Export failed:" `
        -ForegroundColor Red

    Write-Host $_.Exception.Message `
        -ForegroundColor Red

    throw
}
finally
{
    if ($null -ne $web)
    {
        $web.Dispose()
    }
}

