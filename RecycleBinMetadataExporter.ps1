# Create a Shell COM object
$shell = New-Object -ComObject Shell.Application

# Get the Recycle Bin folder
$recycleBin = $shell.Namespace(0xA)

# Get computer name
$computerName = $env:COMPUTERNAME

# Total size of the recycle bin to be calculated
$totalByteSize = 0

# Get number of items in the recycle bin
$numItems = $recycleBin.Items().Count

# Get current date
$dateRun = Get-Date -Format "yyyy-MM-dd hh:mm:ss tt"
$fileDate = [datetime]::parse($dateRun, $null).ToString("yyyyMMdd-HHmmss")

# Get all items in the Recycle Bin
$items = $recycleBin.Items()

# Create an empty array to store the metadata
$metadata = @()

# Iterate through each item and retrieve metadata
foreach ($item in $items) {
    $byteSize = $item.ExtendedProperty("size")
    if (![string]::IsNullOrEmpty($byteSize)) {
        $totalByteSize += [double]$byteSize
    }

    $dateDeleted = $recycleBin.GetDetailsOf($item, 2)
    $dateDeleted = $dateDeleted -creplace '\P{IsBasicLatin}' # Remove any non-latin characters
    $dateDeleted = [datetime]::parse($dateDeleted, $null).ToString("yyyy-MM-dd hh:mm:ss tt")

    $dateModified = $recycleBin.GetDetailsOf($item, 5)
    $dateModified = $dateModified -creplace '\P{IsBasicLatin}' # Remove any non-latin characters
    $dateModified = [datetime]::parse($dateModified, $null).ToString("yyyy-MM-dd hh:mm:ss tt")

    $dateCreated = $recycleBin.GetDetailsOf($item, 6)
    $dateCreated = $dateCreated -creplace '\P{IsBasicLatin}' # Remove any non-latin characters
    $dateCreated = [datetime]::parse($dateCreated, $null).ToString("yyyy-MM-dd hh:mm:ss tt")

    $dateAccessed = $recycleBin.GetDetailsOf($item, 7)
    $dateAccessed = $dateAccessed -creplace '\P{IsBasicLatin}' # Remove any non-latin characters
    $dateAccessed = [datetime]::parse($dateAccessed, $null).ToString("yyyy-MM-dd hh:mm:ss tt")

    $metadata += [PSCustomObject]@{
        Name = $recycleBin.GetDetailsOf($item, 0)
        OriginalLocation = $recycleBin.GetDetailsOf($item, 1)
        DateDeleted = $dateDeleted
        DateModified = $dateModified
        DateCreated = $dateCreated
        DateAccessed = $dateAccessed
        Size = $recycleBin.GetDetailsOf($item, 3)
        Type = $recycleBin.GetDetailsOf($item, 4)
        ByteSize = $byteSize
    }
}

# Headers and first row data with export metadata
$firstRow = [PSCustomObject]@{
    Name = $metadata[0].Name
    OriginalLocation = $metadata[0].OriginalLocation
    DateDeleted = $metadata[0].DateDeleted
    DateModified = $metadata[0].DateModified
    DateCreated = $metadata[0].DateCreated
    DateAccessed = $metadata[0].DateAccessed
    Size = $metadata[0].Size
    Type = $metadata[0].Type
    ByteSize = $metadata[0].ByteSize
    "RecycleBinMetadataExport`n$dateRun`n$totalByteSize Bytes`n$numItems Items`n$computerName" = ""
}
$metadata = $metadata[1..($metadata.Length - 1)]

# Combine headers, first row, and recycle bin metadata
$csvData = @($firstRow) + $metadata

# Specify the path to the CSV file
$csvPath = "C:\Users\power\Coding Projects\Python Projects\RecycleBinMetaExporter\RecycleBinMetadataExport_$fileDate.csv"

# Export metadata to CSV
$csvData | Export-Csv -Path $csvPath -Encoding UTF8 -NoTypeInformation

# Set read-only
Set-ItemProperty -Path $csvPath -Name IsReadOnly -Value $true

# Display a success message
"Recycle Bin metadata exported to $csvPath."
