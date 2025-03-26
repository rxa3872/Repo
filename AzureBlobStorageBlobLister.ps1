<#
.SYNOPSIS
    AzureBlobStorageBlobLister.ps1

    Author: Rajesh Adapa

    This PowerShell script dynamically retrieves blobs from an Azure Storage Account container using a provided Shared Access Signature (SAS) token and Storage Account URL. It filters blobs based on user-defined search terms and exports the filtered blob URLs and details (Name, Last Modified, Size) into both a plain text (.txt) file and an Excel (.xlsx) file.

.DESCRIPTION
    The script performs the following operations:

    - Dynamically extracts the Azure Storage Account Name from the provided URL.
    - Connects securely to Azure Blob Storage using a SAS token.
    - Fetches all blob files from the specified container.
    - Filters blobs based on predefined search terms.
    - Exports the filtered list of blob URLs to a TXT file and detailed blob information to an Excel file.
    - Includes automated error handling, ensuring the script halts gracefully on errors.
    - Installs required modules (Az.Storage and ImportExcel) if they're not already installed.
    - Closes any open Excel instances automatically to avoid file lock conflicts.

.PARAMETERS
    ContainerName - Specifies the Azure Storage container name.
    SasToken      - Shared Access Signature token for secure storage access.
    StorageAccountUrl - URL for the Azure Storage account.
    SearchTerms   - Terms used to filter blob files.
    # Define search terms dynamically
    ##$SearchTerms = @("Production")
    #$SearchTerms = @("AZUSEYUMDBSQ468")
    #$SearchTerms = @("FinalBackup")
    #$SearchTerms = @("GETS")
    #$SearchTerms = @("ITSS")
    #$SearchTerms = @("DB2")
    ##$SearchTerms = @("YSSRepDB")

.EXAMPLE
    Update the parameters ($ContainerName, $SasToken, $StorageAccountUrl, and $SearchTerms) within the script and run it in PowerShell.

.NOTES
    Ensure you have adequate permissions to connect and read data from Azure Storage. Close any active Excel windows before running to prevent file locks.
#>

# Author: Rajesh Adapa
# Define storage account details dynamically
$ContainerName = "sqlfiles"
$SasToken =  "sv=2022-11-02&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2035-02-21T22:40:03Z&st=2025-02-21T14:40:03Z&spr=https&sig=H2MjNyEcwqzPkWYaiD97Ua0%2BQtEg4O%2FywUj6XkJdpRE%3D"

# Dynamically extract the Storage Account Name from the SAS Token URL
$StorageAccountUrl = "https://ybazusetyumdbusep001.blob.core.windows.net/"
$StorageAccountName = ($StorageAccountUrl -split "\.")[0] -replace "https://", ""

#---
# Define storage account details dynamically
#$ContainerName = "sqlfiles"
#$SasToken = "sv=2022-11-02&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2035-02-21T04:55:22Z&st=2025-02-20T20:55:22Z&spr=https&sig=rGb3Sh0xjXXcNDNY3qzQ7NNf6%2BIGYaRFGmxZJnhvcGQ%3D"

# Dynamically extract the Storage Account Name from the SAS Token URL
#$StorageAccountUrl = "https://azeastsqlstoragefileshar.blob.core.windows.net/"
#$StorageAccountName = ($StorageAccountUrl -split "\.")[0] -replace "https://", ""
#--



# Define storage account details
#$StorageAccountName = "ybazusetyumdbusep001"
#$ContainerName = "sqlfiles"
#$SasToken = "sv=2022-11-02&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2035-02-21T22:40:03Z&st=2025-02-21T14:40:03Z&spr=https&sig=H2MjNyEcwqzPkWYaiD97Ua0%2BQtEg4O%2FywUj6XkJdpRE%3D"

# Define variables
#$StorageAccountName = "ybazfrcstyumdbfrcp001"
#$StorageAccountName = "azeastsqlstoragefileshar"
#$StorageAccountName = "ybazusestyumnetbkp001"
#$StorageAccountName = "ybazusetyumdbusep001"
#$ContainerName = "sqlfiles"  # Container name
#$BlobPrefix = "Production/AZUSETBDB411N03/"  # Root folder (includes FULL, DIFF, LOG)
#$BlobPrefix = "FinalBackup/YUM ITSS Trintech ReconNET/"  # Root folder (includes FULL, DIFF, LOG)#
#$BlobPrefix = "uardb"  # Root folder (includes FULL, DIFF, LOG)#
#$BlobPrefix = "FinalBackup/.+GETS.*/"  # Root folder (includes FULL, DIFF, LOG)
#$BlobPrefix = "Production/AZESXDUSE414A/EPM_HFM/"  # Root folder (includes FULL, DIFF, LOG)
#$BlobPrefix = "Production/AZUSEYUMDBSP486/"  # Root folder (includes FULL, DIFF, LOG)
#$StorageAccountName = "azeastsqlstoragefileshar"
#$SasToken = "sv=2022-11-02&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2035-02-21T04:55:22Z&st=2025-02-20T20:55:22Z&spr=https&sig=rGb3Sh0xjXXcNDNY3qzQ7NNf6%2BIGYaRFGmxZJnhvcGQ%3D"
#$StorageAccountName = "ybazfrcstyumdbfrcp001"
#$SasToken =  "sv=2022-11-02&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2035-02-21T22:27:36Z&st=2025-02-21T14:27:36Z&spr=https&sig=47K1xGz9dMkfuLyPzLHlEkhEWXccle2UYpt10kLgf%2Bw%3D"
#$StorageAccountName = "ybazusestyumnetbkp001"
#$SasToken =  "sv=2022-11-02&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2035-02-21T22:49:57Z&st=2025-02-21T14:49:57Z&spr=https&sig=1Gmfc5GP5H55DWgTOwSEsHQRqwRkHDArpP2hUSE4llA%3D"
#$StorageAccountName = "ybazusetyumdbusep001"
#$SasToken =  "sv=2022-11-02&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2035-02-21T22:40:03Z&st=2025-02-21T14:40:03Z&spr=https&sig=H2MjNyEcwqzPkWYaiD97Ua0%2BQtEg4O%2FywUj6XkJdpRE%3D"



# Construct the base blob URL
$StorageUrl = "$StorageAccountUrl$ContainerName"

# Output file paths
$OutputFileTxt = "D:\rajesha\BlobList.txt"
$OutputFileXlsx = "D:\rajesha\BlobList.xlsx"
$TempXlsx = "D:\rajesha\BlobList_Temp.xlsx"

# Ensure the output directory exists
$OutputDir = Split-Path -Path $OutputFileTxt
if (!(Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force }

# Close any open Excel instances to avoid file lock issues
Write-Host "Checking for open Excel instances..."
Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2  # Ensures Excel closes properly

# Install necessary PowerShell modules if not already installed
if (!(Get-Module -ListAvailable -Name Az.Storage)) {
    Install-Module -Name Az.Storage -Force -Scope CurrentUser
}
if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Import required modules
Import-Module Az.Storage
Import-Module ImportExcel

# Get the storage context with error handling
try {
    Write-Host "Connecting to Azure Storage Account: $StorageAccountName ..."
    $StorageContext = New-AzStorageContext -StorageAccountName $StorageAccountName -SasToken $SasToken
} catch {
    Write-Host "ERROR: Unable to connect to Azure Storage. Check your StorageAccountName and SasToken."
    exit
}

# Fetch all blobs in the container
try {
    Write-Host "Fetching blobs from container '$ContainerName'... This may take a while..."
    $Blobs = Get-AzStorageBlob -Container $ContainerName -Context $StorageContext
} catch {
    Write-Host "ERROR: Unable to retrieve blobs. Check your container name and permissions."
    exit
}

# Verify if any blobs were retrieved
if ($Blobs.Count -eq 0) {
    Write-Host "No blobs found in container '$ContainerName'."
    exit
}

Write-Host "Total blobs found: $($Blobs.Count)"

# Display a few sample blob names for verification
Write-Host "Sample blob names:"
$Blobs | Select-Object -First 10 | ForEach-Object { Write-Host $_.Name }

# Define search terms dynamically
#$SearchTerms = @("Production")
#$SearchTerms = @("AZUSEYUMDBSQ468")
#$SearchTerms = @("FinalBackup")
#$SearchTerms = @("GETS")
#$SearchTerms = @("ITSS")
#$SearchTerms = @("DB2")
$SearchTerms = @("YSSRepDB")



Write-Host "Searching blobs for hardcoded values: $($SearchTerms -join ', ')"

# Filter blobs that match search terms
$FilteredBlobs = $Blobs | Where-Object { 
    $BlobName = $_.Name
    $SearchTerms | ForEach-Object { if ($BlobName -like "*$_*") { return $true } }
}

# Check if filtering found any results
if ($FilteredBlobs.Count -eq 0) {
    Write-Host "No blobs matched any of the search terms: $($SearchTerms -join ', ')"
    exit
}

Write-Host "Matching blobs found: $($FilteredBlobs.Count)"

# Construct full URLs for filtered blobs ensuring correct format
$BlobList = $FilteredBlobs | ForEach-Object { 
    [PSCustomObject]@{
        FullUrl = "$StorageUrl/$($_.Name)"  # Ensures full URL matches Azure blob structure
        Name = $_.Name
        LastModified = $_.LastModified
        Size_MB = [math]::Round($_.Length / 1MB, 2)
    }
}

# Export results to a TXT file (Only Full URLs)
$BlobList | ForEach-Object { $_.FullUrl } | Out-File -FilePath $OutputFileTxt -Encoding utf8

# Remove existing Excel file to prevent lock issues
if (Test-Path $OutputFileXlsx) {
    Write-Host "Removing old Excel file to prevent lock issues..."
    Remove-Item -Path $OutputFileXlsx -Force -ErrorAction SilentlyContinue
}

# Export results to an Excel file using ImportExcel module
try {
    Write-Host "Exporting to Excel..."
    $BlobList | Export-Excel -Path $TempXlsx -WorksheetName "BlobList" -AutoSize -FreezeTopRow
    Rename-Item -Path $TempXlsx -NewName $OutputFileXlsx -Force
} catch {
    Write-Host "ERROR: Failed to save the Excel file. Ensure no other programs (Excel, OneDrive, etc.) are using it."
    exit
}

Write-Host "Blob list saved to:"
Write-Host " - $OutputFileTxt (Plain Text - Full URLs Only)"
Write-Host " - $OutputFileXlsx (Excel File - Open in Microsoft 365)"
