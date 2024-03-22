param (
    [object]$WebhookData
)

$containerName = "applications"
$containerNameDestination = "intunewin"

try {
    $bodyJson = $WebhookData.RequestBody

    # Convert body
    $body = $bodyJson | ConvertFrom-Json

    # F
    $topic = $body.topic
    $subject = $body.subject
    $url = $body.data.url
} catch {
    Write-Error "An error occurred processing the webhook data: $_"
}

# Split the subject path
$splitSubject = $subject -split '/'

# Extract the blob file name (last element)
$blobFileName = $splitSubject[-1]

# Determine the index for 'blobs' to identify the start of the actual blob path
$blobsIndex = [array]::IndexOf($splitSubject, "blobs")

# Initialize blobFolderPath with foldername 'app'
$blobFolderPath = "app"

# Check if there are elements after 'blobs' before the file name
if ($blobsIndex -lt ($splitSubject.Length - 2)) {
    # Extract the folder path
    $folderPathElements = $splitSubject[($blobsIndex + 1)..($splitSubject.Length - 2)]

    # Rejoin the path elements to form the full folder path
    $blobFolderPath = ($folderPathElements -join '/')
}

Write-Output "Blob File Name: $blobFileName"
Write-Output "Full Folder Path: $blobFolderPath"

# Check if the filename ends with .exe or .msi
if (-not ($blobFileName.EndsWith(".exe") -or $blobFileName.EndsWith(".msi") -or $blobFileName.EndsWith(".zip"))) {
    Write-Output "The file is not an exe, msi or zip file. Exiting intune packaging tool."
    exit
}

# Split the topic string by '/'
$splitTopic = $topic -split '/'

# Extract resource group and storage account names
$resourceGroupNameIndex = [array]::IndexOf($splitTopic, "resourceGroups") + 1
$storageAccountNameIndex = [array]::IndexOf($splitTopic, "storageAccounts") + 1

$resourceGroupName = $splitTopic[$resourceGroupNameIndex]
$storageAccountName = $splitTopic[$storageAccountNameIndex]

Write-Output "Resource Group Name: $resourceGroupName"
Write-Output "Storage Account Name: $storageAccountName"

# Connect to Azure with system-assigned managed identity
$AzureLogin = (Connect-AzAccount -Identity).Context

# Set context
Set-AzContext -Subscription $AzureLogin.Subscription -DefaultProfile $AzureLogin

# Get the storage account key
$storageAccountKey = (Get-AzStorageAccountKey -ResourceGroupName $resourceGroupName -Name $storageAccountName)[0].Value

# Create the storage context
$context = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccountKey

# Download file from blob storage
$localPath = $env:TEMP
$blobToDownload = $blobFileName
$blobDestination = Join-Path $localPath $blobFolderPath
Write-Output "BlobDestination: $blobDestination"
$blobDestinationFilePath = Join-Path $blobDestination $blobToDownload
Write-Output "BlobDestinationPath: $blobDestinationFilePath"

if (-not (Test-Path -Path $blobDestination)) {
    New-Item -ItemType Directory -Path $blobDestination | Out-Null
}

# Download blob content
try {
    Get-AzStorageBlobContent -Blob $blobToDownload -Container $containerName -Context $context -Destination $blobDestinationFilePath
    Write-Output "Downloaded '$blobToDownload' to '$blobDestinationFilePath'"

    # Check if the file is a .zip file
    if ($blobFileName -like "*.zip") {
        # Define the destination directory for the extracted contents
        $extractDestination = Join-Path $blobDestination $blobFileName.TrimEnd(".zip")

        # Create the directory if it doesn't exist
        if (-not (Test-Path -Path $extractDestination)) {
            New-Item -ItemType Directory -Path $extractDestination | Out-Null
        }

        # Unpack the zip file contents
        Expand-Archive -Path $blobDestinationFilePath -DestinationPath $extractDestination -Force
        Write-Output "Unpacked '$blobToDownload' to '$extractDestination'"
    }
} catch {
    Write-Output "Failed to download or unpack: '$blobToDownload'. Error: $_"
    exit
}

# Download latest Microsoft Win32 Content Prep Tool
$latestReleaseUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/releases/latest"
$response = Invoke-WebRequest -Uri $latestReleaseUrl -UseBasicParsing
$urlWithTag = $response.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
$tagName = $urlWithTag -split '/' | Select-Object -Last 1
$downloadUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/archive/refs/tags/$tagName.zip"
Write-Output "Following version of Win32 Prep Tool will be downloaded: $downloadUrl"

$tempZipPath = Join-Path $env:TEMP ($tagName + ".zip")

try {
    Invoke-WebRequest -Uri $downloadUrl -OutFile $tempZipPath -UseBasicParsing
    Expand-Archive -Path $tempZipPath -DestinationPath $localPath -Force
} catch {
    Write-Output "Failed to download or expand archive $tagName.zip : $_"
    exit
}

# Locate IntuneWinAppUtil.exe
$prepToolPath = Get-ChildItem -Path $localPath -Filter "IntuneWinAppUtil.exe" -Recurse | Select-Object -First 1 -ExpandProperty FullName

# Define the output folder for the .intunewin file
$outputFolder = Join-Path $localPath $containerNameDestination

# Ensure output folder exists
if (-not (Test-Path -Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
    Write-Output "Output folder created: $outputFolder."
}

# Check if the downloaded file was a zip
if ($blobFileName -like "*.zip") {
    $extractDestination = Join-Path $blobDestination $blobFileName.TrimEnd(".zip")
    $blobDestination = $extractDestination
    Get-ChildItem -Path $blobDestination -Recurse

    # Search for the first .exe or .msi file in the extraction directory
    $extractedFile = Get-ChildItem -Path $extractDestination -Filter "*.exe" -Recurse | Select-Object -First 1
    if (-not $extractedFile) {
        $extractedFile = Get-ChildItem -Path $extractDestination -Filter "*.msi" -Recurse | Select-Object -First 1
    }

    # Update $blobToDownload if a file is found
    if ($extractedFile) {
        $blobToDownload = $extractedFile.Name
    } else {
        Write-Output "No .exe or .msi file found in the extracted contents."
        exit
    }
}

# Construct the command line arguments
$arguments = "-c `"$blobDestination`" -s `"$blobToDownload`" -o `"$outputFolder`" -q"

Write-Output "Prep Tool Arguments: $arguments"



# Execute the Win32 Content Prep Tool
Write-output "Running Win32 Prep Tool:"
try {
    Start-Process -FilePath "$prepToolPath" -ArgumentList $arguments -NoNewWindow -Wait
    $intunewinFile = Get-ChildItem -Path $outputFolder -Filter "*.intunewin" | Select-Object -First 1
    if ($null -ne $intunewinFile) {
        Write-Output "Successfully created .intunewin file: $($intunewinFile.FullName)"
        
        # Define the blob name for the upload
        $blobName = $intunewinFile.Name
        
        # Upload the .intunewin file
        try {
            $blobUploadResult = Set-AzStorageBlobContent -File $intunewinFile.FullName -Container $containerNameDestination -Blob $blobName -Context $context -Force
            Write-Output "Successfully uploaded .intunewin file to blob storage: $($blobUploadResult.Name)"
        } catch {
            Write-Output "Failed to upload .intunewin file to blob storage: $_"
            exit
        }
    } else {
        throw "Failed to create .intunewin file."
    }
} catch {
    Write-Output "Error during packaging: $_"
    exit
}
