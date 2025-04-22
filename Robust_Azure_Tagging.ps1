<#
.SYNOPSIS
    Azure Tagging Script for Normalizing and Updating Resource Tags.

.DESCRIPTION
    This script processes Azure resources across multiple subscriptions and resource groups to normalize and update their tags.
    It supports dry run mode for testing changes without applying them and can log results either locally or to an Azure Storage account.
    The script allows you to define tag variations and their normalized forms to ensure consistent tagging across resources.

.PARAMETERS
    $DryRun
        Enables dry run mode. When set to $true, the script simulates tagging operations without making any changes.
        Default: $true

    $OutputDirectory
        Optional. Specifies a custom directory for the output CSV file. The script will generate the filename automatically.
        Default: The user's home directory or "Documents" folder.

    $UseAzureStorage
        Enables logging to an Azure Storage account. When set to $true, logs are uploaded to the specified storage account.
        Default: $false

    $targetSubscriptionIds
        An array of Azure subscription IDs where the script will process resources.

    $targetResourceGroupNames
        An array of resource group names to target. Set to $null to target all resource groups in the specified subscriptions.

.NOTES
    Script Version: 1.1.0
    Authors: Dan Wittenberg and Bryan Wilson
    Date: March 25, 2025

.EXAMPLE
    # Run the script in dry run mode (default behavior)
    .\Robust_Azure_Tagging.ps1

    # Run the script and apply changes
    $DryRun = $false
    .\Robust_Azure_Tagging.ps1

    # Specify a custom output directory
    $OutputDirectory = "C:\Logs"
    .\Robust_Azure_Tagging.ps1

    # Enable Azure Storage logging
    $UseAzureStorage = $true
    .\Robust_Azure_Tagging.ps1

    # Target specific resource groups
    $targetResourceGroupNames = @("ResourceGroup1", "ResourceGroup2")
    .\Robust_Azure_Tagging.ps1

.HOW TO SET TAG VARIATIONS
    The script uses a `$tagVariations` array to define tag variations and their normalized forms. Each entry in the array is a hashtable
    with two keys:
        - `Variations`: An array of possible variations of a tag name (e.g., different spellings).  The script will atuomatically process all variations of capitalization.
        - `NormalizedTagKey`: The standardized form of the tag name.

    Example:
        $tagVariations = @(
            @{ Variations = @("env", "environmentid", "enviornment", "environemnt", "x_environment"); NormalizedTagKey = "environment" },
            @{ Variations = @("Cost Center", "costcenter", "cost center"); NormalizedTagKey = "cost-center" },
            @{ Variations = @("CreatedBy"); NormalizedTagKey = "created-by" }
        )

    In this example:
        - Tags like "env", "environmentid", or "environemnt" will be normalized to "environment".
        - Tags like "Cost Center" or "costcenter" will be normalized to "cost-center".
        - Tags like "CreatedBy" will be normalized to "created-by".

    To add more tag variations, simply append new hashtables to the `$tagVariations` array.


.KNOWN ISSUES
    - **Azure Resource Manager Throttling**: If processing a large number of resources, the script may encounter Azure API throttling. 
      To mitigate this, consider adding delays between API calls or reducing the number of subscriptions processed in a single run.
    
    - This script cannot update resources of type "Microsoft.CognitiveServices/accounts" due to Azure API limitations. 
      If you need to update tags for these resources, consider using the Azure portal.
        TODO: Add logic to use "Set-AzCognitiveServicesAccount" for updating tags on Cognitive Services accounts.
        
.IMPORTANT NOTE
    If multiple tags exist that are normalized to the same key but have different values, only one of the values will be assigned to the normalized tag key. 
    This means that:
    - The value of one variation could overwrite the value of another variation.
    - A correctly normalized tag could be overwritten by the value of a variation.

    To avoid unintended overwrites, carefully review the tag variations defined in the `$tagVariations` array 
    and ensure that conflicting values for the same normalized key are not present.
    Recommend performing a DryRun and reviewing the output CSV file before applying changes to resources.
#>

# Script version
$scriptVersion = "1.1.0"

Write-Output "Running Azure Tag Update Script - Version $scriptVersion"

# Enable or disable dry run mode
$DryRun = $true # Set to $false to apply changes to resources

# Optional: Specify a custom output directory
# Set to $null to use the user's home directory
$OutputDirectory = $null  # Set to a custom directory, e.g., "C:\Logs"

# Define an array of target subscription IDs
# These are the subscriptions where the script will process resources
$targetSubscriptionIds = @(
    "f2ee03c4-c49c-4ca3-9c7a-48ca004e0c5d"
)

# Specify target resource group names
# Set to $null to target all resource groups in the subscriptions
$targetResourceGroupNames = $null  # Example: @('ResourceGroup1', 'ResourceGroup2')


######################################################################
# UNDER DEVELOPMENT - DO NOT ENABLE THIS FEATURE
# Azure Storage Configuration
# Configure whether to use Azure Storage for logging processed resources
######################################################################
$UseAzureStorage = $false  # Set to $true to use Azure Storage, $false to log locally
$storageAccountName = "taggingoutput"  # Name of the Azure Storage account
$storageContainerName = "azuretag"  # Name of the container in the storage account
$storageAccountSubscriptionId = "f2ee03c4-c49c-4ca3-9c7a-48ca004e0c5d" # Subscription ID for the storage account
$resourceGroupForStorage = "StorageAccountTest1"  # Resource group of the storage account
######################################################################



# Obtain date to include in output filenames
$currentDate = Get-Date -Format "yyyyMMdd_HHmmss"

# Define tag variations and normalized tags
# This maps variations of tag names to a normalized key
$tagVariations = @(
    @{ Variations = @("env", "environmentid", "enviornment", "environemnt", "x_environment"); NormalizedTagKey = "environment" },
    @{ Variations = @("Cost Center", "costcenter", "cost center"); NormalizedTagKey = "cost-center" },
    @{ Variations = @("cost center number"); NormalizedTagKey = "cost-center-number" },
    @{ Variations = @("cost center number billing"); NormalizedTagKey = "cost-center-number-billing" },
    @{ Variations = @("CreatedBy"); NormalizedTagKey = "created-by" },
    @{ Variations = @("Dept"); NormalizedTagKey = "department" },
    @{ Variations = @("decription", "Description", "descripton"); NormalizedTagKey = "description" }
)



# Function to determine the output CSV file path
function Get-OutputCsvFilePath {
    param (
        [string]$CustomDirectory
    )
    # Generate the filename
    $fileName = "output_TagUpdateResults_${scriptVersion}_${currentDate}.csv"

    if ($CustomDirectory) {
        # Ensure the directory ends with a backslash
        if (-not ($CustomDirectory -like "*\")) {
            $CustomDirectory += "\"
        }

        # Combine the directory and filename
        return "$CustomDirectory$fileName"
    } else {
        # Default logic for determining the output file path
        if (Test-Path "$HOME") {
            return "$HOME/$fileName"
        } else {
            return "C:\Users\$env:USERNAME\Documents\$fileName"
        }
    }
}

# Set the output CSV file path
$outputCsvFilePath = Get-OutputCsvFilePath -CustomDirectory $OutputDirectory
Write-Output "Output CSV file path: $outputCsvFilePath"

# Validate the output file path
function Validate-OutputFilePath {
    param (
        [string]$FilePath
    )

    try {
        # Check if the directory exists
        $directory = Split-Path -Path $FilePath -Parent
        if (-not (Test-Path $directory)) {
            throw "The directory '$directory' does not exist."
        }

        # Test if the file is writable
        if (Test-Path $FilePath) {
            # Try to open the file for writing
            $stream = [System.IO.File]::OpenWrite($FilePath)
            $stream.Close()
        } else {
            # Try to create the file
            @() | Export-Csv -Path $FilePath -NoTypeInformation
        }
    } catch {
        Write-Output "Error: The output file path '$FilePath' is not valid or writable. $_"
        exit 1
    }
}

# Validate the output CSV file path
Validate-OutputFilePath -FilePath $outputCsvFilePath
Write-Output "Validated output file path: $outputCsvFilePath"

# Initialize Azure Storage context if enabled
if ($UseAzureStorage) {
 
    $processedResourcesFilePath = "$env:TEMP\processed_resources_${currentDate}.txt"
    $processedResourceGroupsFilePath = "$env:TEMP\processed_resource_groups_${currentDate}.txt"


    write-output "Processed resources file path: $processedResourcesFilePath"
    write-output "Processed resource groups file path: $processedResourceGroupsFilePath"

    # Initialize storage context
    try {
        Write-Output "Switching context to Storage Account subscription: $storageAccountSubscriptionId"
        Set-AzContext -SubscriptionId $storageAccountSubscriptionId -ErrorAction Stop

        $storageAccountKey = (Get-AzStorageAccountKey -ResourceGroupName $resourceGroupForStorage -Name $storageAccountName).Value[0]
        $storageContext = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccountKey
        Write-Output "Initialized storage context for account: $storageAccountName"
        Write-Output "Storage context details: $storageContext"
    }
    catch {
        Write-Output "Failed to initialize storage account context. Ensure the subscription and permissions are configured correctly."
        throw $_
    }
}

# Function to check or create blobs in Azure Storage
function Ensure-BlobExists {
    param (
        [string]$blobName
    )
    try {
        $blob = Get-AzStorageBlob -Container $storageContainerName -Blob $blobName -Context $storageContext -ErrorAction SilentlyContinue
        if (-not $blob) {
            Write-Output "Blob $blobName does not exist. Attempting to create it."
            $tempFile = New-TemporaryFile
            Set-Content -Path $tempFile -Value "Initialized blob: $blobName"
            Set-AzStorageBlobContent -File $tempFile -Container $storageContainerName -Blob $blobName -Context $storageContext -Force
            Remove-Item -Path $tempFile -Force
            Write-Output "Blob $blobName successfully created."
        }
        else {
            Write-Output "Blob $blobName already exists."
        }
    }
    catch {
        Write-Output "Failed to ensure blob $blobName exists in Azure Storage. Error: $_"
        throw $_
    }
}

# Function to download existing processed files from Azure Storage
function Download-FromStorage {
    param (
        [string]$blobName,
        [string]$localFilePath
    )
    try {
        Get-AzStorageBlobContent -Container $storageContainerName -Blob $blobName -Destination $localFilePath -Context $storageContext -Force -ErrorAction SilentlyContinue
        Write-Output "Downloaded existing file from Azure Storage: $blobName"
    }
    catch {
        Write-Output "File not found in Azure Storage for: $blobName. Creating a new local file."
        New-Item -Path $localFilePath -ItemType File | Out-Null
    }
}

# Function to append data to Azure Storage
function Append-ToStorage {
    param (
        [string]$filePath,
        [string]$blobName
    )
    try {
        Set-AzStorageBlobContent -File $filePath -Container $storageContainerName -Blob $blobName -Context $storageContext -Force
        Write-Output "Successfully uploaded or appended to Azure Storage: $blobName"
    }
    catch {
        Write-Output "Failed to append to Azure Storage. Error: $_"
        throw $_
    }
}

# Function to display a text progress bar
function Show-ProgressBar {
    param (
        [int]$Current,
        [int]$Total,
        [string]$TaskDescription
    )
    $progress = [math]::Floor(($Current / $Total) * 100)
    Write-Progress -Activity $TaskDescription -Status "$progress% Complete" -PercentComplete $progress
}

# Helper function to log results
function Log-Result {
    param (
        [string]$subId,
        [string]$resourceGroupName,
        [string]$resourceId,
        [hashtable]$oldTags,
        [hashtable]$newTags,
        [string]$addStatus,
        [string]$removeStatus,
        [string]$outputError,
        [string]$httpResponse
    )

    $oldTagsJson = $oldTags | ConvertTo-Json -Compress
    $newTagsJson = $newTags | ConvertTo-Json -Compress

    $logEntry = [PSCustomObject]@{
        SubscriptionId    = $subId
        ResourceGroupName = $resourceGroupName
        ResourceId        = $resourceId
        OldTags           = $oldTagsJson
        NewTags           = $newTagsJson
        AddStatus         = $addStatus
        RemoveStatus      = $removeStatus
        Error             = $outputError
        HttpResponse      = $httpResponse
    }

    $logEntry | Export-Csv -Path $outputCsvFilePath -NoTypeInformation -Append
    Write-Output "Logged result for resource: $resourceId, addStatus: $addStatus, removeStatus: $removeStatus"
}

if ($UseAzureStorage) {

    # Ensure processed resources and resource groups files exist in Azure Storage
    try {
        Ensure-BlobExists -blobName "processed_resources_${currentDate}_sub.txt"
        Ensure-BlobExists -blobName "processed_resource_groups_${currentDate}_sub.txt"
    }
    catch {
        Write-Output "Critical failure: Unable to create required blobs in Azure Storage. Aborting script."
        exit 1
    }

    # Check and initialize local processed files
    Download-FromStorage -blobName "processed_resources_${currentDate}_sub.txt" -localFilePath $processedResourcesFilePath
    Download-FromStorage -blobName "processed_resource_groups_${currentDate}_sub.txt" -localFilePath $processedResourceGroupsFilePath

    # Load processed resources and resource groups into memory
    $processedResources = if (Test-Path $processedResourcesFilePath) {
        Get-Content $processedResourcesFilePath
    }
    else {
        @()
    }

    $processedResourceGroups = if (Test-Path $processedResourceGroupsFilePath) {
        Get-Content $processedResourceGroupsFilePath
    }
    else {
        @()
    }

}

foreach ($tagVar in $tagVariations) {
    # Check to the correctly cased normaliezed key is not present in the variations list.
    $tagVar.Variations = $tagVar.Variations | Where-Object { $_ -ne $tagVar.NormalizedTagKey }
    # Convert all variations to lowercase and remove duplicates
    $tagVar.Variations = $tagVar.Variations | ForEach-Object { $_.ToLower() } | Select-Object -Unique 
}

try {
    foreach ($subId in $targetSubscriptionIds) {
        Write-Output "Processing subscription: $subId"
        Set-AzContext -SubscriptionId $subId -ErrorAction Stop

        $resourceGroups = if ($null -eq $targetResourceGroupNames) {
            Get-AzResourceGroup -ErrorAction SilentlyContinue
        }
        else {
            $targetResourceGroupNames | ForEach-Object { Get-AzResourceGroup -Name $_ -ErrorAction SilentlyContinue }
        }

        foreach ($resourceGroup in $resourceGroups) {
            Write-Output ""
            Write-Output "============================================================================================="
            Write-Output "======== Processing resources in resource group: $($resourceGroup.ResourceGroupName) ========"
            Write-Output "============================================================================================="

            $resources = Get-AzResource -ResourceGroupName $resourceGroup.ResourceGroupName -ErrorAction SilentlyContinue
            $resourceCount = $resources.Count
            $currentResourceIndex = 0

            foreach ($resource in $resources) {
                $currentResourceIndex++
                Show-ProgressBar -Current $currentResourceIndex -Total $resourceCount -TaskDescription "Processing Resources"

                Write-Output ""
                Write-Output "------------- Processing resource: $($resource.Name)"

                try {
                    $oldTags = @{}
                    if ($resource.Tags) {
                        $resource.Tags.GetEnumerator() | ForEach-Object { $oldTags[$_.Key] = $_.Value }
                    }
                    # Initialize new variables
                    $tagsToRemove = @{}
                    $normalizedTags = @{}
                    $finalNormalizedTags = @{}

                    # Update tags based on defined variations
                    foreach ($tagVar in $tagVariations) {
                        $normalizedTagKey = $tagVar.NormalizedTagKey
                        foreach ($variation in $tagVar.Variations) {
                            foreach ($key in $oldTags.Keys) {
                                if ($key -ieq $variation) {
                                    # Case-insensitive comparison
                                    write-output "Matched: Key: $key | Normalize to: $normalizedTagKey"
                                    $tagsToRemove[$key] = $oldTags[$key]
                                    $normalizedTags[$normalizedTagKey] = $oldTags[$key]
                                }
                            }
                        }
                        # Check for tags that are spelled correctly but have the wrong capitalization
                        foreach ($key in $oldTags.Keys) {
                            if ($key -ieq $normalizedTagKey -and $key -cne $normalizedTagKey) {
                                write-output "Matched: Key: $key | Normalize to: $normalizedTagKey"
                                $tagsToRemove[$key] = $oldTags[$key]
                                $normalizedTags[$normalizedTagKey] = $oldTags[$key]
                            }
                        }
                    }

                    # If a Normalized version of a tag already exists, do not atttempt to add it again
                    # Create a new hashtable for final normalized tags
                    foreach ($key in $normalizedTags.Keys) {
                        if (-not $oldTags.ContainsKey($key) -or ($oldTags.ContainsKey($key) -and $oldTags.GetEnumerator() | Where-Object { $_.Key -ceq $key } -eq $null)) {
                            $finalNormalizedTags[$key] = $normalizedTags[$key]
                        }
                    }
                    $addStatus = "NoChange"
                    $removeStatus = "NoChange"
                    # Apply the new tags
                    if ($finalNormalizedTags.Count -gt 0) {
                        if ($DryRun) {
                            $removeResponse = Update-AzTag -ResourceId $resource.ResourceId -Tag $tagsToRemove -Operation Delete -ErrorAction Stop -WhatIf  
                            $removeStatus = "DryRun"  
                            $errormsg = ""
                        }
                        else {
                            $removeResponse = Update-AzTag -ResourceId $resource.ResourceId -Tag $tagsToRemove -Operation Delete -ErrorAction Stop

                            $updatedTags = (Get-AzResource -ResourceId $resource.ResourceId).Tags
                            $remainingOldKeys = $tagsToRemove.Keys | Where-Object { $updatedTags.ContainsKey($_) }
                            if ($remainingOldKeys.Count -gt 0) {
                                $removeStatus = "Failed"
                                $errormsg += "Old tags still present.  Remaining tags: $($remainingOldKeys -join ', ')  "
                            }
                            else {
                                $removeStatus = "Success"
                                $errormsg = ""
                            }
                        }
                    
                        if ($DryRun) {
                            $addResponse = Update-AzTag -ResourceId $resource.ResourceId -Tag $finalNormalizedTags -Operation Merge -ErrorAction Stop -WhatIf
                            $addStatus = "DryRun"
                            $errormsg = ""
                        }
                        else {
                            $addResponse = Update-AzTag -ResourceId $resource.ResourceId -Tag $finalNormalizedTags -Operation Merge -ErrorAction Stop

                            $updatedTags = (Get-AzResource -ResourceId $resource.ResourceId).Tags
                            $missingNewKeys = $finalNormalizedTags.Keys | Where-Object { -not $updatedTags.ContainsKey($_) }
                            if ($missingNewKeys.Count -gt 0) {
                                $addStatus = "Failed"
                                $errormsg = "New tags not added successfully. Missing tags: $($missingNewKeys -join ', ')  "
                            }
                            else {
                                $addStatus = "Success"
                                $errormsg += ""
                            }
                        }
                        Log-Result -subId $subId -resourceGroupName $resourceGroup.ResourceGroupName -resourceId $resource.ResourceId -oldTags $tagsToRemove -newTags $finalNormalizedTags -addStatus $addStatus -removeStatus $removeStatus -outputError $errormsg -httpResponse "200"
                    }
                }
                catch {
                    Write-Output "Error processing resource: $($resource.Name). Error: $_"
                    if ($removeStatus -eq "NoChange" -or $removeStatus -eq "Success") {
                        Log-Result -subId $subId -resourceGroupName $resourceGroup.ResourceGroupName -resourceId $resource.ResourceId -oldTags $tagsToRemove -newTags $null -addStatus $addStatus -removeStatus $removeStatus -outputError $_ -httpResponse "N/A"
                    }
                    else {
                        Log-Result -subId $subId -resourceGroupName $resourceGroup.ResourceGroupName -resourceId $resource.ResourceId -oldTags $null -newTags $null -removeStatus "Error" -addStatus "Error" -outputError $_ -httpResponse "N/A"
                    }
                    continue
                }
            }
        }
    }
}
finally {
    Write-Output "Processing completed for all subscriptions."
    write-output "Output CSV file path: $outputCsvFilePath"
}
