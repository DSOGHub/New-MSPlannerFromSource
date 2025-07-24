<#
.SYNOPSIS
    Creates a Microsoft Planner plan by copying structure and tasks from an existing source plan.

.DESCRIPTION
    This script copies tasks, buckets, and structure from a source Microsoft Planner plan to create a new plan.
    It preserves task details including descriptions, checklists, priorities, and bucket organization.
    Authentication is meant to be handled via ClientSecretCredentials, currently utilizing the IO Collab application, 
    as it requires tasks.ReadWrite.All, tasks.ReadWrite.Shared, Group.Read.All, and User.Read.All permissions.

.PARAMETER SourcePlanId
    The ID of the source Planner plan to copy from
.PARAMETER UserId
    The objectId of the user who will owner the new plan
.PARAMETER PlanTitle
    Title for the new Planner plan

.EXAMPLE
    .\New-MSPlannerPlanFromSource.ps1 -SourcePlanId "abc123..." -UserId "def456..." -PlanTitle "My New Plan"

.NOTES
    Author: David Sorenson
    Version: 1.0
    Requires: Microsoft.Graph.Planner PowerShell module
.TODO:
    Fix task order. Bucket order works fine, but not task order. 
    Consider removing URL decode/encode now that I've figured out the original issue with double encoding on original test case - not sure there's any point now.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePlanId,
    [Parameter(Mandatory = $true)]
    [string]$UserId,
    [Parameter(Mandatory = $true)]
    [string]$PlanTitle
)

Add-Type -AssemblyName System.Web

function Get-FirstDifferencePosition {
    param(
        [Parameter(Mandatory)]
        [string[]]$OrderHints
    )
    
    if ($OrderHints.Count -eq 0) { return $null }
    if ($OrderHints.Count -eq 1) { return 0 }
    
    $minLength = ($OrderHints | Measure-Object -Property Length -Minimum).Minimum
    
    for ($i = 0; $i -lt $minLength; $i++) {
        $firstChar = $OrderHints[0][$i]
        for ($j = 1; $j -lt $OrderHints.Count; $j++) {
            if ($OrderHints[$j][$i] -ne $firstChar) {
                return $i
            }
        }
    }
    
    return $minLength
}

function Get-OrdinalValuesAtPosition {
    param(
        [Parameter(Mandatory)]
        [string[]]$OrderHints,
        [Parameter(Mandatory)]
        [int]$Position
    )
    
    $ordinalData = @()
    
    foreach ($hint in $OrderHints) {
        if ($Position -lt $hint.Length) {
            $char = $hint[$Position]
            $ordinal = [int][char]$char
            $ordinalData += [PSCustomObject]@{
                String = $hint
                Character = $char
                OrdinalValue = $ordinal
            }
        } else {
            $ordinalData += [PSCustomObject]@{
                String = $hint
                Character = "(end of string)"
                OrdinalValue = -1
            }
        }
    }
    
    return $ordinalData | Sort-Object -Property OrdinalValue -Descending
}

# Main script execution
Write-Information "Starting Planner copy process" -InformationAction Continue
Write-Information "Source Plan ID: $SourcePlanId" -InformationAction Continue
Write-Information "Plan Title: $PlanTitle" -InformationAction Continue
Write-Information "User ID: $UserId" -InformationAction Continue

# Verify Microsoft Graph connection
try {
    $context = Get-MgContext
    if (-not $context) {
        Write-Error "No existing Graph context found. Ensure authentication is handled externally." -ErrorAction Stop
    }
    Write-Information "Microsoft Graph context verified" -InformationAction Continue
}
catch {
    Write-Error "Failed to verify Microsoft Graph connection: $_" -ErrorAction Stop
}

# Verify source plan exists
try {
    $sourcePlan = Get-MgPlannerPlan -PlannerPlanId $SourcePlanId
    Write-Information "Source plan verified: $($sourcePlan.Title)" -InformationAction Continue
}
catch {
    Write-Error "Failed to access source plan: $_" -ErrorAction Stop
}

# Get source tasks and details
try {
    $sourceTasks = Get-MgPlannerPlanTask -PlannerPlanId $SourcePlanId
    Write-Information "Retrieved $($sourceTasks.Count) tasks from source plan" -InformationAction Continue
}
catch {
    Write-Error "Failed to retrieve source tasks: $_" -ErrorAction Stop
}

$sourceTaskData = [System.Collections.Generic.List[object]]::new()

foreach ($task in $sourceTasks) {
    try {
        $taskDetails = Get-MgPlannerTaskDetail -PlannerTaskId $task.Id

        $checklistItems = @()
        if ($taskDetails.Checklist -and $taskDetails.Checklist.AdditionalProperties) {
            foreach ($checklistItemId in $taskDetails.Checklist.AdditionalProperties.Keys) {
                $checklistItem = $taskDetails.Checklist.AdditionalProperties[$checklistItemId]
                $checklistItems += [PSCustomObject]@{
                    Title = $checklistItem.title
                    IsChecked = $checklistItem.isChecked
                }
            }
        }

        # Get task attachments from task details
        $attachments = @()
        if ($taskDetails.References -and $taskDetails.References.AdditionalProperties) {
            foreach ($referenceId in $taskDetails.References.AdditionalProperties.Keys) {
                $reference = $taskDetails.References.AdditionalProperties[$referenceId]
                
                # The referenceId IS the URL (often encoded), reference object contains metadata
                $originalUrl = $referenceId
                
                # Decode URL - handle multiple layers of encoding
                $decodedUrl = $null
                if (-not [string]::IsNullOrEmpty($originalUrl)) {
                    try {
                        $decodedUrl = $originalUrl
                        $previousUrl = ""
                        
                        # Decode multiple times until no more changes occur
                        while ($decodedUrl -ne $previousUrl) {
                            $previousUrl = $decodedUrl
                            $decodedUrl = [System.Web.HttpUtility]::UrlDecode($decodedUrl)
                        }
                        
                        # Validate the final decoded URL
                        if ([string]::IsNullOrEmpty($decodedUrl)) {
                            Write-Warning "URL became null or empty after decoding for attachment '$($reference.alias)' in task '$($task.Title)'"
                            $decodedUrl = $originalUrl  # Use original if decode results in empty
                        }
                    }
                    catch {
                        Write-Warning "Failed to decode URL for attachment '$($reference.alias)' in task '$($task.Title)': $originalUrl"
                        $decodedUrl = $originalUrl  # Use original if decode fails
                    }
                }
                
                $attachments += [PSCustomObject]@{
                    Id = $referenceId  # This is actually the URL
                    Alias = $reference.alias
                    Type = $reference.type
                    PreviewType = $reference.previewType
                    LastModifiedDateTime = $reference.lastModifiedDateTime
                    Url = $decodedUrl
                    OriginalUrl = $originalUrl  # The encoded URL from the key
                }
            }
        }

        $sourceTaskData.Add([PSCustomObject]@{
            TaskId = $task.Id
            Title = $task.Title
            Description = $taskDetails.Description
            Priority = $task.Priority
            OrderHint = $task.OrderHint
            BucketId = $task.BucketId
            PreviewType = $taskDetails.PreviewType
            ChecklistItems = $checklistItems
            Attachments = $attachments
        })
    }
    catch {
        Write-Warning "Failed to get details for task '$($task.Title)': $_"
    }
}

# Get and sort buckets
[Array][System.Collections.Generic.HashSet[string]]$sourceBucketIds = $sourceTasks.BucketId
$sourceBuckets = [System.Collections.Generic.List[object]]::new()

foreach ($bucketId in $sourceBucketIds) {
    try {
        $bucket = Get-MgPlannerBucket -PlannerBucketId $bucketId
        if ($bucket.PSObject.Properties['OrderHint']) {
            $bucket | Add-Member -MemberType NoteProperty -Name BucketOrderHint -Value $bucket.OrderHint -Force
        }
        $sourceBuckets.Add($bucket)
    }
    catch {
        Write-Warning "Failed to get bucket details for ID '$bucketId': $_"
    }
}

# Sort buckets using OrderHint analysis
if ($sourceBuckets.Count -gt 1) {
    $orderHints = $sourceBuckets | ForEach-Object { $_.OrderHint }
    $firstDifferencePosition = Get-FirstDifferencePosition -OrderHints $orderHints
    $sortedOrdinals = Get-OrdinalValuesAtPosition -OrderHints $orderHints -Position $firstDifferencePosition

    $sortOrder = @{}
    for ($i = 0; $i -lt $sortedOrdinals.Count; $i++) {
        $sortOrder[$sortedOrdinals[$i].String] = $i
    }

    $sourceBuckets = $sourceBuckets | Sort-Object { $sortOrder[$_.OrderHint] }
}

Write-Information "Retrieved $($sourceBuckets.Count) buckets from source plan" -InformationAction Continue

# Create the new plan
try {
    $newPlan = New-MgPlannerPlan -BodyParameter @{
        container = @{
            url = "https://graph.microsoft.com/v1.0/users/$UserId"
        }
        title = $PlanTitle
    }
    
    $newPlanId = $newPlan.id
    Write-Information "Created new plan: $PlanTitle (ID: $newPlanId)" -InformationAction Continue
}
catch {
    Write-Error "Failed to create new plan: $_" -ErrorAction Stop
}

# Create buckets in the new plan
$bucketMapping = @{}

foreach ($bucket in $sourceBuckets) {
    try {
        $newBucket = New-MgPlannerBucket -Name $bucket.Name -PlanId $newPlanId
        $bucketMapping[$bucket.Id] = $newBucket.Id
        Write-Information "Created bucket: $($bucket.Name)" -InformationAction Continue
    }
    catch {
        Write-Error "Failed to create bucket '$($bucket.Name)': $_" -ErrorAction Stop
    }
}

# Create tasks
$taskCounter = 0
$successCount = 0
$tasksByBucket = $sourceTaskData | Group-Object -Property BucketId

foreach ($bucketGroup in $tasksByBucket) {
    $bucketTasks = $bucketGroup.Group
    
    # Sort tasks within bucket using OrderHint analysis
    if ($bucketTasks.Count -gt 1) {
        $taskOrderHints = $bucketTasks | ForEach-Object { $_.OrderHint }
        $taskFirstDifferencePosition = Get-FirstDifferencePosition -OrderHints $taskOrderHints
        $taskSortedOrdinals = Get-OrdinalValuesAtPosition -OrderHints $taskOrderHints -Position $taskFirstDifferencePosition
        
        $taskSortOrder = @{}
        for ($i = 0; $i -lt $taskSortedOrdinals.Count; $i++) {
            $taskSortOrder[$taskSortedOrdinals[$i].String] = $i
        }
        
        $bucketTasks = $bucketTasks | Sort-Object { $taskSortOrder[$_.OrderHint] }
    }
    
    foreach ($task in $bucketTasks) {
        $taskCounter++
        
        try {
            $taskParams = @{
                title = $task.Title
                bucketId = $bucketMapping[$task.BucketId]
                planId = $newPlanId
            }
            
            if ($task.Priority) {
                $taskParams.priority = $task.Priority
            }
            
            $newTask = New-MgPlannerTask -BodyParameter $taskParams
            
            # Add task details
            $needsDetailUpdate = $false
            $detailParams = @{}
            
            if ($task.Description -and $task.Description.Trim() -ne '') {
                $detailParams.description = $task.Description.Trim()
                $needsDetailUpdate = $true
            }
            
            if ($task.ChecklistItems.Count -gt 0) {
                $checklistItems = @{}
                foreach ($checklistItem in $task.ChecklistItems) {
                    $checklistId = [guid]::NewGuid().ToString()
                    $checklistItems[$checklistId] = @{
                        '@odata.type' = "#microsoft.graph.plannerChecklistItem"
                        title = $checklistItem.Title
                        isChecked = $false
                    }
                }
                $detailParams.checklist = $checklistItems
                $detailParams.previewType = "checklist"
                $needsDetailUpdate = $true
            } elseif ($task.Description) {
                $detailParams.previewType = "description"
                $needsDetailUpdate = $true
            }
            
            if ($needsDetailUpdate) {
                Start-Sleep -Milliseconds 500
                
                $taskForUpdate = Get-MgPlannerTaskDetail -PlannerTaskId $newTask.Id
                $etag = $taskForUpdate.AdditionalProperties.'@odata.etag'
                
                if ($etag) {
                    try {
                        Update-MgPlannerTaskDetail -PlannerTaskId $newTask.Id -IfMatch $etag -BodyParameter $detailParams
                    }
                    catch {
                        Write-Warning "Failed to update task details for '$($task.Title)': $_"
                    }
                }
            }
            
            # Copy attachments (references)
            if ($task.Attachments.Count -gt 0) {
                Start-Sleep -Milliseconds 500
                
                $taskForUpdate = Get-MgPlannerTaskDetail -PlannerTaskId $newTask.Id
                $etag = $taskForUpdate.AdditionalProperties.'@odata.etag'
                
                if ($etag) {
                    $references = @{}
                    foreach ($attachment in $task.Attachments) {
                        # Validate the decoded URL before adding it to references
                        if ([string]::IsNullOrEmpty($attachment.Url)) {
                            Write-Warning "Skipping attachment '$($attachment.Alias)' for task '$($task.Title)' - URL is null or empty"
                            continue
                        }
                        
                        # Try to parse the decoded URL to ensure it's valid
                        try {
                            $uri = [System.Uri]::new($attachment.Url)
                            if (-not $uri.IsAbsoluteUri) {
                                Write-Warning "Skipping attachment '$($attachment.Alias)' for task '$($task.Title)' - URL is not absolute: $($attachment.Url)"
                                continue
                            }
                        }
                        catch {
                            Write-Warning "Skipping attachment '$($attachment.Alias)' for task '$($task.Title)' - Invalid URL format: $($attachment.Url). Original: $($attachment.OriginalUrl)"
                            continue
                        }
                        
                        # URL-encode the URL for use as the key
                        $encodedUrl = [System.Web.HttpUtility]::UrlEncode($attachment.Url)
                        
                        $references[$encodedUrl] = @{
                            '@odata.type' = "microsoft.graph.plannerExternalReference"
                            alias = $attachment.Alias
                            type = $attachment.Type
                            previewPriority = "!" # Default priority, can be customized if needed
                        }
                    }
                    
                    # Only proceed if we have valid references
                    if ($references.Count -gt 0) {
                        $attachmentParams = @{
                            references = $references
                        }
                        
                        try {
                            Update-MgPlannerTaskDetail -PlannerTaskId $newTask.Id -IfMatch $etag -BodyParameter $attachmentParams
                            Write-Information "Copied $($references.Count) valid attachment(s) for task '$($task.Title)'" -InformationAction Continue
                        }
                        catch {
                            Write-Warning "Failed to copy attachments for task '$($task.Title)': $_"
                        }
                    } else {
                        Write-Information "No valid attachments to copy for task '$($task.Title)'" -InformationAction Continue
                    }
                }
            }
            
            $successCount++
        }
        catch {
            Write-Warning "Failed to create task '$($task.Title)': $_"
        }
        
        Start-Sleep -Milliseconds 200
    }
}

# Output final results
$result = @{
    SourcePlanTitle = $sourcePlan.Title
    NewPlanId = $newPlanId
    NewPlanUrl = "https://planner.cloud.microsoft.com/webui/plan/$newPlanId"
    TasksCreated = $successCount
    BucketsCreated = $sourceBuckets.Count
    Status = "Completed"
}

Write-Information "Plan copy completed successfully" -InformationAction Continue
Write-Information "New Plan ID: $newPlanId" -InformationAction Continue # !Export this to sharepoint hub list
Write-Information "Tasks created: $successCount of $($sourceTaskData.Count)" -InformationAction Continue

# Return structured result for automation scenarios
return $result
