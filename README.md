# New-MSPlannerPlanFromSource.ps1

A PowerShell script that creates a new Microsoft Planner plan by copying the complete structure and content from an existing source plan.

## Features

- ✅ Copies all tasks with details (descriptions, checklists, priorities)
- ✅ Preserves bucket organization and ordering
- ✅ Maintains task ordering within buckets
- ✅ Copies file attachments and references
- ✅ Handles URL encoding/decoding for attachments
- ✅ Provides detailed progress reporting

## Prerequisites

- **PowerShell Module**: `Microsoft.Graph.Planner`
- **Authentication**: Must be pre-authenticated with Microsoft Graph
- **Permissions Required**:
  - `Tasks.ReadWrite.All`
  - `Tasks.ReadWrite.Shared`
  - `Group.Read.All`
  - `User.Read.All`

## Usage

```powershell
.\New-MSPlannerPlanFromSource.ps1 -SourcePlanId "abc123..." -UserId "def456..." -PlanTitle "My New Plan"
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `SourcePlanId` | String | Yes | ID of the existing plan to copy from |
| `UserId` | String | Yes | Object ID of the user who will own the new plan |
| `PlanTitle` | String | Yes | Title for the new Planner plan |

## Output

Returns a structured object containing:
- Source plan title
- New plan ID and URL
- Count of tasks and buckets created
- Completion status

## Authentication Setup

Ensure you're authenticated with Microsoft Graph before running:

```powershell
Connect-MgGraph -Scopes "Tasks.ReadWrite.All", "Tasks.ReadWrite.Shared", "Group.Read.All", "User.Read.All"
```

## Notes

- Script includes rate limiting (200-500ms delays) to avoid API throttling
- Validates URLs before copying attachments
- Maintains original task and bucket ordering using OrderHint analysis
- Designed for automation scenarios with structured output

## Author

David Sorenson - Version
