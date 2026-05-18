<#
.SYNOPSIS
    Active DMC Contact Report Agent — Option A (REST API Pull)

.DESCRIPTION
    Fetches the latest (or a specific) meeting from Read AI via REST API,
    sends the data to Claude, and saves a polished Active DMC contact report.

.PARAMETER MeetingId
    Optional. Pull a specific meeting by ID instead of the most recent one.

.PARAMETER OutputPath
    Optional. Full file path for the saved report.
    Defaults to Desktop\ADMC_ContactReport_<MeetingTitle>_<Date>.txt

.PARAMETER OpenAfterSave
    Optional switch. Opens the report in Notepad when done.

.EXAMPLE
    .\Contact-Report-Agent.ps1
    .\Contact-Report-Agent.ps1 -MeetingId "mtg_abc123"
    .\Contact-Report-Agent.ps1 -OutputPath "C:\Reports\ClientX.txt" -OpenAfterSave
#>

param(
    [string]$MeetingId,
    [string]$OutputPath,
    [switch]$OpenAfterSave
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Environment variables ────────────────────────────────────────────────────
$ReadToken    = $env:READ_AI_TOKEN
$AnthropicKey = $env:ANTHROPIC_API_KEY

if (-not $ReadToken)    { throw "READ_AI_TOKEN environment variable is not set." }
if (-not $AnthropicKey) { throw "ANTHROPIC_API_KEY environment variable is not set." }

# ── Active DMC system prompt ─────────────────────────────────────────────────
$SystemPrompt = @"
You are the Active DMC Contact Report Agent.

Your job is to take Read AI meeting data (summary, action items, topics, transcript, chapter summaries) and produce a polished client contact report in Active DMC's exact format.

Use this structure every time:

Hi [Client Name],

Thanks for your time – following our discussion please see the contact report below.

Present on behalf of [Client Company]:                       [Client attendees]

Present on behalf of Active DMC:                       [Active DMC attendees]

Date:                                                                [Date]

Time:                                                               [Time] UAE time

Location:                                                          [Teams/Zoom/In person]

Summary of Discussion

•	[Topic title]
o	[Summary point]
o	[Summary point]

•	[Topic title]
o	[Summary point]
o	[Summary point]

Next Steps

1.	Active DMC
o	[Action point owned by Active DMC]
o	[Action point owned by Active DMC]

2.	[Client Name/Client Company]
o	[Action point owned by client]
o	[Action point owned by client]

Please let us know if we missed anything or if you have any questions.

Rules:
- Follow the Active DMC contact report format exactly.
- Use "•" for main discussion points and "o" for sub-points.
- Keep the tone professional, concise, and client-friendly.
- Extract action items from the data and place them under the correct owner.
- Active DMC staff are identified by @activedmc.com email addresses or by context clues.
- Do not invent missing information.
- If attendees, date, time, or location are missing, use placeholders like [Date], [Time], [Location].
- Preserve all names, deadlines, publications, clients, campaign details, and deliverables accurately.
- Write like a post-meeting follow-up email, not a meeting transcript.
"@

# ── Step 1: Fetch meeting from Read AI ───────────────────────────────────────
$ReadHeaders = @{
    "Authorization" = "Bearer $ReadToken"
    "Accept"        = "application/json"
}

if (-not $MeetingId) {
    Write-Host "Fetching latest meeting from Read AI..." -ForegroundColor Cyan
    $list = Invoke-RestMethod `
        -Uri "https://api.read.ai/v1/meetings?limit=1" `
        -Headers $ReadHeaders `
        -Method GET
    $MeetingId = $list.data[0].id
    Write-Host "Found meeting: $MeetingId" -ForegroundColor Green
}

Write-Host "Retrieving full meeting data..." -ForegroundColor Cyan
$expand = "expand[]=summary&expand[]=transcript&expand[]=action_items&expand[]=topics&expand[]=chapter_summaries"
$meeting = Invoke-RestMethod `
    -Uri "https://api.read.ai/v1/meetings/$MeetingId`?$expand" `
    -Headers $ReadHeaders `
    -Method GET

$meetingTitle = if ($meeting.title) { $meeting.title } else { $MeetingId }
Write-Host "Retrieved: $meetingTitle" -ForegroundColor Green

# ── Step 2: Build Claude request ─────────────────────────────────────────────
$meetingJson = $meeting | ConvertTo-Json -Depth 20

$requestBody = @{
    model      = "claude-opus-4-7"
    max_tokens = 4096
    system     = $SystemPrompt
    messages   = @(
        @{
            role    = "user"
            content = "Please generate a contact report from this Read AI meeting data:`n`n$meetingJson"
        }
    )
} | ConvertTo-Json -Depth 20

$AnthropicHeaders = @{
    "x-api-key"         = $AnthropicKey
    "anthropic-version" = "2023-06-01"
    "content-type"      = "application/json"
}

# ── Step 3: Call Claude ───────────────────────────────────────────────────────
Write-Host "Sending to Claude..." -ForegroundColor Cyan
$response = Invoke-RestMethod `
    -Uri "https://api.anthropic.com/v1/messages" `
    -Method POST `
    -Headers $AnthropicHeaders `
    -Body $requestBody

$contactReport = $response.content[0].text

# ── Step 4: Save report ───────────────────────────────────────────────────────
if (-not $OutputPath) {
    $safeName  = ($meetingTitle -replace '[\\/:*?"<>|]', '_').Trim()
    $datestamp = Get-Date -Format "yyyy-MM-dd"
    $OutputPath = Join-Path "$env:USERPROFILE\Desktop" "ADMC_ContactReport_${safeName}_${datestamp}.txt"
}

$contactReport | Out-File -FilePath $OutputPath -Encoding utf8
Write-Host "`nContact report saved:" -ForegroundColor Green
Write-Host $OutputPath -ForegroundColor White

if ($OpenAfterSave) {
    Start-Process notepad $OutputPath
}
