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

# ── Active DMC prompt template (meeting data injected at runtime) ─────────────
# Single user-message design with adaptive thinking — matches the reference
# Python implementation from the Anthropic Console.
$PromptTemplate = @"
You are the Active DMC Contact Report Agent. Your job is to transform meeting notes into a polished contact report following Active DMC's exact format.

The meeting data may come from any source — a transcript, a Read AI report, hand-written notes, an email summary, bullet points, a voice memo transcription, or any other format. Work with whatever is provided.

Here is the meeting data you will be working with:

<meeting_data>
{{MEETING_DATA}}
</meeting_data>

Your output must follow this exact structure:

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

Important rules to follow:

- Follow the Active DMC contact report format exactly as shown above
- Use "•" (bullet points) for main discussion topics and "o" (sub-bullets) for summary points under each topic
- Use numbered lists (1., 2.) for Next Steps owners, then "o" for individual action items
- Keep the tone professional, concise, and client-friendly
- Extract and synthesize information from the meeting data; do not copy the transcript word-for-word
- Organize discussion points by topic, not chronologically
- Preserve all specific details accurately: names, dates, deadlines, publications, clients, campaign details, deliverables, numbers, etc.
- For action items, carefully determine whether Active DMC or the client is responsible and place them under the correct owner
- If attendees, date, time, or location information is missing from the meeting data, use placeholders like "[Client Name]", "[Date]", "[Time]", or "[Location - Teams/Zoom/In person]"
- Always specify "UAE time" after the time
- Do not invent or assume information that is not present in the meeting data
- Write this as a professional follow-up email, not a meeting transcript or verbatim summary

Before writing your final contact report, use the scratchpad below to:
1. Identify the client name and company
2. List attendees from both sides
3. Extract meeting logistics (date, time, location)
4. Identify main discussion topics and organize key points under each
5. Extract action items and assign them to the correct owner (Active DMC or client)

<scratchpad>
[Your analysis and organization of the meeting data goes here]
</scratchpad>

Now write the complete contact report following the exact format specified above. Your final output should be the complete, polished contact report ready to send to the client - do not include the scratchpad in your final answer.

<contact_report>
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
$meetingJson  = $meeting | ConvertTo-Json -Depth 20
$userMessage  = $PromptTemplate -replace '{{MEETING_DATA}}', $meetingJson

$requestBody = @{
    model      = "claude-opus-4-7"
    max_tokens = 20000
    thinking   = @{ type = "adaptive" }
    messages   = @(
        @{
            role    = "user"
            content = @(
                @{
                    type = "text"
                    text = $userMessage
                }
            )
        }
    )
} | ConvertTo-Json -Depth 20

$AnthropicHeaders = @{
    "x-api-key"         = $AnthropicKey
    "anthropic-version" = "2023-06-01"
    "content-type"      = "application/json"
}

# ── Step 3: Call Claude ───────────────────────────────────────────────────────
Write-Host "Sending to Claude (adaptive thinking enabled)..." -ForegroundColor Cyan
$response = Invoke-RestMethod `
    -Uri "https://api.anthropic.com/v1/messages" `
    -Method POST `
    -Headers $AnthropicHeaders `
    -Body $requestBody

# Extract text blocks only (skip thinking blocks produced by adaptive thinking)
$textBlocks = $response.content | Where-Object { $_.type -eq "text" }
$rawText    = ($textBlocks | ForEach-Object { $_.text }) -join ""

# Strip the closing </contact_report> tag if Claude included it
$contactReport = $rawText -replace '\s*</contact_report>\s*$', '' |
                 ForEach-Object { $_.Trim() }

# ── Step 4: Save report as .docx via Word COM ────────────────────────────────
if (-not $OutputPath) {
    $safeName  = ($meetingTitle -replace '[\\/:*?"<>|]', '_').Trim()
    $datestamp = Get-Date -Format "yyyy-MM-dd"
    $OutputPath = Join-Path "$env:USERPROFILE\Desktop" "ADMC_ContactReport_${safeName}_${datestamp}.docx"
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false

try {
    $doc       = $word.Documents.Add()
    $selection = $word.Selection

    foreach ($line in $contactReport -split "`n") {
        $trimmed = $line.TrimEnd()

        if ($trimmed -eq "") {
            $selection.Style = $doc.Styles("Normal")
            $selection.TypeParagraph()
            continue
        }

        if ($trimmed -eq "Summary of Discussion" -or $trimmed -eq "Next Steps") {
            $selection.Style = $doc.Styles("Heading 2")
            $selection.TypeText($trimmed)
            $selection.TypeParagraph()
            continue
        }

        if ($trimmed -match "^•") {
            $selection.Style = $doc.Styles("List Bullet")
            $selection.TypeText(($trimmed -replace '^•[\t ]*', ''))
            $selection.TypeParagraph()
            continue
        }

        if ($trimmed -match "^o[\t ]") {
            $selection.Style = $doc.Styles("List Bullet 2")
            $selection.TypeText(($trimmed -replace '^o[\t ]*', ''))
            $selection.TypeParagraph()
            continue
        }

        if ($trimmed -match "^\d+\.\t") {
            $selection.Style = $doc.Styles("Normal")
            $selection.Font.Bold = $true
            $selection.TypeText(($trimmed -replace '\t', '  '))
            $selection.Font.Bold = $false
            $selection.TypeParagraph()
            continue
        }

        $selection.Style = $doc.Styles("Normal")
        $selection.TypeText($trimmed)
        $selection.TypeParagraph()
    }

    # wdFormatXMLDocument = 12
    $doc.SaveAs([ref]$OutputPath, [ref]12)
    $doc.Close($false)
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

Write-Host "`nContact report saved:" -ForegroundColor Green
Write-Host $OutputPath -ForegroundColor White

if ($OpenAfterSave) {
    Start-Process $OutputPath
}
