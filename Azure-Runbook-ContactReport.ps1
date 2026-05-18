<#
.SYNOPSIS
    Active DMC Contact Report Agent — Azure Automation Runbook

.DESCRIPTION
    Designed to run inside Azure Automation (called by Power Automate or
    a Logic App that receives the Read AI webhook).

    Power Automate flow:
      1. "When an HTTP request is received" trigger
      2. Parse the Read AI JSON payload
      3. "Create job" in Azure Automation, passing the payload as a parameter
      4. (Optional) "Send an email (V2)" with the completed report

    Azure Automation setup:
      1. Create an Automation Account.
      2. Add Credentials or Variables for ANTHROPIC_API_KEY.
      3. Import this file as a PowerShell Runbook.
      4. Publish the Runbook.
      5. Link it from Power Automate using the Azure Automation connector.

.PARAMETER WebhookPayload
    The raw JSON string sent by Read AI (passed in by Power Automate).
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$WebhookPayload
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Retrieve API key from Azure Automation credential store ──────────────────
# Store your key in: Automation Account → Shared Resources → Credentials
# Name the credential "AnthropicApiKey", username = "key", password = <your key>
$credential   = Get-AutomationPSCredential -Name "AnthropicApiKey"
$AnthropicKey = $credential.GetNetworkCredential().Password

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

# ── Parse payload ─────────────────────────────────────────────────────────────
$payload      = $WebhookPayload | ConvertFrom-Json
$meetingData  = if ($payload.data) { $payload.data } else { $payload }
$meetingTitle = if ($meetingData.title) { $meetingData.title } else { "Meeting" }
$meetingJson  = $meetingData | ConvertTo-Json -Depth 20

Write-Output "Processing: $meetingTitle"

# ── Call Claude ───────────────────────────────────────────────────────────────
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

$response = Invoke-RestMethod `
    -Uri "https://api.anthropic.com/v1/messages" `
    -Method POST `
    -Headers @{
        "x-api-key"         = $AnthropicKey
        "anthropic-version" = "2023-06-01"
        "content-type"      = "application/json"
    } `
    -Body $requestBody

$contactReport = $response.content[0].text

Write-Output "Report generated successfully."

# ── Output the report (Power Automate reads this as the job output) ───────────
# Power Automate: use "Get job output" action after this runbook completes,
# then pipe $contactReport into "Send an email" or "Create file in SharePoint".
Write-Output $contactReport
