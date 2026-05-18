<#
.SYNOPSIS
    Active DMC Contact Report Agent — Option B (Read AI Webhook Listener)

.DESCRIPTION
    Runs a local HTTP listener that receives Read AI webhook payloads,
    verifies the X-Read-Signature header, generates a contact report via
    Claude, and saves it to your Desktop.

    For production use, replace this with Azure Automation, Azure Function,
    or Power Automate. This script is ideal for local testing and small
    on-premises setups.

.PARAMETER Port
    Port to listen on. Default: 8080.
    For HTTPS on a server, run behind a reverse proxy (nginx/IIS) with a TLS cert.

.PARAMETER OutputFolder
    Folder where reports are saved. Default: Desktop.

.EXAMPLE
    .\Webhook-Listener.ps1
    .\Webhook-Listener.ps1 -Port 9000 -OutputFolder "C:\Reports"

.NOTES
    Read AI webhook setup:
      1. Go to Read AI → Integrations → Webhooks → Add Webhook.
      2. Set URL to http://<your-ip>:<Port>/webhook  (use ngrok for external testing).
      3. Copy the signing key into the READ_WEBHOOK_SIGNING_KEY env variable.
      4. Enable events: report.created (sends full report payload).
      5. Click "Test Webhook" to verify.

    Required environment variables:
      ANTHROPIC_API_KEY
      READ_WEBHOOK_SIGNING_KEY
#>

param(
    [int]    $Port         = 8080,
    [string] $OutputFolder = "$env:USERPROFILE\Desktop"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Environment variables ────────────────────────────────────────────────────
$AnthropicKey = $env:ANTHROPIC_API_KEY
$SigningKey    = $env:READ_WEBHOOK_SIGNING_KEY

if (-not $AnthropicKey) { throw "ANTHROPIC_API_KEY environment variable is not set." }
if (-not $SigningKey)    { throw "READ_WEBHOOK_SIGNING_KEY environment variable is not set." }

# ── Shared: Active DMC system prompt ─────────────────────────────────────────
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

# ── Helper: verify Read AI webhook signature ──────────────────────────────────
function Test-ReadSignature {
    param(
        [string]$RawBody,
        [string]$ReceivedSignature,
        [string]$Secret
    )
    $hmac    = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = [System.Text.Encoding]::UTF8.GetBytes($Secret)
    $hash    = $hmac.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($RawBody))
    $computed = "sha256=" + ([System.BitConverter]::ToString($hash) -replace '-', '').ToLower()
    return $computed -eq $ReceivedSignature
}

# ── Helper: generate contact report via Claude ────────────────────────────────
function Invoke-ContactReport {
    param([string]$MeetingJson)

    $requestBody = @{
        model      = "claude-opus-4-7"
        max_tokens = 4096
        system     = $SystemPrompt
        messages   = @(
            @{
                role    = "user"
                content = "Please generate a contact report from this Read AI meeting data:`n`n$MeetingJson"
            }
        )
    } | ConvertTo-Json -Depth 20

    $headers = @{
        "x-api-key"         = $AnthropicKey
        "anthropic-version" = "2023-06-01"
        "content-type"      = "application/json"
    }

    $response = Invoke-RestMethod `
        -Uri "https://api.anthropic.com/v1/messages" `
        -Method POST `
        -Headers $headers `
        -Body $requestBody

    return $response.content[0].text
}

# ── Helper: save report ───────────────────────────────────────────────────────
function Save-Report {
    param(
        [string]$ReportText,
        [string]$MeetingTitle
    )
    $safeName  = ($MeetingTitle -replace '[\\/:*?"<>|]', '_').Trim()
    $datestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
    $filename  = "ADMC_ContactReport_${safeName}_${datestamp}.txt"
    $fullPath  = Join-Path $OutputFolder $filename
    $ReportText | Out-File -FilePath $fullPath -Encoding utf8
    return $fullPath
}

# ── HTTP listener ─────────────────────────────────────────────────────────────
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$Port/")
$listener.Start()

Write-Host "Webhook listener running on port $Port" -ForegroundColor Green
Write-Host "Endpoint : http://localhost:$Port/webhook" -ForegroundColor Cyan
Write-Host "Reports  : $OutputFolder" -ForegroundColor Cyan
Write-Host "Press Ctrl+C to stop.`n" -ForegroundColor Yellow

try {
    while ($listener.IsListening) {

        $context  = $listener.GetContext()
        $request  = $context.Request
        $response = $context.Response

        # Read raw body (needed for signature verification)
        $reader  = New-Object System.IO.StreamReader($request.InputStream)
        $rawBody = $reader.ReadToEnd()
        $reader.Close()

        $path   = $request.Url.LocalPath
        $method = $request.HttpMethod

        Write-Host "$(Get-Date -Format 'HH:mm:ss') $method $path" -ForegroundColor Gray

        # ── Health check ──────────────────────────────────────────────────────
        if ($path -eq "/health" -and $method -eq "GET") {
            $response.StatusCode = 200
            $bytes = [System.Text.Encoding]::UTF8.GetBytes('{"status":"ok"}')
            $response.ContentType   = "application/json"
            $response.ContentLength64 = $bytes.Length
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
            $response.Close()
            continue
        }

        # ── Webhook endpoint ──────────────────────────────────────────────────
        if ($path -eq "/webhook" -and $method -eq "POST") {

            # 1. Verify signature
            $sig = $request.Headers["X-Read-Signature"]
            if (-not $sig) {
                Write-Host "  Missing X-Read-Signature — rejected" -ForegroundColor Red
                $response.StatusCode = 401
                $response.Close()
                continue
            }

            if (-not (Test-ReadSignature -RawBody $rawBody -ReceivedSignature $sig -Secret $SigningKey)) {
                Write-Host "  Invalid signature — rejected" -ForegroundColor Red
                $response.StatusCode = 401
                $response.Close()
                continue
            }

            # 2. Acknowledge immediately (Read AI expects 2xx within a few seconds)
            $response.StatusCode = 200
            $ackBytes = [System.Text.Encoding]::UTF8.GetBytes('{"received":true}')
            $response.ContentType     = "application/json"
            $response.ContentLength64 = $ackBytes.Length
            $response.OutputStream.Write($ackBytes, 0, $ackBytes.Length)
            $response.Close()

            # 3. Parse payload
            try {
                $payload = $rawBody | ConvertFrom-Json
                $eventType = $payload.event

                Write-Host "  Event    : $eventType" -ForegroundColor Cyan

                # Only process report.created events
                if ($eventType -ne "report.created") {
                    Write-Host "  Skipped (not report.created)" -ForegroundColor Yellow
                    continue
                }

                $meetingData  = $payload.data
                $meetingTitle = if ($meetingData.title) { $meetingData.title } else { "Meeting" }
                Write-Host "  Meeting  : $meetingTitle" -ForegroundColor Cyan

                # 4. Generate report
                Write-Host "  Sending to Claude..." -ForegroundColor Cyan
                $meetingJson   = $meetingData | ConvertTo-Json -Depth 20
                $contactReport = Invoke-ContactReport -MeetingJson $meetingJson

                # 5. Save report
                $savedPath = Save-Report -ReportText $contactReport -MeetingTitle $meetingTitle
                Write-Host "  Saved    : $savedPath" -ForegroundColor Green

            } catch {
                Write-Host "  Error processing payload: $_" -ForegroundColor Red
            }

            continue
        }

        # ── 404 for everything else ───────────────────────────────────────────
        $response.StatusCode = 404
        $response.Close()
    }
} finally {
    $listener.Stop()
    Write-Host "`nListener stopped." -ForegroundColor Yellow
}
