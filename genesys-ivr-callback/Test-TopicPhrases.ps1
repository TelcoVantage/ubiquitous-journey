#Requires -Version 5.1
<#
.SYNOPSIS
    Genesys Cloud Speech & Text Analytics: run the "test phrase" job for a list of
    candidate topic phrases and report how many transcripts each one matches.

.DESCRIPTION
    For every phrase in $Phrases (CONFIG block below) the script:
      1. POST /api/v2/speechandtextanalytics/topics/testphrase/jobs
         (same body the Genesys UI sends from Topic > Add phrase > Test)
      2. Polls GET /api/v2/speechandtextanalytics/topics/testphrase/jobs/{jobId}
         until the job finishes
      3. Records processedTranscriptsCount / matchedTranscriptsCount
      4. Matched conversations (skip with -CountsOnly): the exact list the UI shows.
         Genesys publishes it only on the notification topic
         v2.speechandtextanalytics.topics.testphrase.jobs.{jobId}, over a WebSocket.
         .NET WebSockets are blocked in Constrained Language Mode, so the script
         creates a channel (POST /api/v2/notifications/channels), runs the Windows
         built-in curl.exe (8.11+ has WebSocket support) as a child process to read
         it, subscribes each job's topic, and parses the frames curl writes.
      5. Optional (-TranscriptSearch): a wording-based list of conversations from
         POST /api/v2/speechandtextanalytics/transcripts/search + each transcript.

    Output:
      TopicPhraseTest_<stamp>.csv     one row per phrase: counts, conversations received
      TopicPhraseMatches_<stamp>.csv  one row per matched conversation: phrase,
                                      conversation id, time, found phrase, snippet,
                                      confidence, communication/transcript id
      TopicPhraseSearch_<stamp>.csv   only with -TranscriptSearch

    Constrained Language Mode safe (Windows PowerShell 5.1 under AppLocker/WDAC):
    no .NET static calls, no ::new(), no [pscustomobject] casts, no Add-Type.

    OAuth client (Client Credentials grant) role permissions:
      speechAndTextAnalytics:topic:testPhrase     phrase test jobs + notifications
      analytics:conversationDetail:view           -TranscriptSearch only
      recording:recording:view                    -TranscriptSearch only
      speechAndTextAnalytics:data:view            -TranscriptSearch only

    Requires curl.exe 8.11 or later with WebSocket support (Windows 11 / Server 2025
    ship it; check "curl.exe -V" lists WebSockets under Features). Without it the
    script still reports the counts.

.EXAMPLE
    # Everything set in the CONFIG block
    .\Test-TopicPhrases.ps1

.EXAMPLE
    # Same phrases, different window and strictness, lexical matching
    .\Test-TopicPhrases.ps1 -StartDate '2026-09-01' -EndDate '2026-09-30' -Strictness 60 -MatchingType Lexical

.EXAMPLE
    # Phrases from a text file (one per line) instead of the CONFIG list
    .\Test-TopicPhrases.ps1 -PhraseFile .\phrases.txt
#>
[CmdletBinding()]
param(
    # ==============================================================================
    # CONFIG - fill these in. Anything passed on the command line overrides them.
    # Do not commit a real client secret to source control.
    # ==============================================================================

    # Region domain: mypurecloud.com, mypurecloud.ie, mypurecloud.de, mypurecloud.com.au,
    # mypurecloud.jp, usw2.pure.cloud, cac1.pure.cloud, euw2.pure.cloud, aps1.pure.cloud, ...
    [string]$Environment = 'mypurecloud.com.au',

    # Client Credentials OAuth client.
    [string]$ClientId = 'PASTE-CLIENT-ID-HERE',
    [string]$ClientSecret = 'PASTE-CLIENT-SECRET-HERE',

    # Speech & Text Analytics program IDs to test against (from the UI payload).
    [string[]]$ProgramIds = @(
        'b808f776-8f0c-4e3f-a932-11b144af4ddf',
        'd54ba0a0-0cd1-4a9f-ba5b-c9afa1d05c8b'
    ),

    # Candidate phrases (Vulnerable client consent).
    [string[]]$Phrases = @(
        "before i do that i need your consent to record this information",
        "can i add this to your profile so you won't have to repeat it in future",
        "can i obtain your consent to record this information",
        "i just need your agreement to add those details to your account",
        "i need your consent to add this information",
        "i will need your permission to add this",
        "i would like to put a note here so that you can get some extra help",
        "is it alright to record this",
        "is it okay if i add a note that you may need extra support",
        "let me add that with your permission",
        "with your permission i would like to add that to your profile",
        "would it be okay if i added a note so you don't need to explain this again",
        "would it help if future agents were aware of this",
        "would you be comfortable if i added that",
        "would you like me to record that so you don't have to tell us again",
        "would you like us to be aware of that for future calls",
        "would you like us to capture that in our system",
        "would you like us to make a note so we can provide extra support",
        "would you like us to note that so we can better assist you",
        "would you like us to record that so we can take that into account in future"
    ),

    # Topic settings (same as the UI test dialog).
    [string]$Dialect = 'en-AU',
    [ValidateSet('Semantic', 'Lexical')]
    [string]$MatchingType = 'Semantic',
    # Internal = agent side, External = customer side, Both = either.
    [ValidateSet('Internal', 'External', 'Both')]
    [string]$Participants = 'Internal',
    [ValidateRange(1, 100)]
    [int]$Strictness = 72,

    # Transcript filters. The UI default window is the last 29 days.
    [ValidateSet('call', 'chat', 'email', 'message', 'all')]
    [string]$MediaType = 'call',
    [datetime]$StartDate = (Get-Date).AddDays(-29),
    [datetime]$EndDate = (Get-Date),
    [string[]]$QueueIds = @(),
    [string[]]$FlowIds = @(),

    # ==============================================================================
    # END OF CONFIG
    # ==============================================================================

    # Optional: read phrases from a text file (one per line) instead of $Phrases.
    [string]$PhraseFile,
    # Or an existing bearer token (skips OAuth).
    [string]$AccessToken,

    # How many test jobs to keep running at once.
    [ValidateRange(1, 10)]
    [int]$MaxConcurrentJobs = 3,
    # Seconds between status polls, and how long to wait for one job before giving up.
    [ValidateRange(1, 60)]
    [int]$PollSeconds = 3,
    [ValidateRange(30, 3600)]
    [int]$JobTimeoutSeconds = 600,
    [ValidateRange(0, 10)]
    [int]$MaxRetries = 6,

    # --- Matched conversations: the exact list the UI shows ---
    # Genesys publishes a test job's matched conversations only on the notification
    # topic v2.speechandtextanalytics.topics.testphrase.jobs.{jobId}, over a WebSocket.
    # .NET WebSockets are blocked in Constrained Language Mode, so the script runs the
    # Windows built-in curl.exe (8.11 or later has WebSocket support) as a child process
    # and parses the frames it writes to a file.
    # Skip this (counts only):
    [switch]$CountsOnly,
    [string]$CurlPath = 'curl.exe',
    # curl does not use the Windows system proxy; set this if you need one, e.g. http://proxy:8080
    [string]$ProxyUrl,
    # Seconds to wait for a finished job's notification before giving up on it.
    [ValidateRange(5, 900)]
    [int]$NotificationWaitSeconds = 120,
    # Hard limit on how long the listener process may run.
    [ValidateRange(60, 14400)]
    [int]$ListenerMaxSeconds = 3600,
    # Keep the raw notification frames file (troubleshooting).
    [switch]$KeepListenerLog,

    # --- Optional extra: wording-based transcript search (writes TopicPhraseSearch_*.csv) ---
    # Finds calls whose transcript contains the phrase's WORDING via
    # POST /api/v2/speechandtextanalytics/transcripts/search. Off by default.
    [switch]$TranscriptSearch,
    # Max conversations to pull per phrase (each costs 2-3 API calls).
    [ValidateRange(1, 1000)]
    [int]$MaxMatchesPerPhrase = 25,
    # PHRASE = words in order, loose; EXACT_PHRASE = exact wording.
    [ValidateSet('PHRASE', 'EXACT_PHRASE')]
    [string]$SearchMatchType = 'PHRASE',
    # Transcript search field names. Genesys doesn't document these well; if the search
    # returns HTTP 400 or no hits, these are the first thing to adjust.
    [string]$SearchTextField = 'transcript.content',
    [string]$SearchDateField = 'conversationStartTime',
    [string]$SearchMediaTypeField = 'mediaType',
    # The search REQUIRES a language criterion (REQUIRED_SEARCH_FIELD: language).
    # Default: the -Dialect value. Try lower-case (en-au) if hits are unexpectedly 0.
    [string]$SearchLanguageField = 'language',
    [string]$SearchLanguage,
    # Minimum word overlap (%) between the phrase and a transcript sentence to report it.
    [ValidateRange(1, 100)]
    [int]$MinMatchScore = 60,

    [string]$OutputFolder = '.'
)

$ErrorActionPreference = 'Stop'

# ----------------------------------------------------------------------------------
# Generic helpers (CLM safe)
# ----------------------------------------------------------------------------------

function Get-CleanList {
    param([string[]]$Values, [switch]$SplitCommas)
    foreach ($v in $Values) {
        if ($null -eq $v) { continue }
        $pieces = @([string]$v)
        if ($SplitCommas) { $pieces = @(([string]$v) -split ',') }
        foreach ($p in $pieces) {
            $t = $p.Trim()
            if ($t) { $t }
        }
    }
}

function Get-GcDomain {
    param([string]$Value)
    $d = ([string]$Value).Trim().ToLower()
    $d = $d.Replace('https://', '').Replace('http://', '')
    $slash = $d.IndexOf('/')
    if ($slash -ge 0) { $d = $d.Substring(0, $slash) }
    foreach ($prefix in @('api.', 'login.', 'apps.')) {
        if ($d.StartsWith($prefix)) { $d = $d.Substring($prefix.Length) }
    }
    if (-not $d) { throw 'Environment is empty. Use e.g. mypurecloud.com, mypurecloud.ie, mypurecloud.com.au, euw2.pure.cloud' }
    return $d
}

# UTF-8 encode a string into byte values (ints), handling surrogate pairs.
function ConvertTo-Utf8ByteValues {
    param([string]$Text)
    $chars = $Text.ToCharArray()
    $n = $chars.Length
    for ($i = 0; $i -lt $n; $i++) {
        $c = [int]$chars[$i]
        if ($c -ge 0xD800 -and $c -le 0xDBFF -and ($i + 1) -lt $n) {
            $low = [int]$chars[$i + 1]
            if ($low -ge 0xDC00 -and $low -le 0xDFFF) {
                $cp = 0x10000 + (($c - 0xD800) -shl 10) + ($low - 0xDC00)
                $i++
                0xF0 -bor ($cp -shr 18)
                0x80 -bor (($cp -shr 12) -band 0x3F)
                0x80 -bor (($cp -shr 6) -band 0x3F)
                0x80 -bor ($cp -band 0x3F)
                continue
            }
        }
        if ($c -lt 0x80) {
            $c
        }
        elseif ($c -lt 0x800) {
            0xC0 -bor ($c -shr 6)
            0x80 -bor ($c -band 0x3F)
        }
        else {
            0xE0 -bor ($c -shr 12)
            0x80 -bor (($c -shr 6) -band 0x3F)
            0x80 -bor ($c -band 0x3F)
        }
    }
}

# Base64 without [Convert]::ToBase64String / [Text.Encoding].
function ConvertTo-Base64Clm {
    param([string]$Text)
    $alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
    $bytes = @(ConvertTo-Utf8ByteValues -Text $Text)
    $n = $bytes.Count
    $out = ''
    for ($i = 0; $i -lt $n; $i += 3) {
        $b0 = [int]$bytes[$i]
        $b1 = 0
        $b2 = 0
        if (($i + 1) -lt $n) { $b1 = [int]$bytes[$i + 1] }
        if (($i + 2) -lt $n) { $b2 = [int]$bytes[$i + 2] }
        $triple = ($b0 -shl 16) -bor ($b1 -shl 8) -bor $b2
        $out += $alphabet[($triple -shr 18) -band 0x3F]
        $out += $alphabet[($triple -shr 12) -band 0x3F]
        if (($i + 1) -lt $n) { $out += $alphabet[($triple -shr 6) -band 0x3F] } else { $out += '=' }
        if (($i + 2) -lt $n) { $out += $alphabet[$triple -band 0x3F] } else { $out += '=' }
    }
    return $out
}

function Get-Padded {
    param([int]$Value, [int]$Width)
    return ([string]$Value).PadLeft($Width, '0')
}

function Format-Local {
    param($Date)
    if ($null -eq $Date) { return '' }
    $l = ([datetime]$Date).ToLocalTime()
    return (Get-Padded $l.Year 4) + '-' + (Get-Padded $l.Month 2) + '-' + (Get-Padded $l.Day 2) + ' ' +
           (Get-Padded $l.Hour 2) + ':' + (Get-Padded $l.Minute 2) + ':' + (Get-Padded $l.Second 2)
}

# Unix epoch milliseconds. '1970-01-01' with no offset parses as Kind=Unspecified,
# and DateTime subtraction ignores Kind, so this is a plain UTC difference.
function ConvertTo-EpochMs {
    param([datetime]$Date)
    $epoch = [datetime]'1970-01-01'
    return [long]($Date.ToUniversalTime() - $epoch).TotalMilliseconds
}

function Format-Pct {
    param($Part, $Whole)
    if ($null -eq $Whole -or [double]$Whole -le 0) { return '' }
    $v = 100.0 * [double]$Part / [double]$Whole
    return [string]([double]([long]($v * 100)) / 100)
}

# ----------------------------------------------------------------------------------
# HTTP / Genesys Cloud API
# ----------------------------------------------------------------------------------

function Get-HttpStatus {
    param($ErrorRecord)
    $code = 0
    try { $code = [int]$ErrorRecord.Exception.Response.StatusCode } catch { $code = 0 }
    if ($code -gt 0) { return $code }
    $msg = ''
    try { $msg = [string]$ErrorRecord.Exception.Message } catch { $msg = '' }
    # PS 5.1: "The remote server returned an error: (429) Too Many Requests."
    # PS 7  : "Response status code does not indicate success: 429 (Too Many Requests)."
    foreach ($c in @(400, 401, 403, 404, 408, 409, 413, 429, 500, 502, 503, 504)) {
        if ($msg -like "*($c)*" -or $msg -like "*: $c *" -or $msg -like "* $c (*") { return $c }
    }
    return 0
}

function Get-ErrorText {
    param($ErrorRecord)
    $t = ''
    try {
        if ($null -ne $ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
            $t = [string]$ErrorRecord.ErrorDetails.Message
        }
    } catch { $t = '' }
    if (-not $t) {
        try { $t = [string]$ErrorRecord.Exception.Message } catch { $t = '' }
    }
    if ($t.Length -gt 600) { $t = $t.Substring(0, 600) + '...' }
    return $t
}

function Get-GcToken {
    $basic = ConvertTo-Base64Clm -Text ($script:OAuthClientId + ':' + $script:OAuthClientSecret)
    $uri = $script:LoginBase + '/oauth/token'
    $resp = $null
    try {
        $resp = Invoke-RestMethod -Method Post -Uri $uri -Headers @{ Authorization = ('Basic ' + $basic) } `
            -Body 'grant_type=client_credentials' -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    }
    catch {
        $status = Get-HttpStatus $_
        $text = Get-ErrorText $_
        throw ('OAuth token request to {0} failed: HTTP {1}. {2} -- check -Environment, the client id/secret, and that the OAuth client uses the Client Credentials grant.' -f $uri, $status, $text)
    }
    if ($null -eq $resp -or -not $resp.access_token) {
        throw ('OAuth response from {0} did not contain an access_token.' -f $uri)
    }
    $script:ApiHeaders = @{ Authorization = ('Bearer ' + [string]$resp.access_token) }
    Write-Host ('Authenticated against {0} (token valid {1} s).' -f $script:LoginBase, [string]$resp.expires_in)
}

function Invoke-GcApi {
    param([string]$Method, [string]$Path, $Body)
    $uri = $script:ApiBase + $Path
    $json = $null
    if ($null -ne $Body) { $json = ConvertTo-Json -InputObject $Body -Depth 20 -Compress }
    $attempt = 0
    $reauthDone = $false
    while ($true) {
        $attempt++
        $status = 0
        $errText = ''
        try {
            if ($null -ne $json) {
                $result = Invoke-RestMethod -Method $Method -Uri $uri -Headers $script:ApiHeaders -Body $json `
                    -ContentType 'application/json; charset=utf-8' -ErrorAction Stop
            }
            else {
                $result = Invoke-RestMethod -Method $Method -Uri $uri -Headers $script:ApiHeaders -ErrorAction Stop
            }
            return $result
        }
        catch {
            $status = Get-HttpStatus $_
            $errText = Get-ErrorText $_
        }

        if ($status -eq 401 -and $script:CanReauth -and -not $reauthDone) {
            Write-Host '  Access token rejected (401) - requesting a new one...'
            Get-GcToken
            $reauthDone = $true
            continue
        }

        $retryable = ($status -eq 429 -or $status -eq 408 -or $status -ge 500 -or ($status -eq 0 -and $attempt -le 2))
        if ($retryable -and $attempt -le $script:MaxRetries) {
            $wait = 2
            for ($k = 1; $k -lt $attempt; $k++) { $wait = $wait * 2 }
            if ($wait -gt 60) { $wait = 60 }
            Write-Warning ('{0} {1} returned HTTP {2}; retry {3}/{4} in {5}s' -f $Method.ToUpper(), $Path, $status, $attempt, $script:MaxRetries, $wait)
            Start-Sleep -Seconds $wait
            continue
        }

        $hint = ''
        if ($status -eq 403) {
            $hint = ' -- the OAuth client role needs speechAndTextAnalytics:topic:testPhrase (plus analytics:conversationDetail:view, recording:recording:view and speechAndTextAnalytics:data:view for matched conversations) and access to the divisions involved.'
        }
        throw ('Genesys API {0} {1} failed: HTTP {2}. {3}{4}' -f $Method.ToUpper(), $Path, $status, $errText, $hint)
    }
}

# ----------------------------------------------------------------------------------
# Test-phrase jobs
# ----------------------------------------------------------------------------------

function Test-JobFinished {
    param([string]$State)
    $s = ([string]$State).ToLower()
    return ($s -in @('completed', 'complete', 'fulfilled', 'succeeded', 'success', 'done', 'finished',
                     'failed', 'error', 'cancelled', 'canceled', 'expired'))
}

function Test-JobFailed {
    param([string]$State)
    $s = ([string]$State).ToLower()
    return ($s -in @('failed', 'error', 'cancelled', 'canceled', 'expired'))
}

function Submit-PhraseJob {
    param([string]$Phrase)
    $body = @{
        topic = @{
            phrase       = @{ text = $Phrase }
            dialect      = $script:DialectValue
            matchingType = $script:MatchingTypeValue
            participants = $script:ParticipantsValue
            strictness   = $script:StrictnessValue
        }
        transcriptsFilters = @{
            mediaType   = $script:MediaTypeValue
            startTimeMs = $script:StartMs
            endTimeMs   = $script:EndMs
            programs    = @($script:ProgramList)
            queues      = @($script:QueueList)
            flows       = @($script:FlowList)
        }
    }
    $resp = Invoke-GcApi -Method Post -Path '/api/v2/speechandtextanalytics/topics/testphrase/jobs' -Body $body
    if ($null -eq $resp -or -not $resp.id) {
        throw ('Test-phrase job for "{0}" was accepted but the response had no job id.' -f $Phrase)
    }
    return $resp
}

# ----------------------------------------------------------------------------------
# Matched conversations: notification channel read through curl.exe (WebSocket)
# ----------------------------------------------------------------------------------

# Splits a stream of concatenated JSON objects ("{..}{..}{..}") into complete objects.
# An incomplete trailing object (still being written) is left out.
function Split-JsonObjects {
    param([string]$Text)
    $objects = @()
    if (-not $Text) { return $objects }
    $chars = $Text.ToCharArray()
    $depth = 0
    $inString = $false
    $escape = $false
    $start = -1
    for ($i = 0; $i -lt $chars.Length; $i++) {
        $c = $chars[$i]
        if ($inString) {
            if ($escape) { $escape = $false }
            elseif ($c -eq '\') { $escape = $true }
            elseif ($c -eq '"') { $inString = $false }
            continue
        }
        if ($c -eq '"') { $inString = $true; continue }
        if ($c -eq '{') {
            if ($depth -eq 0) { $start = $i }
            $depth++
        }
        elseif ($c -eq '}') {
            if ($depth -gt 0) {
                $depth--
                if ($depth -eq 0 -and $start -ge 0) {
                    $objects += $Text.Substring($start, $i - $start + 1)
                    $start = -1
                }
            }
        }
    }
    return $objects
}

function Test-CurlWebSocket {
    $text = ''
    try { $text = (& $script:CurlPathValue -V 2>&1) | Out-String }
    catch { return @{ Ok = $false; Reason = ('could not run "{0}": {1}' -f $script:CurlPathValue, $_.Exception.Message); Version = '' } }
    $version = ''
    if ($text -match 'curl\s+(\d+\.\d+(\.\d+)?)') { $version = $matches[1] }
    if ($text -notmatch '(?i)websocket') {
        return @{ Ok = $false; Version = $version; Reason = ('curl {0} at "{1}" was built without WebSocket support (needs curl 8.11 or later; check "curl.exe -V" lists WebSockets under Features)' -f $version, $script:CurlPathValue) }
    }
    return @{ Ok = $true; Version = $version; Reason = '' }
}

# Creates a notification channel and starts curl.exe reading its WebSocket into a file.
function Start-NotificationListener {
    $channel = Invoke-GcApi -Method Post -Path '/api/v2/notifications/channels'
    if ($null -eq $channel -or -not $channel.id -or -not $channel.connectUri) {
        throw 'POST /api/v2/notifications/channels did not return a channel id and connectUri.'
    }
    $script:ChannelId = [string]$channel.id
    $script:ListenerOut = Join-Path -Path $script:TempFolder -ChildPath ('gc_testphrase_ws_' + $script:RunStamp + '.log')
    $script:ListenerErr = Join-Path -Path $script:TempFolder -ChildPath ('gc_testphrase_ws_' + $script:RunStamp + '.err')
    $args = @('--no-buffer', '--silent', '--show-error', '--max-time', [string]$script:ListenerMaxSecondsValue)
    if ($script:ProxyUrlValue) { $args += @('--proxy', $script:ProxyUrlValue) }
    $args += [string]$channel.connectUri
    $proc = Start-Process -FilePath $script:CurlPathValue -ArgumentList $args -NoNewWindow -PassThru `
        -RedirectStandardOutput $script:ListenerOut -RedirectStandardError $script:ListenerErr
    Start-Sleep -Seconds 2
    if ($proc.HasExited) {
        $err = ''
        try { $err = (Get-Content -Path $script:ListenerErr -Raw -ErrorAction SilentlyContinue) } catch { $err = '' }
        throw ('curl exited straight away (code {0}). {1}' -f $proc.ExitCode, ([string]$err).Trim())
    }
    $script:ListenerProc = $proc
    $script:FramesSeen = 0
    Write-Host ('Notification listener: channel {0}, curl pid {1}' -f $script:ChannelId, $proc.Id)
}

function Stop-NotificationListener {
    if ($null -eq $script:ListenerProc) { return }
    try { if (-not $script:ListenerProc.HasExited) { Stop-Process -Id $script:ListenerProc.Id -Force -ErrorAction SilentlyContinue } } catch { }
    $script:ListenerProc = $null
    if (-not $script:KeepListenerLogValue) {
        Remove-Item -Path $script:ListenerOut -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $script:ListenerErr -Force -ErrorAction SilentlyContinue
    }
}

function Add-JobSubscription {
    param([string]$JobId)
    $topic = 'v2.speechandtextanalytics.topics.testphrase.jobs.' + $JobId
    $resp = Invoke-GcApi -Method Post -Path ('/api/v2/notifications/channels/' + $script:ChannelId + '/subscriptions?ignoreErrors=true') -Body @(@{ id = $topic })
    foreach ($e in $resp.entities) {
        if ($null -ne $e -and [string]$e.id -eq $topic -and $e.rejectionReason) {
            Write-Warning ('Subscription to {0} rejected: {1} (missing permissions: {2})' -f $topic, $e.rejectionReason, (@($e.missingPermissions) -join ', '))
        }
    }
}

# Reads new frames from the listener file; stores matched transcripts per job id.
function Read-JobNotifications {
    if ($null -eq $script:ListenerProc) { return }
    $raw = ''
    try { $raw = Get-Content -Path $script:ListenerOut -Raw -ErrorAction SilentlyContinue } catch { $raw = '' }
    if (-not $raw) { return }
    $frames = @(Split-JsonObjects -Text $raw)
    for ($i = $script:FramesSeen; $i -lt $frames.Count; $i++) {
        $msg = $null
        try { $msg = $frames[$i] | ConvertFrom-Json } catch { $msg = $null }
        if ($null -eq $msg -or -not $msg.topicName) { continue }
        $topic = [string]$msg.topicName
        if ($topic -notlike '*testphrase.jobs.*') { continue }
        $body = $msg.eventBody
        $jobId = ''
        if ($null -ne $body -and $body.id) { $jobId = [string]$body.id }
        if (-not $jobId) { $jobId = $topic.Substring($topic.LastIndexOf('.') + 1) }
        if (-not $script:JobResults.ContainsKey($jobId)) { $script:JobResults[$jobId] = @{ State = ''; Transcripts = @(); Notifications = 0 } }
        $entry = $script:JobResults[$jobId]
        $entry.Notifications = $entry.Notifications + 1
        if ($null -ne $body -and $body.state) { $entry.State = [string]$body.state }
        foreach ($res in @($body.testTopicPhraseResults)) {
            if ($null -eq $res) { continue }
            foreach ($t in @($res.matchedTranscripts)) { if ($null -ne $t) { $entry.Transcripts += , $t } }
        }
        Write-Verbose ('notification for job {0}: state {1}, {2} matched transcripts so far' -f $jobId, $entry.State, $entry.Transcripts.Count)
    }
    $script:FramesSeen = $frames.Count
}

# ----------------------------------------------------------------------------------
# Optional: wording-based transcript search + transcript download
# ----------------------------------------------------------------------------------

# ISO-8601 UTC built from components: culture and calendar independent.
function Format-IsoUtc {
    param([datetime]$Date)
    $u = $Date.ToUniversalTime()
    return (Get-Padded $u.Year 4) + '-' + (Get-Padded $u.Month 2) + '-' + (Get-Padded $u.Day 2) + 'T' +
           (Get-Padded $u.Hour 2) + ':' + (Get-Padded $u.Minute 2) + ':' + (Get-Padded $u.Second 2) + '.' +
           (Get-Padded $u.Millisecond 3) + 'Z'
}

# First non-empty value among property names (dotted paths allowed, e.g. 'conversation.id').
function Get-Prop {
    param($Obj, [string[]]$Names)
    foreach ($n in $Names) {
        $cur = $Obj
        foreach ($part in $n.Split('.')) {
            if ($null -eq $cur) { break }
            $cur = $cur.$part
        }
        if ($null -ne $cur -and [string]$cur -ne '') { return $cur }
    }
    return $null
}

function Get-Words {
    param([string]$Text)
    $t = ([string]$Text).ToLower() -replace "[^a-z0-9' ]", ' '
    return @($t -split '\s+' | Where-Object { $_ })
}

# % of the phrase's distinct words that appear in the text; 100 when the phrase
# appears verbatim (ignoring case/punctuation).
function Get-MatchScore {
    param([string[]]$PhraseWords, [string]$Text)
    $textWords = @(Get-Words $Text)
    if ($PhraseWords.Count -eq 0 -or $textWords.Count -eq 0) { return 0 }
    if ((' ' + ($textWords -join ' ') + ' ').Contains(' ' + ($PhraseWords -join ' ') + ' ')) { return 100 }
    $set = @{}
    foreach ($w in $textWords) { $set[$w] = $true }
    $distinct = @{}
    foreach ($w in $PhraseWords) { $distinct[$w] = $true }
    $hit = 0
    foreach ($w in $distinct.Keys) { if ($set.ContainsKey($w)) { $hit++ } }
    return [int](100 * $hit / $distinct.Count)
}

function Search-PhraseTranscripts {
    param([string]$Phrase, [int]$Max)
    $hits = @()
    $page = 1
    $total = 0
    while ($hits.Count -lt $Max) {
        $size = $Max - $hits.Count
        if ($size -gt 100) { $size = 100 }
        $query = @(
            @{ type = 'DATE_RANGE'; fields = @($script:SearchDateFieldValue); startValue = $script:StartIso; endValue = $script:EndIso; dateFormat = "yyyy-MM-dd'T'HH:mm:ss.SSSX" },
            @{ type = 'EXACT'; fields = @($script:SearchLanguageFieldValue); value = $script:SearchLanguageValue },
            @{ type = $script:SearchMatchTypeValue; fields = @($script:SearchTextFieldValue); value = $Phrase }
        )
        if ($script:MediaTypeValue -ne 'all') {
            $query += , @{ type = 'EXACT'; fields = @($script:SearchMediaTypeFieldValue); value = $script:MediaTypeValue }
        }
        $body = @{ types = @('transcripts'); pageSize = $size; pageNumber = $page; sortOrder = 'SCORE'; query = $query }
        $resp = Invoke-GcApi -Method Post -Path '/api/v2/speechandtextanalytics/transcripts/search' -Body $body
        if ($null -ne $resp.total) { $total = [int]$resp.total }
        $got = 0
        foreach ($r in $resp.results) {
            if ($null -eq $r) { continue }
            $got++
            $hits += , $r
            if ($hits.Count -ge $Max) { break }
        }
        if ($got -lt $size) { break }
        if ($null -ne $resp.pageCount -and $page -ge [int]$resp.pageCount) { break }
        $page++
    }
    return @{ Total = $total; Hits = $hits }
}

# Downloads a transcript JSON (pre-signed URL: no Authorization header) via a temp file,
# so it works whatever content type the storage returns.
function Get-TranscriptJson {
    param([string]$Url)
    $tmp = Join-Path -Path $script:TempFolder -ChildPath 'gc_transcript_tmp.json'
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            Invoke-WebRequest -Uri $Url -OutFile $tmp -UseBasicParsing -ErrorAction Stop | Out-Null
            break
        }
        catch {
            $status = Get-HttpStatus $_
            if (($status -eq 429 -or $status -ge 500 -or $status -eq 0) -and $attempt -le 3) { Start-Sleep -Seconds (2 * $attempt); continue }
            throw ('Transcript download failed: HTTP {0}. {1}' -f $status, (Get-ErrorText $_))
        }
    }
    $raw = Get-Content -Path $tmp -Raw -Encoding UTF8
    Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
    return ($raw | ConvertFrom-Json)
}

# Candidate communication ids for a conversation, when the search hit doesn't carry one:
# every voice session in the analytics record, customer side first.
function Get-CommunicationIds {
    param([string]$ConversationId)
    $conv = Invoke-GcApi -Method Get -Path ('/api/v2/analytics/conversations/' + $ConversationId + '/details')
    $first = @()
    $rest = @()
    foreach ($p in $conv.participants) {
        foreach ($s in $p.sessions) {
            if (-not $s.sessionId) { continue }
            if ([string]$p.purpose -eq 'customer' -or [string]$p.purpose -eq 'external') { $first += [string]$s.sessionId }
            else { $rest += [string]$s.sessionId }
        }
    }
    return @($first + $rest)
}

function Get-TranscriptForHit {
    param([string]$ConversationId, [string]$CommunicationId)
    $ids = @()
    if ($CommunicationId) { $ids = @($CommunicationId) } else { $ids = @(Get-CommunicationIds -ConversationId $ConversationId) }
    $lastErr = 'no communication id found'
    foreach ($cid in $ids) {
        try {
            $u = Invoke-GcApi -Method Get -Path ('/api/v2/speechandtextanalytics/conversations/' + $ConversationId + '/communications/' + $cid + '/transcripturl')
            if ($null -eq $u -or -not $u.url) { $lastErr = 'transcripturl returned no url'; continue }
            return @{ CommunicationId = $cid; Transcript = (Get-TranscriptJson -Url ([string]$u.url)); Error = '' }
        }
        catch { $lastErr = [string]$_.Exception.Message }
    }
    return @{ CommunicationId = ''; Transcript = $null; Error = $lastErr }
}

# Best-matching sentence in a transcript for one phrase (optionally one speaker side).
function Find-BestSentence {
    param($Transcript, [string[]]$PhraseWords, [string]$Side)
    $best = @{ Score = -1; Text = ''; Speaker = ''; OffsetSec = '' }
    foreach ($t in $Transcript.transcripts) {
        $sentences = @(foreach ($ph in $t.phrases) { if ($null -ne $ph -and $ph.text) { $ph } })
        for ($i = 0; $i -lt $sentences.Count; $i++) {
            $ph = $sentences[$i]
            $speaker = ([string]$ph.participantPurpose).ToLower()
            if ($Side -ne 'both' -and $speaker -and $speaker -ne $Side) { continue }
            # Also try this sentence joined with the next one from the same speaker,
            # in case the phrase was split across two transcript segments.
            $candidates = @([string]$ph.text)
            if (($i + 1) -lt $sentences.Count -and ([string]$sentences[$i + 1].participantPurpose).ToLower() -eq $speaker) {
                $candidates += ([string]$ph.text + ' ' + [string]$sentences[$i + 1].text)
            }
            foreach ($c in $candidates) {
                $score = Get-MatchScore -PhraseWords $PhraseWords -Text $c
                if ($score -gt $best.Score) {
                    $offset = ''
                    $ms = Get-Prop $ph @('startTimeMs', 'offsetMs', 'startTime')
                    if ($null -ne $ms) { try { $offset = [string]([int]([double]$ms / 1000)) } catch { $offset = '' } }
                    $best = @{ Score = $score; Text = $c; Speaker = $speaker; OffsetSec = $offset }
                }
            }
        }
    }
    return $best
}

# ----------------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------------

$domain = Get-GcDomain $Environment
$script:ApiBase = 'https://api.' + $domain
$script:LoginBase = 'https://login.' + $domain
$script:MaxRetries = $MaxRetries
$script:CanReauth = $false
$script:ApiHeaders = @{}

# Drop config placeholders that were never filled in.
if ($ClientId -like 'PASTE-*') { $ClientId = $env:GC_CLIENT_ID }
if ($ClientSecret -like 'PASTE-*') { $ClientSecret = '' }

# --- Phrases
$phraseList = @()
if ($PhraseFile) {
    if (-not (Test-Path -Path $PhraseFile)) { throw ('Phrase file not found: {0}' -f $PhraseFile) }
    $phraseList = @(Get-CleanList -Values @(Get-Content -Path $PhraseFile))
}
else {
    $phraseList = @(Get-CleanList -Values $Phrases)
}
# De-duplicate (case-insensitive), keep first occurrence.
$seen = @{}
$phraseList = @(foreach ($p in $phraseList) { $k = $p.ToLower(); if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; $p } })
if ($phraseList.Count -eq 0) { throw 'No phrases to test. Fill in $Phrases in the CONFIG block or pass -PhraseFile.' }

$script:ProgramList = @(Get-CleanList -Values $ProgramIds -SplitCommas)
$script:QueueList = @(Get-CleanList -Values $QueueIds -SplitCommas)
$script:FlowList = @(Get-CleanList -Values $FlowIds -SplitCommas)
if ($script:ProgramList.Count -eq 0) { throw 'Fill in $ProgramIds in the CONFIG block (at least one Speech & Text Analytics program id).' }

$script:DialectValue = $Dialect
$script:MatchingTypeValue = $MatchingType
$script:ParticipantsValue = $Participants
$script:StrictnessValue = $Strictness
$script:MediaTypeValue = $MediaType

$startUtc = $StartDate.ToUniversalTime()
$endUtc = $EndDate.ToUniversalTime()
$nowUtc = (Get-Date).ToUniversalTime()
if ($endUtc -gt $nowUtc) { $endUtc = $nowUtc }
if ($endUtc -le $startUtc) { throw '-EndDate must be after -StartDate (and -StartDate must be in the past).' }
$script:StartMs = ConvertTo-EpochMs -Date $startUtc
$script:EndMs = ConvertTo-EpochMs -Date $endUtc
$script:StartIso = Format-IsoUtc $startUtc
$script:EndIso = Format-IsoUtc $endUtc
$script:SearchMatchTypeValue = $SearchMatchType
$script:SearchTextFieldValue = $SearchTextField
$script:SearchDateFieldValue = $SearchDateField
$script:SearchMediaTypeFieldValue = $SearchMediaTypeField
$script:SearchLanguageFieldValue = $SearchLanguageField
$script:SearchLanguageValue = $SearchLanguage
if (-not $script:SearchLanguageValue) { $script:SearchLanguageValue = $Dialect }
$script:TempFolder = $env:TEMP
if (-not $script:TempFolder) { $script:TempFolder = $OutputFolder }
$script:CurlPathValue = $CurlPath
$script:ProxyUrlValue = $ProxyUrl
$script:ListenerMaxSecondsValue = $ListenerMaxSeconds
$script:KeepListenerLogValue = [bool]$KeepListenerLog
$script:ListenerProc = $null
$script:ChannelId = ''
$script:JobResults = @{}
$rs = Get-Date
$script:RunStamp = (Get-Padded $rs.Year 4) + (Get-Padded $rs.Month 2) + (Get-Padded $rs.Day 2) + '_' + (Get-Padded $rs.Hour 2) + (Get-Padded $rs.Minute 2) + (Get-Padded $rs.Second 2)
if (-not (Test-Path -Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }

# --- Authentication
if ($AccessToken) {
    $script:ApiHeaders = @{ Authorization = ('Bearer ' + $AccessToken.Trim()) }
}
else {
    if (-not $ClientId) { throw 'Fill in $ClientId in the CONFIG block at the top of the script (or pass -ClientId / -AccessToken).' }
    if (-not $ClientSecret) { $ClientSecret = $env:GC_CLIENT_SECRET }
    if (-not $ClientSecret) {
        $cred = Get-Credential -UserName $ClientId -Message 'Genesys Cloud OAuth client secret (enter it as the password)'
        if ($null -eq $cred) { throw 'No client secret supplied.' }
        $ClientSecret = $cred.GetNetworkCredential().Password
    }
    $script:OAuthClientId = $ClientId.Trim()
    $script:OAuthClientSecret = $ClientSecret.Trim()
    Get-GcToken
    $script:CanReauth = $true
}

Write-Host ''
Write-Host ('Transcripts (local) : {0}  ->  {1}   mediaType={2}   search language={3}' -f (Format-Local $startUtc), (Format-Local $endUtc), $MediaType, $script:SearchLanguageValue)
Write-Host ('Programs            : {0}' -f ($script:ProgramList -join ', '))
if ($script:QueueList.Count -gt 0) { Write-Host ('Queues              : {0}' -f ($script:QueueList -join ', ')) }
if ($script:FlowList.Count -gt 0) { Write-Host ('Flows               : {0}' -f ($script:FlowList -join ', ')) }
Write-Host ('Topic settings      : dialect={0} matching={1} participants={2} strictness={3}' -f $Dialect, $MatchingType, $Participants, $Strictness)
Write-Host ('Phrases             : {0}   (up to {1} jobs at a time)' -f $phraseList.Count, $MaxConcurrentJobs)
Write-Host ''

# --- Notification listener (matched conversations)
$listening = $false
if (-not $CountsOnly) {
    $curlCheck = Test-CurlWebSocket
    if ($curlCheck.Ok) {
        try {
            Start-NotificationListener
            $listening = $true
        }
        catch {
            Write-Warning ('Could not start the notification listener: {0}' -f $_.Exception.Message)
        }
    }
    else {
        Write-Warning ('Matched conversations unavailable: {0}.' -f $curlCheck.Reason)
    }
    if (-not $listening) {
        Write-Warning 'Continuing with counts only. Use -TranscriptSearch for a wording-based list of conversations instead.'
    }
    Write-Host ''
}

# --- Run the jobs: keep up to MaxConcurrentJobs in flight, poll until each finishes.
$results = @()
$pending = @()       # hashtables: Phrase, JobId, Submitted, LastState, Row
$nextIndex = 0
$done = 0

while ($nextIndex -lt $phraseList.Count -or $pending.Count -gt 0) {
    # Submit more jobs while there is room.
    while ($pending.Count -lt $MaxConcurrentJobs -and $nextIndex -lt $phraseList.Count) {
        $phrase = $phraseList[$nextIndex]
        $nextIndex++
        $row = @{
            Phrase              = $phrase
            JobId               = ''
            State               = ''
            ProcessedTranscripts = ''
            MatchedTranscripts  = ''
            MatchPct            = ''
            SecondsToComplete   = ''
            MatchedConversationsReceived = ''
            FinishedAt          = $null
            Error               = ''
        }
        try {
            $job = Submit-PhraseJob -Phrase $phrase
            $row.JobId = [string]$job.id
            $row.State = [string]$job.state
            Write-Host ('[{0}/{1}] submitted  {2}  "{3}"' -f $nextIndex, $phraseList.Count, $row.JobId, $phrase)
            if ($listening) {
                try { Add-JobSubscription -JobId $row.JobId }
                catch { Write-Warning ('Could not subscribe to notifications for job {0}: {1}' -f $row.JobId, $_.Exception.Message) }
            }
            $pending += , @{ Phrase = $phrase; JobId = $row.JobId; Submitted = (Get-Date); LastState = [string]$job.state; Row = $row }
        }
        catch {
            $row.State = 'SubmitFailed'
            $row.Error = [string]$_.Exception.Message
            Write-Warning ('[{0}/{1}] submit failed for "{2}": {3}' -f $nextIndex, $phraseList.Count, $phrase, $row.Error)
            $results += , $row
            $done++
        }
    }
    if ($pending.Count -eq 0) { continue }

    Start-Sleep -Seconds $PollSeconds

    $stillPending = @()
    foreach ($p in $pending) {
        $row = $p.Row
        $status = $null
        try {
            $status = Invoke-GcApi -Method Get -Path ('/api/v2/speechandtextanalytics/topics/testphrase/jobs/' + $p.JobId)
        }
        catch {
            $row.State = 'PollFailed'
            $row.Error = [string]$_.Exception.Message
            Write-Warning ('job {0} poll failed: {1}' -f $p.JobId, $row.Error)
            $results += , $row
            $done++
            continue
        }
        $state = ''
        if ($null -ne $status -and $null -ne $status.state) { $state = [string]$status.state }
        $row.State = $state
        if ($null -ne $status.processedTranscriptsCount) { $row.ProcessedTranscripts = [string]([int]$status.processedTranscriptsCount) }
        if ($null -ne $status.matchedTranscriptsCount) { $row.MatchedTranscripts = [string]([int]$status.matchedTranscriptsCount) }

        $elapsed = ((Get-Date) - $p.Submitted).TotalSeconds
        if (Test-JobFinished -State $state) {
            $row.SecondsToComplete = [string]([int]$elapsed)
            $row.FinishedAt = Get-Date
            if (Test-JobFailed -State $state) {
                $row.Error = ('Job ended in state {0}' -f $state)
                Write-Warning ('job {0} "{1}" ended in state {2}' -f $p.JobId, $p.Phrase, $state)
            }
            else {
                $row.MatchPct = Format-Pct $row.MatchedTranscripts $row.ProcessedTranscripts
                Write-Host ('    finished   {0}  matched {1} of {2}  ({3}s)  "{4}"' -f $p.JobId, $row.MatchedTranscripts, $row.ProcessedTranscripts, [int]$elapsed, $p.Phrase)
            }
            $results += , $row
            $done++
        }
        elseif ($elapsed -gt $JobTimeoutSeconds) {
            $row.Error = ('Timed out after {0}s in state {1}' -f [int]$elapsed, $state)
            Write-Warning ('job {0} "{1}" timed out (state {2})' -f $p.JobId, $p.Phrase, $state)
            $results += , $row
            $done++
        }
        else {
            if ($state -ne $p.LastState) {
                Write-Verbose ('job {0} state {1} -> {2}' -f $p.JobId, $p.LastState, $state)
                $p.LastState = $state
            }
            $stillPending += , $p
        }
    }
    $pending = $stillPending
    if ($listening) { Read-JobNotifications }
}

# --- Wait for the notifications of finished jobs, then build the matched-conversation rows
$matchRows = @()
if ($listening) {
    while ($true) {
        Read-JobNotifications
        $waiting = @(foreach ($r in $results) {
                if (-not $r.JobId -or $null -eq $r.FinishedAt -or $r.Error) { continue }
                if ($r.MatchedTranscripts -eq '' -or [int]$r.MatchedTranscripts -eq 0) { continue }
                if ($script:JobResults.ContainsKey($r.JobId) -and $script:JobResults[$r.JobId].Transcripts.Count -ge [int]$r.MatchedTranscripts) { continue }
                if (((Get-Date) - $r.FinishedAt).TotalSeconds -gt $NotificationWaitSeconds) { continue }
                $r
            })
        if ($waiting.Count -eq 0) { break }
        if ($script:ListenerProc.HasExited) {
            Write-Warning 'The curl listener process exited before all notifications arrived.'
            break
        }
        Start-Sleep -Seconds $PollSeconds
    }
    Stop-NotificationListener

    foreach ($r in $results) {
        if (-not $r.JobId) { continue }
        $entry = $null
        if ($script:JobResults.ContainsKey($r.JobId)) { $entry = $script:JobResults[$r.JobId] }
        if ($null -eq $entry) {
            $r.MatchedConversationsReceived = '0'
            if ($r.MatchedTranscripts -ne '' -and [int]$r.MatchedTranscripts -gt 0 -and -not $r.Error) {
                $r.Error = ('No notification received within {0}s of the job finishing' -f $NotificationWaitSeconds)
            }
            continue
        }
        $r.MatchedConversationsReceived = [string]$entry.Transcripts.Count
        foreach ($t in $entry.Transcripts) {
            $when = ''
            if ($null -ne $t.timestamp) {
                try {
                    $ms = [double]$t.timestamp
                    $when = Format-Local (([datetime]'1970-01-01').AddMilliseconds($ms))
                } catch { $when = [string]$t.timestamp }
            }
            $detected = @($t.detectedPhrases)
            if ($detected.Count -eq 0) { $detected = @($null) }
            foreach ($d in $detected) {
                $matchRows += , @{
                    Phrase           = $r.Phrase
                    ConversationId   = [string]$t.conversationId
                    ConversationTime = $when
                    MediaType        = [string]$t.mediaType
                    FoundPhrase      = $(if ($null -ne $d) { [string]$d.foundPhrase } else { '' })
                    Snippet          = $(if ($null -ne $d) { [string]$d.snippet } else { '' })
                    Confidence       = $(if ($null -ne $d -and $null -ne $d.confidence) { [string]$d.confidence } else { '' })
                    CommunicationId  = [string]$t.communicationId
                    TranscriptId     = [string]$t.transcriptId
                    JobId            = $r.JobId
                }
            }
        }
    }
}

# --- Optional: wording-based transcript search
$searchRows = @()
$searchTotals = @{}
if ($TranscriptSearch) {
    if (-not (Test-Path -Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
    $side = $Participants.ToLower()
    Write-Host ''
    Write-Host ('Finding matched conversations (transcript search, {0}, up to {1} per phrase)...' -f $SearchMatchType, $MaxMatchesPerPhrase)
    $transcriptCache = @{}
    $searchFailures = 0
    $pi = 0
    foreach ($phrase in $phraseList) {
        $pi++
        $search = $null
        try { $search = Search-PhraseTranscripts -Phrase $phrase -Max $MaxMatchesPerPhrase }
        catch {
            Write-Warning ('[{0}/{1}] transcript search failed for "{2}": {3}' -f $pi, $phraseList.Count, $phrase, $_.Exception.Message)
            if ($_.Exception.Message -like '*HTTP 400*') {
                $msg = [string]$_.Exception.Message
                if ($msg -match '"errorCode"\s*:\s*"([A-Z_]+)"\s*,\s*"fieldName"\s*:\s*"([^"]+)"') {
                    Write-Warning ('Transcript search rejected the query: {0} on field "{1}". Adjust the matching -Search*Field / -SearchLanguage parameter.' -f $matches[1], $matches[2])
                }
                else {
                    Write-Warning 'HTTP 400 from transcript search usually means a field name or value is wrong: see -SearchTextField / -SearchDateField / -SearchMediaTypeField / -SearchLanguage.'
                }
                if ($searchFailures -ge 2) {
                    Write-Warning 'Transcript search failed for 3 phrases in a row; skipping the matched-conversation step for the remaining phrases.'
                    break
                }
                $searchFailures++
            }
            $searchTotals[$phrase.ToLower()] = 'error'
            continue
        }
        $searchFailures = 0
        $searchTotals[$phrase.ToLower()] = [string]$search.Total
        Write-Host ('[{0}/{1}] {2} search hits (pulling {3})  "{4}"' -f $pi, $phraseList.Count, $search.Total, $search.Hits.Count, $phrase)
        $phraseWords = @(Get-Words $phrase)
        foreach ($h in $search.Hits) {
            $convId = [string](Get-Prop $h @('conversationId', 'conversation.id', 'conversation_id'))
            $commId = [string](Get-Prop $h @('communicationId', 'communication.id', 'communication_id'))
            $convStart = [string](Get-Prop $h @('conversationStartTime', 'startTime', 'conversationStart'))
            $row = @{
                Phrase             = $phrase
                ConversationId     = $convId
                ConversationStart  = ''
                CommunicationId    = $commId
                Speaker            = ''
                DetectedTranscript = ''
                MatchScorePct      = ''
                OffsetSec          = ''
                Error              = ''
            }
            if ($convStart) { try { $row.ConversationStart = Format-Local ([datetime]$convStart) } catch { $row.ConversationStart = $convStart } }
            if (-not $convId) {
                $row.Error = 'Search hit had no conversation id'
                $searchRows += , $row
                continue
            }
            $key = $convId + '|' + $commId
            if (-not $transcriptCache.ContainsKey($key)) { $transcriptCache[$key] = Get-TranscriptForHit -ConversationId $convId -CommunicationId $commId }
            $tr = $transcriptCache[$key]
            if (-not $row.CommunicationId) { $row.CommunicationId = $tr.CommunicationId }
            if ($null -eq $tr.Transcript) {
                $row.Error = $tr.Error
            }
            else {
                $best = Find-BestSentence -Transcript $tr.Transcript -PhraseWords $phraseWords -Side $side
                if ($best.Score -ge $MinMatchScore) {
                    $row.Speaker = $best.Speaker
                    $row.DetectedTranscript = $best.Text
                    $row.MatchScorePct = [string]$best.Score
                    $row.OffsetSec = $best.OffsetSec
                }
                elseif ($best.Score -ge 0) {
                    $row.Error = ('Best sentence only {0}% similar (below -MinMatchScore {1}): {2}' -f $best.Score, $MinMatchScore, $best.Text)
                }
                else {
                    $row.Error = ('No {0} sentences in transcript' -f $side)
                }
            }
            $searchRows += , $row
        }
    }
}

# --- Output
foreach ($r in $results) {
    $k = $r.Phrase.ToLower()
    if ($searchTotals.ContainsKey($k)) { $r.SearchHits = $searchTotals[$k] } else { $r.SearchHits = '' }
}
$columns = @('Phrase', 'MatchedTranscripts', 'ProcessedTranscripts', 'MatchPct', 'MatchedConversationsReceived', 'SearchHits', 'State', 'SecondsToComplete', 'JobId', 'Error')
$rows = @(foreach ($r in $results) { New-Object PSObject -Property $r })

# Sort: most matches first, then by phrase.
$sorted = @($rows | Sort-Object -Property @{ Expression = { if ($_.MatchedTranscripts -eq '') { -1 } else { [int]$_.MatchedTranscripts } }; Descending = $true }, Phrase)

# Export-Csv with no rows writes only a byte-order mark, which Excel shows as "ï»¿",
# so an empty result still gets a header row.
function Write-CsvRows {
    param([string]$Path, $Rows, [string[]]$Columns)
    if (@($Rows).Count -gt 0) {
        @(foreach ($m in $Rows) { New-Object PSObject -Property $m }) | Select-Object $Columns |
            Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    }
    else {
        Set-Content -Path $Path -Value ('"' + ($Columns -join '","') + '"') -Encoding UTF8
    }
}

$stamp = $script:RunStamp
$csvPath = Join-Path -Path $OutputFolder -ChildPath ('TopicPhraseTest_' + $stamp + '.csv')
$sorted | Select-Object $columns | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
$matchPath = ''
if ($listening) {
    $matchPath = Join-Path -Path $OutputFolder -ChildPath ('TopicPhraseMatches_' + $stamp + '.csv')
    $matchColumns = @('Phrase', 'ConversationId', 'ConversationTime', 'MediaType', 'FoundPhrase', 'Snippet', 'Confidence', 'CommunicationId', 'TranscriptId', 'JobId')
    $sortedMatches = @($matchRows | Sort-Object { $_.Phrase }, { $_.ConversationTime })
    Write-CsvRows -Path $matchPath -Rows $sortedMatches -Columns $matchColumns
}
$searchPath = ''
if ($TranscriptSearch) {
    $searchPath = Join-Path -Path $OutputFolder -ChildPath ('TopicPhraseSearch_' + $stamp + '.csv')
    $searchColumns = @('Phrase', 'ConversationId', 'ConversationStart', 'Speaker', 'DetectedTranscript', 'MatchScorePct', 'OffsetSec', 'CommunicationId', 'Error')
    Write-CsvRows -Path $searchPath -Rows $searchRows -Columns $searchColumns
}

Write-Host ''
Write-Host '================ Topic phrase test results ================'
$sorted | Select-Object MatchedTranscripts, MatchedConversationsReceived, ProcessedTranscripts, MatchPct, State, Phrase | Format-Table -AutoSize -Wrap | Out-String -Width 200 | Write-Host
$failed = @($rows | Where-Object { $_.Error })
if ($failed.Count -gt 0) { Write-Warning ('{0} phrase(s) have an error; see the Error column.' -f $failed.Count) }
if ($listening) {
    $convCount = @($matchRows | ForEach-Object { $_.ConversationId } | Sort-Object -Unique).Count
    Write-Host ('Matched conversations received from Genesys: {0} rows, {1} distinct conversations.' -f $matchRows.Count, $convCount)
    if ($script:KeepListenerLogValue) { Write-Host ('Raw notification frames: {0}' -f $script:ListenerOut) }
}
if ($TranscriptSearch) {
    $found = @($searchRows | Where-Object { $_.DetectedTranscript }).Count
    Write-Host ('Transcript search (wording-based): {0} conversations with a detected sentence, {1} rows in total.' -f $found, $searchRows.Count)
}
Write-Host ('Counts CSV : {0}' -f $csvPath)
if ($matchPath) { Write-Host ('Matches CSV: {0}' -f $matchPath) }
if ($searchPath) { Write-Host ('Search CSV : {0}' -f $searchPath) }
