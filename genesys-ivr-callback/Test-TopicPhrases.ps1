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

    Output: a CSV with one row per phrase (matched, processed, match %, job id, state)
    plus a console table sorted by matches.

    Constrained Language Mode safe (Windows PowerShell 5.1 under AppLocker/WDAC):
    no .NET static calls, no ::new(), no [pscustomobject] casts, no Add-Type.

    OAuth client (Client Credentials grant) role permission:
      speechAndTextAnalytics:topic:testPhrase

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
            $hint = ' -- the OAuth client role needs speechAndTextAnalytics:topic:testPhrase and access to the divisions the programs are in.'
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
Write-Host ('Transcripts (local) : {0}  ->  {1}   mediaType={2}' -f (Format-Local $startUtc), (Format-Local $endUtc), $MediaType)
Write-Host ('Programs            : {0}' -f ($script:ProgramList -join ', '))
if ($script:QueueList.Count -gt 0) { Write-Host ('Queues              : {0}' -f ($script:QueueList -join ', ')) }
if ($script:FlowList.Count -gt 0) { Write-Host ('Flows               : {0}' -f ($script:FlowList -join ', ')) }
Write-Host ('Topic settings      : dialect={0} matching={1} participants={2} strictness={3}' -f $Dialect, $MatchingType, $Participants, $Strictness)
Write-Host ('Phrases             : {0}   (up to {1} jobs at a time)' -f $phraseList.Count, $MaxConcurrentJobs)
Write-Host ''

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
            Error               = ''
        }
        try {
            $job = Submit-PhraseJob -Phrase $phrase
            $row.JobId = [string]$job.id
            $row.State = [string]$job.state
            Write-Host ('[{0}/{1}] submitted  {2}  "{3}"' -f $nextIndex, $phraseList.Count, $row.JobId, $phrase)
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
}

# --- Output
$columns = @('Phrase', 'MatchedTranscripts', 'ProcessedTranscripts', 'MatchPct', 'State', 'SecondsToComplete', 'JobId', 'Error')
$rows = @(foreach ($r in $results) { New-Object PSObject -Property $r })

# Sort: most matches first, then by phrase.
$sorted = @($rows | Sort-Object -Property @{ Expression = { if ($_.MatchedTranscripts -eq '') { -1 } else { [int]$_.MatchedTranscripts } }; Descending = $true }, Phrase)

if (-not (Test-Path -Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
$n = Get-Date
$stamp = (Get-Padded $n.Year 4) + (Get-Padded $n.Month 2) + (Get-Padded $n.Day 2) + '_' + (Get-Padded $n.Hour 2) + (Get-Padded $n.Minute 2) + (Get-Padded $n.Second 2)
$csvPath = Join-Path -Path $OutputFolder -ChildPath ('TopicPhraseTest_' + $stamp + '.csv')
$sorted | Select-Object $columns | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host '================ Topic phrase test results ================'
$sorted | Select-Object MatchedTranscripts, ProcessedTranscripts, MatchPct, State, Phrase | Format-Table -AutoSize -Wrap | Out-String -Width 200 | Write-Host
$failed = @($rows | Where-Object { $_.Error })
if ($failed.Count -gt 0) { Write-Warning ('{0} phrase(s) did not complete; see the Error column.' -f $failed.Count) }
Write-Host ('CSV: {0}' -f $csvPath)
