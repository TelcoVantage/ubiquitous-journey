#Requires -Version 5.1
<#
.SYNOPSIS
    Genesys Cloud: success rate of the IVR "hang up and call back the authenticated
    way" message. Finds callers who ended their call in the IVR of the given inbound
    call flows and checks whether the same number called back within N minutes.

.DESCRIPTION
    1. Pulls every inbound voice conversation that went through the source flows in
       the reporting period (POST /api/v2/analytics/conversations/details/query,
       filtered server side on flowId).
    2. Keeps the calls that ENDED IN THE IVR: no queue (acd), no agent, no external /
       voicemail / group transfer. These are the "hung up" calls.
    3. For each caller number (ANI) looks for another inbound call from the same ANI
       that started within -CallbackWindowMinutes after the first call ended
       (second analytics query, filtered server side on ani, in batches).
    4. Decides whether the callback came in "the authenticated way":
         route  : callback went through -CallbackFlowIds/-CallbackFlowNames OR hit -CallbackDnis
         proof  : callback flow outcome -AuthOutcomeId = SUCCESS OR participant data
                  -AuthAttributeName (= -AuthAttributeValue)
         qualified = route AND proof   (a group with no criteria supplied is ignored;
                                        no criteria at all = any callback qualifies)
    5. Writes a per-call detail CSV and a summary CSV (overall / per flow / per day)
       and prints the headline rates.

    Constrained Language Mode safe (Windows PowerShell 5.1 under AppLocker/WDAC):
    no .NET static calls, no ::new(), no [pscustomobject] casts, no Add-Type.
    Base64 for the OAuth Basic header is done with bit operators.

    OAuth client (Client Credentials grant) role permissions:
      analytics:conversationDetail:view   always
      architect:flow:view                 when you pass flow NAMES
      conversation:communication:view     when you use -AuthAttributeName

.EXAMPLE
    .\Get-IvrHangupCallbackRate.ps1 -Environment mypurecloud.com.au -ClientId 'xxxx' `
        -SourceFlowNames 'Main Inbound','Billing Inbound' `
        -StartDate '2026-09-01' -EndDate '2026-09-08'

.EXAMPLE
    # Success = called back within 30 min on the authenticated number and passed ID&V
    .\Get-IvrHangupCallbackRate.ps1 -Environment mypurecloud.ie -ClientId 'xxxx' `
        -SourceFlowIds 'a1b2...','c3d4...' -CallbackDnis '+35315550100' `
        -AuthOutcomeId 'e5f6...' -CustomerHangupOnly
#>
[CmdletBinding()]
param(
    # Region domain: mypurecloud.com, mypurecloud.ie, mypurecloud.de, mypurecloud.com.au,
    # mypurecloud.jp, usw2.pure.cloud, cac1.pure.cloud, euw2.pure.cloud, aps1.pure.cloud, ...
    [string]$Environment = 'mypurecloud.com',

    # Client Credentials OAuth client. Secret falls back to $env:GC_CLIENT_SECRET, then a prompt.
    [string]$ClientId = $env:GC_CLIENT_ID,
    [string]$ClientSecret,
    # Or an existing bearer token (skips OAuth).
    [string]$AccessToken,

    # The inbound call flows that play the "hang up and call back" message.
    [string[]]$SourceFlowIds,
    [string[]]$SourceFlowNames,

    # Reporting period for the first (hung-up) call. Local machine time unless the value
    # ends in Z. EndDate is exclusive.
    [datetime]$StartDate = (Get-Date).Date.AddDays(-7),
    [datetime]$EndDate = (Get-Date).Date,

    [ValidateRange(1, 1440)]
    [int]$CallbackWindowMinutes = 30,

    # --- "Authenticated way" criteria for the callback (all optional) ---
    [string[]]$CallbackFlowIds,
    [string[]]$CallbackFlowNames,
    [string[]]$CallbackDnis,
    [string]$AuthOutcomeId,
    [string]$AuthAttributeName,
    [string]$AuthAttributeValue,

    # --- Which IVR-ended calls count as "hung up after the message" ---
    # Only count calls where the customer hung up (customer disconnectType = endpoint).
    [switch]$CustomerHangupOnly,
    # Only count calls where the source flow reached this flow outcome (put a
    # "Set Flow Outcome" right after the message plays).
    [string]$MessageOutcomeId,
    # Ignore calls shorter than this many seconds (hung up before hearing the message).
    [int]$MinIvrSeconds = 0,

    # AniFilter: query callbacks by the exact ANI strings seen (fast).
    # AllInbound: pull every inbound voice call and match on digits (slow, format tolerant).
    [ValidateSet('AniFilter', 'AllInbound')]
    [string]$CallbackSearch = 'AniFilter',
    # 0 = compare all digits of the ANI. e.g. 9 = compare the last 9 digits only
    # (tolerates +61 / 0 style prefix differences in AllInbound mode).
    [ValidateRange(0, 15)]
    [int]$AniMatchDigits = 0,

    [ValidateRange(1, 168)]
    [int]$ChunkHours = 24,
    [ValidateRange(1, 100)]
    [int]$PageSize = 100,
    [ValidateRange(1, 100)]
    [int]$AniBatchSize = 50,
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

# ISO-8601 UTC built from components: culture and calendar independent.
function Format-IsoUtc {
    param([datetime]$Date)
    $u = $Date.ToUniversalTime()
    return (Get-Padded $u.Year 4) + '-' + (Get-Padded $u.Month 2) + '-' + (Get-Padded $u.Day 2) + 'T' +
           (Get-Padded $u.Hour 2) + ':' + (Get-Padded $u.Minute 2) + ':' + (Get-Padded $u.Second 2) + '.' +
           (Get-Padded $u.Millisecond 3) + 'Z'
}

function Format-Local {
    param($Date)
    if ($null -eq $Date) { return '' }
    $l = ([datetime]$Date).ToLocalTime()
    return (Get-Padded $l.Year 4) + '-' + (Get-Padded $l.Month 2) + '-' + (Get-Padded $l.Day 2) + ' ' +
           (Get-Padded $l.Hour 2) + ':' + (Get-Padded $l.Minute 2) + ':' + (Get-Padded $l.Second 2)
}

function Format-LocalDay {
    param($Date)
    $l = ([datetime]$Date).ToLocalTime()
    return (Get-Padded $l.Year 4) + '-' + (Get-Padded $l.Month 2) + '-' + (Get-Padded $l.Day 2)
}

# Analytics timestamps are strings in PS 5.1 (ConvertFrom-Json leaves ISO dates alone)
# and DateTime in PS 7. Normalise both to a UTC DateTime.
function ConvertTo-UtcDate {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    $s = [string]$Value
    if ($s.Length -eq 0) { return $null }
    try { return ([datetime]$s).ToUniversalTime() } catch { return $null }
}

# Round and return as an invariant-culture string (PowerShell [string] casts are
# culture invariant, so CSV numbers always use '.' as the decimal separator).
function Format-Number {
    param($Value, [int]$Decimals = 1)
    if ($null -eq $Value) { return '' }
    $m = 1
    for ($k = 0; $k -lt $Decimals; $k++) { $m = $m * 10 }
    return [string]([double]([long]([double]$Value * $m)) / $m)
}

function Format-Pct {
    param([int]$Part, [int]$Whole)
    if ($Whole -le 0) { return '' }
    return Format-Number -Value (100.0 * $Part / $Whole) -Decimals 1
}

function Get-Median {
    param([double[]]$Values)
    $sorted = @($Values | Sort-Object)
    $n = $sorted.Count
    if ($n -eq 0) { return $null }
    $mid = $n -shr 1
    if (($n % 2) -eq 1) { return $sorted[$mid] }
    return ($sorted[$mid - 1] + $sorted[$mid]) / 2
}

function Get-Average {
    param([double[]]$Values)
    $n = @($Values).Count
    if ($n -eq 0) { return $null }
    $sum = 0.0
    foreach ($v in $Values) { $sum += $v }
    return $sum / $n
}

# Phone number -> digits key. '' for withheld/anonymous/too short.
function Get-DigitKey {
    param([string]$Address, [int]$LastDigits = 0)
    if (-not $Address) { return '' }
    $a = $Address
    $at = $a.IndexOf('@')
    if ($at -gt 0) { $a = $a.Substring(0, $at) }
    $semi = $a.IndexOf(';')
    if ($semi -gt 0) { $a = $a.Substring(0, $semi) }
    $digits = $a -replace '[^0-9]', ''
    if ($digits.Length -lt 6) { return '' }
    if ($LastDigits -gt 0 -and $digits.Length -gt $LastDigits) {
        $digits = $digits.Substring($digits.Length - $LastDigits)
    }
    return $digits
}

# DNIS compare: equal digits, or same last 9 digits (tolerates +CC vs national 0 prefix).
function Test-NumberMatch {
    param([string]$A, [string]$B)
    if (-not $A -or -not $B) { return $false }
    if ($A -eq $B) { return $true }
    if ($A.Length -ge 9 -and $B.Length -ge 9) {
        return ($A.Substring($A.Length - 9) -eq $B.Substring($B.Length - 9))
    }
    return $false
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
            $hint = ' -- the OAuth client role needs analytics:conversationDetail:view (plus architect:flow:view for flow names, conversation:communication:view for -AuthAttributeName) and access to the divisions involved.'
        }
        throw ('Genesys API {0} {1} failed: HTTP {2}. {3}{4}' -f $Method.ToUpper(), $Path, $status, $errText, $hint)
    }
}

function Get-GcInboundCallFlows {
    $page = 1
    $pageCount = 1
    do {
        $resp = Invoke-GcApi -Method Get -Path ('/api/v2/flows?type=inboundcall&pageSize=100&pageNumber=' + $page)
        foreach ($f in $resp.entities) { $f }
        if ($null -ne $resp.pageCount) { $pageCount = [int]$resp.pageCount } else { $pageCount = $page }
        $page++
    } while ($page -le $pageCount)
}

# Returns hashtable flowId(lower) -> flow name ('' when only the id was supplied).
function Resolve-FlowSet {
    param([string[]]$Ids, [string[]]$Names, [string]$Label)
    $set = @{}
    foreach ($id in @(Get-CleanList -Values $Ids -SplitCommas)) { $set[$id.ToLower()] = '' }
    $nameList = @(Get-CleanList -Values $Names)
    if ($nameList.Count -eq 0) { return $set }

    if ($null -eq $script:FlowCache) {
        Write-Host 'Loading inbound call flow list (architect:flow:view)...'
        $script:FlowCache = @(Get-GcInboundCallFlows)
    }
    foreach ($name in $nameList) {
        $hits = @(foreach ($f in $script:FlowCache) { if ([string]$f.name -eq $name) { $f } })
        if ($hits.Count -eq 0) {
            $needle = $name.ToLower()
            $similar = @(foreach ($f in $script:FlowCache) { if (([string]$f.name).ToLower().Contains($needle)) { [string]$f.name } })
            $msg = ('{0} flow "{1}" was not found among {2} inbound call flows.' -f $Label, $name, $script:FlowCache.Count)
            if ($similar.Count -gt 0) { $msg += ' Did you mean: ' + (($similar | Select-Object -First 10) -join ' | ') }
            throw $msg
        }
        foreach ($h in $hits) { $set[([string]$h.id).ToLower()] = [string]$h.name }
    }
    return $set
}

# ----------------------------------------------------------------------------------
# Analytics conversation details -> compact per-conversation facts
# ----------------------------------------------------------------------------------

function Get-ConversationFacts {
    param($Conversation, [hashtable]$SourceFlowSet)

    $participants = @(foreach ($p in $Conversation.participants) { if ($null -ne $p) { $p } })

    # The caller: purpose 'customer', else the first 'external'.
    $customer = $null
    foreach ($p in $participants) { if ([string]$p.purpose -eq 'customer') { $customer = $p; break } }
    if ($null -eq $customer) {
        foreach ($p in $participants) { if ([string]$p.purpose -eq 'external') { $customer = $p; break } }
    }
    $customerId = ''
    if ($null -ne $customer) { $customerId = [string]$customer.participantId }

    $aniRaw = ''
    $dnisRaw = ''
    $otherAni = ''
    $otherDnis = ''
    $custDisc = ''
    $custDiscEnd = $null
    $reachedQueue = $false
    $answered = $false
    $otherEndpoint = $false
    $answerMs = $null
    $abandonMs = $null
    $acdMs = $null
    $flowIds = @()
    $flowNames = @()
    $srcId = ''
    $srcName = ''
    $srcExit = ''
    $outcomes = @{}

    foreach ($p in $participants) {
        $purpose = [string]$p.purpose
        $isCustomer = ($customerId -and [string]$p.participantId -eq $customerId)
        if ($purpose -eq 'acd') { $reachedQueue = $true }
        if (-not $isCustomer -and @('external', 'voicemail', 'group', 'station') -contains $purpose) { $otherEndpoint = $true }

        foreach ($s in $p.sessions) {
            if ($null -eq $s) { continue }

            if ($isCustomer) {
                if (-not $aniRaw -and $s.ani) { $aniRaw = [string]$s.ani }
                if (-not $dnisRaw -and $s.dnis) { $dnisRaw = [string]$s.dnis }
                # Who hung up: the last customer segment that carries a disconnectType.
                foreach ($seg in $s.segments) {
                    if ($null -eq $seg -or -not $seg.disconnectType) { continue }
                    $segEnd = ConvertTo-UtcDate $seg.segmentEnd
                    if ($null -eq $custDiscEnd -or ($null -ne $segEnd -and $segEnd -ge $custDiscEnd)) {
                        $custDiscEnd = $segEnd
                        $custDisc = [string]$seg.disconnectType
                    }
                }
            }
            else {
                if (-not $otherAni -and $s.ani) { $otherAni = [string]$s.ani }
                if (-not $otherDnis -and $s.dnis) { $otherDnis = [string]$s.dnis }
            }

            if ($purpose -eq 'agent' -or $purpose -eq 'user') {
                foreach ($seg in $s.segments) {
                    if ($null -ne $seg -and [string]$seg.segmentType -eq 'interact') { $answered = $true }
                }
            }

            foreach ($m in $s.metrics) {
                if ($null -eq $m) { continue }
                $mn = [string]$m.name
                if ($mn -eq 'tAnswered' -and $null -eq $answerMs) { $answerMs = [double]$m.value; $answered = $true }
                elseif ($mn -eq 'tAbandon' -and $null -eq $abandonMs) { $abandonMs = [double]$m.value }
                elseif ($mn -eq 'tAcd' -and $null -eq $acdMs) { $acdMs = [double]$m.value }
            }

            $flow = $s.flow
            if ($null -ne $flow -and $flow.flowId) {
                $fid = ([string]$flow.flowId).ToLower()
                $fname = [string]$flow.flowName
                if ($flowIds -notcontains $fid) {
                    $flowIds += $fid
                    $flowNames += $fname
                }
                if (-not $srcId -and $SourceFlowSet.ContainsKey($fid)) {
                    $srcId = $fid
                    $srcName = $fname
                    if (-not $srcName) { $srcName = [string]$SourceFlowSet[$fid] }
                    $srcExit = [string]$flow.exitReason
                }
                foreach ($o in $flow.outcomes) {
                    if ($null -eq $o -or -not $o.flowOutcomeId) { continue }
                    $ok = ([string]$o.flowOutcomeId).ToLower()
                    $ov = [string]$o.flowOutcomeValue
                    if (-not $outcomes.ContainsKey($ok) -or $ov -eq 'SUCCESS') { $outcomes[$ok] = $ov }
                }
            }
        }
    }

    if (-not $aniRaw) { $aniRaw = $otherAni }
    if (-not $dnisRaw) { $dnisRaw = $otherDnis }

    $start = ConvertTo-UtcDate $Conversation.conversationStart
    $end = ConvertTo-UtcDate $Conversation.conversationEnd

    if ($answered) { $outcome = 'Answered' }
    elseif ($null -eq $end) { $outcome = 'InProgress' }
    elseif ($reachedQueue) { $outcome = 'QueuedNotAnswered' }
    elseif ($otherEndpoint) { $outcome = 'TransferredOut' }
    else { $outcome = 'EndedInIvr' }

    $waitSec = $null
    if ($null -ne $answerMs) { $waitSec = $answerMs / 1000 }
    elseif ($null -ne $abandonMs) { $waitSec = $abandonMs / 1000 }
    elseif ($null -ne $acdMs) { $waitSec = $acdMs / 1000 }

    $durationSec = $null
    if ($null -ne $start -and $null -ne $end) { $durationSec = ($end - $start).TotalSeconds }

    return New-Object PSObject -Property @{
        ConversationId = [string]$Conversation.conversationId
        StartUtc       = $start
        EndUtc         = $end
        DurationSec    = $durationSec
        AniRaw         = $aniRaw
        AniKey         = (Get-DigitKey -Address $aniRaw -LastDigits $script:AniMatchDigitsValue)
        DnisRaw        = $dnisRaw
        DnisKey        = (Get-DigitKey -Address $dnisRaw)
        FlowIds        = $flowIds
        FlowPath       = ($flowNames -join ' > ')
        SourceFlowId   = $srcId
        SourceFlowName = $srcName
        SourceExit     = $srcExit
        Outcomes       = $outcomes
        CustomerDisc   = $custDisc
        CustomerHungUp = ($custDisc -eq 'endpoint')
        Outcome        = $outcome
        WaitSec        = $waitSec
    }
}

# Runs the details query over [From, To) in -ChunkHours slices, pages each slice,
# and stores compact facts in $Store (conversationId -> facts). Duplicates across
# slices overwrite, so a conversation is only counted once.
function Import-GcConversationDetails {
    param([datetime]$From, [datetime]$To, $SegmentFilters, $ConversationFilters, [hashtable]$Store, [string]$Label, [switch]$Quiet)
    $chunkStart = $From
    while ($chunkStart -lt $To) {
        $chunkEnd = $chunkStart.AddHours($script:ChunkHoursValue)
        if ($chunkEnd -gt $To) { $chunkEnd = $To }
        $interval = (Format-IsoUtc $chunkStart) + '/' + (Format-IsoUtc $chunkEnd)
        $pageNumber = 1
        $chunkCount = 0
        while ($true) {
            $body = @{
                interval = $interval
                order    = 'asc'
                orderBy  = 'conversationStart'
                paging   = @{ pageSize = $script:PageSizeValue; pageNumber = $pageNumber }
            }
            if ($null -ne $SegmentFilters) { $body['segmentFilters'] = @($SegmentFilters) }
            if ($null -ne $ConversationFilters) { $body['conversationFilters'] = @($ConversationFilters) }

            $resp = Invoke-GcApi -Method Post -Path '/api/v2/analytics/conversations/details/query' -Body $body

            $got = 0
            if ($null -ne $resp -and $null -ne $resp.conversations) {
                foreach ($c in $resp.conversations) {
                    if ($null -eq $c -or -not $c.conversationId) { continue }
                    $got++
                    $Store[[string]$c.conversationId] = Get-ConversationFacts -Conversation $c -SourceFlowSet $script:SourceFlowSet
                }
            }
            $chunkCount += $got
            $total = 0
            if ($null -ne $resp -and $null -ne $resp.totalHits) { $total = [int]$resp.totalHits }
            Write-Verbose ('{0} {1} page {2}: {3} conversations (totalHits {4})' -f $Label, $interval, $pageNumber, $got, $total)

            if ($got -lt $script:PageSizeValue) { break }
            if ($total -gt 0 -and ($pageNumber * $script:PageSizeValue) -ge $total) { break }
            $pageNumber++
        }
        if (-not $Quiet) {
            Write-Host ('  {0} {1} -> {2}: {3} conversations' -f $Label, (Format-Local $chunkStart), (Format-Local $chunkEnd), $chunkCount)
        }
        $chunkStart = $chunkEnd
    }
}

function New-DimensionPredicate {
    param([string]$Dimension, [string]$Value)
    return @{ type = 'dimension'; dimension = $Dimension; operator = 'matches'; value = $Value }
}

function Test-ConversationAttribute {
    param([string]$ConversationId, [string]$Name, [string]$Value)
    $conv = $null
    try {
        $conv = Invoke-GcApi -Method Get -Path ('/api/v2/conversations/' + $ConversationId)
    }
    catch {
        Write-Warning ('Could not read participant data for {0}: {1}' -f $ConversationId, $_.Exception.Message)
        return $null
    }
    foreach ($p in $conv.participants) {
        if ($null -eq $p -or $null -eq $p.attributes) { continue }
        foreach ($prop in $p.attributes.PSObject.Properties) {
            if ($prop.Name -eq $Name) {
                if (-not $Value -or [string]$prop.Value -eq $Value) { return $true }
            }
        }
    }
    return $false
}

function New-Batch {
    param($FirstEvent)
    $b = @{ Anis = @(); AniSet = @{}; MinEnd = $FirstEvent.EndUtc; MaxEnd = $FirstEvent.EndUtc }
    $b.Anis += $FirstEvent.AniRaw
    $b.AniSet[$FirstEvent.AniRaw] = $true
    return $b
}

function New-SummaryRow {
    param([string]$Scope, $Events, [int]$TotalCalls)
    $ivr = 0; $excluded = 0; $eligible = 0; $none = 0; $any = 0; $qual = 0; $notQual = 0; $qualAnswered = 0
    $mins = @()
    $waits = @()
    $callers = @{}
    $callersQual = @{}
    foreach ($e in $Events) {
        $ivr++
        if ($e.Status -like 'Excluded*') { $excluded++; continue }
        $eligible++
        if (-not $callers.ContainsKey($e.AniKey)) { $callers[$e.AniKey] = $false }
        if ($e.Status -eq 'NoCallback') { $none++; continue }
        $any++
        if ($e.Status -eq 'CalledBackQualified') {
            $qual++
            $callers[$e.AniKey] = $true
            $mins += [double]$e.QualMinutes
            if ($e.QualCallback.Outcome -eq 'Answered') {
                $qualAnswered++
                if ($null -ne $e.QualCallback.WaitSec) { $waits += [double]$e.QualCallback.WaitSec }
            }
        }
        else { $notQual++ }
    }
    $uniqueQual = 0
    foreach ($k in $callers.Keys) { if ($callers[$k]) { $uniqueQual++ } }

    $row = New-Object PSObject -Property @{
        Scope                          = $Scope
        CallsThroughSourceFlows        = $TotalCalls
        EndedInIvr                     = $ivr
        EndedInIvrPct                  = (Format-Pct $ivr $TotalCalls)
        Excluded                       = $excluded
        EligibleHangups                = $eligible
        NoCallback                     = $none
        CalledBackAny                  = $any
        CalledBackAnyPct               = (Format-Pct $any $eligible)
        CalledBackQualified            = $qual
        QualifiedSuccessRatePct        = (Format-Pct $qual $eligible)
        CalledBackNotQualifiedOnly     = $notQual
        QualifiedAnsweredByAgent       = $qualAnswered
        QualifiedAnsweredPct           = (Format-Pct $qualAnswered $qual)
        AvgMinutesToQualifiedCallback  = (Format-Number (Get-Average $mins) 1)
        MedianMinutesToQualifiedCallback = (Format-Number (Get-Median $mins) 1)
        MedianQualifiedAnswerWaitSec   = (Format-Number (Get-Median $waits) 0)
        UniqueCallers                  = $callers.Count
        UniqueCallersQualified         = $uniqueQual
        UniqueCallerQualifiedPct       = (Format-Pct $uniqueQual $callers.Count)
    }
    return $row
}

# ----------------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------------

$domain = Get-GcDomain $Environment
$script:ApiBase = 'https://api.' + $domain
$script:LoginBase = 'https://login.' + $domain
$script:MaxRetries = $MaxRetries
$script:ChunkHoursValue = $ChunkHours
$script:PageSizeValue = $PageSize
$script:AniMatchDigitsValue = $AniMatchDigits
$script:CanReauth = $false
$script:FlowCache = $null
$script:ApiHeaders = @{}

$startUtc = $StartDate.ToUniversalTime()
$endUtc = $EndDate.ToUniversalTime()
$nowUtc = (Get-Date).ToUniversalTime()
if ($endUtc -gt $nowUtc) { $endUtc = $nowUtc }
if ($endUtc -le $startUtc) { throw '-EndDate must be after -StartDate (and -StartDate must be in the past).' }

if (@(Get-CleanList -Values $SourceFlowIds).Count -eq 0 -and @(Get-CleanList -Values $SourceFlowNames).Count -eq 0) {
    throw 'Supply the flows that play the message with -SourceFlowIds and/or -SourceFlowNames.'
}

# --- Authentication
if ($AccessToken) {
    $script:ApiHeaders = @{ Authorization = ('Bearer ' + $AccessToken.Trim()) }
}
else {
    if (-not $ClientId) { throw 'Supply -ClientId (or $env:GC_CLIENT_ID), or -AccessToken.' }
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

# --- Flows and criteria
$script:SourceFlowSet = Resolve-FlowSet -Ids $SourceFlowIds -Names $SourceFlowNames -Label 'Source'
$routeFlowSet = Resolve-FlowSet -Ids $CallbackFlowIds -Names $CallbackFlowNames -Label 'Callback'
$routeDnisKeys = @(foreach ($d in @(Get-CleanList -Values $CallbackDnis -SplitCommas)) {
        $k = Get-DigitKey -Address $d
        if ($k) { $k } else { Write-Warning ('Ignoring callback DNIS "{0}" (fewer than 6 digits).' -f $d) }
    })
$authOutcomeKey = ([string]$AuthOutcomeId).Trim().ToLower()
$msgOutcomeKey = ([string]$MessageOutcomeId).Trim().ToLower()
$authAttrName = ([string]$AuthAttributeName).Trim()
$hasRoute = ($routeFlowSet.Count -gt 0 -or $routeDnisKeys.Count -gt 0)
$hasProof = ([bool]$authOutcomeKey -or [bool]$authAttrName)

Write-Host ''
Write-Host ('Period (local)   : {0}  ->  {1}' -f (Format-Local $startUtc), (Format-Local $endUtc))
Write-Host ('Source flows     : {0}' -f (($script:SourceFlowSet.Keys | ForEach-Object { if ($script:SourceFlowSet[$_]) { $script:SourceFlowSet[$_] + ' (' + $_ + ')' } else { $_ } }) -join '; '))
Write-Host ('Callback window  : {0} minutes after the first call ended' -f $CallbackWindowMinutes)
if ($hasRoute -or $hasProof) {
    $crit = @()
    if ($routeFlowSet.Count -gt 0) { $crit += ('callback flow in [' + (($routeFlowSet.Keys | ForEach-Object { if ($routeFlowSet[$_]) { $routeFlowSet[$_] } else { $_ } }) -join ', ') + ']') }
    if ($routeDnisKeys.Count -gt 0) { $crit += ('callback DNIS in [' + ($routeDnisKeys -join ', ') + ']') }
    if ($authOutcomeKey) { $crit += ('flow outcome ' + $authOutcomeKey + ' = SUCCESS') }
    if ($authAttrName) {
        if ($AuthAttributeValue) { $crit += ('participant data ' + $authAttrName + ' = ' + $AuthAttributeValue) }
        else { $crit += ('participant data ' + $authAttrName + ' present') }
    }
    Write-Host ('Qualified if     : {0}' -f ($crit -join '; '))
}
else {
    Write-Warning 'No authenticated-route criteria supplied (-CallbackDnis / -CallbackFlowIds / -CallbackFlowNames / -AuthOutcomeId / -AuthAttributeName): ANY callback from the same number counts as qualified.'
}
Write-Host ''

# --- Pass 1: every inbound call through the source flows
$inboundFilter = @{ type = 'and'; predicates = @(New-DimensionPredicate -Dimension 'originatingDirection' -Value 'inbound') }
$voiceFilter = @{ type = 'and'; predicates = @(New-DimensionPredicate -Dimension 'mediaType' -Value 'voice') }
$sourcePredicates = @(foreach ($id in $script:SourceFlowSet.Keys) { New-DimensionPredicate -Dimension 'flowId' -Value $id })

$factStore = @{}
Write-Host 'Pass 1: calls through the source flows...'
Import-GcConversationDetails -From $startUtc -To $endUtc -SegmentFilters @(@{ type = 'or'; predicates = $sourcePredicates }) `
    -ConversationFilters @($inboundFilter) -Store $factStore -Label 'source'

$sourceCalls = @(foreach ($f in $factStore.Values) {
        if ($f.SourceFlowId -and $null -ne $f.StartUtc -and $f.StartUtc -ge $startUtc -and $f.StartUtc -lt $endUtc) { $f }
    })

# --- Classify the IVR-ended ("hung up") calls
$events = @(foreach ($f in $sourceCalls) {
        if ($f.Outcome -ne 'EndedInIvr') { continue }
        $status = ''
        if ($CustomerHangupOnly -and -not $f.CustomerHungUp) { $status = 'ExcludedFlowDisconnect' }
        elseif ($msgOutcomeKey -and -not $f.Outcomes.ContainsKey($msgOutcomeKey)) { $status = 'ExcludedNoMessageOutcome' }
        elseif ($MinIvrSeconds -gt 0 -and $null -ne $f.DurationSec -and $f.DurationSec -lt $MinIvrSeconds) { $status = 'ExcludedShortIvr' }
        elseif (-not $f.AniKey) { $status = 'ExcludedWithheldAni' }
        elseif ($f.EndUtc.AddMinutes($CallbackWindowMinutes) -gt $nowUtc) { $status = 'ExcludedWindowStillOpen' }
        @{
            Fact            = $f
            ConversationId  = $f.ConversationId
            AniRaw          = $f.AniRaw
            AniKey          = $f.AniKey
            EndUtc          = $f.EndUtc
            Status          = $status
            CallbacksInWindow = 0
            FirstCallback   = $null
            FirstMinutes    = $null
            QualCallback    = $null
            QualMinutes     = $null
            QualReason      = ''
        }
    })
$eligible = @(foreach ($e in $events) { if (-not $e.Status) { $e } })

$inProgress = 0
foreach ($f in $sourceCalls) { if ($f.Outcome -eq 'InProgress') { $inProgress++ } }
Write-Host ('  {0} calls through the source flows, {1} ended in the IVR, {2} eligible for callback matching.' -f $sourceCalls.Count, $events.Count, $eligible.Count)
if ($inProgress -gt 0) { Write-Host ('  ({0} calls still in progress were ignored.)' -f $inProgress) }

# --- Pass 2: calls from the same numbers after the hang-up
if ($eligible.Count -gt 0) {
    $lastEnd = $startUtc
    foreach ($e in $eligible) { if ($e.EndUtc -gt $lastEnd) { $lastEnd = $e.EndUtc } }

    if ($CallbackSearch -eq 'AllInbound') {
        $to = $lastEnd.AddMinutes($CallbackWindowMinutes + 1)
        if ($to -gt $nowUtc) { $to = $nowUtc }
        Write-Host 'Pass 2: all inbound voice calls (AllInbound mode)...'
        Import-GcConversationDetails -From $startUtc -To $to -SegmentFilters @($voiceFilter) `
            -ConversationFilters @($inboundFilter) -Store $factStore -Label 'inbound'
    }
    else {
        # Batch ANIs by time so each query only covers a short interval.
        $sorted = @($eligible | Sort-Object { $_.EndUtc })
        $batches = @()
        $cur = $null
        foreach ($e in $sorted) {
            if ($null -eq $cur) { $cur = New-Batch -FirstEvent $e; continue }
            if (-not $cur.AniSet.ContainsKey($e.AniRaw)) {
                if ($cur.Anis.Count -ge $AniBatchSize -or ($e.EndUtc - $cur.MinEnd).TotalHours -ge $ChunkHours) {
                    $batches += , $cur
                    $cur = New-Batch -FirstEvent $e
                    continue
                }
                $cur.Anis += $e.AniRaw
                $cur.AniSet[$e.AniRaw] = $true
            }
            $cur.MaxEnd = $e.EndUtc
        }
        if ($null -ne $cur) { $batches += , $cur }

        Write-Host ('Pass 2: callbacks from {0} hang-ups, {1} ANI batch queries...' -f $eligible.Count, $batches.Count)
        $work = @($batches)
        $i = 0
        while ($i -lt $work.Count) {
            $b = $work[$i]
            $i++
            if ($i % 25 -eq 0 -or $i -eq $work.Count) { Write-Host ('  batch {0}/{1}' -f $i, $work.Count) }
            $from = $b.MinEnd.AddMinutes(-1)
            $to = $b.MaxEnd.AddMinutes($CallbackWindowMinutes + 1)
            if ($to -gt $nowUtc) { $to = $nowUtc }
            if ($to -le $from) { continue }
            $aniPredicates = @(foreach ($a in $b.Anis) { New-DimensionPredicate -Dimension 'ani' -Value $a })
            try {
                Import-GcConversationDetails -From $from -To $to -SegmentFilters @(@{ type = 'or'; predicates = $aniPredicates }, $voiceFilter) `
                    -ConversationFilters @($inboundFilter) -Store $factStore -Label 'callbacks' -Quiet
            }
            catch {
                $n = $b.Anis.Count
                if ($_.Exception.Message -like '*HTTP 400*' -and $n -gt 1) {
                    # Too many predicates for one query: split the batch and retry both halves.
                    $half = $n -shr 1
                    Write-Warning ('ANI batch of {0} rejected (HTTP 400); splitting. Consider a smaller -AniBatchSize.' -f $n)
                    $first = @{ Anis = @($b.Anis[0..($half - 1)]); MinEnd = $b.MinEnd; MaxEnd = $b.MaxEnd }
                    $second = @{ Anis = @($b.Anis[$half..($n - 1)]); MinEnd = $b.MinEnd; MaxEnd = $b.MaxEnd }
                    $work += , $first
                    $work += , $second
                }
                else { throw }
            }
        }
    }
}

# --- Index every known inbound call by caller number
$byAni = @{}
foreach ($f in $factStore.Values) {
    if (-not $f.AniKey -or $null -eq $f.StartUtc) { continue }
    if (-not $byAni.ContainsKey($f.AniKey)) { $byAni[$f.AniKey] = @() }
    $byAni[$f.AniKey] += $f
}

# --- Find callbacks in the window
$candidateIds = @{}
foreach ($e in $eligible) {
    $cands = $byAni[$e.AniKey]
    $windowEnd = $e.EndUtc.AddMinutes($CallbackWindowMinutes)
    $inWindow = @(foreach ($c in $cands) {
            if ($c.ConversationId -ne $e.ConversationId -and $c.StartUtc -ge $e.EndUtc -and $c.StartUtc -le $windowEnd) { $c }
        })
    $inWindow = @($inWindow | Sort-Object { $_.StartUtc })
    $e.InWindow = $inWindow
    foreach ($c in $inWindow) { $candidateIds[$c.ConversationId] = $true }
}

# --- Participant data check (only for the callbacks found)
$attrResults = @{}
if ($authAttrName -and $candidateIds.Count -gt 0) {
    Write-Host ('Reading participant data for {0} callback conversations (conversation:communication:view)...' -f $candidateIds.Count)
    $n = 0
    foreach ($id in @($candidateIds.Keys)) {
        $n++
        if ($n % 50 -eq 0) { Write-Host ('  {0}/{1}' -f $n, $candidateIds.Count) }
        $attrResults[$id] = Test-ConversationAttribute -ConversationId $id -Name $authAttrName -Value $AuthAttributeValue
    }
}

# --- Qualify each callback (cached per conversation)
$qualCache = @{}
foreach ($id in @($candidateIds.Keys)) {
    $c = $factStore[$id]
    $reasons = @()
    $routeOk = -not $hasRoute
    if ($hasRoute) {
        foreach ($fid in $c.FlowIds) {
            if ($routeFlowSet.ContainsKey($fid)) { $routeOk = $true; $reasons += 'flow'; break }
        }
        foreach ($k in $routeDnisKeys) {
            if (Test-NumberMatch -A $c.DnisKey -B $k) { $routeOk = $true; $reasons += 'dnis'; break }
        }
    }
    $proofOk = -not $hasProof
    if ($hasProof) {
        if ($authOutcomeKey -and $c.Outcomes.ContainsKey($authOutcomeKey) -and $c.Outcomes[$authOutcomeKey] -eq 'SUCCESS') {
            $proofOk = $true
            $reasons += 'outcome'
        }
        if ($authAttrName -and $attrResults[$id] -eq $true) {
            $proofOk = $true
            $reasons += 'attribute'
        }
    }
    if (-not $hasRoute -and -not $hasProof) { $reasons += 'anyCallback' }
    $qualCache[$id] = @{ Qualified = ($routeOk -and $proofOk); Reason = ($reasons -join '+') }
}

foreach ($e in $eligible) {
    $e.CallbacksInWindow = $e.InWindow.Count
    if ($e.InWindow.Count -eq 0) { $e.Status = 'NoCallback'; continue }
    $e.FirstCallback = $e.InWindow[0]
    $e.FirstMinutes = ($e.InWindow[0].StartUtc - $e.EndUtc).TotalMinutes
    foreach ($c in $e.InWindow) {
        $q = $qualCache[$c.ConversationId]
        if ($q.Qualified) {
            $e.QualCallback = $c
            $e.QualMinutes = ($c.StartUtc - $e.EndUtc).TotalMinutes
            $e.QualReason = $q.Reason
            break
        }
    }
    if ($null -ne $e.QualCallback) { $e.Status = 'CalledBackQualified' } else { $e.Status = 'CalledBackNotQualified' }
}

# --- Detail CSV
$detailColumns = @(
    'HangupConversationId', 'HangupStart', 'HangupEnd', 'SourceFlowName', 'SourceFlowId', 'Ani', 'Dnis',
    'IvrSeconds', 'CustomerHungUp', 'CustomerDisconnectType', 'FlowExitReason', 'Status', 'CallbacksInWindow',
    'FirstCallbackConversationId', 'FirstCallbackStart', 'MinutesToFirstCallback', 'FirstCallbackDnis',
    'FirstCallbackFlows', 'FirstCallbackOutcome',
    'QualifiedCallbackConversationId', 'QualifiedCallbackStart', 'MinutesToQualifiedCallback', 'QualifiedBy',
    'QualifiedCallbackDnis', 'QualifiedCallbackFlows', 'QualifiedCallbackOutcome', 'QualifiedCallbackWaitSec'
)
$sortedEvents = @($events | Sort-Object { $_.Fact.StartUtc })
$detailRows = @(foreach ($e in $sortedEvents) {
        $f = $e.Fact
        $fc = $e.FirstCallback
        $qc = $e.QualCallback
        $row = @{
            HangupConversationId            = $f.ConversationId
            HangupStart                     = (Format-Local $f.StartUtc)
            HangupEnd                       = (Format-Local $f.EndUtc)
            SourceFlowName                  = $f.SourceFlowName
            SourceFlowId                    = $f.SourceFlowId
            Ani                             = $f.AniRaw
            Dnis                            = $f.DnisRaw
            IvrSeconds                      = (Format-Number $f.DurationSec 0)
            CustomerHungUp                  = [string]$f.CustomerHungUp
            CustomerDisconnectType          = $f.CustomerDisc
            FlowExitReason                  = $f.SourceExit
            Status                          = $e.Status
            CallbacksInWindow               = [string]$e.CallbacksInWindow
            FirstCallbackConversationId     = ''
            FirstCallbackStart              = ''
            MinutesToFirstCallback          = ''
            FirstCallbackDnis               = ''
            FirstCallbackFlows              = ''
            FirstCallbackOutcome            = ''
            QualifiedCallbackConversationId = ''
            QualifiedCallbackStart          = ''
            MinutesToQualifiedCallback      = ''
            QualifiedBy                     = ''
            QualifiedCallbackDnis           = ''
            QualifiedCallbackFlows          = ''
            QualifiedCallbackOutcome        = ''
            QualifiedCallbackWaitSec        = ''
        }
        if ($null -ne $fc) {
            $row.FirstCallbackConversationId = $fc.ConversationId
            $row.FirstCallbackStart = (Format-Local $fc.StartUtc)
            $row.MinutesToFirstCallback = (Format-Number $e.FirstMinutes 1)
            $row.FirstCallbackDnis = $fc.DnisRaw
            $row.FirstCallbackFlows = $fc.FlowPath
            $row.FirstCallbackOutcome = $fc.Outcome
        }
        if ($null -ne $qc) {
            $row.QualifiedCallbackConversationId = $qc.ConversationId
            $row.QualifiedCallbackStart = (Format-Local $qc.StartUtc)
            $row.MinutesToQualifiedCallback = (Format-Number $e.QualMinutes 1)
            $row.QualifiedBy = $e.QualReason
            $row.QualifiedCallbackDnis = $qc.DnisRaw
            $row.QualifiedCallbackFlows = $qc.FlowPath
            $row.QualifiedCallbackOutcome = $qc.Outcome
            $row.QualifiedCallbackWaitSec = (Format-Number $qc.WaitSec 0)
        }
        New-Object PSObject -Property $row
    })

# --- Summary rows: overall, per source flow, per local day
$summaryRows = @()
$summaryRows += New-SummaryRow -Scope 'ALL' -Events $events -TotalCalls $sourceCalls.Count

# Grouped by flow id so a flow renamed mid-period stays one row.
$flowIdsSeen = @($sourceCalls | ForEach-Object { $_.SourceFlowId } | Sort-Object -Unique)
foreach ($fid in $flowIdsSeen) {
    $callsInFlow = @(foreach ($f in $sourceCalls) { if ($f.SourceFlowId -eq $fid) { $f } })
    $eventsInFlow = @(foreach ($e in $events) { if ($e.Fact.SourceFlowId -eq $fid) { $e } })
    $label = [string]$script:SourceFlowSet[$fid]
    if (-not $label) { $label = [string]$callsInFlow[$callsInFlow.Count - 1].SourceFlowName }
    if (-not $label) { $label = $fid }
    $summaryRows += New-SummaryRow -Scope ('Flow: ' + $label) -Events $eventsInFlow -TotalCalls $callsInFlow.Count
}

$callsByDay = @{}
$eventsByDay = @{}
foreach ($f in $sourceCalls) {
    $d = Format-LocalDay $f.StartUtc
    if (-not $callsByDay.ContainsKey($d)) { $callsByDay[$d] = 0; $eventsByDay[$d] = @() }
    $callsByDay[$d] = $callsByDay[$d] + 1
}
foreach ($e in $events) { $eventsByDay[(Format-LocalDay $e.Fact.StartUtc)] += , $e }
foreach ($d in @($callsByDay.Keys | Sort-Object)) {
    $summaryRows += New-SummaryRow -Scope ('Day: ' + $d) -Events $eventsByDay[$d] -TotalCalls $callsByDay[$d]
}

$summaryColumns = @(
    'Scope', 'CallsThroughSourceFlows', 'EndedInIvr', 'EndedInIvrPct', 'Excluded', 'EligibleHangups',
    'NoCallback', 'CalledBackAny', 'CalledBackAnyPct', 'CalledBackQualified', 'QualifiedSuccessRatePct',
    'CalledBackNotQualifiedOnly', 'QualifiedAnsweredByAgent', 'QualifiedAnsweredPct',
    'AvgMinutesToQualifiedCallback', 'MedianMinutesToQualifiedCallback', 'MedianQualifiedAnswerWaitSec',
    'UniqueCallers', 'UniqueCallersQualified', 'UniqueCallerQualifiedPct'
)

# --- Write files
if (-not (Test-Path -Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
$n = Get-Date
$stamp = (Get-Padded $n.Year 4) + (Get-Padded $n.Month 2) + (Get-Padded $n.Day 2) + '_' + (Get-Padded $n.Hour 2) + (Get-Padded $n.Minute 2) + (Get-Padded $n.Second 2)
$detailPath = Join-Path -Path $OutputFolder -ChildPath ('IvrHangupCallback_Detail_' + $stamp + '.csv')
$summaryPath = Join-Path -Path $OutputFolder -ChildPath ('IvrHangupCallback_Summary_' + $stamp + '.csv')
$detailRows | Select-Object $detailColumns | Export-Csv -Path $detailPath -NoTypeInformation -Encoding UTF8
$summaryRows | Select-Object $summaryColumns | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8

# --- Console summary
$statusCounts = @{}
foreach ($e in $events) {
    if (-not $statusCounts.ContainsKey($e.Status)) { $statusCounts[$e.Status] = 0 }
    $statusCounts[$e.Status] = $statusCounts[$e.Status] + 1
}

Write-Host ''
Write-Host '================ IVR hang-up -> callback ================'
foreach ($r in $summaryRows) {
    if ($r.Scope -like 'Day:*') { continue }
    Write-Host ''
    Write-Host ('[{0}]' -f $r.Scope)
    Write-Host ('  Calls through source flows        : {0}' -f $r.CallsThroughSourceFlows)
    Write-Host ('  Ended in IVR                      : {0} ({1}%)' -f $r.EndedInIvr, $r.EndedInIvrPct)
    Write-Host ('  Eligible hang-ups (denominator)   : {0}   (excluded {1})' -f $r.EligibleHangups, $r.Excluded)
    Write-Host ('  Called back within {0} min (any)   : {1} ({2}%)' -f $CallbackWindowMinutes, $r.CalledBackAny, $r.CalledBackAnyPct)
    Write-Host ('  Called back QUALIFIED (success)   : {0} ({1}%)' -f $r.CalledBackQualified, $r.QualifiedSuccessRatePct)
    Write-Host ('  Called back, other route only     : {0}' -f $r.CalledBackNotQualifiedOnly)
    Write-Host ('  Qualified callbacks answered      : {0} ({1}%), median wait {2}s' -f $r.QualifiedAnsweredByAgent, $r.QualifiedAnsweredPct, $r.MedianQualifiedAnswerWaitSec)
    Write-Host ('  Minutes to qualified callback     : avg {0}, median {1}' -f $r.AvgMinutesToQualifiedCallback, $r.MedianMinutesToQualifiedCallback)
    Write-Host ('  Unique callers qualified          : {0} of {1} ({2}%)' -f $r.UniqueCallersQualified, $r.UniqueCallers, $r.UniqueCallerQualifiedPct)
}
Write-Host ''
Write-Host 'Status breakdown (all IVR-ended calls):'
foreach ($k in @($statusCounts.Keys | Sort-Object)) { Write-Host ('  {0,-28} {1}' -f $k, $statusCounts[$k]) }
Write-Host ''
Write-Host ('Detail : {0}' -f $detailPath)
Write-Host ('Summary: {0}' -f $summaryPath)
