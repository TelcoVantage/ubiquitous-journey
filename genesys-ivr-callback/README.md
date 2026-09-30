# Genesys Cloud PowerShell scripts (Constrained Language Mode safe)

| Script | What it does |
|---|---|
| `Get-IvrHangupCallbackRate.ps1` | IVR "hang up and call back" success rate (see below) |
| `Test-TopicPhrases.ps1` | Runs the Speech & Text Analytics **test phrase** job for a list of candidate topic phrases and reports how many transcripts each one matches |

## Test-TopicPhrases.ps1

Same thing as *Topic > Add phrase > Test* in the Genesys UI, for a whole list of phrases at once. For each phrase it calls `POST /api/v2/speechandtextanalytics/topics/testphrase/jobs` with the same body the UI sends, polls `GET .../jobs/{jobId}` until the job finishes, and records `matchedTranscriptsCount` / `processedTranscriptsCount`.

**CONFIG block** (top of the script): region, client id/secret, the program IDs, the phrase list, and the topic settings (dialect `en-AU`, `Semantic`, participants `Internal`, strictness 72, media type `call`, last 29 days). Anything passed on the command line overrides it. The OAuth client needs `speechAndTextAnalytics:topic:testPhrase`.

```powershell
.\Test-TopicPhrases.ps1                                             # everything from CONFIG
.\Test-TopicPhrases.ps1 -StartDate '2026-09-01' -EndDate '2026-09-30' -Strictness 60
.\Test-TopicPhrases.ps1 -PhraseFile .\phrases.txt -Participants Both -MatchingType Lexical
```

Output: `TopicPhraseTest_<stamp>.csv` with `Phrase, MatchedTranscripts, ProcessedTranscripts, MatchPct, State, SecondsToComplete, JobId, Error`, sorted by matches, plus the same table on the console. Up to `-MaxConcurrentJobs` (default 3) jobs run at a time; a job that hasn't finished after `-JobTimeoutSeconds` (default 600) is reported with an error rather than blocking the rest.

The public API only exposes the match counts for a test job, not the transcript snippets the UI shows; use the UI for those once you've narrowed the list.

---

# IVR hang-up → authenticated callback success rate (Genesys Cloud)

`Get-IvrHangupCallbackRate.ps1` measures how well the IVR message *"hang up and call back the authenticated way for faster service"* works. It checks:

> Of the callers who **ended their call in the IVR** of flow X or Y, how many **called back from the same number within 30 minutes**, and how many of those callbacks came in **the authenticated way**?

The script is safe to run in **Constrained Language Mode** (Windows PowerShell 5.1 under AppLocker/WDAC). It makes no .NET static calls and uses no `::new()`, no `[pscustomobject]` casts and no `Add-Type`. It builds the Base64 for the OAuth Basic header with bit operators, and it builds timestamps from date components, so it doesn't depend on culture or calendar settings.

## How it works

| Step | What happens | API |
|---|---|---|
| 1 | Pull every inbound call that went through the source flows in the period. The server filters on `flowId`. | `POST /api/v2/analytics/conversations/details/query` |
| 2 | Keep the calls that **ended in the IVR**: no `acd` participant, no agent, and no external, voicemail or group transfer. These are the "hung-up" calls. | – |
| 3 | Look for inbound voice calls from the same ANIs, starting from the end of each hung-up call through the window. The server filters on `ani`, and ANIs are batched by time so each query covers a short interval. | same query endpoint |
| 4 | Decide whether each callback is **qualified**, i.e. came in the authenticated way (see below). | optional `GET /api/v2/conversations/{id}`, run only for the callbacks found |
| 5 | Write a detail CSV (one row per hung-up call) and a summary CSV (overall, per flow and per day), and print the headline rates. | – |

**Headline KPI:** `QualifiedSuccessRatePct` = eligible hang-ups followed by a qualified callback within the window ÷ eligible hang-ups.

### Quick start: fill in the CONFIG block

The top of the script's `param(...)` block has a **CONFIG** section. Paste your values there and run the script with no arguments:

```powershell
[string]$Environment = 'mypurecloud.com.au',
[string]$ClientId = '<your client id>',
[string]$ClientSecret = '<your client secret>',
[string[]]$SourceFlowIds = @(
    '<flow 1 id>',
    '<flow 2 id>'
),
```

Anything you pass on the command line overrides the CONFIG values. Keep the real secret out of source control: fill it in on your machine only, or leave the placeholder and set `$env:GC_CLIENT_SECRET`.

### What counts as a successful callback

**Default (no criteria passed):** a callback counts when the **same number calls one of the source flows again** within the window, **on any DNIS**. Calls from that number to other flows show up as `CalledBackNotQualified`. Pass `-AnyCallback` to count a call to **any** flow.

You can tighten this with optional criteria, which fall into two groups:

- **Route**: the callback went through `-CallbackFlowIds` / `-CallbackFlowNames`, **or** it dialled `-CallbackDnis`.
- **Proof**: the callback reached flow outcome `-AuthOutcomeId` with value `SUCCESS`, **or** it has participant data `-AuthAttributeName` (optionally equal to `-AuthAttributeValue`).

A callback is **qualified** when it matches **route AND proof**. A group you supply no criteria for is ignored.

Pick the criteria that match how your authenticated path works:

| Your setup | Parameters |
|---|---|
| Separate "authenticated" number or flow | `-CallbackDnis '+61299990000'` or `-CallbackFlowNames 'Authenticated Line'` |
| Same number, and the flow does ID&V | `-AuthOutcomeId <outcome id>` (recommended), or `-AuthAttributeName Auth_Status -AuthAttributeValue Success` |
| Both | Combine them, e.g. DNIS **and** outcome |

`-AuthAttributeName` makes one extra API call per callback found. The synchronous analytics query doesn't return participant data, so a flow outcome is the cheaper choice.

### Which hung-up calls count (eligibility)

| Switch | Effect |
|---|---|
| *(default)* | Every call through the source flows that ended in the IVR |
| `-CustomerHangupOnly` | Only calls where the customer hung up (customer segment `disconnectType = endpoint`). This drops calls the flow disconnected itself. |
| `-MessageOutcomeId <id>` | Only calls where the source flow reached that flow outcome. **Recommended:** put `Set Flow Outcome` right after the "hang up and call back" prompt so you count only callers who actually heard it. |
| `-MinIvrSeconds 20` | Ignore calls shorter than N seconds, as a rough "hung up before the message" filter |

These calls are always excluded and are listed in the detail CSV with an `Excluded…` status:

- Withheld or anonymous ANIs (fewer than 6 digits)
- Hang-ups whose callback window hasn't closed yet
- Calls still in progress

## Setup

1. **OAuth client**: Admin > Integrations > OAuth > Add client > **Client Credentials**. Give its role:
   - `analytics:conversationDetail:view` (always needed)
   - `architect:flow:view` if you pass flow **names**
   - `conversation:communication:view` if you use `-AuthAttributeName`
   - access to the divisions your flows and queues are in
2. **Region**: pass `-Environment`, for example `mypurecloud.com`, `mypurecloud.ie`, `mypurecloud.de`, `mypurecloud.com.au`, `mypurecloud.jp`, `usw2.pure.cloud`, `euw2.pure.cloud` or `aps1.pure.cloud`.
3. **Credentials and flows**: fill in the CONFIG block (see Quick start), or pass `-ClientId` / `-ClientSecret` / `-SourceFlowIds`. If the secret is left as the placeholder, the script uses `$env:GC_CLIENT_SECRET`, then shows a `Get-Credential` prompt; enter the secret as the password. You can also pass `-AccessToken` to skip OAuth.

## Examples

```powershell
# CONFIG block filled in: callbacks into the same two flows within 30 min, last 7 days
.\Get-IvrHangupCallbackRate.ps1

# Same, for a given period
.\Get-IvrHangupCallbackRate.ps1 -StartDate '2026-09-01' -EndDate '2026-09-08'

# Any callback within 30 min counts, whichever flow it reached
.\Get-IvrHangupCallbackRate.ps1 -Environment mypurecloud.com.au -ClientId $id -ClientSecret $secret `
    -SourceFlowNames 'Main Inbound','Billing Inbound' -AnyCallback -StartDate '2026-09-01' -EndDate '2026-09-08'

# Success = called back within 30 min on the authenticated number AND passed ID&V,
# counting only callers who heard the message and hung up themselves
.\Get-IvrHangupCallbackRate.ps1 -Environment mypurecloud.com.au -ClientId $id `
    -SourceFlowIds 'a1b2c3d4-....','e5f6a7b8-....' `
    -CallbackDnis '+61299990000' -AuthOutcomeId '9f8e7d6c-....' `
    -MessageOutcomeId '1a2b3c4d-....' -CustomerHangupOnly `
    -StartDate '2026-09-01' -EndDate '2026-10-01' -CallbackWindowMinutes 30 -OutputFolder .\reports
```

When you call the script with `powershell.exe -File`, arrays arrive as one string. `-SourceFlowIds`, `-CallbackFlowIds` and `-CallbackDnis` therefore also accept a comma-separated list, e.g. `-SourceFlowIds "id1,id2"`.

Dates are read as local machine time unless the value ends in `Z`. `-EndDate` is exclusive.

## Output

**`IvrHangupCallback_Detail_<stamp>.csv`**: one row per call that ended in the IVR.

| Column | Meaning |
|---|---|
| `Status` | `CalledBackQualified`, `CalledBackNotQualified` (called back, but only via another route), `NoCallback`, or `Excluded…` |
| `MinutesToFirstCallback`, `FirstCallbackFlows`, `FirstCallbackDnis`, `FirstCallbackOutcome` | The first call back from that number, whatever route it took |
| `QualifiedCallback…`, `QualifiedBy` | The first qualified callback, and which criteria matched (`flow`, `dnis`, `outcome`, `attribute`, `anyCallback`) |
| `…Outcome` | `Answered`, `QueuedNotAnswered`, `EndedInIvr`, `TransferredOut` or `InProgress` |
| `QualifiedCallbackWaitSec` | Queue time before answer (or before abandon) on the callback: evidence for the "faster service" promise |
| `CustomerDisconnectType`, `FlowExitReason` | The raw Genesys values, so you can check the classification |

**`IvrHangupCallback_Summary_<stamp>.csv`**: one row each for `ALL`, `Flow: <name>` and `Day: <yyyy-MM-dd>`.

Each row has: calls through the flows, IVR-ended %, eligible hang-ups, callback counts and rates (any and qualified), how many qualified callbacks were answered, average and median minutes to the callback, median answer wait, and a **unique callers** view. The unique-callers view counts a caller who hung up three times and then called back once as one success rather than one-in-three.

## Tuning and limits

| Parameter | Default | When to change it |
|---|---|---|
| `-CallbackSearch AllInbound` | `AniFilter` | `AniFilter` matches the exact ANI string. If callbacks can arrive with a different ANI format (another carrier or trunk), use `AllInbound` together with `-AniMatchDigits 9`, which compares only the last 9 digits. |
| `-ChunkHours` | 24 | Lower it (e.g. 6) for very busy flows, to keep each query's result set small. |
| `-AniBatchSize` | 50 | ANIs per query. If the API rejects a batch with HTTP 400, the script halves the batch automatically. |
| `-MaxRetries` | 6 | Retries on 429, 408 and 5xx with exponential backoff (2 s doubling to a 60 s cap). A 401 triggers one token refresh. |

## Reading the result honestly

- **Measure a baseline.** Some callers call back anyway. Run the same query for a period *before* the message went live, or on a flow without the message, and compare the rates. The difference between the two is what the message is responsible for.
- **"Ended in IVR" is broader than "hung up because of the message"** unless you use `-MessageOutcomeId` (best) or `-MinIvrSeconds`.
- **Withheld numbers can't be matched.** They are reported separately and left out of the denominator.
- **Spot-check the matching.** Open a few conversation IDs from the detail CSV in *Performance > Interactions* and confirm the classification before you share the numbers.

## Testing

I tested the script offline with **PowerShell 7.5 on Linux**, with `ConstrainedLanguage` enforced and a mocked Genesys API:

- **Scenarios**: no criteria; route + outcome criteria with every eligibility filter; participant data with 404s; `AllInbound` mode; access-token auth; comma-separated IDs.
- **Error handling**: a forced 429 retry and a forced HTTP 400 that splits the ANI batch.
- **Dates**: both string dates (like Windows PowerShell 5.1) and `DateTime` dates (like PowerShell 7).
- **Base64**: checked against Python for 306 strings, including multi-byte characters.

It has **not** been run against a live Genesys Cloud org or on Windows PowerShell 5.1 under WDAC. Run a short period first, e.g. one day.
