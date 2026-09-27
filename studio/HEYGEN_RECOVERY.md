# Existing HeyGen video recovery

The Exception UI offers **Recover Existing HeyGen Video**. Paste the video page
URL and select **Recover & Continue**. This uses `POST /api/avatar-recover`
locally, and only `GET /v3/videos/{id}` at HeyGen. No upload or generation occurs.

The URL suffix is a candidate, never sufficient evidence: the API must return
the same `id`, a matching canonical page ID when present, and the exact expected
LectureForge project/slide/take title. Unknown or renamed titles fail closed.
Malformed URLs, missing videos, network errors and mismatches leave the original
job unchanged. Retrying a successful recovery performs no further API calls or
state changes. Failed assets remain exceptions; replacements still need explicit
paid authorization.

Recovery follows a blocked row's original submission, including legacy rows whose
lineage exists only in the duplicate-protection error. It preserves authorization,
take, narration, placement and prior history, adding recovery evidence and history.
Only the original job runs. The dashboard follows it without scheduling another
conversion for the newer blocked row.

Recovered processing uses the existing download, white-avatar conversion and
validation stages. Its immutable authorization and consumed receipt remain
mandatory. A later project-wide authorization does not prevent completion if the
currently selected take, narration hash and preset still match. This exception
cannot enter Queued/Submitting or a non-GET provider request. Changed narration,
take or preset still blocks processing.

## Incident findings

The reported Slide 16 case was not a lost create response: the original job had
persisted the provider ID and polled successfully. A later authorization made it
stale, and newer blocked rows displayed their own empty IDs. The original ID was
retained throughout recovery.

A lost response or interrupted persistence remains possible in general. Existing
submission intent already prevents a second generation. Automatic title matching
is unsafe because separate authorizations can share a title. Future automatic
reconciliation should require a provider-supported immutable correlation value
or verified idempotency contract, with ambiguous matches left for manual recovery.
No automatic guessing or paid retries were added.

## Tests

Run `tests/v2/Test-HeyGenReconciliation.ps1` in PowerShell and
`tests/v2/Test-HeyGenReconciliationUI.cjs` with Node. Provider transport in the
PowerShell fixture rejects every method/endpoint except GET video status. The
tests exercise the production worker's polling branch, original duplicate guard,
immutable authorization checks and idempotency. The UI test exercises shipped
handlers, recovery-only submission, and draft/focus preservation during polling.
