# Milestone D: background avatar production

D consumes confirmed C authorization snapshots and ends at AVATAR READY.
There is no Recording PowerPoint assembly, slide insertion, or provider cancellation.
The existing shared `WhiteAvatar.psm1` and the proven HeyGen v3 request shapes are
retained. API keys stay in environment variables and are never sent to the browser.

## Runtime workflow

`Worker.ps1` discovers confirmed authorization, atomically claims its immutable
consumption receipt and creates deterministic per-slide jobs. It schedules up to
three child processes. Each child runs one bounded submission, poll, download,
conversion, or validation step; polling waits are represented by durable timestamps.
The HTTP server does not wait for provider work. It reads job metadata for the live
dashboard and automatically recovers a worker process that exits.

Source components:

- `AvatarProduction.psm1`: durable queues, claims, input checks, concurrency,
  compatible asset reuse, exception recovery, pause/resume and status.
- `HeyGenProvider.psm1`: small PowerShell adapter for the existing v3 look check,
  audio upload, video submission and status request shapes. Multipart upload keeps
  the key out of subprocess command lines.
- `Invoke-AvatarStep.ps1`: one provider/local stage, explicit selected audio,
  canonical white-avatar conversion and technical validation.
- `avatars.js`: dashboard, automatic display after authorization, pause/resume,
  details and targeted recovery controls.
- `Production.psm1`: shared JSON reader with delete-sharing and brief I/O retry
  for runtime records replaced atomically during restart.

The normal root remains `%LOCALAPPDATA%\LectureForge\Projects\<id>`.
D acceptance uses short isolated roots such as `LectureForgeD2`; all evidence and
media stay outside Git. New provider/conversion files use short generated names.
Every successful asset is preserved. Local retries use new evidence/output names.

## Durable state and safety

- `avatar-queue.json`: paused flag, authorization batches and job IDs.
- `a/<16-hex-job>/job.json`: immutable authorization inputs plus mutable stage,
  provider intent/ID, lease owner/PID/start time, polling time, paths/hashes,
  placement, stage history, validation and exception.
- `approvals/<authorization>.consumed.json`: immutable consumer identity and
  original authorization-file hash. A crash between receipt and queue creation
  recovers the same deterministic job files; incompatible receipts fail closed.
- `a/<job>/run-*.log` / `.err`: separate executor diagnostics, preserving earlier logs.

Exact project/intent revision, enabled slides, narration asset ID/hash, script
revision/hash, preparation, avatar preset/version and job inputs are checked before
work. An altered job record is rejected against the immutable snapshot. Unsupported
avatar presets fail rather than silently using the hard-coded proven converter.
The original authorization is never rewritten. Stale inputs require Review Selections.

A submission intent is durable before POST /v3/videos. The returned job ID is
written before polling. An interrupted submission without an ID becomes an
exception and is never submitted again automatically. Known IDs resume polling.
Three nonterminal provider submissions (including uncertain ones) prevent a fourth
paid start. Process leases prevent duplicate work during worker/server recovery.
A new authorization cannot silently duplicate a prior live submission for the same
slide, narration and avatar preset. A completed compatible result can be reused.

Pause Queue prevents starting new steps, including new polls; it does not cancel
already-submitted jobs at HeyGen. Clean stop finishes active bounded child steps.
Abrupt restart recovers existing leases/IDs; interrupted local steps become targeted
exceptions. No generic retry action can create a paid replacement. Reconciliation
accepts an existing provider job ID, and does not submit one.

## Avatar/audio validation

The canonical converter retains its 640x560 white RGB 255/255/255 tile, approved
crop, 70-pixel torso fade, motion, H.264 yuv420p/AAC encoding and explicit `1:a:0`
selected narration mapping. Normalized placement is stored separately for the
later final assembly stage. It is never a full-slide composite.

Ready validation checks media SHA-256, H.264/AAC profile, full decode, nonzero
narration duration, output duration within 0.08 seconds, independently encoded
selected-narration AAC packet identity, and changing presenter pixels. Prepared
pause metadata is bound to the reviewed asset hash and the decoded selected
narration tail is checked for the declared silence samples. Trailing narration is
not trimmed. No subjective voice/lip-sync score or extra approval gate is added.

Solar Thermal Slide 8 binds to prepared Take 2:
`E234676D7F92B42D20E98A3CB6CADF5BAA89595C2DCA906005EB7913607D69A5`.
Its 6.000 seconds / 264600 samples at 44100 Hz are validated before Avatar Ready.
Rejected historical Slide 4 Take 1 is retained and excluded from approved reuse.

## Acceptance and provider boundary

`Test-AvatarStudio.ps1` copies the verified 39 narration fixtures and 13 existing
white avatars, preserves rejected history, exercises queue/recovery guards and
uses one existing raw avatar for a real local conversion. This regenerates no paid
media. `Test-AvatarRecovery.ps1` continues from existing assets to test local retry,
audio/pause validation, concurrent HTTP reads and server/worker restart.
`Test-AvatarGuards.ps1` tests four-slide reservations, three-provider capacity,
receipt-to-queue crash recovery, stale inputs and duplicate/tampered-job rejection.
`Test-AvatarBrowser.cjs` uses a separate headless Edge profile to exercise the real
local dashboard, pause/resume, details and page refresh without another dependency.

Acceptance runtimes contain `avatar-policy.json` with
`{"provider_calls_enabled": false}`. The executor checks this before transport,
including look requests, uploads, polling and downloads. This is a test/runtime
provider block, not a fake successful provider response. Existing records provide
known-job recovery fixtures; live HeyGen transport is intentionally not exercised.
Do not enable provider transport on these fixture projects.

The C regression now pauses its fixture queue and blocks provider transport before
startup, so testing authorization alone cannot accidentally begin D production.

## Scope limits

No new provider request is needed for this checkpoint. Real v3 transport remains
unexercised in D; its request/response shapes match the completed Solar Thermal run.
Existing approved assets are recognized through the explicit catalog or prior
ready queue records. C's pre-authorization cost summary can conservatively overcount
uncataloged prior queue outputs; D reuses exact matches without spending for them.
There is no new paid-replacement retry button: use narration review and a new
explicit authorization after reconciling existing submissions. Final PowerPoint
assembly, placement editing and export remain later milestones.

## Checkpoint results

All 102 suite checks passed:

| Suite | Result | Evidence under LOCALAPPDATA |
|---|---:|---|
| D full fixture production | 13/13 | LectureForgeD4/results.json |
| D local retry and recovery | 13/13 | LectureForgeD2/recovery-results.json |
| D submission/queue guards | 11/11 | LectureForgeDG3/results.json |
| Headless Edge dashboard | 6/6 | LectureForgeD4/browser-results.json |
| C authorization regression | 25/25 | LectureForgeDC/results.json |
| B narration regression | 19/19 | LectureForgeDB/Tests/d99a8fde/results.json |
| A foundation regression | 15/15 | LectureForgeDA/Tests/08a80d43/results.json |

The final fixture reached 13/13 AVATAR READY with zero exceptions. Its review
project is `LectureForgeD4/Projects/906821db088d`, served at
http://127.0.0.1:8783/. The runtime provider block remains enabled for review.
The final restart additionally preserved byte-identical completed job records.

Thirteen existing Solar Thermal white avatars and their provider records were
reused. Known provider ID recovery was exercised without remote polling. Two
isolated local conversion proof runs (initial/repaired fixture runs) reused the
same already-paid raw Slide 2 video; original narration/provider/media files were
never overwritten. Every resulting ready avatar passed selected-audio packet
identity and duration checks. Slide 8 passed decoded prepared-silence verification.

During development a local child interruption was flagged on Slide 8 before
validation. Its successful MP4 was preserved and targeted validation passed.
Per-step logs now preserve executor diagnostics. The final uninterrupted fixture
run completed all 13 slides without exceptions. Restart testing also exposed a
transient runtime-file read during replacement; the shared reader fixes this and
server/worker restart tests passed. Test-only setup issues (UTF-8 scripts and a
synthetic uncertain job that initially retained a known ID) were corrected.

ElevenLabs calls during D: **0**. HeyGen calls during D: **0**, including status
polls, uploads and downloads. No paid assets regenerated. Canonical source hash:
`6340C137FD85021957721341298D640C8FAF59D36507D904DF09948DCB576C01`.
No final Recording deck assembly was performed.

Created files:

- studio/AvatarProduction.psm1
- studio/HeyGenProvider.psm1
- studio/Invoke-AvatarStep.ps1
- studio/MILESTONE_D.md
- studio/web/avatars.js
- tests/v2/Test-AvatarBrowser.cjs
- tests/v2/Test-AvatarGuards.ps1
- tests/v2/Test-AvatarRecovery.ps1
- tests/v2/Test-AvatarStudio.ps1

Modified files:

- Start-LectureForge.ps1
- studio/Authorization.psm1
- studio/Production.psm1
- studio/Server.ps1
- studio/Worker.ps1
- studio/web/index.html
- studio/web/selections.js
- tests/v2/Test-SelectionAuthorization.ps1

No generated modules or binary assets belong in this change.

## Instructor-approved commit checkpoint

The instructor approved Milestone D and accepted the previously proven Solar
Thermal transport as the live-provider integration evidence. The commit checkpoint
reruns the offline suites in fresh short runtime roots, preserving prior evidence:

| Suite | Result | Evidence under LOCALAPPDATA |
|---|---:|---|
| D full fixture production | 13/13 | LectureForgeDF/results.json |
| D recovery | 13/13 | LectureForgeDF/recovery-results.json |
| D submission/queue guards | 11/11 | LectureForgeDGF/results.json |
| Headless Edge dashboard | 6/6 | LectureForgeDF/browser-results.json |
| C authorization regression | 25/25 | LectureForgeDCF/results.json |
| B narration regression | 19/19 | LectureForgeDBF/Tests/9848c813/results.json |
| A foundation regression | 15/15 | LectureForgeDAF/Tests/e64322ec/results.json |

The fixture project `LectureForgeDF/Projects/9a6884f7a78e` reached 13 Avatar Ready
at http://127.0.0.1:8784/. Checkpoint provider calls: ElevenLabs **0**, HeyGen **0**.
The existing 39 narration takes and paid provider assets were reused. One local
conversion proof used the existing raw Slide 2 avatar. Source hash remains the
canonical hash above. PowerShell/JavaScript syntax, credential-literal and Git
whitespace checks passed. No Milestone E work or PowerPoint assembly was performed.

Preserved Git references:

- master: `e8f9ff57c2dcf54ac81d65f34ced9c8b8c71c9cd`
- lectureforge-v1-baseline: `a3bbc92a0f499e29df0ea966c2d1794ed95fb249`
