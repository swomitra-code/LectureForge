# Milestone B: narration generation and review

Scope: ElevenLabs narration only. No HeyGen endpoints, avatar jobs, or Recording
PowerPoint assembly. The production preset and existing media/Recording engine
are unchanged. Existing safer narration diagnostics moved verbatim into
`NarrationProvider.psm1`, shared by Studio and `generate-narration.ps1`.

## Instructor workflow

Run `Start-LectureForge.ps1`, open a lecture, save scripts and enabled slides,
then click **GENERATE 3 TAKES**. Confirm the displayed slide/generation count.
Only missing narration revisions are queued. Existing revisions and successful
takes are preserved. A confirmed failed take has its own one-request retry button.

Open **Review Narration** to see a read-only slide thumbnail, exact saved script,
three players, durations, and selection controls. **Select Take** and **Change
Selection** only update narration intent. Use Previous/Next or the slide list.
Selections remain editable. No avatar authorization exists in this milestone.

Normal production has one canonical project root:
`%LOCALAPPDATA%\LectureForge\Projects\`. Both the launcher default and the backend
root resolver retain that Milestone A default; normal startup is simply
`Start-LectureForge.ps1` (port 8770).

`%LOCALAPPDATA%\LectureForgeB\Projects\` is intentionally an isolated Milestone B
test namespace, selected only through an explicit `-RuntimeRoot` override. The
`LectureForgeBA`, `LectureForgeBF`, `LectureForgeAF`, and `LectureForgeAF2` roots
likewise contain regression evidence only. They are not alternate production
defaults. No runtime-root migration, move, or deletion is required or performed.

To reopen the Milestone B test copies (port 8771):

```powershell
.\Start-LectureForge.ps1 -RuntimeRoot "$env:LOCALAPPDATA\LectureForgeB" -Port 8771
```

## State and recovery

- `project.json` retains source/preset/scripts; `narration.json` contains narration
  revisions, selections, and attempt history. A named per-project mutex protects
  state mutations; JSON replacement is atomic. Media paths are project-relative.
- Binding includes exact UTF-8 script hash, saved project revision, pause, and
  preset settings. Revised scripts/preparation receive new revision IDs and take
  paths. Old assets/selections remain in history; returning to the same binding
  restores the corresponding prior selection.
- Each attempt gets a unique directory: `n/<revision>/<attempt>/raw.mp3` and
  `take.mp3` or `take.wav`. Outputs use create-new/no-overwrite behavior. Decode,
  duration, and SHA-256 are recorded before a candidate becomes ready. Hashes are
  checked on selection, media requests, and generation planning.
- A background worker dispatches one narration executor at a time while keeping
  the HTTP server responsive. Dispatch and calling state are persisted before
  the request. PID plus process start time identify active executors after restart.
  Abandoned attempts become uncertain and are never automatically resubmitted.
- A clean stop stops dispatch and lets an already-authorized request finish.
  Previously authorized queued takes may resume on restart; successful takes and
  selections themselves never trigger requests.
- Confirmed HTTP rejections can be explicitly retried individually with a new
  attempt path. Timeouts, interrupted requests, and ambiguous server failures need
  investigation. Local preparation failures retain received media as exceptions;
  they do not automatically purchase replacement audio.
- Paused candidates use PCM WAV: decoded speech plus an exact integer number of
  zero samples at 44100 Hz. This avoids MP3 encoder padding in the reviewed pause.
  Non-paused candidates preserve the provider MP3 bytes. Pause metadata is saved
  before review. The maximum supported pause is 120 seconds.
- Thumbnails use the same read-only desktop PowerPoint COM approach as Render,
  exporting all slides in one session without save/embedding operations.

## Acceptance evidence (2026-09-12)

- `Test-NarrationStudio.ps1`: 19 checks passed. Copied 39 Solar Thermal candidates
  with exact script/hash binding; verified decode, editable selection, revision
  history, targeted retry, uncertain-attempt protection, HTTP 206 seeking, 14
  thumbnails, restart identity, and canonical source preservation. Provider calls: 0.
  Results: `%LOCALAPPDATA%\LectureForgeB\Tests\f0c9667e\results.json`.
- `Test-NarrationLive.ps1`: three independent real ElevenLabs calls for approved
  Slide 8 text in a separate deck copy. Takes ready at 20.489252, 18.027937, and
  21.000091 seconds. No retries. Results: `LectureForgeB\live-narration-results.json`.
- `verify_narration_pcm.py`: all three live candidates and the preparation fixture
  equal decoded speech plus exactly 264600 zero samples (6.000 seconds).
- `Test-NarrationResume.ps1`: clean stop/restart, byte-identical live narration JSON,
  all hashes unchanged, Take 2 retained, HTTP audio seeking works, diagnostic
  redaction passes, zero new provider calls.
- Existing Milestone A suite: 15/15 passed against a separate short test root
  `%LOCALAPPDATA%\LectureForgeBA`; results in `Tests\175ce55f\results.json`.
- Edge UI: project opening, thumbnail/script display, three players, playback,
  seeking, selection of Take 1 and change back to Take 2 verified.
- Total new ElevenLabs calls: **3**. HeyGen calls: **0**.
- Canonical Solar Thermal SHA-256 remains
  `6340C137FD85021957721341298D640C8FAF59D36507D904DF09948DCB576C01`.

Review fixture project: `LectureForgeB\Projects\4b9d067085b3`.
Live test project: `LectureForgeB\Projects\31cd01c7e329` (Slide 8; final UI test selects Take 3).
All fixture/live media and test outputs remain outside Git.

## Test usage and current limits

Fixture tests require a **fresh short runtime root** with no running worker;
synthetic retry states must never be consumed by a production worker. Live tests
are explicitly opt-in and checkpoint their project before queue submission;
rerunning never repeats successful or uncertain requests.

Generation targets all enabled slides; use enabled checkboxes for a subset.
There is no additional multi-select generation UI. No keyboard shortcuts beyond
normal browser navigation. Thumbnails are static; interrupted/failed thumbnail
exports and local media-preparation exceptions currently require local diagnosis.
Ambiguous paid attempts deliberately have no one-click retry. Subjective narration
quality remains the instructor's listening decision.

## Final approval checks

The dedicated browser **Change Selection** button was exercised on the existing
live-proof project: Take 2 -> no selection -> Select Take 3. The browser showed
"No take selected" after clearing; persisted intent matched the UI. All three
asset hashes and complete take/attempt metadata stayed identical. Take 3 survived
restart. `Test-ChangeSelection.ps1` records the four observation phases without
mutating selections or calling providers, and asserts the canonical launcher and
backend runtime-root defaults. Evidence: `LectureForgeB/change-selection-results.json`.

The fixture suite passed again (19/19), using the existing 39 narration takes in
`LectureForgeBF/Tests/c7cdc732`. The Milestone A suite passed (15/15) in
`LectureForgeAF2/Tests/5a56873f`. No live-generation test was repeated.

An initial regression exposed a Windows sharing conflict when reading the worker
heartbeat during atomic replacement. The server now reads with delete-sharing and
briefly retries transient I/O conflicts; persistent errors remain visible. Only
health/status reading changed, not narration/provider behavior.

ElevenLabs calls during final approval checks: **0**. HeyGen calls: **0**.
The earlier isolated three-request proof remains the only new paid narration.

Final controlled restart validation passed, including 100/100 local health reads,
byte-identical narration state, preserved Take 3, and audio seeking. Earlier stress
runs observed transient heartbeat freshness failures; the worker remained alive.
Results: `LectureForgeB/resume-narration-take-3-results.json`.
