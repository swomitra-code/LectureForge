# Milestone C: selection review and avatar authorization

Milestone C adds review and durable paid-generation authorization around the
existing narration engine. It submits no HeyGen requests and does not consume
authorizations, generate avatars, or assemble PowerPoint decks.

## Instructor workflow

1. Select or change narration takes freely in Narration Review.
2. Click REVIEW SELECTIONS. Each enabled slide shows its title, selected take,
   duration, prepared pause, selection validity, and expected avatar work.
3. Use CHANGE to return to that slide's narration review without changing its
   selection. Select another take or clear intent with Change Selection.
4. Click GENERATE AVATARS FOR SELECTED TAKES to inspect the final exact mapping,
   new-job count, reusable-asset count, exclusions, and paid-generation warning.
5. Only CONFIRM GENERATE FOR THESE SELECTIONS writes the authorization record.
   The UI reports: Avatar generation authorized — waiting for production worker.

The instructor manually accepted CHANGE, editable selection, correct review
mappings and expected job count at http://127.0.0.1:8772/. No provider generation
was authorized during that manual check. This closes the browser acceptance gap;
programmatic confirmation tests use separate disposable fixture projects.

## State and safety boundary

`Authorization.psm1` validates the current project, exact script/narration revision,
selected file SHA-256, recorded successful decode and positive duration, and
prepared-pause metadata. Missing, stale, changed-script, and missing/invalid-asset
selections block final confirmation. Hash checks reuse the technical validation
record from narration generation; review does not re-encode or regenerate audio.

All authorization operations use the existing per-project narration mutex.
Project edits through the HTTP server use the same mutex. A selection intent
revision advances when intent changes, including clearing it. Changing a take and
then changing it back cannot revive the previous authorization.

Project-relative files:

- `selection-review.json`: resumable UI phase, slide navigation, review identifier.
- `approvals/<binding-sha256>.review.json`: immutable exact reviewed input snapshot.
- `approvals/<binding-sha256>.authorization.json`: immutable confirmed snapshot,
  project/revision, UTC timestamp, expected new-job count and provider boundary.
- `approvals/<binding-sha256>.consumed.json`: reserved receipt checked fail-closed;
  Milestone C only creates synthetic receipts in acceptance tests.

The binding uses sorted-key canonical JSON and SHA-256. It includes project and
intent revisions, enabled/excluded slides, source hash, narration asset IDs and
hashes, script revisions/hashes, preparation, avatar preset/version and encoding,
placement metadata and reusable-asset registry hash. On read, snapshot integrity
is checked. Files are created by atomic move without overwriting existing records.
Repeated confirmation returns the same record and file hash. A stale tab is
rejected against current inputs inside the mutex. Old records remain history.

Milestone D must revalidate exact authorization inputs and atomically claim an
unconsumed authorization under the same project mutex before any provider request.
It must persist attempts/job IDs and never retry an uncertain paid submission.
There is no authorization consumer in this milestone. The existing worker still
handles only explicitly requested narration and local thumbnails.

## Exact reusable assets

An optional project `avatars.json` catalog contains an `assets` array. Each entry
records `id`, `slide_number`, `status`, `rejected`, `narration_sha256`,
`preset_sha256`, project-relative `path`, media `sha256`, and
`validation.passed` / `validation.authoritative_audio_verified`.

Reuse requires status `approved`, not rejected, the exact selected prepared
narration hash, exact avatar-preset hash, both validation flags, a safe local
`avatars/<short-name>.mp4` path, and a matching media file hash. No catalog means
no reuse credit. Legacy directories are not automatically inferred or migrated.
The fixture explicitly catalogs one approved Slide 2 avatar and rejected Slide 4
Take 1 history: 13 enabled slides yield 12 expected NEW jobs and 1 reusable asset.
Rejected Slide 4 Take 1 earns no credit, even when its narration hash matches.

Slide 8 binds to the prepared Take 2 asset SHA-256
`E234676D7F92B42D20E98A3CB6CADF5BAA89595C2DCA906005EB7913607D69A5`,
including 6.000 seconds (264600 samples at 44100 Hz) of appended silence.
No pause is added after approval.

## Runtime and testing

Normal startup remains `Start-LectureForge.ps1`, with projects under
`%LOCALAPPDATA%\LectureForge\Projects\`. LectureForgeC/CF/CBF/CA are explicit,
isolated acceptance namespaces, not production defaults. All runtime evidence,
PPTX copies and media remain outside Git. Existing Solar Thermal takes and paid
avatar files are copied as fixtures; no provider-generation proof is repeated.

Run the suites with fresh short roots and unused ports:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/v2/Test-SelectionAuthorization.ps1 -RuntimeRoot "$env:LOCALAPPDATA\LectureForgeCF" -Port 8775
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/v2/Test-SelectionAuthorizationHttp.ps1 -RuntimeRoot "$env:LOCALAPPDATA\LectureForgeCF" -Port 8775
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/v2/Test-NarrationStudio.ps1 -RuntimeRoot "$env:LOCALAPPDATA\LectureForgeCBF" -Port 8777
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/v2/Test-ProductionStudio.ps1 -RuntimeRoot "$env:LOCALAPPDATA\LectureForgeCA" -Port 8776
```

The C fixture suite requires the verified B fixture project
`LectureForgeB/Projects/4b9d067085b3` and existing Solar Thermal production history.
It leaves its server running for supplemental HTTP tests. Test outputs contain
local evidence only and are not source files. No API keys are written to snapshots,
presets, browser code or test reports.

## Final acceptance results

Final pre-commit rerun: **63/63 checks passed** (C 25/25, supplemental HTTP 4/4,
B narration 19/19, A foundation 15/15). PowerShell parsing (six changed/new
scripts/modules), both JavaScript syntax checks and `git diff --check` passed.

Evidence stays outside Git:

- `LectureForgeCF/results.json` and `LectureForgeCF/http-results.json`
- `LectureForgeCBF/Tests/c4f29b53/results.json`
- `LectureForgeCA/Tests/abde51e1/results.json`

Confirmed: immutable snapshot bytes after duplicate confirmation and restart;
missing/stale/invalid selection blocks; stale-tab rejection; consumed-record
rejection; unchanged 39 narration hashes; editable narration intent; exact Slide 8
prepared-pause binding; rejected Slide 4 history excluded. Unconfirmed review and
confirmed authorization both survive restart. Supplemental HTTP checks leave their
fixture awaiting review after testing CHANGE; this does not alter old snapshots.

ElevenLabs calls: **0**. HeyGen calls: **0**. No paid assets regenerated.
Canonical source SHA-256 remained
`6340C137FD85021957721341298D640C8FAF59D36507D904DF09948DCB576C01`.

Known limits: authorization consumption and provider recovery are Milestone D;
final PowerPoint assembly is a later distinct stage. The legacy reusable catalog
currently requires explicit population, exercised here only through fixture setup.
