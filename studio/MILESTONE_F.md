# Milestone F: V2 dress rehearsal and packaging

Baseline: `907f25aa1166e8baebcdfa2c2a41aa2a94bf0577` on
`lectureforge-v2-production-studio`. This milestone adds no production stage and
does not replace the proven pipeline. Milestone F received final instructor approval;
the release checkpoint reruns validation before committing and tagging V2.0.

## Package

- Root `LectureForge.cmd`: double-click entry point to the existing launcher;
  quoted repository path, visible prerequisite failures, default browser, and
  reuse of an existing instance. No installer, service, shortcut creation, or
  persistent security-policy change.
- `LECTUREFORGE_USER_GUIDE.md`: concise instructor workflow, costs, recovery,
  file locations, optional manual shortcut, and accepted V2 limits.
- `LECTUREFORGE_DEVELOPER_GUIDE.md`: architecture, state, safety, baselines,
  tests, and safe debugging, separate from instructor instructions.
- Studio's stopped message now points to the double-click launcher.

## Test setup and boundaries

`Prepare-DressRehearsal.ps1` creates an isolated source/scripts copy under
`%LOCALAPPDATA%\LectureForgeF`. After the UI creates the project and imports
scripts/pause metadata, it imports 42 existing verified narration candidates
(Slide 1's three plus Slides 2-14's 39), approved white avatars, and rejected
historical assets. This helper is test setup only, not an instructor action or a
new production import feature. It creates no selection or authorization.

`Test-DressRehearsal.cjs` uses a dedicated visible Edge profile and actual Studio
controls. Generate 3 Takes proves successful-asset reuse with zero new requests;
new paid transport is intentionally not retested. All narration choices, CHANGE,
selection review, explicit zero-new-job confirmation, worker reuse/validation,
assembly, Open actions, stop, restart, and project reopening use Studio boundaries.
Provider transport is disabled in the isolated fixture root.

The test can resume its preserved project after a harness failure. Initial
harness corrections covered Windows CMD quoting, asynchronous screen/button
readiness, and waiting for actual audio clock advancement after browser startup.
No production audio or authorization workaround was added. The A regression's
stop-state read now uses the existing shared JSON reader to avoid a transient
read race during atomic state replacement.

The 14-slide rehearsal exposed real project-lock contention: three concurrent
local validators plus browser reads could exceed the former ten-second lock
wait and produce `Narration state busy` exceptions. A small engine adjustment
extends that bounded wait to sixty seconds. It does not repeat provider calls,
change authorization, alter media, or change concurrency. The fixture reloads
the change through a clean UI stop/restart, retaining the existing job identities,
and uses the existing local-validation retry control for affected jobs.

Two exceptions occurred before a reusable output path had been recorded. The
existing local-validation retry now also accepts that exact approved reusable
asset: it rechecks authorization/job binding and the media hash, records the
existing path, and enters Validating directly. Studio exposes the same retry
button for this case. No provider resume/download or white-media rebuild is used.

## Regression evidence

| Suite | Passed | Local evidence |
| --- | ---: | --- |
| A | 15 | LectureForgeF2A/Tests/1f4d0eac/results.json |
| B | 19 | LectureForgeF2B/Tests/fc7cde91/results.json |
| C | 25 | LectureForgeF3C/results.json |
| D guards | 11 | LectureForgeF3G/results.json |
| D recovery | 13 | LectureForgeDF/recovery-results.json |
| E integration | 22 | LectureForgeF3E/results.json |
| E supplemental | 8 | LectureForgeF3E/supplemental-results.json |
| E browser | 8 | LectureForgeF3E/browser-results.json |
| E final restart | 4 | LectureForgeF3E/checkpoint-results.json |

All 125 relevant A-E checks passed. Generated files/evidence remain outside Git.
The E regression Recording is
`%LOCALAPPDATA%\LectureForgeF3E\Output\Solar Thermal_RECORDING_12ef3f4c.pptx`.

These are the final-code results; earlier fixture evidence remains preserved.
A/C stop-state test readers were corrected to use the existing shared reader.
An attempted overlap between E and F assembly was correctly refused by the
existing global serialization guard. F's failed local attempt is retained and
its UI Retry Assembly is exercised after the E test server stops.

Canonical source SHA-256:
`6340C137FD85021957721341298D640C8FAF59D36507D904DF09948DCB576C01`.
Master and the V1 baseline tag remain unchanged.

## Accepted limits

Live provider transport is not exercised in F (both provider call counts must
remain zero). Desktop PowerPoint and local prerequisites are still required.
Native folder-dialog selection remains a manual acceptance limitation; explicit
output folders are proven and limited to 160 characters. Technical validation
does not claim subjective teaching or visual quality. The instructor performs
the final cursor/laser/ink recording and export in PowerPoint. No desktop shortcut
was automatically created.

## Final checkpoint

F: **30/30 checks passed**. A-E: **125/125 relevant checks passed**. Total: **155**.
Additional packaging checks passed for 40 PowerShell files, 7 JavaScript files,
61 credential-scanned text files, the launcher, and Git whitespace.

Studio: `http://127.0.0.1:8797/`, left running for instructor review.
Fixture project: `%LOCALAPPDATA%\LectureForgeF\Projects\51a0c9465c73`.
Recording: `%LOCALAPPDATA%\LectureForgeF\Output\Solar Thermal V2 Dress Rehearsal_RECORDING.pptx`.

The derivative contains 14 native slides and 14 separate embedded autoplay avatar
objects, saved in one assembly session and reopened successfully. Slide 4 retains
approved Take 2; rejected Take 1 remains history. Slide 8 retains its authoritative
prepared six-second pause. Source SHA-256 before and after is the value above.

All 42 imported narration candidates and approved/historical avatar file hashes
match both their original fixtures and their project copies. ElevenLabs calls: 0.
HeyGen calls: 0. No narration or avatar media was regenerated.

Clean Stop LectureForge, launcher restart, and Open Existing Lecture passed.
Project, narration, authorization/review snapshots, ready-job records, assembly
state, and final Recording hash remain byte-identical across the final restart.
Open Recording PowerPoint and Open Output Folder passed through the browser UI.
The final screenshot shows the correct project identity, preset, and 14 ready
narration slides without provider requests.

Evidence: `LectureForgeF/results.json` (30 browser checks), `verified.json`
(complete-deck validation and recovery evidence), `before-restart.json`, and
`dress-rehearsal.png`. Test setup/resume and defect corrections are described
above; ordinary instructor operations use Studio, without JSON edits or Codex.

No unresolved production exception remains. Master and the V1 baseline tag remain
untouched; the approved release is committed only on the V2 production branch.

## Repeatable release recovery checks

The F browser harness accepts `--recovery-checks` for a fresh release rehearsal.
After verifying all reusable avatars, it cleanly stops its isolated Studio and
injects a local-validation exception before a reusable output path is recorded.
Restart must preserve the authorized job IDs, and the browser's targeted local
retry must return that existing asset to Avatar Ready. A second stopped-fixture
setup simulates a dead local assembly executor; Retry Assembly must build a fresh
derivative. These tests replace dependence on incidental lock contention or COM
overlap. Only test metadata changes, with provider transport disabled; media and
immutable authorization records remain intact. No production feature was added.

The release rerun also exposed Edge's deferred audio loading when the test window
was occluded. The F harness restores its own window, auditions each existing take
with a user gesture, and disables native-window occlusion throttling only in its
isolated test Edge process. Playback and seeking then pass. The instructor launcher,
Studio media handling, and Windows/security configuration are unchanged.

## Final release validation

After instructor approval, all **155/155** checks were rerun successfully:
A 15, B 19, C 25, D guards/recovery 24, E 42, F 30. Evidence is under
`%LOCALAPPDATA%`: `LFReleaseA/Tests/8a0ee3ea`,
`LFReleaseB/Tests/84b952bd`, `LFReleaseC`, `LFReleaseG`,
`LectureForgeDF/recovery-results.json`, `LFReleaseE`, and `LFReleaseF`.
An E browser CDP timeout passed on an isolated rerun against the same validated
deck. F resumed its preserved fixture after the browser-harness corrections
above. Its final Stop assertion uses durable server/worker state rather than a
transient message that an in-flight browser request can replace. Completed
recovery checks are not reinjected on resume. No successful media was regenerated.

Release F project: `LFReleaseF/Projects/c835765d79f3`.
Studio: `http://127.0.0.1:8798/`.
Recording: `LFReleaseF/Output/Solar Thermal V2 Dress Rehearsal_RECORDING.pptx`.
All 14 native slides and 14 separate embedded avatars pass validation; exact
selection/authorization/job/Recording hashes survive restart. Canonical source
hash remains the value above. ElevenLabs calls **0**; HeyGen calls **0**.

PowerShell syntax (40 files), JavaScript syntax (7 files), credential-literal
scan (61 text files), launcher validation, and Git whitespace checks pass.
Release tag: annotated `lectureforge-v2.0`, on the V2 production branch only.
Master and `lectureforge-v1-baseline` retain their documented targets.
