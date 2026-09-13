# Milestone E: Recording PowerPoint checkpoint

The Studio assembles a derivative Recording PowerPoint only after every enabled
slide passes current narration, authorization, Avatar Ready, audio, pause, hash,
exception and placement checks. Provider calls during E: ElevenLabs **0**, HeyGen
**0**. All paid narration and avatar assets were reused without regeneration.

## Implementation

- `Assembly.psm1`: readiness, versioned placement overrides, immutable attempt
  manifest, durable local state, retry/recovery and validated-output open actions.
- `Invoke-Assembly.ps1`: local background executor with a session-wide COM mutex,
  exact manifest rechecks before assembly and publication, and atomic export.
- `Recording.psm1`: extends the existing package validator to the requested slide
  and adds one batch insertion/save followed by read-only reopen validation.
  Existing Slide 1 proof entry points remain available.
- `Worker.ps1`, `Server.ps1`, `Start-LectureForge.ps1`: assembly dispatch/status
  routes and E startup identity. Child processes use named script files and visible
  windows, following the instructor's security-compatible execution requirement.
- `Select-OutputFolder.ps1`, `web/assembly.js`, `web/index.html`: native folder
  chooser, output-folder entry, readiness summary, placement and reset controls,
  Create/Retry progress, final path, Open Recording and Open Output Folder.

Placement overrides live in `placement.json`, separate from provider authorization.
Changing placement never modifies paid-media inputs or regenerates media. Values
are normalized to slide dimensions, preserve the 640:560 aspect ratio, remain
inside the slide and are limited to 30% width/height. Existing per-slide overrides
and project defaults are inherited when no assembly override exists.

Each attempt has short paths `r/<8hex>/input.json`, `deck.pptx` and `v/sN/embedded.mp4`.
Parent directories are created before extraction. `assembly.json` retains the
current attempt, previous attempts and last successful output. A failed/dead local
executor requires Retry Assembly and a fresh derivative. Successful prior outputs
and Avatar Ready records are preserved. Repeated requests for the identical
completed recording return that recording. Existing filenames receive a revision
suffix. Export copies are hash-verified before atomic publication, never overwritten.

The production worker runs one assembly child at a time. It does not start its
thumbnail child during assembly. The assembly executor's named mutex serializes
assemblies across runtime roots. Cleanup closes only the presentations it opens;
it never kills PowerPoint processes or closes instructor-owned decks.

## Complete-deck validation

The copied source is opened once for all insertions and saved once. The derivative
is reopened read-only. Every native shape signature and native-only slide render
is compared with the pre-insertion copy. Video objects have no border/shadow, use
the approved geometry and autoplay, and pass unsaved independent movement/resize
probes. Package relationships must resolve to embedded video bytes with exact
Avatar Ready hashes. Each extracted video/audio is decoded. Selected narration
and preparation binding follow the independently validated D media hash; no audio
is replaced or trimmed in E. The source hash is checked before and after.

The first integration attempt exposed a COM cast error while restoring a geometry
value during the unsaved movement probe. Explicit Single casts, matching the proven
Slide 1 validator, fixed it. The failed derivative remains historical. A real Retry
Assembly then passed the complete deck validation. No provider asset was rebuilt.

## Acceptance evidence

Runtime root: `%LOCALAPPDATA%\LectureForgeE` (isolated acceptance namespace).
Normal production root remains `%LOCALAPPDATA%\LectureForge\Projects`.
Project: `9a6884f7a78e`. Studio: http://127.0.0.1:8788/.

Final fixture deck:
`C:\Users\swomi\AppData\Local\LectureForgeE\Output\Solar Thermal_RECORDING_6c7a4c15.pptx`

| Suite | Passed | Evidence under LOCALAPPDATA |
|---|---:|---|
| E initial readiness, placement and pre-assembly recovery checks | 8 | LectureForgeE/results.json |
| E repaired assembly and recovery | 15 | LectureForgeE/finish-results.json |
| E subset/disabled-slide and stale-request checks | 8 | LectureForgeE/supplemental-results.json |
| E visible Edge browser checks | 8 | LectureForgeE/browser-results.json |
| E final restart/integrity checkpoint | 4 | LectureForgeE/checkpoint-results.json |
| D queue/submission guards | 11 | LectureForgeEG/results.json |
| D provider-record/server/worker recovery | 13 | LectureForgeDF/recovery-results.json |
| C authorization regression | 25 | LectureForgeEC/results.json |
| B narration regression | 19 | LectureForgeEB/Tests/ecf66f6e/results.json |
| A foundation regression | 15 | LectureForgeEA/Tests/0e822b73/results.json |

E: 43 passing checks across these scripts. Relevant A-D regressions: 83 passing
checks. Source checks additionally parse changed PowerShell/JavaScript, scan for
credential literals and run `git diff --check`. No unresolved production exception.

The final fixture contains 14 native slides and 14 independently movable/resizable
embedded autoplay avatars. Slide 4 uses approved Take 2. Slide 8 binds to
`E234676D7F92B42D20E98A3CB6CADF5BAA89595C2DCA906005EB7913607D69A5`
and its exact 6.000-second prepared pause. The fixture deliberately retains a
Slide 6 test override (left .84, top .80); all other slides inherit the approved
default (left .875, top .8111111111). A separate one-avatar subset deck proved
all 13 unproduced slides unchanged. The completed full-deck state was unaffected.

Source SHA-256 before/after:
`6340C137FD85021957721341298D640C8FAF59D36507D904DF09948DCB576C01`.

Open Recording PowerPoint and Open Output Folder passed HTTP and actual browser
control tests. Browser refresh and Studio/server restart preserve the completed
path, output hash and byte-identical state. Screenshots and media remain in the
ignored local runtime root. Fixture avatar transport remains disabled and paused.

## Limits and review boundary

- Technical checks do not claim subjective teaching quality, lip-sync quality,
  or absence of content obstruction for arbitrary instructor placement overrides.
- The native folder chooser is implemented; automated acceptance exercised explicit
  output-folder paths and Open actions, not manual selection inside the Windows dialog.
- Output folders are limited to 160 characters, with short generated filenames,
  to keep the tested Windows path handling safe. New revisions preserve prior decks.
- The original full fixture script halted at the initial COM validation error;
  `Test-AssemblyFinish.ps1` continued acceptance on the same preserved fixture after
  the fix, rather than recreating any paid assets or discarding evidence.
- The restored worker is readable/writable. No antivirus configuration or Windows
  security settings were changed. No blocked long inline rewrite was replayed.

## Instructor-approved commit checkpoint

The fresh checkpoint fixture uses `%LOCALAPPDATA%/LectureForgeEF` and runs the
full assembly suite uninterrupted: 22 checks, plus 8 supplemental, 8 actual-browser,
and 4 restart/checkpoint checks (42 passing checks). This replaces the earlier
split 8-check initial run and 15-check recovery continuation; the original
43-check acceptance evidence remains preserved above.

Two final-checkpoint restart attempts timed out during local startup with aborted
loopback HTTP responses. After cleanly stopping the completed regression fixtures,
the isolated launcher and clean restart passed without production code changes.

Relevant A-D regressions were rerun: A 15 (`LectureForgeEFA`), B 19
(`LectureForgeEFB`), C 25 (`LectureForgeEFC`), D guards 11 (`LectureForgeEFG`),
and D recovery 13 (`LectureForgeDF`): 83 passing checks. All evidence and media
remain outside Git. ElevenLabs calls: 0. HeyGen calls: 0.

The new full fixture output is
`%LOCALAPPDATA%/LectureForgeEF/Output/Solar Thermal_RECORDING_beb095e2.pptx`.
The canonical source retains the SHA-256 recorded above. Master remains
`e8f9ff57c2dcf54ac81d65f34ced9c8b8c71c9cd`; the V1 baseline tag remains
`a3bbc92a0f499e29df0ea966c2d1794ed95fb249`.

Milestone E is instructor-approved for commit. Milestone F has not begun.
