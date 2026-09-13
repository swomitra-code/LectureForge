# LectureForge V2 maintenance and recovery

Instructor operations belong in [the user guide](LECTUREFORGE_USER_GUIDE.md).
Codex is not a production dependency. Do not turn debugging steps into instructor
workflow or add another provider approval between authorized successful stages.

## Baselines and runtime

- Branch: `lectureforge-v2-production-studio`.
- Milestone E baseline: `907f25aa1166e8baebcdfa2c2a41aa2a94bf0577`.
- Preserve master `e8f9ff57c2dcf54ac81d65f34ced9c8b8c71c9cd` and
  `lectureforge-v1-baseline` target `a3bbc92a0f499e29df0ea966c2d1794ed95fb249`.
- Normal root: `%LOCALAPPDATA%\LectureForge`; projects are under `Projects\<id>`.
  Suffixed roots such as `LectureForgeF` are isolated test evidence, not production
  path changes. Do not move working projects under Dropbox or the repository.
- `LectureForge.cmd` invokes `Start-LectureForge.ps1` with a process-only execution
  policy. No service, installer, security exclusions, or persistent policy changes.
  PowerPoint requires the logged-in interactive Windows session.

## Components and state

`Server.ps1` serves plain HTML/JS and loopback JSON endpoints. Mutations require a
session token; revisions reject stale tabs. `Worker.ps1` dispatches durable work
to named PowerShell scripts. Browser reads do not call providers.

| Stage | Modules | Durable state |
| --- | --- | --- |
| Project import | Production / Project | project.json, source.pptx, source hash, copied preset |
| Narration | Narration / NarrationProvider / Invoke-NarrationTake | narration.json, immutable n/ revisions and takes |
| Selection authorization | Authorization | selection-review.json, immutable approvals/ snapshots |
| Avatar queue | AvatarProduction / HeyGenProvider / Invoke-AvatarStep | avatar queue, a/ job records, avatars.json reuse registry |
| White media | WhiteAvatar | validated H.264/AAC output and authoritative audio evidence |
| Recording assembly | Assembly / Invoke-Assembly / Recording | placement.json, assembly.json, immutable r/ input and validation |

Use existing module entry points and test patterns rather than writing production
JSON by hand. Placement is separate from provider inputs. Assembly opens a fresh
source derivative, inserts all enabled videos in one serialized session, saves
once, and reopens for validation. Native slides remain native, with separate
embedded autoplay media. Short paths and parent-directory creation are required.

## Provider and source safety

Only environment variables supply credentials. Never log values, headers, or
credential-bearing requests. Preserve the existing sanitized provider diagnostics.
Narration selection is intent only. HeyGen work requires the exact immutable
confirmed snapshot. Persist submission intent before POST and ID before polling.
Unknown submission outcomes are exceptions, never automatic retries. Known jobs
resume by ID; historical/rejected assets remain history. New paid replacements
require explicit new authorization.

Selected ElevenLabs audio, including prepared pauses, is authoritative through
final media. Never substitute HeyGen audio or append pauses after approval.
Never regenerate successful paid media merely to test UI. Protect the canonical
source hash and prior valid Recording decks. COM cleanup owns only its own decks;
never kill unrelated instructor PowerPoint sessions.

## Tests

Run named scripts from the repository in an interactive Windows session, with
prerequisites available. Use fresh short roots and run COM suites sequentially.
Stop completed test servers before another restart suite to avoid local load.
Node.js is needed for the browser test harness, not for ordinary Studio production.
`Run-FRegressions.ps1` sequences the relevant offline suites and stops their servers;
its `-From C` continuation preserves already-passing A/B evidence after a test interruption.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/v2/Test-ProductionStudio.ps1 -RuntimeRoot "$env:LOCALAPPDATA\LectureForgeFA" -Port 8801
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/v2/Test-NarrationStudio.ps1 -RuntimeRoot "$env:LOCALAPPDATA\LectureForgeFB" -Port 8802
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/v2/Test-SelectionAuthorization.ps1 -RuntimeRoot "$env:LOCALAPPDATA\LectureForgeFC" -Port 8803
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/v2/Test-AvatarGuards.ps1 -RuntimeRoot "$env:LOCALAPPDATA\LectureForgeFG"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/v2/Test-AssemblyStudio.ps1 -RuntimeRoot "$env:LOCALAPPDATA\LectureForgeFE" -Port 8804
node tests/v2/Test-DressRehearsal.cjs "$env:LOCALAPPDATA\LectureForgeF" 8797
```

Follow assembly integration with `Test-AssemblySupplemental.ps1`,
`Test-AssemblyBrowser.cjs`, and `Test-AssemblyCheckpoint.ps1` using the same root
and port. `Test-AvatarRecovery.ps1` accepts a completed offline fixture root/id.
`Test-AssemblyFinish.ps1` is a continuation for a failed historical run, not a
second assembly of an already completed attempt. Do not run `Test-NarrationLive.ps1`
during zero-provider acceptance. The original AvatarStudio test includes a local
conversion; use guards/recovery plus F's full reuse worker path when media must
not be rebuilt. Run `Test-ProductionPackaging.ps1` for syntax/secret/whitespace checks.

F's fixture helper copies verified assets from the local completed Solar Thermal
fixture into a newly created Studio project. This is test setup, not a production
import feature. The browser test then uses real controls and production workers.
The Generate 3 Takes action is intentionally a no-op for existing valid takes;
live transport was proven earlier. Provider transport is disabled in the F root.
Test output, screenshots, decks, and fixture state stay outside Git.

For a fresh release rehearsal, add `--recovery-checks` to the browser command.
This also exercises restart with the same authorized job IDs, local-validation
retry before a reusable output path was recorded, and interrupted-assembly retry.
The fixture helper injects only local exception metadata while its Studio is
stopped and provider transport is disabled. Existing media and immutable
authorization snapshots remain intact. This makes recovery tests repeatable
without relying on incidental disk contention or overlapping COM sessions.

## Resume debugging safely

Read Git status/diff, runtime.json, project state, and existing checkpoint evidence
before changing anything. Verify source hashes and record the failing stage/slide.
Preserve completed jobs, media, selections, and immutable authorization records.
Reconcile uncertain jobs; retry only a failed local stage when appropriate.
The project lock has a bounded sixty-second wait for full-lecture validation
contention. Local-validation retry can recover an authorization-bound reusable
MP4 even if failure occurred before its output path was recorded; it verifies the
binding/hash and never routes that retry through provider submission or download.
Use direct patches or short named scripts. Do not use long inline PowerShell,
encoded commands, Python here-string rewrites, hidden development launches, or
antivirus changes. Inspect intended source files and scan for secrets before staging.
Do not commit checkpoint work until instructor review authorizes it.
