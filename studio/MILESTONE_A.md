# Production Studio — Milestone A checkpoint

Implemented: launcher, prerequisite checks, interactive-session server and idle worker, default-browser launch, New/Open Lecture, PPTX upload/copy and package inspection, saved Renewable Energy preset, script import/editor, enabled slides, pause metadata, atomic versioned project state, Open Project Folder, and clean Stop LectureForge.

Start from the repository directory:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Start-LectureForge.ps1"
```

The browser opens at http://127.0.0.1:8770/. Stop LectureForge saves pending script edits and stops the local server and worker. Close the launcher window without stopping the background processes; run the same launcher again to return. This is an ordinary logged-in desktop application, not a Windows service. No Codex runtime is required.

Runtime root: `%LOCALAPPDATA%\LectureForge`.

Projects: `%LOCALAPPDATA%\LectureForge\Projects\<12-character-id>\`.

Each project currently contains `source.pptx` and `project.json`. Original PPTX files are never opened for writing. Source hash, slide order/titles/dimensions, copied preset, scripts, slide enablement, and optional post-speech silence are saved. Scripts use `## Slide N` or `Slide N` headings followed by narration text. A pause is only metadata in this milestone; no audio preparation runs.

The launcher reads API keys from the inherited Windows environment and reports presence only. It also checks registered PowerPoint and FFmpeg/FFprobe availability. Keys are not stored in projects, presets, browser code, or logs. Runtime state and startup diagnostics stay outside the repository. Project updates reject stale revisions and use atomic replacement. Local HTTP mutations require the session token and accept only the local host/origin.

## Preserved refinements

- Future background production stops at AVATAR READY. It will not open/save PowerPoint per slide.
- Create Recording PowerPoint will be a distinct explicit final stage, using one serialized production session, one save, then reopen/validation.
- Preset placement is normalized from the approved 960x540-point Solar Thermal placement (840,438,110,96.25). Future per-slide overrides are insertion/preview metadata only.
- The exact approved 70-pixel torso fade was promoted into shared `WhiteAvatar.psm1`. No conversion was run.
- Existing `generate-narration.ps1` diagnostics and the prior Recording path repair were not modified.
- No generated production-specific modules were added to source control.

## Acceptance evidence

`tests/v2/Test-ProductionStudio.ps1` passed 15 checks using a copied Solar Thermal PPTX, including Unicode/exact script preservation, source hash and slide inspection, per-slide settings, stale revision rejection, token enforcement, absent generation endpoint, clean server/worker stop, and identical state after launcher restart. Evidence:

`C:\Users\swomi\AppData\Local\LectureForge\Tests\8e2ab64c\results.json`

Passing project:

`C:\Users\swomi\AppData\Local\LectureForge\Projects\883e079bd9d3\`

Default-browser launch, New Lecture/preset controls, Open Existing Lecture, restored editor, and Open Project Folder were checked in desktop Edge/File Explorer. Earlier failed test fixtures and startup diagnostics were retained.

Canonical Solar Thermal SHA-256 remains:

`6340C137FD85021957721341298D640C8FAF59D36507D904DF09948DCB576C01`

Provider calls: **0**. Narration generation, HeyGen submission, avatar conversion, and final PowerPoint assembly are not implemented/enabled in this Studio milestone. No generated paid media was modified. No commit was made.

The automated restart test is an application stop/relaunch test; no physical Windows reboot was performed. Placement controls, thumbnails, production queues, and final-output export belong to later milestones. The initial local server accepts PPTX uploads up to 100 MB.
