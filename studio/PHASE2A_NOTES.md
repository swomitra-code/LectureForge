# Phase 2A: one offline Solar Thermal slide

This is a technical engine proof only. It adds no Studio UI, approval workflow,
locking, provider integration, stable export publication, or lecture assembly.
The approved architecture is unchanged. No V1 script is called or modified.

## Inputs and local project

- Source: `source/RE-M2-Solar-Thermal-Overview.pptx`, 22,196,828 bytes,
  14 slides, 16:9 (12,192,000 x 6,858,000 EMU).
- Source SHA-256:
  `6340C137FD85021957721341298D640C8FAF59D36507D904DF09948DCB576C01`.
- Explicit narration fixture: `media/energy-management/slide-1-take-1.mp3`.
  This was selected in the retained Energy Management manifest. It is not approved
  Solar Thermal narration.
- Avatar fixture: `media/energy-management/slide-1-SKM-BLUE-7-production-transparent.webm`.
  This is a technical fixture only. No Solar Thermal lip-sync acceptance is claimed.

`Project.psm1` copies these local files byte-for-byte into a new project directory,
checks hashes, and records their provenance. Default project location:
`work/v2-phase2a/solar-thermal/`. That location is already ignored by V1's `.gitignore`.
Binary assets and generated proof records are not intended for Git.

The minimal `project.json` records source geometry/hashes, explicit media bindings,
Slide 1, crop and independent candidate placement metadata. It explicitly prohibits
provider calls and marks the placement and lip-sync as unapproved. There is no
Phase 3 revision/approval state implementation.

## Run a fresh proof

Use Windows PowerShell with installed desktop PowerPoint and FFmpeg/FFprobe on PATH.
No packages or dependencies are installed by these scripts.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\studio\Invoke-Phase2A.ps1
```

The launcher requires a new project directory and refuses to overwrite an existing
project. For a second independent run, supply `-ProjectDirectory` with a new path
under `work/`. Existing project recomposition is a separate command:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module .\studio\Render.psm1 -Force
Build-V2Slide .\work\v2-phase2a\solar-thermal\project.json `
  slide_01_another_candidate @{ x=0.82; y=0.68; width=0.1458333333333333; approved=$false }
```

The sample coordinates are provisional, not a permanently approved project default.
Every candidate name must be unique. No existing candidate or source is overwritten.

## Rendering and composition contract

`Export-V2Slide` opens the imported PPTX read-only, exports Slide 1 to a 1920x1080
PNG, and closes only that presentation without saving. It does not quit an existing
user PowerPoint application. Source hashes are checked in cleanup and COM/package
slide counts are recorded. No save/embed or slide-object mutation code exists.

`Build-V2Slide` does not invoke COM. It validates the imported asset hashes and cached
render, checks full media decoding and decoded alpha, then invokes FFmpeg with:

- input 0: static slide PNG;
- input 1: transparent avatar WebM, video only;
- input 2: the explicitly bound narration;
- output maps: `[v]` and `2:a:0`, with no automatic audio selection or `-shortest`.

The V1 crop/despill/alpha-restore technique is retained. Crop is in avatar pixels;
placement uses normalized slide x/y/width and derives height from crop aspect ratio.
The output explicitly converts to limited-range yuv420p, H.264, 30 fps, 1920x1080,
and AAC (192 kbps, 48 kHz stereo). Audio is neither sourced from HeyGen nor modified
with pause editing, speed changes, or normalization.

The narration's probed duration determines output length. Fixtures must differ in
duration by no more than one second. A short avatar can hold its last video frame;
this is video-only tail handling and does not establish lip-sync acceptance. The
80 ms technical duration tolerance accounts for frame/packet boundaries. This proof
does not implement the Slide 8 pause or any new narration preparation.

## Validation

Per-candidate JSON records include exact FFmpeg arguments, input/output hashes,
placement, profile/decode checks, and an independently encoded narration reference.
The final AAC packet hash must equal the reference AAC packet hash. This verifies
the actual output audio, not merely the command string. It does not claim that AAC
bytes equal the input MP3 bytes.

The MP4 itself is opaque, as intended. Transparency is consumed during composition:
the slide background remains visible around the presenter. Tests compare sampled
background pixels through fully transparent portions of the crop against the source
PNG, allowing for lossy H.264 differences. Human edge/placement review is still needed.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\v2\Test-Phase2A.ps1
```

For an existing candidate with another name, pass `-BaseCandidateName`. The tests
create a second full-length composition with changed placement, retain its artifacts,
and verify that source/render/media/project hashes and final audio remain unchanged.
They also exercise rejection of a changed source hash, wrong dimensions, wrong
duration and off-slide placement. Source-code checks prohibit network/provider and
PowerPoint save/embed calls. No network API is part of the execution path.

The initial local run exposed FFmpeg's full-range `yuvj420p` output tagging despite
the original pixel-format request. The validator rejected `slide_01_candidate_001.mp4`.
That rejected artifact is retained; explicit limited-range conversion is now used.
Candidate 002 passed all 23 technical checks. Visual review then tightened the
provisional crop from 640x620 to 640x400 for head-and-shoulders framing and moved
the provisional y coordinate from 0.70 to 0.78. No source/media asset changed.
Use `-BaseCandidateName slide_01_candidate_003` for the final local checkpoint.

The final local run passed 23 focused checks, including a second full composition.
Candidate 003 is 32.833333 seconds versus narration duration 32.833016 seconds;
its AAC packet hash matches the independently encoded narration. The alpha test
sampled 34 transparent locations and found a maximum background channel difference
of 3 (allowed: 20). Visual spot checks at 0, 16, and 32 seconds found readable text,
head-and-shoulders framing, and no covered text or opaque presenter rectangle.
These are technical observations, not instructor approval. Machine evidence is in
`work/v2-phase2a/solar-thermal/validation/phase2a-tests.json`,
`source-render.json`, and `slide_01_candidate_003.json`.

## Boundaries and known limitations

- Instructor technical approval of candidate 003 is recorded. Project defaults remain
  unset; nothing is published to stable approved exports.
- The narration and presenter are unrelated technical fixtures. No final Solar
  Thermal content or lip-sync approval is implied.
- Static slide rendering only: animations/builds/transitions are not captured.
- COM export is synchronous and requires a responsive interactive Windows desktop.
  A general worker timeout/recovery system is outside Phase 2A; do not kill unrelated
  PowerPoint sessions to recover a failed proof.
- The engine refuses changed inputs and existing output names rather than providing
  project migration, regeneration scheduling or approval state.
- Hash/decode/pixel tests do not replace instructor listening and visual review.
- No API credentials are read and no ElevenLabs or HeyGen API is contacted.
- Phase 2A remains a standalone proof. The instructor has separately authorized the white-avatar checkpoint; no later phases or provider calls.

## Accepted Recording PPTX clarification

The instructor approved native slides plus separate compact white-background avatar
MP4s, with independent movement/resizing and autoplay. Full-slide composites remain
standalone outputs and must not cover native content in a Recording PPTX.
The Phase 2A engine remains read-only for canonical sources and routes explicit
narration audio. See RECORDING_PHASE2_NOTES.md for the accepted separate packager.
