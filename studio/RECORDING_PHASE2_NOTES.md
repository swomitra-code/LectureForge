# Accepted V2.0 Recording PowerPoint checkpoint

The instructor approved native slides plus a separate white-background avatar
video, independently movable/resizable and autoplaying with narration. The white
background blends into the reserved white area. Source, conversion, embedding and
desktop slideshow motion passed; instructor subsequently confirmed visible motion.
The earlier apparent freeze is treated as a transient PowerPoint session condition.

Full-slide composite video is rejected for Recording PPTX. PowerPoint alpha playback
will not be pursued further in V2.0. Standalone MP4 composition remains supported.

## Accepted local proof

- Source: `source/RE-M2-Solar-Thermal-Overview.pptx`, 14 slides.
- Source SHA-256: `6340C137FD85021957721341298D640C8FAF59D36507D904DF09948DCB576C01`.
- Tile: `work/avatar-white-proof-20260912/slide-1-avatar-white.mp4`.
- Tile SHA-256: `3D9420617F53F0172021B22665A81BCB4D0FA83683CF06C29C5954AE2FA8E2B8`.
- Review deck: `output/RE-M2-Solar-Thermal-Overview_RECORDING_AVATAR_WHITE_TEST_20260912.pptx`.
- Tile: 640x560, H.264/AAC, 25 fps, RGB 255,255,255 background, no border/shadow.
- Placement in points: left=438, top=337, width=110, height=96.25 on a 960x540 slide.
  This is inside the white content panel above the closing question. The far-right
  photograph is unchanged. Placement is accepted for this technical test only.
- Existing WebM audio was explicitly retained for that technical fixture and encoded
  to AAC. It is not final Solar Thermal narration approval.

## Reproduce the separate Recording proof

`Recording.psm1` accepts only the compact white-avatar profile, an approved tile hash,
explicit placement and a new output name. It refuses full-slide placement, existing
outputs and changed input hashes. It copies the native source, adds only Slide 1's
avatar, saves the derivative, closes/reopens it, and verifies the package and desktop
player. It never composes media, calls providers or saves the canonical source.

```powershell
Import-Module .\studio\Recording.psm1 -Force
New-V2RecordingProof `
  -Source .\source\RE-M2-Solar-Thermal-Overview.pptx `
  -ApprovedVideo .\work\avatar-white-proof-20260912\slide-1-avatar-white.mp4 `
  -ApprovedVideoSha256 3D9420617F53F0172021B22665A81BCB4D0FA83683CF06C29C5954AE2FA8E2B8 `
  -Placement @{left=438;top=337;width=110;height=96.25} `
  -Output .\output\SolarThermal_REVIEW_RECORDING_AVATAR_WHITE.pptx `
  -EvidenceDirectory .\work\recording-white-review
```

Use new paths for each run. Preserve all prior instructor-reviewed files.

## Authoritative narration for future real production

`WhiteAvatar.psm1` records the locally proven crop/despill/white composition technique.
`New-V2WhiteAvatar` requires separate avatar and narration paths with matching hashes.
It maps only filtered avatar video and explicit narration input 1:a:0, preserves the
narration duration, and compares final AAC packets with an independent encode of that
narration. There is no automatic avatar-audio fallback and no provider call.
The fixed crop/profile is intentionally limited to this proof, not a generalized UI.

## Regression checks

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\v2\Test-Phase2A.ps1 `
  -BaseCandidateName slide_01_candidate_003
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\v2\Test-Recording.ps1
```

The Recording suite creates fresh work-only artifacts and checks source/deck/media
immutability, authoritative audio, white pixels at several timestamps, decoded
mouth-frame changes, hash identity after embedding, rejection of full-slide media,
independent move/resize, autoplay after reopen and live player advancement. It also
compares all 13 unrelated slide renderings and recursive native-object signatures.
Local binary fixtures must exist; tests make no network calls to obtain them.

## Acceptance and limits

White blending, technical placement, independent editing and presenter motion are
accepted. Final audible narration confirmation is required during real Slide 1
production. Machine frame differences are technical checks, not lip-sync scoring.
The straight lower-torso crop edge is a minor framing improvement for real production.
Cursor/laser/ink recording, replay and export remain instructor workflow checks;
do not infer their success from autoplay properties. Desktop COM requires a responsive
Windows session; stop on playback/embedding failures rather than changing global
settings or killing unrelated PowerPoint sessions.

This checkpoint includes only code/docs/tests, not source decks, generated binaries,
provider media or work evidence. No Studio UI or later phase is part of it.
