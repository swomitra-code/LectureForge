# LectureForge V1 Production Baseline

Status: approved production reference as of 2026-08-26.

This document freezes the proven Windows command-line workflow. The five production tools listed under **Baseline hashes** define V1 behavior. Do not refactor, merge, or change them without creating a new baseline.

## Proven workflow

`lectureforge.ps1` reads one manifest and resumes at the first incomplete stage:

1. For each enabled slide whose narration is enabled, inspect the three generated takes.
2. If narration is missing, call `generate-narration.ps1`, generate exactly three MP3 takes, and stop for instructor review.
3. If valid takes exist but `instructor_selected_take` is null, stop for instructor selection. Never choose automatically.
4. If a selected take exists and avatar generation is enabled but no valid avatar WebM is bound, call `generate-avatar.ps1` and wait for technical validation.
5. If all enabled slides have valid avatar media, call `build-deck.ps1`.
6. If the configured PPTX and validation report already prove all enabled slides complete, report them and make no provider or build calls.

The workflow is resumable. Valid deterministic outputs are reused, including recovery when provider output exists but a prior run ended before its manifest binding was saved.

## Required folders

```text
config/   JSON deck manifests
source/   original editable PPTX files
scripts/  exact instructor-approved narration scripts
media/    narration MP3s and transparent avatar WebMs
work/     rendered slide PNGs and temporary build artifacts
output/   composed MP4s, final PPTX files, and validation JSON
```

Run commands from the repository root. Keep the source PPTX unchanged; outputs always use a different path.

## Requirements and environment

- Windows desktop PowerPoint with COM automation available.
- PowerShell, FFmpeg, and FFprobe on `PATH`.
- Network access and provider account capacity for generation stages.
- `ELEVENLABS_API_KEY` for narration generation.
- `HEYGEN_API_KEY` for avatar generation.

API keys are read from the process environment and must never be printed, logged, or stored in files.

## Exact commands

Normal orchestration:

```powershell
.\lectureforge.ps1 .\config\module1.json
```

Instructor narration selection:

```powershell
.\select-narration.ps1 .\config\module1.json <slide-number> <take-number>
```

Individual proven stages, primarily for diagnosis or controlled production:

```powershell
.\generate-narration.ps1 .\config\module1.json
.\generate-avatar.ps1 .\config\module1.json
.\build-deck.ps1 .\config\module1.json
```

When local execution policy blocks scripts, use a process-only bypass:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\lectureforge.ps1 .\config\module1.json
```

## Manifest format

Paths are absolute or relative to the manifest unless a tool explicitly records a repository-relative `media/...` selection. The orchestrator resolves both forms and normalizes generated avatar bindings for the deck builder.

```json
{
  "source_pptx": "../source/input.pptx",
  "output_pptx": "../output/finished.pptx",
  "output_media_dir": "../output/finished-media",
  "defaults": {
    "presenter": {
      "crop": { "width": 640, "height": 620, "x": 870, "y": 80 },
      "scale": { "width": 280, "height": 271 },
      "position": { "x": 1620, "y": 430 }
    },
    "narration": {
      "voice_id": "PqwAVcrzgTuCTbgkDpWl",
      "model_id": "eleven_multilingual_v2",
      "output_dir": "../media",
      "stability": 0.34,
      "similarity_boost": 0.80,
      "style": 0.50,
      "use_speaker_boost": true,
      "speed": 1.00
    },
    "avatar_generation": {
      "avatar_name": "SKM-BLUE-7",
      "avatar_id": "f04bc24f9fd7480885c266c4f1d71f98",
      "output_dir": "media",
      "remove_background": true,
      "output_format": "webm",
      "aspect_ratio": "16:9",
      "resolution": "1080p",
      "burn_captions": false
    }
  },
  "slides": [
    {
      "slide_number": 15,
      "enabled": true,
      "avatar_webm": "../media/slide-15-SKM-BLUE-7-production-transparent.webm",
      "avatar_generation": { "enabled": true },
      "narration": {
        "enabled": true,
        "speech_script": "../scripts/slide-15.txt",
        "generated_takes": [
          "../media/slide-15-take-1.mp3",
          "../media/slide-15-take-2.mp3",
          "../media/slide-15-take-3.mp3"
        ],
        "instructor_selected_take": "media/slide-15-take-1.mp3"
      }
    }
  ]
}
```

Optional per-slide `presenter.scale` and `presenter.position` override the defaults only when needed to avoid meaningful slide content.

## Narration selection

Narration generation reads the approved script verbatim and produces:

```text
media/slide-<N>-take-1.mp3
media/slide-<N>-take-2.mp3
media/slide-<N>-take-3.mp3
```

Technical checks are limited to existence, nonzero size, successful decode, and readable positive duration. The instructor listens and selects one take:

```powershell
.\select-narration.ps1 .\config\module1.json 15 1
```

The selector accepts only takes 1–3, verifies that the take is recorded and decodes, updates only that slide's `instructor_selected_take`, and calls no downstream stage.

## ElevenLabs production defaults

- Voice ID: `PqwAVcrzgTuCTbgkDpWl`
- Model: `eleven_multilingual_v2`
- Output: MP3, 44.1 kHz, 128 kbps
- Stability: `0.34`
- Similarity boost: `0.80`
- Style: `0.50`
- Speaker boost: `true`
- Speed: `1.00`

No automatic script rewriting, take selection, WPM gate, or subjective scoring is permitted.

## HeyGen avatar defaults

- Avatar/look: `SKM-BLUE-7`
- Avatar ID: `f04bc24f9fd7480885c266c4f1d71f98`
- Input: exact instructor-selected MP3 uploaded as an asset
- Output: transparent WebM, 1920×1080, 16:9
- Background removal: `true`
- Burned captions: none
- Playback-speed alteration or voice regeneration: none

Technical checks require a nonempty, decodable video with audio, WebM alpha metadata, and duration within 5% or one second of the selected narration.

## PowerPoint production

For every enabled slide, `build-deck.ps1`:

1. Opens the original deck and renders the editable slide at 1920×1080.
2. Decodes the VP9 alpha avatar explicitly.
3. Uses the frozen presenter crop, RGB-only green despill, preserved alpha, Lanczos scale, and lower-right overlay.
4. Encodes a 1920×1080 H.264 MP4 with AAC audio, using the avatar's approved audio.
5. Copies the original deck to the configured output path.
6. Embeds—not links—the full-slide MP4 over the original editable slide.
7. Configures automatic playback and PlayOnEntry.
8. Saves, closes, reopens, and validates the desktop PowerPoint journey.

The original editable content remains underneath the media overlay. Unrelated slides are compared against source signatures. Final MP4s must be H.264/AAC, 1920×1080, and fully decodable.

## Bounded PowerPoint playback validation

Structural failures are never retried: missing media, unexpected media count, linked media, missing autoplay, missing PlayOnEntry, invalid MP4 structure, decode failure, slide-count changes, damaged underlays, or changed unrelated slides fail the build.

Live playback runs in a fresh PowerPoint/presentation/slideshow session. If the live Slide Show view alone fails to expose a media shape while edit-view structure and autoplay are valid, the validator makes exactly one retry in another fresh session. Maximum: two total playback probes. Any other playback failure is not generally retried, and two inconclusive missing-media observations fail validation.

The validation JSON records probe count and results for each slide.

## Human approval gates

1. Instructor approves the exact speech script.
2. Narration generation produces three takes and stops.
3. Instructor listens and selects one take; selection is never automatic.
4. The selected audio is used unchanged for avatar production.
5. Technical checks gate production. Subjective lip-sync and visual review remain optional instructor checks.

## Recovery and resume behavior

- Re-running `lectureforge.ps1` inspects the manifest and existing files before acting.
- Three valid narration takes plus no selection stops at the human gate without ElevenLabs calls.
- A selected take plus a valid avatar WebM skips HeyGen.
- Deterministic narration and avatar outputs can be recovered after an interrupted run without creating duplicate provider jobs.
- A PPTX is considered complete only when its validation report covers exactly the enabled slides and all required checks pass.
- Provider failures leave completed upstream assets intact. Build validation failure does not delete or roll back the saved candidate or MP4s.

## Known limitations

- Windows desktop PowerPoint is required; headless PowerPoint validation is not the V1 reference.
- PowerPoint COM live Slide Show state can be transient; V1 handles only the proven missing-live-media observation with one bounded retry.
- Provider calls require network access, credentials, account credits, and provider availability.
- Narration and visual quality remain human judgments; machine checks are technical only.
- V1 has no screen-cue animation layer, captions, web UI, database, provider abstraction, or automatic take selection.
- Output names are manifest-controlled and are not inferred from enabled slide numbers.

## Baseline hashes

SHA-256:

| Production tool | SHA-256 |
|---|---|
| `generate-narration.ps1` | `AC41943F44225FB05F951ACB0136AAA575C78B83190698ED67D19B7FBA5E3C05` |
| `select-narration.ps1` | `637451FA8BF7E7737659832640741A322F389E599301EB1C2B89FD72008B73BF` |
| `generate-avatar.ps1` | `832C3C3639CAB91103211EE4081BF4C125BB2B0FCFF740BD28B3A5FC19B4D3F2` |
| `build-deck.ps1` | `0616BE21ADF8CCAA4B31249947058440A463BCA7C8BD0B363A7B110634E632C9` |
| `lectureforge.ps1` | `0078BCC716951338D8896F7E0BAC13C3A0F73CD5A4B9D821F30532401764204C` |

Any hash change requires review and a new baseline version.

## Proposed optional screen-cue layer — not implemented

### Proposed manifest format

Coordinates use the final 1920×1080 slide pixel space. Times are seconds relative to the start of the approved narration.

```json
{
  "slide_number": 15,
  "screen_cues": {
    "enabled": true,
    "version": 1,
    "cues": [
      {
        "type": "highlight",
        "start_seconds": 4.2,
        "end_seconds": 8.0,
        "rect": { "x": 180, "y": 260, "width": 720, "height": 120 },
        "color": "#FFD54A",
        "line_width": 6
      },
      {
        "type": "arrow",
        "start_seconds": 9.0,
        "end_seconds": 12.5,
        "from": { "x": 1050, "y": 540 },
        "to": { "x": 1320, "y": 410 },
        "color": "#D32F2F",
        "line_width": 8
      },
      {
        "type": "spotlight",
        "start_seconds": 14.0,
        "end_seconds": 19.0,
        "rect": { "x": 120, "y": 180, "width": 980, "height": 560 },
        "dim_opacity": 0.55,
        "feather_pixels": 24
      },
      {
        "type": "zoom",
        "start_seconds": 22.0,
        "end_seconds": 28.0,
        "rect": { "x": 920, "y": 280, "width": 760, "height": 500 },
        "scale": 1.12,
        "easing": "smoothstep"
      }
    ]
  }
}
```

Allowed V1 cue types should be strictly limited to:

- `highlight` or `outline`: animated or static rectangle outline.
- `arrow` or `underline`: simple line geometry with optional arrowhead.
- `spotlight`: preserve one rectangle while dimming surrounding content.
- `zoom`: gentle crop/scale toward one rectangle, with no abrupt camera movement.

### Smallest safe implementation path

1. Define and validate the optional `screen_cues` schema. Reject unknown cue types, invalid coordinates, overlaps in time that the renderer cannot combine, and times outside the selected narration duration.
2. Add one isolated FFmpeg cue helper that converts the slide PNG plus cue JSON into a 1920×1080 cue-enhanced background stream or a transparent cue-overlay stream. It must not touch audio.
3. Add one guarded build branch for cue-enabled slides only. It supplies the helper output to final composition before the existing presenter overlay and retains the existing H.264/AAC encoding and embedding path.
4. When `screen_cues` is absent, null, disabled, or empty, execute the current V1 FFmpeg command and filter graph unchanged.
5. Extend technical validation only to confirm cue configuration validity, output decode, dimensions, duration, and unchanged approved audio. Cue timing and instructional usefulness remain instructor review items.
6. Pilot one slide first, compare the no-cue output hash against V1 to prove the unchanged path, then visually approve the cue-enabled result before generalizing.

This design requires no Camtasia for normal production. Cues become pixels in the final full-slide MP4 and remain optional and slide-specific.

## TTS-ready script preparation rule

Editorial directions in square brackets are never sent literally to TTS. Supported pause directions must be converted into ElevenLabs-compatible break tags during preparation of the TTS-ready speech script.
