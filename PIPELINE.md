# LectureForge manifest-driven production

The local production pipeline is one PowerShell command plus one JSON manifest. It preserves the approved Slides 10–12 FFmpeg filter graph, encoding settings, PowerPoint COM embedding, autoplay, close/reopen, and playback validation.

## Folder conventions

```text
config/   deck manifests
source/   original editable PPTX files
media/    approved transparent avatar WebMs
work/     rendered slide PNGs and temporary build files
output/   finished PPTX, composed MP4s, and validation JSON
```

Manifest paths may be absolute or relative to the manifest file.

## Build command

From the repository root:

```powershell
.\build-deck.ps1 .\config\module1.json
```

If the current PowerShell execution policy blocks local scripts:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build-deck.ps1 .\config\module1.json
```

## Manifest format

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
    }
  },
  "slides": [
    {
      "slide_number": 10,
      "enabled": true,
      "avatar_webm": "../media/slide-10-approved.webm"
    },
    {
      "slide_number": 11,
      "enabled": true,
      "avatar_webm": "../media/slide-11-approved.webm",
      "presenter": {
        "scale": { "width": 260, "height": 252 },
        "position": { "x": 1640, "y": 450 }
      }
    },
    {
      "slide_number": 12,
      "enabled": false,
      "avatar_webm": "../media/slide-12-approved.webm"
    }
  ]
}
```

`presenter.scale` and `presenter.position` are optional per-slide overrides. Omitted values inherit `defaults.presenter`. Set `enabled` to `false` to leave a configured slide unprocessed.

## Add another slide

1. Put its approved transparent WebM in `media/`.
2. Add one enabled slide object to the manifest with `slide_number` and `avatar_webm`.
3. Add a presenter override only when the default lower-right placement would cover meaningful content.
4. Run the same build command.

No PowerShell or FFmpeg changes are needed for another slide.

## What the command validates

- source and output slide counts match;
- every enabled slide receives one new full-slide media object;
- media is embedded, not linked;
- autoplay and PlayOnEntry are enabled;
- desktop PowerPoint playback position advances;
- the original editable slide remains underneath;
- unrelated slides retain matching shape, text, position, and size signatures;
- each composed MP4 is 1920×1080 H.264/AAC and fully decodes.

The validation report is written beside the output PPTX with `.validation.json` appended to the deck name.
