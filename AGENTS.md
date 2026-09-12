# AGENTS.md — LectureForge Local Production Rules

## Goal
Build a simple, repeatable Windows workflow that turns an existing PowerPoint deck plus approved narration/avatar media into a self-contained PowerPoint with embedded videos.

## Production priority
1. Produce a real working slide first.
2. Verify it in desktop PowerPoint.
3. Only then generalize the pipeline.
4. Complexity must earn its way into the system.

## Source-of-truth rules
- Preserve instructor slide content, facts, charts, labels, sources, and instructional intent.
- Do not rewrite slide content unless explicitly requested.
- Keep the original PPTX intact. Always write to a new output file.
- Use only approved narration takes for final production.
- Human listening is the subjective narration-quality gate.
- Machine checks should only catch technical failures such as missing files, decode failures, wrong slide/media binding, or broken embedding.

## Presenter standard
- Avatar: SKM-BLUE-7.
- Presenter appears in the lower-right.
- Crop from top of head to just below shoulders.
- Presenter stays small so slide content remains dominant.
- Reduce green spill around hair and edges.
- Never cover chart labels, legends, captions, equations, or source lines.

## Audio/video standard
- Preserve the exact selected ElevenLabs narration audio.
- HeyGen is used for lip-sync/avatar motion, not to regenerate final voice.
- If source avatar video contains alpha, preserve alpha while cleaning RGB edge spill.
- Final composed slide videos: 1920x1080 H.264 MP4 with AAC audio.
- Do not add subjective audio QA, WPM gating, or repeated approval loops.

## PowerPoint delivery standard
- Keep the original editable slide underneath.
- Composite the presenter over a rendered image of the slide to create a full-slide MP4.
- Embed that MP4 over the full slide using desktop PowerPoint automation.
- Media must be embedded, not linked.
- Default playback: automatically when the slide is entered.
- Save to a new PPTX.

## Pilot scope
Start with Slide 10 only. Do not build Slides 11–12 until Slide 10 passes.

## Golden-path test
1. Open source deck.
2. Render Slide 10 exactly.
3. Composite approved SKM-BLUE-7 video lower-right.
4. Preserve exact approved narration audio.
5. Produce PowerPoint-compatible MP4.
6. Copy the deck.
7. Embed MP4 on Slide 10 full-slide and embedded.
8. Set it to start automatically.
9. Save, close PowerPoint, reopen output PPTX.
10. Verify video exists, is embedded, plays with audio, and slide count is unchanged.

## Stop rules
Stop and report instead of adding workarounds if:
- PowerPoint cannot embed video reliably.
- media is linked rather than embedded.
- cleanup destroys presenter edges.
- unrelated slides are modified.
- the final desktop PowerPoint journey cannot be verified.

## Preferred implementation
- Python
- FFmpeg
- pywin32
- Windows desktop PowerPoint COM automation
- simple JSON config
- Git
