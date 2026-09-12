# BUILD_SPEC.md

## Architecture
For each narrated slide:
1. Render the original PowerPoint slide to a 1920x1080 image.
2. Composite the approved transparent avatar over the slide.
3. Encode one normal MP4 containing the slide plus exact approved audio.
4. Embed that MP4 over the entire original slide.

This avoids relying on transparent WebM playback inside PowerPoint. The editable slide remains underneath.

## Prerequisites to verify
- Codex CLI
- Git
- Python 3.10+
- FFmpeg in PATH
- Microsoft PowerPoint desktop
- pywin32

## Workspace
LectureForge/
  AGENTS.md
  BUILD_SPEC.md
  source/
  media/
  config/
  work/
  output/
  src/
  tests/

## Pilot inputs
- source/Module1_Energy_Landscape_2026_Refresh_Cleaned_No_Old_Photo_FINAL2.pptx
- media/Module_1_Slide_10_-_Take_2_-_SKM-BLUE-7.webm

## FFmpeg behavior
Preserve source alpha. For green-spill cleanup, preserve the alpha mask, apply despill only to RGB, then merge alpha back. Use a stable crop from top of head to just below shoulders, scale small, and overlay lower-right. Preserve source audio exactly. Output 1920x1080 H.264/AAC MP4.

## PowerPoint behavior
Use desktop PowerPoint automation.
- Work from a copy.
- Add the MP4 to Slide 10 at 0,0, sized to full slide.
- Embed rather than link.
- Set playback to start automatically.
- Save to output/Module1_Energy_Landscape_VIDEO_PILOT.pptx.
- Close PowerPoint and reopen the result for validation.

## Pass conditions
- output PPTX reopens,
- slide count unchanged,
- Slide 10 has embedded media,
- media is not an external link,
- video fills slide,
- video/audio play,
- no unrelated slide changed,
- separate Slide 10 MP4 also saved.
