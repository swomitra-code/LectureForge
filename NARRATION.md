# ElevenLabs narration takes

Narration generation is a separate upstream step. It does not call HeyGen or PowerPoint.

## Command

```powershell
.\generate-narration.ps1 .\config\module1.json
```

Set `ELEVENLABS_API_KEY` in the environment before running. The key is never written to the manifest or logs. Set `defaults.narration.voice_id` in the manifest to the instructor-approved ElevenLabs voice ID.

## Slide 13 example

1. Save the exact approved script as `scripts/slide-13.txt`.
2. Add this slide entry to the manifest:

```json
{
  "slide_number": 13,
  "enabled": false,
  "avatar_webm": null,
  "narration": {
    "enabled": true,
    "speech_script": "../scripts/slide-13.txt",
    "generated_takes": [],
    "instructor_selected_take": null
  }
}
```

3. Run the narration command. It generates exactly:

```text
media/slide-13-take-1.mp3
media/slide-13-take-2.mp3
media/slide-13-take-3.mp3
```

4. Listen to all three. The instructor chooses the production take by setting, for example:

```json
"instructor_selected_take": "../media/slide-13-take-2.mp3"
```

The tool skips narration-disabled slides and slides that already have an instructor-selected take. It performs only file-size, decode, and duration checks. It does not score naturalness, pacing, pronunciation, prosody, speaking rate, or subjective audio quality.
