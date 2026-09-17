# Lecture-level narration settings

Implemented September 14, 2026.

## Architecture and compatibility

The existing `project.preset.narration` object is the saved lecture-level source of truth. `NarrationSettings.psm1` reads the unchanged production defaults from `presets/renewable-energy.json` and provides one validator, reused by project loading/saving, preview queuing, and the existing ElevenLabs provider.

The project schema remains `lectureforge-studio-1`. No new project schema version or migration is required. Missing narration settings, including missing individual fields, receive defaults in memory; opening a project does not rewrite its files. `/api/save` accepts an optional `narration_settings` object and persists it in `preset.narration` using existing project version checks and the narration lock.

Existing revision binding semantics remain in place: saved settings changes create a different narration revision on the next Generate Takes action. Old takes and selections remain in history. Restoring matching settings restores the matching revision. Default JSON numeric representations are preserved because legacy bindings hash the serialized preset, including decimal formatting.

The existing provider, queue dispatcher, `Invoke-NarrationTake.ps1` executor, audio preparation, cue sanitizer, and spoken-text sidecar writer are reused. There is no second narration provider implementation.

## Studio behavior

- A collapsed Narration Settings panel sits beside the existing narration controls.
- It exposes Voice ID, Model, Stability, Similarity, Style, Speed, Speaker boost, and Takes per slide.
- Voice is editable text. Model offers the known multilingual v2 model and retains a previously saved model if present. No voice/model discovery requests are made.
- Save Changes and Generate Takes save the displayed settings with the lecture. Generate Takes still uses the existing confirmation and generation endpoints.
- Restore Default Preset loads the exact production defaults into the panel and marks the lecture dirty. Save Changes persists them.
- The preset helper offers Current Production Default and More Natural / Less AI. The latter changes stability to 0.47, similarity to 0.80, style to 0.38, speed to 1.00, and speaker boost to true, retaining the chosen voice, model, and take count. Nothing switches automatically.
- Test Take uses the current slide selector, synchronized with narration review. It snapshots the displayed settings and current script text without saving them.
- Preview status displays submission, queued, dispatched, generating, ready, and failure/uncertain states. Ready previews expose the existing audio playback endpoint. Active requests disable the test button, and the backend rejects another pending preview for that slide.

UI and backend validate stability/similarity/style from 0 through 1; speed from 0.7 through 1.2; and an integer take count from 1 through 5. Backend validation also rejects null, numeric strings, booleans in numeric fields, NaN, infinity, and nonboolean speaker boost. Speed uses the documented conservative range from [ElevenLabs' speed help](https://elevenlabs.io/docs/help-center/product/core-capabilities/text-to-speech/can-i-change-the-pace-of-the-voice).

## Exact provider request

For script `Heat [POINT: collector]flows.`, before this change and after it with defaults, the request is semantically identical. JSON object key order is not significant:

```http
POST https://api.elevenlabs.io/v1/text-to-speech/PqwAVcrzgTuCTbgkDpWl?output_format=mp3_44100_128
Content-Type: application/json
Accept: audio/mpeg
xi-api-key: <existing environment credential>
```

```json
{
  "text": "Heat  flows.",
  "model_id": "eleven_multilingual_v2",
  "voice_settings": {
    "stability": 0.34,
    "similarity_boost": 0.80,
    "style": 0.50,
    "use_speaker_boost": true,
    "speed": 1.00
  }
}
```

After the instructor chooses More Natural / Less AI, with the default voice/model retained, the same endpoint receives:

```json
{
  "text": "Heat  flows.",
  "model_id": "eleven_multilingual_v2",
  "voice_settings": {
    "stability": 0.47,
    "similarity_boost": 0.80,
    "style": 0.38,
    "use_speaker_boost": true,
    "speed": 1.00
  }
}
```

Voice changes affect the escaped voice ID in the URL; model changes affect `model_id`. Takes per slide controls how many independent attempts the existing batch queues; it is never sent as a provider payload field. The default remains three. A preview always queues exactly one attempt regardless of this setting, with no automatic provider retry.

## Test artifact storage

Previews are separate revisions in the existing `narration.json`, marked `preview: true`, with no production binding or selected take and exactly one attempt. `/api/narration` exposes these separately as `previews`. Production revision lookup explicitly excludes them.

Each preview uses a new `n/<revision-id>/<attempt-id>/` directory containing `raw.mp3`, `raw.mp3.spoken.txt`, and the prepared `take.mp3`. Preview pause is zero so the test hears the provider take without appended silence. No existing take path is overwritten. Previews remain local for inspection/recovery; there is no automatic expiry or deletion. The panel displays the latest preview for the current slide. Preview audio cannot be selected for production or authorized for HeyGen.

## Files changed for this feature

- `studio/NarrationSettings.psm1` — new shared defaults/validation module.
- `studio/Production.psm1` — backward-compatible read defaults and project persistence.
- `studio/NarrationProvider.psm1` — calls shared validation before its existing provider logic.
- `studio/Narration.psm1` — configurable take count, separate preview revisions, shared dispatch, slots 1–5.
- `studio/Server.ps1` — test-take endpoint and audio access for slots 4–5.
- `studio/web/index.html`, `studio/web/studio.js`, `studio/web/studio.css` — compact settings panel, presets, dynamic take controls, preview status/playback.
- `tests/v2/Test-NarrationSettings.ps1`, `tests/v2/Test-NarrationSettings.cjs` — new provider/executor and HTTP/UI regressions.
- `tests/v2/Test-ScriptImport.cjs`, `tests/v2/Test-NarrationRegeneration.cjs` — DOM fixture support for settings inputs/audio.
- `tests/v2/Run-FRegressions.ps1` — includes the new settings suite.
- `studio/NARRATION_SETTINGS.md` — this report.

Pre-existing uncommitted per-take regeneration work was preserved.

## Verification

Passed: narration settings, narration cues, narration regeneration (including HTTP/UI), script import (including HTTP/UI), avatar guards, and production packaging checks. Settings tests exercise the real executor and provider payload with only HTTP transport mocked; they verify exact request counts, zero HeyGen requests, defaults, save/reopen, restore, custom payload/voice, validation, legacy schema, sidecars, preview isolation, authorization binding, and range-based playback delivery.

No live paid requests were made. Human listening quality was not assessed. A real browser visual check could not run because the browser connector reported no available browsers; UI handlers and HTTP playback delivery were verified by the regression harness. Desktop PowerPoint was not exercised for this narration-only change.

HeyGen provider code, authentication, authorization/gating, project import logic, and PowerPoint recording code were not changed. Avatar guard regressions pass, and preview generation leaves the avatar authorization input hash unchanged.
