# LectureForge V2 Production Architecture

Status: V2.0 white-background avatar Recording workflow approved by the instructor
on 2026-09-12. Native slides remain visible with a separate movable/resizable avatar
video, matching the reserved white area and autoplaying with narration.
Source, white MP4, embedded media and desktop slideshow motion all passed. The
earlier apparent freeze could not be reproduced and is treated as a transient
PowerPoint session condition, not a pipeline failure. Final audible narration
confirmation remains required during real Slide 1 production.

Development branch: `lectureforge-v2-production-studio`.

Protected V1 reference: `master`, commit
`e8f9ff57c2dcf54ac81d65f34ced9c8b8c71c9cd`, annotated tag
`lectureforge-v1-baseline`. V2 development must not move these references.

## Product goal

Evolve the existing Windows production system into a lightweight local Production
Studio. Preserve the native editable PowerPoint lecture and produce reviewable,
independently regenerable slide videos.

```text
Editable PPTX
  -> slide scripts
  -> candidate narration, including intentional pauses
  -> instructor narration approval
  -> transparent avatar generation using approved narration
  -> per-slide MP4 composition using approved narration
  -> instructor video approval and optional locking
  -> stable approved slide exports
  -> optional full-lecture MP4 assembly
Separate Recording path: native PPTX + approved white-avatar tile with narration
  -> derived Recording PPTX -> instructor PowerPoint Record -> final lecture video
```

The instructor may use the slide MP4s directly or open a derived Recording PPTX to
add live cursor movement, laser, ink/annotation, and instructor-controlled emphasis
through PowerPoint Record. PowerPoint can then export the final lecture video.
The Recording PPTX is a playback/recording artifact, never the editable lecture master.

## Non-negotiable rules

### Native PowerPoint preservation and the V2.0 static slide contract

- The imported/source PPTX is canonical and remains editable. Preserve the original
  uploaded file. Import a byte-identical project copy and record its SHA-256.
- Open the canonical project source read-only for rendering; never save it during
  production. Verify its hash before and after rendering.
- Never replace or cover native slide content with full-slide video in the canonical
  source or editable master PowerPoint. Never flatten or save changes to that master.
- Recording PPTX derivatives must also retain visible native slide content. Never
  place a full-slide composite video or raster over native content in these decks.
  Full-slide composition is reserved for standalone MP4 delivery.
- Rendering is permitted for thumbnails, previews, and final video production.
- V2.0 intentionally renders the fully composed static PowerPoint slide. Native
  animations, object builds, transitions, and staged reveals are not reproduced.
- Do not implement animation capture, slideshow recording, or animation timing
  reconstruction in V2.0. Animated-slide capture is only a possible future
  production mode if instructional testing demonstrates a need.
- A static representation that does not communicate the intended lesson requires
  instructor review; do not silently introduce an animation workaround.

### Per-slide output

- Produce one finished 16:9 MP4 per enabled slide, containing the rendered native
  static slide, the approved narration, and the presenter overlay.
- Each slide can be regenerated independently. Approval and locking are explicit.
- Retain immutable revision-specific assets, such as
  `videos/slide_01_video_rev_004.mp4`.
- On explicit approval, publish a byte-identical copy of the approved video to
  `exports/slide_01.mp4` through `exports/slide_14.mp4` for Solar Thermal.
- Stable export names always identify the currently approved video revision.
  Publishing does not delete historical revisions. Unapproved candidates never
  replace stable exports.

### Separate derived Recording PowerPoint

- Accepted architecture: native PowerPoint slide + separate white-background avatar
  video + synchronized narration + autoplay. The avatar remains independently
  selectable, movable and resizable in desktop PowerPoint.
- Build a distinctly named derivative from the canonical native deck, never from
  an instructor's previous recording. Reject existing output paths and canonical
  aliases. Never save changes to the canonical source or master.
- Embed only the compact avatar-only H.264/AAC MP4, never linked media or a full-slide
  composite. Use RGB 255,255,255 behind the presenter, no border and no shadow.
- Place it inside the visually verified reserved white area. Do not cover visible
  instructional content. Keep aspect ratio and head/shoulder padding. Moving or
  resizing within a matching white area needs no provider call or recomposition.
- Verify approved tile hashes before packaging and embedded byte identity afterward.
  Narration is authoritative in tile composition; packaging must not transcode it.
- Configure autoplay; reopen and check saved settings and actual player advancement.
  Validate decoded changing presenter frames separately from playback configuration.
- Preserve all 14 native slides in the proof. Only Slide 1 receives an avatar;
  Slides 2-14 must match structurally and visually. No full-slide image is added.
- V2.0 will not pursue PowerPoint alpha playback. Transparent source assets may
  still be decoded locally for RGB despill and white-background tile composition.
- Standalone slide MP4s and Recording avatar tiles are distinct approved artifacts.
  A standalone composite approval is not approval to embed that composite in PPTX.
- Final audible narration confirmation remains required during real Slide 1
  production. The technical fixture does not approve Solar Thermal content or
  production lip-sync. AAC encoding is allowed, not byte identity with MP3/WAV.
- PowerPoint Record cursor, laser, ink, replay and export behavior remain instructor
  workflow checks. Do not infer ordinary cursor recording from laser support.
- The straight lower-torso crop is a minor framing issue to improve during real
  avatar production; it is not an architecture blocker.

### Approved audio and instructional pauses

- The instructor-approved prepared narration asset is authoritative. Submit that
  exact asset to HeyGen for synchronization and use it as the sole final MP4
  audio source. Do not substitute HeyGen-extracted audio unless explicitly requested.
- Intentional silence is part of the candidate narration before the instructor
  listens and approves it. Do not add, remove, shorten, or lengthen instructional
  silence after approval.
- Do not automatically trim, normalize, change playback speed, or rewrite approved
  narration. AAC encoding is allowed; it does not make encoded bytes identical to
  the source MP3/WAV, so do not claim byte-identical final audio encoding.
- Missing approved narration or a changed asset hash blocks production; there is
  no automatic fallback to avatar audio.
- Human listening is the subjective narration gate. No WPM scoring, automatic
  take selection, or subjective machine approval is introduced.

### Avatar and placement

- Retain SKM-BLUE-7. Transparent source assets support local composition; Recording
  tiles use the explicitly approved pure white background.
- Project placement defaults to the lower-right. Do not adopt a permanent arbitrary
  scale. During Phase 2 Slide 1 review, the instructor approves initial position
  and scale; that approved placement can become the Solar Thermal project default.
- Slides may inherit project placement or store position/scale overrides.
- Standalone placement/crop changes require local recomposition only. Recording
  avatar position/size can be edited directly in PowerPoint. Preserve alpha during
  RGB edge-spill cleanup before flattening the tile onto white.

### State and approvals

Track independently: script ready, narration generated, narration approved, avatar
generated, preview rendered, slide video approved, and locked.

Approved assets must not be regenerated automatically. A lock pins the approved
input/output revision. An explicit unlock is required before changing a locked
slide; unlocking does not itself regenerate anything.

## Smallest proposed architecture

Retain PowerShell, desktop PowerPoint COM, FFmpeg/FFprobe, the existing ElevenLabs
and HeyGen integrations, JSON manifests, and local browser review. Use plain
HTML/CSS/JavaScript, one authoritative project manifest, and one serialized
background worker. No database, frontend framework, general timeline editor,
plugin system, or general provider abstraction is required.

```text
Local browser Studio -> project/state layer -> background worker
                                            |- read-only COM slide rendering
                                            |- ElevenLabs candidates
                                            |- HeyGen synchronization
                                            |- FFmpeg per-slide composition
                                            |- validation and approval publishing
                                            |- optional approved-video assembly
                                            `- approved white-avatar tiles -> native derivative Recording PPTX
```

HTTP requests submit jobs and return immediately; the UI polls progress. Provider
generation and PowerPoint operations never run synchronously inside a request.
There is no embedding into the canonical source. A separate approved-output packager
embeds approved videos into the derivative Recording PPTX only.

## V1 reuse and planned change boundaries

These are future implementation boundaries, not changes authorized by this document.

| Existing component | Treatment |
|---|---|
| `build-deck.ps1` | Retain unchanged as V1. Reuse COM export/composition techniques and bounded embedding/autoplay/reopen validation techniques in separate V2 components. Do not invoke the V1 end-to-end builder on canonical sources or use its HeyGen-audio composition path. |
| `lectureforge.ps1` | Retain unchanged as the V1 deck orchestrator. V2 needs revision-aware per-slide completion. |
| `select-narration.ps1` | Retain for V1. V2 records asset-specific approvals through its state layer rather than regex manifest edits. |
| `src/*.ps1` | Retain unchanged as historical pilot implementation and validation evidence. |
| V1 manifests, reports, external inventory | Preserve unchanged; migrate into separate V2 projects. |
| `generate-narration.ps1` | Later add backward-compatible explicit slide requests, revision destinations, and single audition samples. Preserve the existing default three-take workflow. |
| `generate-avatar.ps1` | Later add explicit job/output handling, persisted provider job IDs, and accurate returned asset paths while retaining provider requests and defaults. |
| `review-narration.ps1` | Retain legacy mode; add explicit Studio mode backed by new modules. Reuse listening and media-range-serving concepts. |
| `.gitignore` and operating documentation | Later add local V2 project exclusions and instructions. |

V1 currently selects avatar audio in final composition and embeds full-slide MP4s
into a copied deck. V2 requires authoritative narration and embeds only independent white-avatar tiles
in the Recording PPTX; standalone composites stay outside the Recording deck. Its
completion checks do not establish input-revision freshness. Its synchronous,
ASCII/character-oriented reviewer transport needs bounded UTF-8 and binary upload
support for Studio.

Provider tools receive job-specific compatibility manifests, never the authoritative
V2 manifest. The V2 state layer alone publishes validated results and approvals.

## Proposed new files

Create only when the corresponding implementation phase needs them.

| File/module | Responsibility |
|---|---|
| `studio/Project.psm1` | Schema, project paths, assets, atomic state changes, approvals and locks |
| `studio/Worker.ps1` | Serialized jobs, progress, recovery, provider invocation |
| `studio/Render.psm1` | Read-only COM export, composition, MP4 validation, optional assembly |
| `studio/WhiteAvatar.psm1` | Compact white-tile composition using explicit narration, decoded motion checks; no providers |
| `studio/Recording.psm1` | Approved-avatar-tile packaging into a separate Recording PPTX, autoplay, embedding/hash checks and desktop validation; no provider or composition calls |
| `studio/Server.psm1` | Local HTTP routes, imports, media serving, job submission |
| `studio/web/index.html`, `studio.css`, `studio.js` | Plain browser Production Studio |
| `studio/Import-V1.ps1` | Non-destructive migration |
| `studio/project.schema.json` | V2 data contract |
| `tests/v2/` | Preservation, dependency, approval, locking, and rendering tests |

## Project layout and data model

```text
project.json
source/lecture.pptx
scripts/
assets/narration/
assets/avatar/
renders/
videos/slide_01_video_rev_004.mp4
exports/slide_01.mp4
exports/<project>_RECORDING.pptx
jobs/
```

Paths are relative to the project root. Credentials remain in environment variables.
Source replacement is an explicit new revision, not an overwrite during production.

The authoritative manifest records:

- Project schema version, stable project ID, name, and manifest revision.
- Source path/hash, source revision, slide count and dimensions.
- Project voice/settings, avatar identity, crop, approved placement defaults, and
  encoding profile. Initial placement approval may be pending.
- Stable slide ID, current slide number, enabled flag, script path/hash/revision.
- Asset ID, role, relative path, SHA-256, duration, generation inputs and provenance.
- Approval asset ID/hash, decision, timestamp, and applicable script/input revision.
- Separate script, narration, avatar, preview, video approval, and lock records.
- Placement inheritance or per-slide override, with effective placement captured
  in every composition revision.
- Pilot slide subset/order and lecture assembly order.
- Stable export path, approved video revision ID/hash, and publication status.
- Recording PPTX path/hash, canonical source hash, source/output slide mapping,
  packaged approved-video hashes, and separate desktop Record/export review results.
- Job state and provider job identifier, without credentials or token-bearing URLs.

Use immutable revision paths. Write temporary output, validate, then publish.
Use one manifest writer, revision checks, and atomic file replacement. Reject stale
browser updates rather than overwriting newer decisions.

Stable export publication is a recoverable operation: prepare and hash-check a
temporary copy, replace the stable export atomically, and reconcile its mapping
with the approval record. A partially completed publication must be shown as pending
and unavailable for assembly until reconciled. Do not claim atomicity across two
separate files. Never delete prior revision assets during publication or recovery.

## UI and routes

| View | Required capabilities |
|---|---|
| Project setup | Choose/upload PPTX, choose/import script source, project name, slide selection |
| Slide list | Thumbnails, script text, duration, independent status and approval indicators |
| Voice Studio | Choose existing voice/clone, generate test sample, compare samples, Approve/Maybe/Reject, save project default; retain three takes where useful |
| Slide production | Slide/avatar preview, drag/resize, Apply to All, per-slide override, narration playback, build/regenerate this slide |
| Pilot Review | Select subset, Build Pilot, watch MP4s, Approve, Regenerate Voice, Regenerate Avatar, Recompose Avatar, Edit Script |
| Production dashboard | All-slide status, build incomplete work, review candidates, lock approvals, create Recording PPTX from approved videos, optionally assemble approved videos |

```text
GET/POST  /api/projects
GET/PATCH /api/projects/{project}
POST      /api/projects/{project}/imports
GET       /api/projects/{project}/slides
GET/PATCH /api/projects/{project}/slides/{slide}
GET/POST  /api/projects/{project}/voice-samples
POST      /api/projects/{project}/slides/{slide}/actions
POST      /api/projects/{project}/pilot
POST      /api/projects/{project}/assembly
POST      /api/projects/{project}/recording-pptx
GET       /api/jobs/{job}
GET       /api/assets/{asset}
```

Actions distinguish generation, recomposition, approval, publication, lock and
unlock. Preview may use an alpha still extracted from the avatar for interactive
placement; the composed MP4 is the authoritative moving review. Do not assume
transparent WebM browser playback is reliable on every local browser.

## Audio authority, duration, and the Solar Thermal pause

Final composition receives three independent inputs: static slide image, avatar
video stream, and approved prepared narration audio. Explicitly map audio only
from the narration input. Record the hashes of all inputs and effective settings.
An avatar generated from an older narration revision is stale for a new selection.

Solar Thermal has 14 slides. Initial pilot slides are 1, 2, 3, and 8. Slide 8 has
an intentional approximately six-second silent prediction interval **after the
final spoken sentence**. The initial preparation target is 6.000 seconds of appended
silence after the defined end of speech, included in the candidate asset before
approval. Preserve preparation provenance, including the spoken source and silence
duration. The instructor hears and approves the complete prepared asset.

Send that same approved prepared asset to HeyGen and use it as the sole final MP4
audio source. Do not append silence only at composition time, strip trailing silence,
or rely on an unverified TTS pause directive. Do not silently extend an already
approved take: prepare a new candidate for approval if its pause is missing.

Approved narration determines final video duration, including all trailing silence.
Do not use `-shortest` to conceal a short avatar and truncate the prediction interval.
If HeyGen omits silent-tail motion, review a documented video-only tail treatment
such as holding a suitable final presenter frame; it must preserve audio unchanged.
Unacceptable presenter behavior or material synchronization drift stops acceptance.

## Placement metadata

Store normalized slide-space `x`, `y`, and `width`, with the anchor at the top-left
of the cropped presenter box. Derive height from the crop aspect ratio. Store crop
coordinates separately in source-avatar pixels. Standalone composition consumes
alpha locally; Recording tile background is pure white. PowerPoint placement is
recorded independently in slide points.

Lower-right is the project default region, but scale and exact position remain
provisional until instructor approval on Slide 1 in Phase 2. Save that accepted
placement as the Solar Thermal default only after approval.

An override changes composition metadata only. Apply to All updates eligible
inheriting slides and shows the affected set. Replacing existing overrides requires
an explicit action. Approved or locked slides retain their approved placement
snapshot; bulk changes must not silently invalidate their published exports.

## Independent regeneration and locking

| Changed input | Required new work |
|---|---|
| Script | New narration candidate; downstream work after its approval |
| Approved narration selection | Avatar and composition |
| Avatar/look | Avatar and composition |
| Position, scale, crop | Composition only |
| Static slide artwork | Slide render and composition |
| Lecture ordering | Assembly and any newly requested Recording PPTX ordering |

Use input hashes and settings to determine freshness. Preserve earlier approved
revisions while new candidates are under review. Existing unapproved candidates
wait for review; repeated Build Incomplete does not regenerate them automatically.
Only missing/stale, authorized stages run. Changing a project default voice does
not regenerate approved narration.

Locks pin the approved source/script/narration/avatar/placement/video revision.
Require explicit unlock before regeneration, reselection or bulk placement changes.
When a candidate is explicitly approved, publish its stable export without deleting
history. A source change must not silently remap stable slide IDs after reordering.

## Pilot and full production

Prove Slide 1 first, including instructor approval of presenter placement. Then
produce/review only slides 1, 2, 3 and 8. Verify the full Slide 8 silent tail in both
approved narration and final MP4. The deck's lower-right avatar-safe design is an
input assumption to verify visually, not permission to cover meaningful content.

Expand to the remaining ten slides only after pilot acceptance. Build Pilot never
changes unrelated slides. Publish only explicitly approved revisions under stable
export names. Optional assembly consumes a recorded ordered snapshot of approved
video revisions with matching stable-export hashes; it does not change slide assets.
Recording PPTX production packages separately approved white-avatar tile revisions. In Phase 2,
approval may be recorded as a narrow proof input; do not build Phase 3's general
approval/publication state system just to conduct the one-slide recording test.

## Migration from V1

- Read V1 manifests without modifying them; create a separate V2 project.
- Resolve mixed V1 path conventions once and write project-relative references.
- Preserve selected narration and selection provenance. Register bound avatars as
  generated, but mark their narration linkage unverified when evidence is missing.
- Import legacy MP4s as candidates, not automatically V2-approved outputs. Technical
  V1 validation is not evidence of subjective instructor video approval.
- Retain original validation reports as historical evidence.
- Never extract avatar audio as an automatic replacement for missing approved
  narration. The V1 pilot WebM has no separately referenced narration asset; locate
  or explicitly approve authoritative audio before claiming V2 compliance.
- Keep V1 code, manifests, tag, and master intact.

## Validation, risks, and stop conditions

| Area | Required validation / response |
|---|---|
| Canonical PPTX preservation | Hash before/after; read-only COM; no save or embed calls on canonical sources. Stop on any mutation. |
| Recording PPTX | Separate output only; verify selected slide mapping/count, compact avatar geometry and independent movement/resizing, embedded bytes/hash, autoplay and audible approved narration after close/reopen. Stop on linking, cropping, changed audio, or wrong video binding. |
| PowerPoint Record | Demonstrate cursor/laser/ink over native slides while the avatar moves, replay the recording and inspect PowerPoint's exported video. Record installed version and control-visibility behavior. Stop and report unsupported or obscured gestures, missing narration, or truncated playback; do not infer success from COM properties. |
| COM behavior | Serialize operations, release only owned objects, bound execution. Do not terminate an instructor PowerPoint session. Stop on unresolved dialogs, hangs, or wrong renders. |
| Static fidelity | Compare exported static slide with desktop PowerPoint. Review missing fonts/links or obscured content. Do not attempt animation capture. |
| Aspect ratio | First acceptance project is 16:9. Do not silently stretch another ratio. |
| Alpha | Check decoded transparency as well as metadata; inspect edges on the real slide. Stop on opaque rectangles or damaged edges. |
| Audio provenance | Verify approved asset hash and explicit stream mapping; missing/mismatched narration blocks build. |
| Timing | Check full duration, timestamps, decoding, and Slide 8's approved silent tail. No post-approval silence edits or hidden truncation. |
| Encoding | Standalone: 1920x1080 H.264, 30 fps, AAC. Recording proof tile: 640x560 H.264, 25 fps, AAC, white background. Check actual streams and full decode. |
| Approval/locks | Test restart, bulk builds, stale UI updates, and overrides cannot replace protected revisions. |
| Stable exports | Check byte/hash identity with approved revision; interrupted publication is pending until recovered. Preserve all history. |
| Provider recovery | Persist returned job IDs before polling. Reconcile uncertain submissions without blindly creating duplicate paid jobs. |
| Browser server | Loopback only, asset-ID routing, bounded uploads, UTF-8 handling, request-origin/session checks, project-path validation; never expose provider keys. |
| Human review | Instructor listens and watches; no automated subjective scoring or approval. |
| Assembly | Approved current revisions only, explicit order, boundary/duration checks. Stream-copy only when compatible; otherwise create a separate normalized lecture export. |

Machine checks address technical failures. They do not replace instructor assessment
of narration, placement, presenter quality, or instructional usefulness.

## Seven implementation phases

### Phase 1 - Architecture

Deliver this approved design and branch separation. Incorporate the static-slide
contract, approved pause audio, placement acceptance, and stable approved exports.
Checkpoint: architecture documentation committed and V1 references verified unchanged.
Architecture approval permits planning, not automatic execution of later checkpoints.

### Phase 2 - One real V2 slide

Phase 2A: implement the minimal project record, read-only static rendering, separate approved
audio composition, and one standalone Slide 1 MP4. Verify source fidelity in desktop
PowerPoint and MP4 playback. Instructor approves initial lower-right position and
scale; save it as the project default after acceptance.

Stop if the source changes, transparency fails, approved audio provenance is wrong,
or position/scale has not passed instructor review. Do not generalize before this
real slide passes.

Phase 2B: after explicit standalone Slide 1 MP4 approval, prove a 14-slide derivative
Recording PPTX with a separately approved white-avatar tile on Slide 1 only. For a technical-fixture approval, retain
the technical-only label; do not imply approval of final Solar Thermal narration.

Checkpoint sequence:

1. Record white-avatar tile approval and SHA-256; hash the source master.
2. Create a separate `<project>_RECORDING_PHASE2.pptx`, preserving all 14 source
   slides. Modify Slide 1 only. Do not save changes to or prune the canonical source.
3. Embed the compact approved white-avatar MP4 in the reserved white area. Keep
   native content visible, preserve independent movement/resizing and set autoplay.
   Preserve audio and embedded bytes; add no border or shadow.
4. Save the derivative, close, reopen, and verify 14 output slides, Slide 1 embedded
   rather than linked media, geometry, autoplay configuration, and its audio stream.
   Compare Slides 2-14 structurally and visually with the source. Return the deck
   for instructor review without claiming recording acceptance from automation.
5. In a separate manual instructor test in desktop PowerPoint Record, exercise ordinary cursor movement, laser pointer,
   and ink/annotation separately over native slide content while the avatar plays. Verify recording replay and
   a PowerPoint-exported MP4, including layering, timing, and complete narration.
   Document ordinary cursor behavior separately from laser gestures; do not silently
   substitute laser for cursor capture or claim support based on documentation alone.
6. Verify canonical source hash and 14-slide count remain unchanged; preserve the
   approved standalone MP4. Present the derivative and exported proof for instructor
   review, including any playback-control limitations.

Stop on unapproved video input, source changes, linked media, failed autoplay/audio,
incorrect alignment, missing/obscured requested recording effects, or an unverifiable
desktop recording journey. Do not add recording workarounds automatically.
**No paid ElevenLabs or HeyGen generation until this Recording PPTX proof is reviewed.**
The seven phases are preserved; 2A and 2B are checkpoints within Phase 2.

### Phase 3 - Safe per-slide state

Implement immutable revisions, independent statuses, locks, resumable jobs, and
approval-to-stable-export publication/recovery. Prove recomposition makes no provider
calls and does not overwrite approved history. Stop on lock bypass, duplicate job
creation, stale approval binding, or inconsistent stable exports.

### Phase 4 - Studio and Voice Studio

Implement project import, slide list, auditions, three-take review, and responsive
background jobs. Prove selections survive restart and stale browser requests cannot
overwrite them. Stop on upload corruption, exposed credentials, or lost approvals.

### Phase 5 - Placement and pilot

Implement drag/resize, project defaults, overrides, Apply to All and Pilot Review.
Approve Solar Thermal slides 1, 2, 3 and 8. Verify Slide 8's six-second post-speech
pause was approved before HeyGen and survives unchanged in the final audio.
Stop before full production if any pilot fails or unrelated assets change.

### Phase 6 - Full production

Produce and review all 14 independent slide MP4s. Publish each approved revision as
`exports/slide_01.mp4` through `exports/slide_14.mp4`. Verify reruns make no unnecessary
generation calls and locked approvals remain intact. Stop on missing approvals,
incorrect stable-export mappings, or unintended regeneration.
Offer a full-project Recording PPTX from approved enabled-slide white-avatar tiles,
preserving the source master and any existing instructor-recorded derivative.

### Phase 7 - Optional assembly

Assemble the approved ordered sequence into a separate lecture MP4. Verify timing,
audio boundaries, and the prediction pause without modifying slide revisions or
the native PPTX. Stop on unapproved/stale inputs or failed boundary validation.

## Current authorization boundary

The instructor approved the white-avatar architecture and manually observed motion.
This milestone authorizes inspection, focused V2 validation, documentation and code
checkpointing on `lectureforge-v2-production-studio`, with the commit message
`feat: establish non-destructive avatar recording workflow`, then push to origin.
Preserve master and lectureforge-v1-baseline. Stage source/code/docs/tests only;
exclude source decks, generated media/decks, work artifacts and secrets.
Stop after checkpoint. No Studio UI, unrelated refactoring, later-phase work or
provider calls are authorized. Real Slide 1 narration audibility is still pending.

### Acceptance evidence

- Canonical Solar Thermal source SHA-256:
  `6340C137FD85021957721341298D640C8FAF59D36507D904DF09948DCB576C01`.
- Accepted white technical tile SHA-256:
  `3D9420617F53F0172021B22665A81BCB4D0FA83683CF06C29C5954AE2FA8E2B8`.
- White tile: 640x560, no perceptible white seam in desktop slideshow; no border or
  shadow. Accepted technical placement: left=438, top=337, width=110, height=96.25 pt
  on the 960x540 pt slide, in the white content panel above the closing question.
  The far-right blue photograph is unchanged; this is not a universal placement.
- Source and embedded video each contain 821 distinct decoded mouth-region frames.
  White conversion preserves face/mouth/body changes. Embedded bytes equal the MP4.
- Desktop autoplay advanced, and instructor confirmed visible motion. Earlier
  apparent freezing is a transient session condition; no speculative repair applied.
- Local evidence lives under ignored `work/avatar-white-proof-20260912/` and
  `work/motion-diagnostic-20260912/`; reproducible checkpoint tests live in `tests/v2/`.

## PowerPoint behavior references and remaining verification

Microsoft documents recording/previewing ink and laser gestures and including them
with media in exported videos. This supports the proposed workflow but does not
prove gesture layering over our embedded MP4 in the installed desktop version.
Ordinary cursor capture must be tested separately.
[Record your presentation](https://support.microsoft.com/en-us/powerpoint/record-your-presentation)

Microsoft also documents embedded-video playback in presentation-to-video export.
Our acceptance still requires actual playback, complete approved narration, and
recorded emphasis in the exported proof.
[Turn your presentation into a video](https://support.microsoft.com/en-us/powerpoint/turn-your-presentation-into-a-video)

Playback controls remain version/view dependent. Verify actual behavior after
reopening. No full-slide video-background technique belongs in this Recording path.
