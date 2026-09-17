# Studio avatar production visibility fix

The selection and authorization components and API routes existed, but visibility
depended on saved review navigation. `loadSelectionReview(true)` hid the summary
when no review existed or the saved phase was narration. Selecting the final take
only refreshed its contents. The dashboard also stayed hidden until a batch
existed. Consequently a lecture with all narration selected could show only the
blocked recording step unless the user discovered REVIEW SELECTIONS elsewhere.
This was a rendering/navigation condition, not a missing provider endpoint or
project schema migration.

Avatar Production now stays visible before Recording PowerPoint. It displays
valid narration selection counts, the current avatar configuration, each enabled
slide and take, reusable assets, expected new HeyGen jobs, and excluded slides.
Incomplete narration leaves the production action disabled. Existing projects
without review state need no migration.

The flow remains:

1. Select narration. Selection and readiness reads authorize nothing.
2. Review the production summary. REVIEW PAID AVATAR GENERATION saves a bound
   confirmation snapshot and opens the final summary; it makes no provider call.
3. Explicitly press CONFIRM GENERATE FOR THESE SELECTIONS. `/api/avatar-authorize`
   still requires `confirm: true`, the confirmation phase, and matching inputs.
4. The worker consumes that authorization once and creates the enabled-slide
   jobs. Approved reusable media follows local validation; required new media
   follows the existing HeyGen submission path and duplicate-submission guards.
5. The dashboard reports queue, processing, ready, and exception states. Changes
   refresh server assembly readiness without resetting the placement form.
6. Only server-validated ready inputs enable Create Recording PowerPoint.

No provider, authorization, encoding, avatar preset, or placement implementation
was changed. SKM-BLUE-7, white background, and selected narration as final audio
remain in force. Cost review reports new provider job counts, not a dollar quote.

Files changed for this fix: `studio/web/index.html`, `studio/web/selections.js`,
`studio/web/avatars.js`, the review-panel visibility line in `studio/web/studio.js`,
`tests/v2/Test-AvatarProductionUI.cjs`, `tests/v2/Test-AvatarGuards.ps1`,
`tests/v2/Run-FRegressions.ps1`, and this note. Pre-existing narration/settings
changes in the working tree were preserved.

Validation on 2026-09-14:

- PASS: new production UI handler regression (VM DOM fixture): legacy/no review
  state, complete/missing selections, preset, summary before authorization,
  read-only readiness, explicit confirmation, queued/processing/partial/error/
  complete states, initial assembly loading, placement save/reset.
- PASS: expanded AvatarGuards: incomplete narration, review-only rejection,
  confirmation preview without queueing, exact enabled slide set, authorization
  consumption/recovery, immutable binding and duplicate-submission guards.
- PASS: NarrationSettings, NarrationRegeneration, ScriptImport, NarrationCues,
  including their applicable HTTP and UI handler checks.
- PASS: ProductionPackaging syntax, credential scan, launcher and whitespace.
- BLOCKED: historical Test-SelectionAuthorization suite after its first two
  checks: missing `work/solar-thermal-slides2-14-production-r001/jobs/slide-2/job.json`.
  The full historical regression runner was therefore not completed.

No live paid calls were made. Narration transports were mocked and avatar guard
fixtures disabled provider calls. No production deck or selected media was
modified. A live paid generation and desktop PowerPoint delivery journey were
not performed for this UI fix.
