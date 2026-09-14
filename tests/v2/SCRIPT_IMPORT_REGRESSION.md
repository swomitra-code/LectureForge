# Import Scripts regression report

## Root cause and trace

The Import button in `studio/web/studio.js` saves pending edits, posts
`script_import` to `/api/save`, and calls `renderProject`. `Server.ps1` runs
`Update-LFProject` under the narration lock. `Production.psm1` parses headings
with `ConvertFrom-LFScripts`, changes matching script text, updates the existing
status/version fields, and saves `project.json`. The UI then reloads
`/api/narration` and renders progress and take review.

The failure is in that reload, not the script parser. In Windows PowerShell,
`takes=if($r){@($r.takes)}else{@()}` enumerates the conditional's output.
The actual JSON for a new slide is `"takes": {}`; one take becomes a bare
object. Thus `s.takes.find` is undefined even for a brand-new lecture.
The fix is `takes=@(if($r){$r.takes})`, preserving zero/one/many cardinality.

## Schema and compatibility

The canonical stored schema remains `lectureforge-narration-1`:
`narration.json.revisions[].takes` is an array of numbered take records.
New revisions already initialize `takes=@()`. Project slides in
`lectureforge-studio-1` contain scripts and metadata, not a second takes store;
no redundant project/session field was added. Every new slide's narration API
projection now initializes `takes` as `[]`.

Before: stored revisions normally used arrays, but the API emitted `{}`, a bare
take object, or an array depending on cardinality. Reads trusted legacy shapes.
After: reads normalize missing/null/empty values to `[]`, singleton records to
`[take]`, and keyed legacy collections to arrays. Valid takes retain all fields
and their original order. Unrecognized non-null values are retained in
`takes_legacy` for recovery and subsequent narration writes. Read-time migration
does not rewrite saved lectures. Import never modifies `narration.json`.
The UI also normalizes narration responses before either renderer uses them.

Sparse or differently ordered legacy arrays require lookup by take number.
Selection, retry, audio serving, execution, and worker completion now use that
identity instead of assuming Take N lives at index N-1.

## Takes-consumer audit and changed files

| File | Uses audited / change |
| --- | --- |
| `studio/Narration.psm1` | Read normalization and API array serialization; `.Count`, revision/take/attempt enumeration, generation planning, batch creation, selection, retry and dispatch audited; selection/retry use number lookup. |
| `studio/web/studio.js` | Response normalization protects progress `.length`, `.every`, `.find`, review `.find`, and take signatures. |
| `studio/Server.ps1` | Audio route resolves a take by number; import route unchanged. |
| `studio/Invoke-NarrationTake.ps1` | Both executor reads resolve takes by number. |
| `studio/Worker.ps1` | Executor completion resolves takes by number. |
| `studio/Authorization.psm1` (unchanged) | Already filters by take number and reads through the normalized narration boundary. |
| `studio/Production.psm1` (unchanged) | Project creation/read/import have no takes array; preserve existing project schema. |
| `tests/v2/Test-ScriptImport.ps1` (new) | Self-contained backend and JSON regression fixtures. |
| `tests/v2/Test-ScriptImport.cjs` (new) | Actual HTTP routes and production UI handlers/renderers in a minimal DOM. |
| `tests/v2/Run-FRegressions.ps1` | Includes the new regression entry point. |
| `tests/v2/SCRIPT_IMPORT_REGRESSION.md` (new) | This report. |

No other production `takes` consumers were found. Other web modules operate on
selection summaries/avatar jobs rather than raw takes.

## Validation on 2026-09-14

- **40 new checks passed:** 35 PowerShell checks and 5 HTTP/UI check groups.
  Coverage includes new/reopened lectures, the exact minimal reproduction,
  missing/null/empty/singleton/keyed/scalar/mixed takes, all three existing takes,
  preservation of fields/order/selection/history, 20 scripts, repeated imports,
  restored bindings, sparse take selection, and preservation on later writes.
- **Packaging checks passed:** syntax for all 41 PowerShell and 8 JavaScript
  files, credential scan, launcher checks, and Git whitespace.
- **Full offline runner attempted:** `Run-FRegressions.ps1` passes the new tests,
  then exits 1 in `Test-ProductionStudio.ps1` because
  `source/RE-M2-Solar-Thermal-Overview.pptx` is absent. The remaining suite stages
  were not reached. Approved media under `work/` is also absent in this checkout.
  This is not a full-suite pass. The opt-in paid `Test-NarrationLive.ps1` was not
  run. No replacement approved media or altered acceptance criteria were used.
- The HTTP/UI test uses the actual server and JavaScript with a minimal DOM;
  it is not a desktop browser or PowerPoint acceptance test.

Run with Windows PowerShell and Node on PATH:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/v2/Test-ScriptImport.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/v2/Test-ProductionPackaging.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/v2/Run-FRegressions.ps1
```

## Provider confirmation

Import maps narration text to existing slides and performs the existing local
save/status/version bookkeeping. It does not call ElevenLabs or HeyGen, create
takes, queue generation, authorize avatars, or rewrite take history. The HTTP/UI
test rejects any request outside bootstrap/new/project/save/narration routes,
runs no worker, and disables provider credentials in its isolated server.
Backend checks verify zero provider calls, no generated jobs, and unchanged
narration-file hashes across import. Existing selection-review reads remain
read-only and do not authorize or submit production.
