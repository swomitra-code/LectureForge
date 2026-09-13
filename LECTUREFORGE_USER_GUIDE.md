# LectureForge V2: instructor guide

## Before you start

Use Windows while signed in to your normal desktop account. Desktop PowerPoint,
FFmpeg/FFprobe, and the `ELEVENLABS_API_KEY` and `HEYGEN_API_KEY` Windows environment
variables must already be configured. LectureForge checks presence without showing
keys. Keep API keys out of scripts, lecture files, and browser fields.

## Quick start

1. Double-click **LectureForge.cmd** in the LectureForge folder. The Studio opens
   in your default browser. Launching it again returns to the running instance.
2. Choose **New Lecture** and enter a project name.
3. Select your PowerPoint. LectureForge copies it; the original stays unchanged.
4. Select **Renewable Energy** and import scripts using headings like `## Slide 1`.
   Edit scripts and enabled slides as needed, then **Save Changes**. Optional
   silence after speech must be set before generating and reviewing candidates.
5. Click **GENERATE 3 TAKES** and confirm the displayed number of new generations.
6. **Review Narration**, play the three takes, and select your preference. Use
   **Previous Slide**, **Next Slide**, or the slide selector to move quickly.
7. Click **REVIEW SELECTIONS**. Check every mapping and any prepared pause.
8. Click **GENERATE AVATARS FOR SELECTED TAKES**, inspect the exact summary and
   new-job count, then **CONFIRM GENERATE FOR THESE SELECTIONS**.
9. Watch the dashboard until every enabled slide is **Avatar Ready**.
10. Click **Check Readiness** if needed, choose an output folder (or leave it blank
    for project output), then **CREATE RECORDING POWERPOINT**. Optional placement
    changes affect insertion only; **RESET TO PROJECT DEFAULT** restores the preset.
11. When ready, click **OPEN RECORDING POWERPOINT**.
12. In PowerPoint, use **Record** to add cursor, laser pointer, ink, and teaching emphasis.
13. Export your final lecture from PowerPoint.

## Selections and costs

Selecting or changing a narration take saves your preference only. **Change
Selection** clears that preference without deleting audio. The **CHANGE** action
in the summary returns to that slide. Neither action calls HeyGen.

Narration generation and confirmed new avatar jobs use provider credits. After
avatar confirmation, background production proceeds without another approval.
Exact compatible approved avatars can be reused with zero new jobs. Changing a
bound script, take, or pause makes an old authorization stale; review and confirm
again. A replacement that needs new media requires new explicit authorization.
Historical assets are retained. Never use a generic retry to replace an uncertain
paid submission.

## Return later or handle a failure

Use **Stop LectureForge** at the top of Studio for a clean stop. Already-submitted
provider jobs are not cancelled. Double-click the launcher to return, then use
**Open Existing Lecture** if the previous project does not reopen automatically.
Selections, authorizations, provider IDs, completed media, and outputs are saved.

For an exception, open **View Details** and use the action matching the failed
stage. Local validation and **RETRY ASSEMBLY** reuse existing media. A known
provider job can resume without submitting a replacement. If submission is
uncertain, reconcile its existing ID before any replacement; get technical help
if the ID is unavailable. Successful slides do not need another approval.

## Files and shortcut

Projects live in `%LOCALAPPDATA%\LectureForge\Projects\<project-id>\`.
Use **Open Project Folder** or **OPEN OUTPUT FOLDER** instead of managing internal
files. Final outputs can go to your chosen folder; existing recordings are
preserved through revision-safe filenames. Keep runtime projects in the short
local location even when source/final files are in Dropbox.

Optional shortcut: right-click `LectureForge.cmd`, choose **Show more options**
if needed, then **Send to > Desktop (create shortcut)**. Keep the LectureForge
folder in place. No installer or Windows service is required.

## V2.0 limits

- Internet access and provider credits are needed for new production; existing
  verified assets can be reused. Routine production does not require Codex.
- Desktop PowerPoint must run in your interactive Windows session. Final
  assembly is serialized; leave it time to finish.
- Output folder paths are limited to 160 characters. Explicit paths are tested;
  native folder-dialog selection has not been manually accepted.
- Technical checks verify media, bindings, and native content preservation.
  They do not judge teaching quality, pronunciation, or visual obstruction.
- The avatar is a separate editable video object. This is not a timeline editor;
  the final cursor/laser/ink pass and export happen in PowerPoint.
