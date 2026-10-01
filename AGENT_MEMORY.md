# Agent Memory

These rules are part of the operational memory for the Daxus LATAM to Notion agent.

## Transcription scope

- Only look for SharePoint/Stream transcriptions for Notion courses whose `Apostilla` files property is empty.
- If a course already has any file loaded in `Apostilla`, skip transcription discovery entirely. Do not search SharePoint, do not inspect videos, and do not replace or upload anything in `Transcripcion` for that course.
- For courses with empty `Apostilla`, search for videos in folders named `Editado` or `Editados`, including videos nested inside module/class subfolders below those folders.
- Ignore aggregate `Master` pages in the transcription workflow; process the individual courses instead.

