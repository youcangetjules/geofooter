Archived VBA module copies - 4 September 2026
=============================================

These are superseded duplicates. The authoritative copy of every AES VBA
module is C:\GeoFooter\VBA\, which is the location IMPORT.txt documents.

Why they existed
----------------
The project layout moved twice and the old copies were never removed:

  1. project root   - the original home. The two oldest surviving files,
                      MSCANModule2.bas and MSCANModuleHeaders.bas, were
                      created 2025-10-20 and exist nowhere else.
  2. MSCAN\         - a complete mirror from Jul-Aug 2026.
  3. VBA\           - current, and the only one that should be imported.

Note: the project root is still the correct home for the Python application
(guri.py and friends). Only the .bas / .cls files there were strays.

What is in here
---------------
  root-modules\   18 .bas/.cls files taken from the project root
  MSCAN\          the entire MSCAN\ mirror, including its icons\ subfolder

Verified before archiving
-------------------------
Every file here is either byte-identical to its VBA\ counterpart or older.
Each apparent exception was checked individually:

  MSCANClassificationDialog.bas  The root copy was dated newer and contained
                                 RunHiddenProcess + VbsQuote, which VBA\ lacks.
                                 Both are dead code - never called - and both
                                 copies still launch via the same
                                 shell.Run(cmd, 1, True). VBA\ is the cleaner
                                 version with the abandoned experiment removed.

  MSCANSettings.bas              VBA\ is definitively correct. The archived
                                 copy still parses account entries with a fixed
                                 500-character window; store IDs can exceed
                                 that, which silently voided every save.

  ThisOutlookSession.cls         VBA\ is ahead: it captures Err.Number and
                                 Err.Description into locals before use.

  MSCANmodLogging.bas            Legacy lower-case "mod" spelling. IMPORT.txt
                                 lists this name under modules to remove.

  MSCANModule2.bas               Dead: depends on MSCANSecurityRatingForm,
                                 which IMPORT.txt says to delete.
  MSCANModuleHeaders.bas         Dead: self-documented as deprecated in favour
                                 of Module1.GetInternetHeaders.

Icons were byte-identical across all three locations. MSCAN\icons was only the
fourth and last fallback in the icon search order, behind VBA\icons, so
removing it changes nothing. The now-dead fallback entry was also stripped from
MSCANToolbar.ResolveIconPath, MSCANAppBootstrap.GetServiceIconPath and the
TARGET_DIRS list in make_aes_icons.py, so nothing recreates the folder.

Safe to delete
--------------
Nothing references this folder. It is kept only as a convenience; the same
content is recoverable from git history and from
_backup\GeoFooter_source_20260904_124007.zip.
