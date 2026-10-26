Update-Osts — double-click .osts updater
=========================================

WHAT IT DOES
Double-click .ts_to_osts_Converter.bat and it rebuilds every .osts (Office
Script) file in the destination folder from the .ts / .tsx TypeScript
masters, using generate-osts.js (the verified converter from
https://github.com/jagodragon/ts-to-osts).

The masters rule: the .ts / .tsx files are the masters. Conversion
only ever flows masters -> .osts. Masters are never edited, moved, or
converted back.

BEFORE IT CHANGES ANYTHING
It shows the masters folder, the destination folder, and the master
files it found, then asks:

  This will replace existing .osts (Office Script) files with updated
  version from the ts or tsx (TypeScript) files. Do you want to
  Continue Yes/[No]

No is the default. Pressing Enter, or typing anything other than Yes,
aborts and changes nothing.

FILES IN THIS FOLDER
The top level holds only the batch file - everything else lives in
ts_to_osts_ConverterBin, out of sight:

  .ts_to_osts_Converter.bat   the double-click launcher
  ts_to_osts_ConverterBin\
    generate-osts.js         the converter
    bin\node.exe             the bundled Node.js runtime - part of
                             the package, do not delete (its license
                             is bin\LICENSE-Node.txt)
    node_modules\            the converter's typescript dependency,
                             bundled - do not delete
    _osts-staging\           created during a run; conversion happens
                             here first, and only a fully successful
                             conversion is copied over the destination
    README-Update-Osts.txt   this file

CONFIGURATION
Two paths at the top of .ts_to_osts_Converter.bat, clearly marked CONFIG:
  SOURCE_DIR   default: the same folder as the batch file -
               drop your .ts / .tsx masters right beside
               .ts_to_osts_Converter.bat
  DEST_DIR     default: the same folder as the batch file -
               the converted .osts files land beside your masters
If you want the .osts files in a different folder (for example a
separate Office Scripts folder), edit DEST_DIR to point at it. The
script stops with a clear message if either folder is missing.

DEPENDENCIES
None to install. This package is SELF-CONTAINED: it carries its own
portable Node.js runtime (bin\node.exe) and the converter's one
dependency (typescript, pinned, in node_modules). Nothing is
installed on the machine and no internet connection is needed - the
folder can travel on a USB stick as-is.
Fallbacks, in order, if bin\node.exe is ever deleted: a Node.js
already installed on the machine is used; failing that, the script
tries to install Node.js LTS with winget, in the open - any installer
prompt is yours to answer; nothing is elevated silently. (If
node_modules is ever deleted, restore it from the original package;
the bundled node.exe alone has no npm to rebuild it.)

SAFETY NOTES
- Conversion runs into the staging folder first. If anything fails,
  the destination folder is untouched.
- Before a destination .osts is replaced, its current version is
  copied to a _osts-backup folder inside the masters folder (the
  same folder as the .ts / .tsx files).
  Each run refreshes the backup saved under the same name - one
  generation deep. To undo an update, copy the file back from
  _osts-backup into the destination folder.
- Destination .osts files with no matching master are listed at the
  end and left untouched — never deleted.
- Paths with spaces (for example inside OneDrive folder names) are
  supported. One trap is handled for you: a source path ending in a
  backslash would escape the closing quote when passed to the
  converter, so the script trims it for that call only.
- The window stays open at the end (pause) so you can read the report.
