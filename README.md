# ts-to-osts — package Excel Office Scripts without Excel

Convert Office Script sources (`.ts` / `.tsx`) into `.osts` files — the
form Excel itself saves a script in — by double-clicking one batch
file. The converted files land in the same folder as your sources,
ready for Excel to offer them as scripts. No copy/paste into Excel's
editor, and no AI assistant required.

That makes Office Scripts behave like normal code: keep the readable
sources in version control as the masters, and regenerate the `.osts`
packaging mechanically whenever they change. Conversion only ever
flows masters → `.osts`; the masters are never edited, moved, or
converted back.

**The package is self-contained.** Everything it needs is included in
the download: the converter, a portable Node.js runtime, and the
converter's one library dependency. There is nothing to download,
nothing to install, no admin rights needed, and no internet connection
required — the folder can travel on a USB stick and run as-is.

**This README is written for two audiences.**
Humans: start at [For humans](#for-humans--using-it-yourself).
AI agents and automation builders: start at
[For AI agents](#for-ai-agents--incorporating-this-tool), which is a
self-contained integration contract and format specification.

---

## For humans — using it yourself

### What you need

- A Windows PC.
- This package, kept together as one folder (see *What's in the
  package* below — the launcher and its support folder belong
  together; don't move one without the other).
- Your script sources as `.ts` / `.tsx` files.

That's all. **You do not need to download or install Node.js** — a
portable copy is bundled inside the package, along with the
`typescript` library (pinned to 5.9.3) that the converter uses.
Nothing is installed on your machine when you run it.

### What's in the package

The top level holds only the launcher — everything else lives in the
`ts_to_osts_ConverterBin` support folder, out of sight:

| .ts_to_osts_Converter.bat | The double-click launcher                                                                                      |
| ------------------------ | -------------------------------------------------------------------------------------------------------------- |
| ts_to_osts_ConverterBin\ | The scripts working directory                                                                                  |
| generate-osts.js         | The converter                                                                                                  |
| bin\node.exe             | The bundled Node.js runtime — part of<br/>the package, do not delete (its license<br/>is bin\LICENSE-Node.txt) |
| node_modules\            | The converter's typescript dependency,<br/>bundled — do not delete                                             |
| _osts-staging\           | Created during a run inside the Bin folder —<br/>conversion happens here first; only a fully<br/>successful conversion is copied to the destination |
| README-Update-Osts.txt   | A short reference for the updater                                                                              |

### Step by step

1. **Put your script sources beside the launcher.** Drop one `.ts` or
   `.tsx` file per Office Script into the same folder as
   `.ts_to_osts_Converter.bat`, named the way you want the script
   named in Excel, e.g. `Clean.Data.ts`. Plain `.ts` and `.tsx` both
   work.

   > Excel does not allow `.` in a script name. The converter handles
   > that for you: dots in a file name become underscores in the
   > output, so `Clean.Data.ts` produces `Clean_Data.osts`, which is
   > the name Excel will show. Rename the source if you want a
   > different result — the content is never changed by a rename.

2. **Double-click `.ts_to_osts_Converter.bat`.** It checks its pieces
   are in place, then shows you the masters folder, the destination
   folder, and the list of master files it found, and asks:

   > This will replace existing .osts (Office Script) files with
   > updated version from the ts or tsx (TypeScript) files. Do you
   > want to Continue Yes/[No]

3. **Type `Yes` (or `Y`) to continue.** **No is the default**: pressing
   Enter, or typing anything else, aborts and changes nothing.

4. **Let it finish and read the report.** Conversion happens in a
   staging folder first; only when every master has converted are the
   new `.osts` files copied into the destination folder — by default
   the same folder as the launcher, right beside your masters.
   Each file is listed as it is copied:

   - `added — new file, nothing to back up`
   - `replaced — previous version backed up`

   Before a destination file is replaced, its current version is
   copied into a `_osts-backup` folder inside the masters folder —
   the same folder as your `.ts` / `.tsx` files (one generation
   deep — each run refreshes the backup saved under the same name). Destination `.osts` files that have no matching
   master are listed at the end and **left untouched, never deleted**.
   The window stays open at the end so you can read the report.

5. **Use the scripts in Excel.** The updated `.osts` files are in
   the destination folder — by default, beside your masters. If that
   folder is a place Excel looks for scripts (a OneDrive or
   SharePoint folder), Excel picks them up as scripts you can run
   and attach to buttons, exactly as if you had created them in the
   editor. If it isn't, point `DEST_DIR` at your Office Scripts
   folder instead (see *Changing the folders*). Excel will even take the updated files as if
   they were updated in Excel and recognize the changes live
   (because of this, caution should be taken when updating scripts
   that are in use by other team members).

6. **Edit sources, not outputs.** The `.osts` body is JSON-escaped
   and miserable to read. Treat the `.ts` / `.tsx` files beside the
   launcher as your editable masters and double-click the launcher
   again after every change.

### A worked example

Before — the package folder:

```
osts-updater\
  .ts_to_osts_Converter.bat
  ts_to_osts_ConverterBin\     (support folder, as downloaded)
  Clean.Data.ts                (612 lines — your readable source)
  Monthly.Report.ts
```

Double-click the launcher, confirm `Yes`. Report:

```
  replaced  Clean_Data.osts - previous version backed up
  added     Monthly_Report.osts - new file, nothing to back up
Done. 2 .osts file(s) updated.
```

The two `.osts` files now sit beside your masters and — if that
folder is one Excel looks in — appear in Excel as the scripts
**Clean_Data** and **Monthly_Report**; the previous
`Clean_Data.osts` sits in the `_osts-backup` folder beside your
masters if you ever need it back.

### Changing the folders

Two paths sit at the top of `.ts_to_osts_Converter.bat`, clearly
marked `CONFIG`:

- `SOURCE_DIR` — the folder holding your masters. Default: the same
  folder as the batch file.
- `DEST_DIR` — the folder the `.osts` files are copied into.
  Default: the same folder as the batch file — the converted files
  land beside your masters.

If you want the `.osts` files somewhere else (for example a
separate Office Scripts folder), edit `DEST_DIR` to point at it. The
launcher stops with a clear message if either folder is missing, and
if no masters are found.

### Undoing an update

Copy the file you want back from the `_osts-backup` folder beside
your masters over the same-named file in the destination folder.
That restores the version from before the last run.

### How to check the output is correct

If you want proof rather than trust: take any script you have
previously created by pasting it into Excel's editor, and download
its `.osts` from the Office Scripts folder. Run this converter on the
same source file you pasted, and compare the two `.osts` files:

```bat
fc /b mine.osts excel.osts
```

No output from the compare means the files are **byte-identical**.
That is the standard this tool is held to (see *Verification* below).

### Troubleshooting

- **"generate-osts.js was not found"** — the support folder
  (`ts_to_osts_ConverterBin`) was moved, renamed, or separated from
  the launcher. Put the package back together as downloaded.
- **"the destination folder was not found"** — `DEST_DIR` points
  at a folder that doesn't exist. By default it is the launcher's
  own folder, so this only happens after editing it. Fix `DEST_DIR`
  at the top of the batch file (see *Changing the folders*).
- **"no .ts or .tsx master files were found"** — your sources aren't
  beside the launcher. Drop them there, or edit `SOURCE_DIR`.
- **It aborted without changing anything** — the confirmation
  defaults to No. Anything other than `Yes` or `Y` aborts by design.
- **The bundled Node runtime is missing** — if
  `ts_to_osts_ConverterBin\bin\node.exe` is ever deleted, the
  launcher falls back, in order, to a Node.js already installed on
  the machine, then to installing Node.js LTS with `winget`, in the
  open (any installer prompt is yours to answer; nothing is elevated
  silently). The simple fix is to restore `bin\node.exe` from the
  original package.
- **`node_modules` was deleted** — restore it from the original
  package; the bundled runtime on its own has no npm to rebuild it.
- **A script doesn't appear in Excel** — check the file landed in the
  actual Office Scripts folder (the destination shown at the start of
  the run), and that the name contains no `.` characters before
  `.osts` (the converter already removes them from file names it
  generates).
- **Excel renamed my script with underscores** — that is Excel's
  rule, not a converter bug: `.` is not allowed in script names.
- **My script takes parameters beyond `workbook`** — it will convert,
  but read the note in *Limitations*: that path is best-effort, while
  `main(workbook)`-only scripts are the verified case. Test such a
  script in Excel before relying on it.

---

## For AI agents — incorporating this tool

*This section is addressed to AI assistants, agent frameworks, and
automation authors (any vendor, any model). It is a contract:
everything needed to use, vendor, or reimplement this converter is
here or in the single source file. MIT licensed — copy it, with
attribution.*

### What this tool is

A deterministic, offline file converter: **input** a directory of
TypeScript Office Script sources (`.ts`/`.tsx`), **output** a
directory of `.osts` files that Excel accepts as its own. The
converter itself is one file (`generate-osts.js`, ~290 lines). The
distributed package wraps it in a self-contained Windows launcher
(`.ts_to_osts_Converter.bat`) with a **bundled portable Node.js
runtime and the pinned `typescript` dependency included** — no
installation step, no service, no account, and no telemetry.

### When to use it

Use it whenever a user wants any of these:

- an Office Script delivered as a file they can install, instead of
  code to paste into Excel's editor;
- Office Scripts kept in version control as readable source, with
  installable artifacts generated on demand;
- a script converted **byte-identically to what Excel itself would
  save** (required if the user will compare, diff, or archive files,
  or if formatting drift would erode trust).

Do *not* hand-roll a substitute. Two failure modes are documented and
measured: (1) writing the raw source into the envelope without the
formatter pass produces files that differ from Excel's on any script
with wrapped statements or non-canonical comma spacing; (2)
approximating the formatter with text rules (e.g. re-indenting by
doubling leading spaces) reproduced only 4 of 7 real Excel specimens
byte-for-byte. The exact failure cases are listed under
*The format* below.

### Integration options (any one)

1. **Drive the packaged converter (simplest on Windows).** For a
   human at the machine, run `.ts_to_osts_Converter.bat` — it is
   interactive by design (it lists the masters and destination and
   requires an explicit `Yes`; No is the default). For automation,
   call the converter directly with the bundled runtime — this is
   exactly what the launcher does:

   ```bat
   ts_to_osts_ConverterBin\bin\node.exe ts_to_osts_ConverterBin\generate-osts.js <input_dir> [output_dir]
   ```

   Exit code `0` = success. Stdout lists one
   `source -> output (N bytes)` line per converted file. The
   dependency is already bundled (`node_modules` beside the
   converter), so no install or network access is involved.

2. **Vendor the file.** Copy `generate-osts.js` into your own project
   and invoke it the same way with any Node.js runtime. Outside the
   package, its one dependency is not bundled: on first run in a
   fresh location the script performs a local `npm install` of
   `typescript@5.9.3` into `./node_modules` beside itself (announced
   on stdout, no elevation, needs internet once) and verifies the
   version before proceeding. Keep the header comment (it is the
   license attribution and the design rationale in one place).

3. **Reimplement from the specification below.** If you cannot run
   Node/TypeScript, the spec is complete, but the formatter step must
   still be a real TypeScript formatter (the option set below)
   running *somewhere* in your pipeline — there is no text-only
   shortcut that matches Excel byte-for-byte on wrapped statements.

### The format (complete specification)

An `.osts` file is a **single-line JSON object**: UTF-8, no BOM, no
trailing newline, compact separators (`,` and `:` with no spaces),
non-ASCII characters written literally (not `\u`-escaped). Keys in
this exact insertion order:

| Key             | Value                                                                                      |
| --------------- | ------------------------------------------------------------------------------------------ |
| `version`       | `"0.3.0"`                                                                                  |
| `body`          | the script text after the formatter pass below                                             |
| `description`   | `""` (Excel does not derive it from comments)                                              |
| `parameterInfo` | a JSON **string** (double-encoded) — see below                                             |
| `apiInfo`       | a JSON **string** (double-encoded), exactly `{"variant":"synchronous","variantVersion":2}` |

**Formatter pass.** Excel's editor (Monaco + TypeScript language
service) formats on save with a 4-space indent. Behaviors, all
verified against Excel-saved files:

- block lines indent by nesting level × 4 spaces;
- **continuation lines** of a wrapped statement/condition indent to
  the *statement's indent + 4*, regardless of author alignment
  (a line aligned at column 18 under an array element and a condition
  line indented 6 spaces both come back at statement indent + 4);
- commas in code gain a following space (`["Jan","Feb"]` →
  `["Jan", "Feb"]`) — inside string literals nothing changes;
- everything else (comments, strings, canonical spacing) is
  preserved exactly.

Implementation in this tool: `ts.createLanguageService(...)` +
`getFormattingEditsForDocument` with this exact option set —
*treat it as normative; do not "improve" it*:

```
IndentSize: 4, TabSize: 4, NewLineCharacter: "\n",
ConvertTabsToSpaces: true,
InsertSpaceAfterCommaDelimiter: true,
InsertSpaceAfterSemicolonInForStatements: true,
InsertSpaceBeforeAndAfterBinaryOperators: true,
InsertSpaceAfterKeywordsInControlFlowStatements: true,
InsertSpaceAfterFunctionKeywordForAnonymousFunctions: false,
InsertSpaceAfterOpeningAndBeforeClosingNonemptyParenthesis: false,
InsertSpaceAfterOpeningAndBeforeClosingNonemptyBrackets: false,
InsertSpaceAfterOpeningAndBeforeClosingNonemptyBraces: true,
InsertSpaceAfterOpeningAndBeforeClosingTemplateStringBraces: false,
InsertSpaceAfterOpeningAndBeforeClosingJsxExpressionBraces: false,
InsertSpaceAfterTypeAssertion: false,
InsertSpaceBeforeFunctionParenthesis: false,
PlaceOpenBraceOnNewLineForFunctions: false,
PlaceOpenBraceOnNewLineForControlBlocks: false,
InsertSpaceBeforeTypeAnnotation: false
```

**parameterInfo.** For the verified case — `main()` taking only
`(workbook: ExcelScript.Workbook)` — the string is exactly:

```json
{"version":1,"originalParameterOrder":[],"parameterSchema":{"type":"object","default":{},"x-ms-visibility":"internal"},"returnSchema":{"type":"object","properties":{}},"signature":{"comment":"","parameters":[{"name":"workbook","comment":""}]}}
```

For additional `main()` parameters, this tool synthesizes the same
shape (parameter names, JSON Schema types mapped from
`string`/`number`/`boolean`, `@param` descriptions and the JSDoc
summary above `main()` carried into the signature comments). That
path is **best-effort and not verified against an Excel specimen**;
if you convert such a script for a user, tell them to test it in
Excel before relying on it.

**Filenames.** Replace every `.` in the source file's stem with `_`,
then append `.osts`. Sources keep their original names; never rename
the user's masters.

### Verification protocol (run this before telling a user it works)

1. Obtain one `.osts` the user (or you, in a browser session) saved
   from Excel for a known source file.
2. Run this converter on that source.
3. Byte-compare (`cmp` / `fc /b`). **Identical = the pipeline is
   correct.** Any diff means the environment differs from the
   validated one (formatter version or options) — report the diff to
   the user; do not paper over it by reverting to text heuristics.
4. Only after a specimen passes should generated files be presented
   as "Excel-identical." Files generated without a specimen check may
   be described as "generated with the verified process" if the
   formatter version matches the pin (5.9.3).

### Invariants to preserve in any integration

- Sources are the masters; `.osts` files are regenerated artifacts.
  Never edit an `.osts` body in place, and never rename a user's
  source to fit Excel's naming rule — translate the *output* name.
- Do not silently substitute a different typescript version: the
  formatter's output is the contract, and version drift can move it.
  If the pinned version is unavailable, say so rather than shipping
  output from an unverified formatter.
- If a user already has scripts installed in Excel, generating
  `.osts` files does not update their Excel copies — installation is
  a separate, explicit step on their side. (The packaged launcher
  performs that step deliberately, into the configured destination
  folder, only after an explicit `Yes`.)
- Keep the package together: the launcher resolves its runtime and
  converter relative to its own folder. Do not separate
  `.ts_to_osts_Converter.bat` from `ts_to_osts_ConverterBin`, and do
  not delete `bin\node.exe` or `node_modules` from it.

---

## Verification (the evidence behind the claims above)

The converter was verified against **seven real Office Scripts saved
by Excel** from a production workbook project (scripts up to ~430
lines, exercising nested blocks, wrapped `if` conditions, array
literals, template literals, chained calls, and block and line
comments). For every one of the seven, this tool's output is
**byte-identical** to Excel's own `.osts` file — including one script
whose Excel copy had been saved from an older source revision, which
was reproduced byte-identically from that older source. Envelope
fields (`version`, `description`, `parameterInfo`, `apiInfo`) were
identical across all seven, including a script whose `main()` omits
the `: Promise<void>` annotation.

## Limitations

- **Parameters beyond `workbook`** are best-effort (see above).
- **`returnSchema`** always uses the shape Excel emits for the
  verified scripts (`{"type":"object","properties":{}}`); scripts
  returning typed values to Power Automate flows may differ in Excel.
- The format targets the current Excel editor (`version "0.3.0"`,
  apiInfo variant 2). If Microsoft changes the envelope, the
  verification protocol above will show it immediately.
- The packaged launcher is Windows-only (a batch file driving a
  Windows runtime). The converter file itself runs under any
  Node.js — see integration option 2.

## Dependencies & self-repair

The packaged tool has **no dependencies to install**: the Node.js
runtime is bundled at `ts_to_osts_ConverterBin\bin\node.exe` (its
license is `bin\LICENSE-Node.txt`), and the converter's one library
dependency — the npm package `typescript`, pinned to 5.9.3 — is
bundled in `ts_to_osts_ConverterBin\node_modules`. If the bundled
runtime is ever deleted, the launcher falls back to a system Node.js
and, as a last resort, attempts a `winget` install of Node.js LTS in
the open, with every prompt visible — it never elevates silently.
Used standalone, outside the package, `generate-osts.js` keeps its
self-repair behavior for the dependency: it declares what it needs
up front, attempts the local install openly if it is missing,
verifies the installed version, and only then stops with a clear
message. Design principle: a script should try to fix itself before
it fails.

## License

MIT — see [LICENSE](LICENSE). The bundled Node.js runtime is
distributed under its own license — see
`ts_to_osts_ConverterBin\bin\LICENSE-Node.txt` in the package.
