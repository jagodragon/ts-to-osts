# ts-to-osts — package Excel Office Scripts without Excel

Convert type script sources (`.ts` / `.tsx`) into `.osts` files — the
form Excel itself saves a script in — from the command line. Drop the
resulting files into the Office Scripts folder in OneDrive or SharePoint
and Excel offers them as scripts or you can save them some where else
and point Excel at them. No copy/paste into Excel's editor, and no AI 
assistant required.

That makes Office Scripts behave like normal code: keep the readable
sources in version control as the masters, and generate the `.osts`
packaging mechanically whenever they change.

**This README is written for two audiences.**
Humans: start at [For humans](#for-humans--using-it-yourself).
AI agents and automation builders: start at
[For AI agents](#for-ai-agents--incorporating-this-tool), which is a
self-contained integration contract and format specification.

---

## For humans — using it yourself

### What you need

- **Node.js** (any current LTS). Get it from
  [nodejs.org](https://nodejs.org) — on Windows, accept the default
  installer options; on macOS you can also use `brew install node`;
  on Linux use your distro's package manager or NodeSource.
  Check it worked by opening a terminal and running `node --version`.
- This folder (clone the repo, or download it as a zip and extract).
- Nothing else. The converter's one library dependency (the
  `typescript` package, pinned to 5.9.3) is **installed automatically
  the first time you run it**, into a `node_modules` folder next to the
  script. That needs an internet connection once. It never asks for
  admin rights and never installs anything system-wide.

### Step by step

1. **Put your script sources in a folder.** One Office Script per file,
   named the way you want the script named in Excel, e.g.
   `Clean.Data.ts`. Plain `.ts` and `.tsx` both work.

   > Excel does not allow `.` in a script name. The converter handles
   > that for you: dots in a file name become underscores in the
   > output, so `Clean.Data.ts` produces `Clean_Data.osts`, which is
   > the name Excel will show. Rename the source if you want a
   > different result — the content is never changed by a rename.

2. **Run the converter**, pointing it at your folder:

   ```sh
   node generate-osts.js path/to/my-scripts
   ```

   Converted files land in `./osts-out` unless you name a second
   folder:

   ```sh
   node generate-osts.js path/to/my-scripts path/to/output
   ```

   On the first run you will see it announce and perform the one-time
   local install of `typescript@5.9.3`, verify the version, and then
   list one line per script: source name → output name and size.

3. **Put the `.osts` files where Excel looks for scripts.** In OneDrive
   (or SharePoint) there is a folder used for Office Scripts
   (typically under `Documents/Office Scripts`). Copy the generated
   files there. Excel picks them up as scripts you can run and attach
   to buttons, exactly as if you had created them in the editor.

4. **Edit sources, not outputs.** The `.osts` body is JSON-escaped and
   miserable to read. Treat `.ts`/`.tsx` files as your editable
   masters and regenerate after every change. Re-running the converter
   overwrites the old outputs.

### A worked example

Before:

```
my-scripts/
  Clean.Data.ts        (612 lines — your readable, reviewed source)
  Monthly.Report.ts
```

Run: `node generate-osts.js my-scripts`

After:

```
osts-out/
  Clean_Data.osts
  Monthly_Report.osts
```

Copy the two `.osts` files to your Office Scripts folder and they
appear in Excel as the scripts **Clean_Data** and **Monthly_Report**.

### How to check the output is correct

If you want proof rather than trust: take any script you have
previously created by pasting it into Excel's editor, and download its
`.osts` from the Office Scripts folder. Run this converter on the same
source file you pasted. Compare the two `.osts` files:

- Windows: `fc /b mine.osts excel.osts`
- macOS/Linux: `cmp mine.osts excel.osts`

No output from the compare means the files are **byte-identical**.
That is the standard this tool is held to (see *Verification* below).

### Troubleshooting

- **`node` is not recognized / not found** — Node.js is not installed
  or the terminal predates the install. Install it, open a *new*
  terminal, and try again.
- **The first run fails during the install step** — you are probably
  offline, or npm is blocked. Connect, or install the dependency
  yourself with `npm install --prefix . typescript@5.9.3` from this
  folder, then re-run. The converter verifies the version before
  doing any work.
- **A script doesn't appear in Excel** — check the file is in the
  actual Office Scripts folder (not a similarly named one), and that
  the name contains no `.` characters before `.osts` (the converter
  already removes them from file names it generates).
- **Excel renamed my script with underscores** — that is Excel's
  rule, not a converter bug: `.` is not allowed in script names.
- **My script takes parameters beyond `workbook`** — it will convert,
  but read the note in *Limitations*: that path is best-effort, while
  `main(workbook)`-only scripts are the verified case. Test such a
  script in Excel before relying on it.

---

## For AI agents — incorporating this tool

*This section is addressed to AI assistants, agent frameworks, and
automation authors (any vendor, any model). It is a contract: everything
needed to use, vendor, or reimplement this converter is here or in the
single source file. MIT licensed — copy it, with attribution.*

### What this tool is

A deterministic, offline-after-setup file converter:
**input** a directory of TypeScript Office Script sources
(`.ts`/`.tsx`), **output** a directory of `.osts` files that Excel
accepts as its own. It is one self-contained Node.js file
(`generate-osts.js`, ~290 lines), plus this README and a license. There
is no service, no account, and no telemetry.

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

1. **Shell out to the CLI (simplest).**
   `node generate-osts.js <input_dir> [output_dir]`
   Exit code `0` = success. Stdout lists one
   `source -> output (N bytes)` line per converted file. First run in
   a fresh location performs a local `npm install` of
   `typescript@5.9.3` into `./node_modules` beside the script
   (announced on stdout, no elevation) and verifies the version before
   proceeding; stage that folder writable, or pre-run it once.
2. **Vendor the file.** Copy `generate-osts.js` into your own project
   and invoke it the same way. Keep the header comment (it is the
   license attribution and the design rationale in one place).
3. **Reimplement from the specification below.** If you cannot run
   Node/TypeScript, the spec is complete, but the formatter step must
   still be a real TypeScript formatter (the option set below) running
   *somewhere* in your pipeline — there is no text-only shortcut that
   matches Excel byte-for-byte on wrapped statements.

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
  a separate, explicit step on their side.

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

## Dependencies & self-repair

One dependency: the npm package `typescript` (pinned, 5.9.3),
installed **locally beside the script** — never globally, never with
elevated privileges. If it is missing, the script says so, attempts
that local install in the open, verifies the installed version, and
only then proceeds. Design principle: a script should try to fix
itself before it fails — declare its dependencies up front, attempt
the fix openly, verify, and only then stop with a clear message.

## License

MIT — see [LICENSE](LICENSE).
