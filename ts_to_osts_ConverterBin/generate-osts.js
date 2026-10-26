#!/usr/bin/env node
/*
 * generate-osts.js — convert Excel Office Script sources (.ts / .tsx) into
 * ready-to-use .osts files, without opening Excel.
 *
 * An .osts file is the form Excel itself saves an Office Script in: drop it
 * into the Office Scripts folder of OneDrive / SharePoint and Excel offers
 * it as a script. This tool lets scripts live as plain, reviewable text in
 * version control and be packaged mechanically, instead of being pasted
 * into Excel's editor by hand.
 *
 * ---------------------------------------------------------------------------
 * The format (reverse-engineered from Excel-saved files, then verified:
 * this tool's output is BYTE-IDENTICAL to Excel's own .osts output for
 * every script in the verification set — see README.md, "Verification").
 *
 * An .osts file is a single-line JSON object (UTF-8, no BOM, no trailing
 * newline, compact separators, non-ASCII written literally). Keys, in
 * order:
 *   version       "0.3.0"
 *   body          the script text, formatted the way Excel's editor
 *                 formats it (see below — this is the non-obvious part)
 *   description   "" (Excel does not derive this from header comments)
 *   parameterInfo a JSON *string* (double-encoded): the signature contract
 *                 Excel / Power Automate use for the script's parameters
 *   apiInfo       a JSON *string* (double-encoded):
 *                 {"variant":"synchronous","variantVersion":2}
 *
 * The body is NOT the raw source. Excel's editor (Monaco + the TypeScript
 * language service) re-formats code when it saves, with a 4-space indent:
 *   - block lines indent by nesting level x 4 spaces;
 *   - continuation lines of a wrapped statement or condition indent to the
 *     statement's indent + 4 — NOT to the author's alignment (a source
 *     line aligned at column 18 under the first array element and a
 *     condition line indented 6 both come back at statement indent + 4);
 *   - commas in code gain a following space (["Jan","Feb"] is saved as
 *     ["Jan", "Feb"]).
 * Everything else — comments, strings, spacing the author already had in
 * canonical form — is preserved exactly. Because the formatting is
 * syntax-aware, this tool runs the actual TypeScript formatter (the same
 * engine family Excel's editor uses) instead of approximating it with
 * text rules; the format options below are the validated set, verified
 * against real Excel output. Do not tweak them casually.
 *
 * Filenames: Excel does not allow "." in an Office Script name, so output
 * names replace every "." in the source stem with "_":
 * Clean.Data.ts -> Clean_Data.osts. Files whose names contain no dots are
 * unaffected apart from the extension.
 *
 * Parameters: scripts whose main() takes only
 * (workbook: ExcelScript.Workbook) — the common case — get the exact
 * parameterInfo Excel itself emits for that shape (verified). For scripts
 * with additional main() parameters this tool builds a best-effort
 * parameterInfo from the parameter names, their TypeScript types
 * (string / number / boolean map to the same JSON Schema types, anything
 * else becomes "string"), and any @param descriptions in the JSDoc block
 * directly above main(); that path is reasonable but is NOT verified
 * against an Excel specimen — see README.md, "Limitations".
 *
 * Dependencies (declared up front):
 *   - Node.js, and npm (npm is only needed for the one-time fix-up below).
 *   - The npm package "typescript", pinned at 5.9.3, installed LOCALLY
 *     beside this script (./node_modules). No global installs, no admin
 *     rights, no privilege elevation of any kind.
 * Self-repair: if the dependency is missing, this script says so, attempts
 * that local install in the open, verifies the installed version, and
 * only then proceeds. If the install fails, it stops with a clear
 * message instead of failing obscurely mid-run.
 *
 * Usage:
 *   node generate-osts.js <input_dir> [output_dir]
 *     input_dir   folder containing .ts / .tsx Office Script sources
 *     output_dir  where .osts files are written (default: ./osts-out)
 * ---------------------------------------------------------------------------
 */
"use strict";

const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const TS_VERSION = "5.9.3";
const HERE = __dirname;

// --- Self-repair: make sure the one declared dependency is present. -------
function loadTypeScript() {
  try {
    const ts = require("typescript");
    if (ts.version === TS_VERSION) return ts;
    console.log(`Found typescript ${ts.version}, need ${TS_VERSION} - reinstalling locally.`);
  } catch (_) {
    console.log(`Missing dependency: typescript@${TS_VERSION} (not found beside this script).`);
  }
  console.log(`Attempting local install: npm install --prefix "${HERE}" typescript@${TS_VERSION}`);
  console.log("(local folder only - no global install, no admin rights, no elevation)");
  try {
    execFileSync("npm", ["install", "--prefix", HERE, "--no-audit", "--no-fund",
      `typescript@${TS_VERSION}`], { stdio: "inherit" });
  } catch (e) {
    console.error("Install failed. Install Node.js/npm, then run the command above yourself.");
    process.exit(1);
  }
  const ts = require(path.join(HERE, "node_modules", "typescript"));
  if (ts.version !== TS_VERSION) {
    console.error(`Installed typescript is ${ts.version}, expected ${TS_VERSION}. Aborting.`);
    process.exit(1);
  }
  console.log(`Verified: typescript ${ts.version} ready.`);
  return ts;
}

const ts = loadTypeScript();

// --- The validated formatter option set (the heart of the map). ------------
const FORMAT_OPTIONS = {
  IndentSize: 4, TabSize: 4, NewLineCharacter: "\n", ConvertTabsToSpaces: true,
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
  InsertSpaceBeforeTypeAnnotation: false,
};

function formatLikeExcel(text) {
  const fileName = "script.ts";
  const host = {
    getCompilationSettings: () => ({ allowNonTsExtensions: true }),
    getScriptFileNames: () => [fileName],
    getScriptVersion: () => "1",
    getScriptSnapshot: (name) =>
      name === fileName ? ts.ScriptSnapshot.fromString(text) : undefined,
    getCurrentDirectory: () => "",
    getDefaultLibFileName: () => "lib.d.ts",
    fileExists: () => true,
    readFile: () => undefined,
  };
  const service = ts.createLanguageService(host);
  const edits = service.getFormattingEditsForDocument(fileName, FORMAT_OPTIONS, text);
  let result = text;
  for (const e of [...edits].sort((a, b) => b.span.start - a.span.start)) {
    result = result.slice(0, e.span.start) + e.newText +
             result.slice(e.span.start + e.span.length);
  }
  return result;
}

// --- parameterInfo ----------------------------------------------------------
// Find main()'s parameter list (text between the parens of "function main(",
// tolerating nested parens in default values) plus the JSDoc block directly
// above it, if any.
function readMainSignature(text) {
  const m = /function\s+main\s*\(/.exec(text);
  if (!m) return null;
  const openIdx = m.index + m[0].length - 1;
  let depth = 0, closeIdx = -1;
  for (let i = openIdx; i < text.length; i++) {
    if (text[i] === "(") depth++;
    else if (text[i] === ")") { depth--; if (depth === 0) { closeIdx = i; break; } }
  }
  if (closeIdx < 0) return null;
  const paramsText = text.slice(openIdx + 1, closeIdx);

  // Split on top-level commas only.
  const rawParams = [];
  let d = 0, start = 0;
  for (let i = 0; i <= paramsText.length; i++) {
    const c = paramsText[i];
    if (c === "(" || c === "[" || c === "{") d++;
    else if (c === ")" || c === "]" || c === "}") d--;
    if (i === paramsText.length || (c === "," && d === 0)) {
      rawParams.push(paramsText.slice(start, i).trim());
      start = i + 1;
    }
  }
  const params = rawParams.filter(Boolean).map((p) => {
    const nameMatch = /^([A-Za-z_$][\w$]*)/.exec(p);
    const typeMatch = /:\s*([^=]+?)\s*(=|$)/.exec(p);
    return {
      name: nameMatch ? nameMatch[1] : p,
      type: typeMatch ? typeMatch[1].trim() : "",
    };
  });

  // JSDoc directly above main(): summary line + @param descriptions.
  let comment = "";
  const docComments = {};
  const before = text.slice(0, m.index).trimEnd();
  const jsdoc = /\/\*\*([\s\S]*?)\*\/$/.exec(before);
  if (jsdoc) {
    const lines = jsdoc[1].split("\n").map((l) => l.replace(/^\s*\*\s?/, "").trim());
    const summaryLines = [];
    for (const line of lines) {
      const pm = /^@param\s+(\S+)\s*(.*)$/.exec(line);
      if (pm) docComments[pm[1]] = pm[2].trim();
      else if (!line.startsWith("@") && line) summaryLines.push(line);
    }
    comment = summaryLines.join(" ");
  }
  return { params, comment, docComments };
}

function schemaType(tsType) {
  if (tsType === "string" || tsType === "number" || tsType === "boolean") return tsType;
  return "string";
}

function buildParameterInfo(text, fileLabel) {
  const sig = readMainSignature(text);
  if (!sig) {
    console.error(`WARNING: ${fileLabel}: no "function main(" found - emitting workbook-only parameterInfo.`);
    return buildParameterInfoFor([{ name: "workbook", type: "ExcelScript.Workbook" }], "", {});
  }
  return buildParameterInfoFor(sig.params, sig.comment, sig.docComments);
}

function buildParameterInfoFor(params, comment, docComments) {
  const extra = params.filter((p) => p.name !== "workbook");
  const parameterSchema = { type: "object", default: {}, "x-ms-visibility": "internal" };
  if (extra.length > 0) {
    parameterSchema.properties = {};
    for (const p of extra) parameterSchema.properties[p.name] = { type: schemaType(p.type) };
  }
  const info = {
    version: 1,
    originalParameterOrder: extra.map((p) => p.name),
    parameterSchema,
    returnSchema: { type: "object", properties: {} },
    signature: {
      comment: comment || "",
      parameters: params.map((p) => ({ name: p.name, comment: docComments[p.name] || "" })),
    },
  };
  if (extra.length > 0) {
    console.error("NOTE: script has parameters beyond (workbook) - parameterInfo is " +
      "best-effort and not verified against an Excel specimen (see README).");
  }
  return JSON.stringify(info);
}

const API_INFO = '{"variant":"synchronous","variantVersion":2}';

function makeOsts(tsxText, fileLabel) {
  const obj = {
    version: "0.3.0",
    body: formatLikeExcel(tsxText),
    description: "",
    parameterInfo: buildParameterInfo(tsxText, fileLabel),
    apiInfo: API_INFO,
  };
  // JSON.stringify gives compact separators and insertion key order; written
  // as UTF-8 it keeps non-ASCII literal - all matching Excel's own output.
  return Buffer.from(JSON.stringify(obj), "utf8");
}

function main() {
  const inputDir = process.argv[2];
  if (!inputDir) {
    console.error("Usage: node generate-osts.js <input_dir> [output_dir]");
    process.exit(1);
  }
  const inDir = path.resolve(inputDir);
  const outDir = process.argv[3] ? path.resolve(process.argv[3])
                                 : path.join(process.cwd(), "osts-out");
  fs.mkdirSync(outDir, { recursive: true });
  const files = fs.readdirSync(inDir)
    .filter((f) => f.endsWith(".ts") || f.endsWith(".tsx")).sort();
  if (files.length === 0) {
    console.error(`No .ts/.tsx sources found in ${inDir}. Aborting.`);
    process.exit(1);
  }
  for (const f of files) {
    const text = fs.readFileSync(path.join(inDir, f), "utf8");
    const stem = f.replace(/\.tsx?$/, "").replace(/\./g, "_");
    const target = path.join(outDir, stem + ".osts");
    fs.writeFileSync(target, makeOsts(text, f));
    console.log(`${f} -> ${path.basename(target)} (${fs.statSync(target).size} bytes)`);
  }
}

main();
