/**
 * Scope the INSTRUMENT design system so it can be loaded next to Bootstrap.
 *
 * The app loads Bootstrap globally and the design system defines its own .btn,
 * .badge, .row and .small — plus a `*` reset and :root variables. Loading it
 * as-is would restyle every existing page. This rewrites every selector under a
 * single `.ins` root so it only applies inside a container that opts in:
 *
 *     :root                  ->  .ins
 *     [data-theme="dark"]    ->  .ins[data-theme="dark"]
 *     html, body             ->  .ins
 *     *, *::before           ->  .ins *, .ins *::before
 *     .btn:hover             ->  .ins .btn:hover
 *
 * `.ins .btn` (0,2,0) outranks Bootstrap's `.btn` (0,1,0), so the design system
 * wins inside the container and changes nothing outside it.
 *
 * Two regions are treated specially, both marked by comments in the source:
 * BRIDGE gets !important, RESET is dropped. See the constants below.
 *
 * Usage: node tools/scope-css.js
 *   in  JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.css
 *   out JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.scoped.css
 *
 * Re-run after every edit to instrument.css — the scoped file is generated,
 * never hand-edited.
 */

const fs = require('fs');
const path = require('path');

const ROOT = '.ins';
const SRC = path.join(__dirname, '..', 'JinoSupporter.Web', 'wwwroot', 'ui-redesign', 'assets', 'instrument.css');
const OUT = path.join(__dirname, '..', 'JinoSupporter.Web', 'wwwroot', 'ui-redesign', 'assets', 'instrument.scoped.css');

/**
 * Declarations between these markers are emitted with !important.
 *
 * app.css re-skins Bootstrap with !important on nearly every rule, so scoping
 * alone does not win inside .ins — `.ins .btn-primary` still loses to
 * `.btn-primary { ... !important }`. With !important on both sides specificity
 * decides again, and `.ins X` (0,2,0) beats `X` (0,1,0). Marking the region
 * keeps the source readable instead of repeating !important by hand.
 */
const BRIDGE_START = 'BRIDGE-START';
const BRIDGE_END = 'BRIDGE-END';

/**
 * Declarations between these markers are dropped entirely.
 *
 * The element reset is right for the static mockup and wrong for the app. Under
 * .ins it becomes `.ins button` / `.ins table` / `.ins a` (0,1,1), and because
 * the shell is the layout for every page, that outranks the (0,1,0) classes the
 * pages style their own buttons, tables, links and inputs with — 242 page rules
 * lose. instrument.css §1b restates the same properties on the design system's
 * own controls, so dropping the region costs the shell nothing.
 */
const RESET_START = 'RESET-START';
const RESET_END = 'RESET-END';

const RESET_NOTE =
  `/* instrument.css §1 element reset omitted — it is scoped to .ins here, where
   \`.ins button\`/\`.ins table\`/\`.ins a\` would outrank every hosted page's own
   classes. See §1b in the source for the design system's own restatement. */`;

/**
 * Append !important to every declaration in a block.
 *
 * Splitting on ';' is wrong: a semicolon is ordinary text inside quotes and
 * inside url(), and `background: url("data:image/svg+xml;charset=utf8,...")`
 * would be cut in half. This walks characters instead, tracking paren depth and
 * quote state, so only a top-level ';' ends a declaration.
 *
 * Comments are carried along rather than split out. A comment is whitespace to
 * CSS, so `color: red /* why *\/ !important` is valid — but a fragment that is
 * only a comment must not get an !important of its own, hence the bare test.
 */
function importantise(block) {
  const uncomment = s => s.replace(/\/\*[\s\S]*?\*\//g, '');

  let out = '';
  let decl = '';
  let depth = 0;      // ( ) nesting
  let quote = null;   // ' or " while open

  const flush = () => {
    const bare = uncomment(decl).trim();
    out += bare && bare.includes(':') && !bare.includes('!important')
      ? decl.replace(/\s+$/, '') + ' !important'
      : decl;
    decl = '';
  };

  for (let i = 0; i < block.length; i++) {
    const ch = block[i];

    if (quote) {
      decl += ch;
      if (ch === '\\') decl += block[++i] ?? '';
      else if (ch === quote) quote = null;
      continue;
    }
    if (ch === '"' || ch === "'") { quote = ch; decl += ch; continue; }
    if (ch === '/' && block[i + 1] === '*') {            // comment: opaque
      const end = block.indexOf('*/', i + 2);
      const stop = end === -1 ? block.length : end + 2;
      decl += block.slice(i, stop);
      i = stop - 1;
      continue;
    }
    if (ch === '(') depth++;
    else if (ch === ')') depth--;
    else if (ch === ';' && depth === 0) { flush(); out += ';'; continue; }

    decl += ch;
  }
  flush();
  return out;
}

/**
 * Rewrite one comma-separated selector list.
 *
 * Deduplicated because several source selectors collapse onto the same scoped
 * one — `html, body` both become `.ins`, and `*` expands to a list that already
 * contains `.ins`. Emitting `.ins, .ins` is harmless but reads like a bug.
 */
function scopeSelectorList(list) {
  const seen = new Set();
  return list
    .split(',')
    .map(s => s.trim())
    .filter(Boolean)
    .flatMap(s => scopeOne(s).split(',').map(x => x.trim()))
    .filter(s => !seen.has(s) && seen.add(s))
    .join(', ');
}

/**
 * `.ins` already leads this selector, so it needs no wrapping. The boundary
 * matters: `.ins-row` merely starts with the same characters and is a normal
 * design-system class that must still be scoped.
 */
const ALREADY_SCOPED = new RegExp(`^${ROOT.replace('.', '\\.')}(?![\\w-])`);

function scopeOne(sel) {
  if (ALREADY_SCOPED.test(sel)) return sel;                   // already scoped
  if (sel === ':root') return ROOT;                           // token block
  if (sel === 'html' || sel === 'body') return ROOT;          // page-level resets
  // `*` alone would scope to `.ins *`, which skips the container itself — the
  // one element that is not a descendant of `.ins`. box-sizing has to reach it.
  if (sel === '*') return `${ROOT}, ${ROOT} *`;
  if (sel.startsWith('[data-theme=')) {
    // theme flag lives on the container itself, not on <html>
    const rest = sel.slice(sel.indexOf(']') + 1).trim();
    const attr = sel.slice(0, sel.indexOf(']') + 1);
    return rest ? `${ROOT}${attr} ${rest}` : `${ROOT}${attr}`;
  }
  return `${ROOT} ${sel}`;
}

/**
 * Walk the stylesheet tracking brace depth. Selectors are rewritten only at the
 * top level and inside conditional at-rules; @keyframes step selectors (from,
 * to, 50%) and declaration blocks are passed through untouched.
 */
function scope(css) {
  let out = '';
  let buf = '';
  let depth = 0;
  let bridge = false;        // inside the !important region
  let reset = false;         // inside the dropped region
  let dropped = 0;           // rules the reset region cost us, for the log
  const stack = [];          // what each open brace belongs to

  // everything the walker writes goes through here, so one flag mutes the
  // whole reset region without threading a condition through each branch
  const emit = s => { if (!reset) out += s; };

  for (let i = 0; i < css.length; i++) {
    const ch = css[i];

    // pass comments through verbatim; they also carry the region markers
    if (ch === '/' && css[i + 1] === '*') {
      const end = css.indexOf('*/', i + 2);
      const stop = end === -1 ? css.length : end + 2;
      const comment = css.slice(i, stop);

      // Inside a declaration block the comment belongs to the buffer, not the
      // output. Flushing here would emit the declarations before it untouched,
      // so a commented bridge rule would silently lose its !important — the
      // block is only importantised as a whole, at its closing brace.
      if (stack[stack.length - 1] === 'decl') {
        buf += comment;
        i = stop - 1;
        continue;
      }

      if (comment.includes(RESET_START)) {
        emit(buf + RESET_NOTE);           // leave a marker where the block was
        reset = true;
      } else if (comment.includes(RESET_END)) {
        reset = false;
      } else {
        if (comment.includes(BRIDGE_START)) bridge = true;
        else if (comment.includes(BRIDGE_END)) bridge = false;
        emit(buf + comment);
      }

      buf = '';
      i = stop - 1;
      continue;
    }

    if (ch === '{') {
      const head = buf.trim();
      const inKeyframes = stack.includes('keyframes');
      if (reset) dropped++;

      if (head.startsWith('@')) {
        const name = head.slice(1).split(/[\s({]/)[0].toLowerCase();
        stack.push(name === 'keyframes' || name.endsWith('keyframes') ? 'keyframes' : 'at');
        emit(buf + '{');
      } else if (inKeyframes || depth > 0 && stack[stack.length - 1] === 'decl') {
        stack.push('decl');
        emit(buf + '{');
      } else {
        stack.push('decl');
        const lead = buf.slice(0, buf.length - buf.trimStart().length);   // keep indentation
        emit(lead + scopeSelectorList(head) + ' {');
      }

      buf = '';
      depth++;
      continue;
    }

    if (ch === '}') {
      const kind = stack.pop();
      depth--;
      emit((bridge && kind === 'decl' ? importantise(buf) : buf) + '}');
      buf = '';
      continue;
    }

    buf += ch;
  }

  // an unclosed RESET-START would silently swallow the rest of the stylesheet
  if (reset) throw new Error(`${RESET_START} was never closed by ${RESET_END}`);

  return { css: out + buf, dropped };
}

const src = fs.readFileSync(SRC, 'utf8');
const header =
`/* GENERATED FILE — do not edit.
   Source: instrument.css · Regenerate: node tools/scope-css.js
   Every selector is scoped under ${ROOT} so this can be loaded alongside Bootstrap. */

`;
const result = scope(src);
const output = header + result.css;
const rules = (src.match(/\{/g) || []).length;
const summary =
  `scoped ${rules - result.dropped} blocks ` +
  `(${result.dropped} dropped with the element reset)`;

// `node tools/scope-css.js check` verifies instead of writing, so the build can
// fail when the committed file no longer matches its source. Line endings are
// normalised first: git may check the file out as CRLF while this writes LF,
// and that difference is not drift.
const nl = s => s.replace(/\r\n/g, '\n');

if (process.argv.slice(2).includes('check')) {
  const current = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
  if (nl(current) !== nl(output)) {
    console.error(
      `${path.basename(OUT)} is out of date with ${path.basename(SRC)}.\n` +
      `Run: node tools/scope-css.js`
    );
    process.exit(1);
  }
  console.log(`${summary} — generated file is up to date`);
} else {
  fs.writeFileSync(OUT, output, 'utf8');
  console.log(`${summary} -> ${path.relative(process.cwd(), OUT)}`);
}
