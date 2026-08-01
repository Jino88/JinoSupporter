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

function importantise(block) {
  return block
    .split(';')
    .map(d => {
      const t = d.trim();
      if (!t || t.includes('!important') || !t.includes(':')) return d;
      return d.replace(/\s+$/, '') + ' !important';
    })
    .join(';');
}

/** Rewrite one comma-separated selector list. */
function scopeSelectorList(list) {
  return list
    .split(',')
    .map(s => s.trim())
    .filter(Boolean)
    .map(scopeOne)
    .join(', ');
}

function scopeOne(sel) {
  if (sel.startsWith(ROOT)) return sel;                       // already scoped
  if (sel === ':root') return ROOT;                           // token block
  if (sel === 'html' || sel === 'body') return ROOT;          // page-level resets
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
  const stack = [];          // what each open brace belongs to

  for (let i = 0; i < css.length; i++) {
    const ch = css[i];

    // pass comments through verbatim; they also carry the bridge markers
    if (ch === '/' && css[i + 1] === '*') {
      const end = css.indexOf('*/', i + 2);
      const stop = end === -1 ? css.length : end + 2;
      const comment = css.slice(i, stop);
      if (comment.includes(BRIDGE_START)) bridge = true;
      else if (comment.includes(BRIDGE_END)) bridge = false;
      out += buf + comment;
      buf = '';
      i = stop - 1;
      continue;
    }

    if (ch === '{') {
      const head = buf.trim();
      const inKeyframes = stack.includes('keyframes');

      if (head.startsWith('@')) {
        const name = head.slice(1).split(/[\s({]/)[0].toLowerCase();
        stack.push(name === 'keyframes' || name.endsWith('keyframes') ? 'keyframes' : 'at');
        out += buf + '{';
      } else if (inKeyframes || depth > 0 && stack[stack.length - 1] === 'decl') {
        stack.push('decl');
        out += buf + '{';
      } else {
        stack.push('decl');
        const lead = buf.slice(0, buf.length - buf.trimStart().length);   // keep indentation
        out += lead + scopeSelectorList(head) + ' {';
      }

      buf = '';
      depth++;
      continue;
    }

    if (ch === '}') {
      const kind = stack.pop();
      depth--;
      out += (bridge && kind === 'decl' ? importantise(buf) : buf) + '}';
      buf = '';
      continue;
    }

    buf += ch;
  }
  return out + buf;
}

const src = fs.readFileSync(SRC, 'utf8');
const header =
`/* GENERATED FILE — do not edit.
   Source: instrument.css · Regenerate: node tools/scope-css.js
   Every selector is scoped under ${ROOT} so this can be loaded alongside Bootstrap. */

`;
fs.writeFileSync(OUT, header + scope(src), 'utf8');

const rules = (src.match(/\{/g) || []).length;
console.log(`scoped ${rules} blocks -> ${path.relative(process.cwd(), OUT)}`);
