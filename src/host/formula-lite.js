/**
 * A small formula evaluator for when LibreOffice is busy or missing. It handles the
 * formulas most workbooks are made of: arithmetic, references (other sheets included)
 * and SUM / AVERAGE / MIN / MAX / COUNT / ROUND / ABS. Anything else is left unknown,
 * never guessed, so the file keeps no result for it and the next program to open the
 * file works it out.
 */

const UNKNOWN = Symbol('unknown');

/** @param {string} letters */
function columnIndex(letters) {
  let index = 0;
  for (const char of letters) index = index * 26 + (char.charCodeAt(0) - 64);
  return index - 1;
}

/** @param {number} index */
function columnLetters(index) {
  let letters = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return letters;
}

/** @param {string} source */
function tokenize(source) {
  const tokens = [];
  let at = 0;
  const text = source.replace(/^=/, '');
  while (at < text.length) {
    const rest = text.slice(at);
    let match;
    if ((match = /^\s+/.exec(rest))) { at += match[0].length; continue; }
    if ((match = /^(?:'((?:[^']|'')+)'|([A-Za-z_\u0080-￿][\w.\u0080-￿]*))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/.exec(rest))) {
      tokens.push({ type: 'ref', sheet: (match[1] ?? match[2]).replace(/''/g, "'"), from: [match[3], match[4]], to: match[5] ? [match[5], match[6]] : undefined });
    } else if ((match = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?(?![\w(])/.exec(rest))) {
      tokens.push({ type: 'ref', from: [match[1], match[2]], to: match[3] ? [match[3], match[4]] : undefined });
    } else if ((match = /^(\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?%?/.exec(rest))) {
      const percent = match[0].endsWith('%');
      tokens.push({ type: 'num', value: Number(percent ? match[0].slice(0, -1) : match[0]) / (percent ? 100 : 1) });
    } else if ((match = /^([A-Za-z][A-Za-z0-9.]*)\s*\(/.exec(rest))) {
      tokens.push({ type: 'fn', name: match[1].toUpperCase().replace(/^_XLFN\./, '') });
      at += match[0].length - 1;
      continue;
    } else if ((match = /^[-+*/^(),]/.exec(rest))) {
      tokens.push({ type: match[0] });
    } else {
      return undefined; // strings, comparisons, names, arrays: not ours
    }
    at += match[0].length;
  }
  return tokens;
}

/**
 * Work out what formulas can be worked out here.
 * @param {(sheet: string, address: string) => { value?: any, formula?: string } | undefined} cellAt
 * @returns {(sheet: string, address: string) => number | undefined}
 */
export function liteEvaluator(cellAt) {
  const memo = new Map();
  const visiting = new Set();

  const valueOf = (sheet, address) => {
    const key = `${sheet}!${address}`;
    if (memo.has(key)) return memo.get(key);
    const cell = cellAt(sheet, address);
    let value;
    if (!cell) value = null;
    else if (cell.formula) {
      if (visiting.has(key)) return UNKNOWN; // a cycle has no answer
      visiting.add(key);
      value = evaluate(cell.formula, sheet);
      visiting.delete(key);
    } else value = cell.value ?? null;
    memo.set(key, value);
    return value;
  };

  const rangeValues = (token, sheet) => {
    const where = token.sheet ?? sheet;
    const [c1, r1] = [columnIndex(token.from[0]), Number(token.from[1])];
    const [c2, r2] = token.to ? [columnIndex(token.to[0]), Number(token.to[1])] : [c1, r1];
    if ((Math.abs(c2 - c1) + 1) * (Math.abs(r2 - r1) + 1) > 100_000) return UNKNOWN;
    const values = [];
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r += 1) {
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c += 1) {
        const value = valueOf(where, `${columnLetters(c)}${r}`);
        if (value === UNKNOWN) return UNKNOWN;
        values.push(value);
      }
    }
    return values;
  };

  const FUNCTIONS = {
    SUM: (numbers) => numbers.reduce((sum, n) => sum + n, 0),
    AVERAGE: (numbers) => (numbers.length ? numbers.reduce((sum, n) => sum + n, 0) / numbers.length : UNKNOWN),
    MIN: (numbers) => (numbers.length ? Math.min(...numbers) : 0),
    MAX: (numbers) => (numbers.length ? Math.max(...numbers) : 0),
    COUNT: (numbers) => numbers.length,
  };

  /** @param {string} formula @param {string} sheet */
  function evaluate(formula, sheet) {
    const tokens = tokenize(formula);
    if (!tokens) return UNKNOWN;
    let at = 0;
    const peek = () => tokens[at];
    const take = (type) => (tokens[at]?.type === type ? tokens[at++] : undefined);

    const scalar = (value) => {
      if (value === UNKNOWN) return UNKNOWN;
      if (value === null || value === '') return 0;
      return typeof value === 'number' && Number.isFinite(value) ? value : UNKNOWN;
    };

    const args = () => {
      const list = [];
      if (take(')')) return list;
      do {
        const token = peek();
        // A bare range is a list of values; anything else is one value.
        if (token?.type === 'ref' && token.to && [',', ')'].includes(tokens[at + 1]?.type)) {
          at += 1;
          list.push({ range: rangeValues(token, sheet) });
        } else list.push({ value: expression() });
      } while (take(','));
      return take(')') ? list : UNKNOWN;
    };

    const call = (name) => {
      const list = args();
      if (list === UNKNOWN) return UNKNOWN;
      if (FUNCTIONS[name]) {
        const numbers = [];
        for (const item of list) {
          if (item.range) {
            if (item.range === UNKNOWN) return UNKNOWN;
            // Like Excel: text and empty cells in a range are skipped.
            for (const value of item.range) if (typeof value === 'number') numbers.push(value);
          } else {
            const value = scalar(item.value);
            if (value === UNKNOWN) return UNKNOWN;
            numbers.push(value);
          }
        }
        return FUNCTIONS[name](numbers);
      }
      const values = list.map((item) => (item.range ? UNKNOWN : scalar(item.value)));
      if (values.includes(UNKNOWN)) return UNKNOWN;
      if (name === 'ABS' && values.length === 1) return Math.abs(values[0]);
      if (name === 'ROUND' && values.length === 2) {
        const factor = 10 ** values[1];
        return Math.sign(values[0]) * Math.round(Math.abs(values[0]) * factor + 1e-9) / factor;
      }
      return UNKNOWN;
    };

    const primary = () => {
      const token = tokens[at++];
      if (!token) return UNKNOWN;
      if (token.type === 'num') return token.value;
      if (token.type === 'ref') {
        if (token.to) return UNKNOWN;
        return scalar(valueOf(token.sheet ?? sheet, `${token.from[0]}${token.from[1]}`));
      }
      if (token.type === 'fn') return take('(') ? call(token.name) : UNKNOWN;
      if (token.type === '(') {
        const value = expression();
        return take(')') ? value : UNKNOWN;
      }
      // As in Excel, a leading minus belongs to the number: -2^2 is 4.
      if (token.type === '-') return negate(primary());
      if (token.type === '+') return primary();
      return UNKNOWN;
    };
    const negate = (value) => (value === UNKNOWN ? UNKNOWN : -value);
    const apply = (a, b, op) => {
      if (a === UNKNOWN || b === UNKNOWN) return UNKNOWN;
      switch (op) {
        case '+': return a + b;
        case '-': return a - b;
        case '*': return a * b;
        case '/': return b === 0 ? UNKNOWN : a / b;
        case '^': return a ** b;
        default: return UNKNOWN;
      }
    };
    const power = () => {
      let value = primary();
      while (peek()?.type === '^') { at += 1; value = apply(value, primary(), '^'); }
      return value;
    };
    const term = () => {
      let value = power();
      while (peek()?.type === '*' || peek()?.type === '/') { const op = tokens[at++].type; value = apply(value, power(), op); }
      return value;
    };
    function expression() {
      let value = term();
      while (peek()?.type === '+' || peek()?.type === '-') { const op = tokens[at++].type; value = apply(value, term(), op); }
      return value;
    }

    const value = expression();
    return at === tokens.length && typeof value === 'number' && Number.isFinite(value) ? value : UNKNOWN;
  }

  return (sheet, address) => {
    const cell = cellAt(sheet, address);
    if (!cell?.formula) return undefined;
    const value = valueOf(sheet, address);
    return typeof value === 'number' ? value : undefined;
  };
}
