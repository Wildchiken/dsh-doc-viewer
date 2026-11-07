/**
 * A small formula evaluator for when LibreOffice is busy or missing. It handles the
 * formulas most workbooks are made of: arithmetic, comparisons, text joined with &,
 * references (other sheets included), SUM / AVERAGE / MIN / MAX / COUNT / COUNTA /
 * ROUND / ROUNDUP / ROUNDDOWN / ABS / IF / IFERROR / AND / OR / NOT, and Excel's error
 * values (#DIV/0!, #VALUE!). Anything else is left unknown, never guessed, so the file
 * keeps no result for it and the next program to open the file works it out.
 */

const UNKNOWN = Symbol('unknown');

/** @param {string} code */
const error = (code) => ({ error: code });
/** @param {any} value */
const isError = (value) => typeof value === 'object' && value !== null && typeof value.error === 'string';
const DIV0 = error('#DIV/0!');
const VALUE = error('#VALUE!');

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
    if ((match = /^"((?:[^"]|"")*)"/.exec(rest))) {
      tokens.push({ type: 'str', value: match[1].replace(/""/g, '"') });
    } else if ((match = /^(?:'((?:[^']|'')+)'|([A-Za-z_\u0080-￿][\w.\u0080-￿]*))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/.exec(rest))) {
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
    } else if ((match = /^(TRUE|FALSE)(?![\w(])/i.exec(rest))) {
      tokens.push({ type: 'bool', value: match[1].toUpperCase() === 'TRUE' });
    } else if ((match = /^(<=|>=|<>|[-+*/^(),&=<>])/.exec(rest))) {
      tokens.push({ type: match[0] });
    } else {
      return undefined; // names, arrays, error literals: not ours
    }
    at += match[0].length;
  }
  return tokens;
}

/** A number as Excel shows it in General format, at most 15 significant digits. */
function numberText(value) {
  return String(Number(value.toPrecision(15)));
}

/**
 * Work out what formulas can be worked out here.
 * @param {(sheet: string, address: string) => { value?: any, formula?: string } | undefined} cellAt
 *   a value that is an Excel error is `{ error: '#N/A' }`, as ExcelJS gives it
 * @returns {(sheet: string, address: string) => number | string | boolean | { error: string } | undefined}
 *   the result of a formula cell, or undefined when it is not one or cannot be worked out
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
    if (value instanceof Date || (typeof value === 'object' && value !== null && !isError(value))) value = UNKNOWN; // dates, rich text, hyperlinks
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

  /** A value used as a number, as Excel coerces it. */
  const toNumber = (value) => {
    if (value === UNKNOWN || isError(value)) return value;
    if (value === null || value === '') return 0;
    if (typeof value === 'boolean') return value ? 1 : 0;
    if (typeof value === 'number') return Number.isFinite(value) ? value : UNKNOWN;
    const number = Number(String(value).trim());
    return String(value).trim() !== '' && Number.isFinite(number) ? number : VALUE;
  };
  const toText = (value) => {
    if (value === UNKNOWN || isError(value)) return value;
    if (value === null) return '';
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    if (typeof value === 'number') return numberText(value);
    return String(value);
  };
  const toBoolean = (value) => {
    if (value === UNKNOWN || isError(value)) return value;
    if (typeof value === 'boolean') return value;
    if (value === null) return false;
    if (typeof value === 'number') return value !== 0;
    return VALUE;
  };
  /** The first problem among values: something unknown, then an error. */
  const problem = (...values) => values.find((value) => value === UNKNOWN) ?? values.find(isError);

  const AGGREGATES = {
    SUM: (numbers) => numbers.reduce((sum, n) => sum + n, 0),
    AVERAGE: (numbers) => (numbers.length ? numbers.reduce((sum, n) => sum + n, 0) / numbers.length : DIV0),
    MIN: (numbers) => (numbers.length ? Math.min(...numbers) : 0),
    MAX: (numbers) => (numbers.length ? Math.max(...numbers) : 0),
  };

  /** Numbers for SUM and the like: in a range only numbers count, an error in it is the answer. */
  const numbersOf = (list) => {
    const numbers = [];
    for (const item of list) {
      if (item.range) {
        if (item.range === UNKNOWN) return UNKNOWN;
        for (const value of item.range) {
          if (isError(value)) return value;
          if (typeof value === 'number') numbers.push(value);
        }
      } else {
        const value = toNumber(item.value);
        if (value === UNKNOWN || isError(value)) return value;
        numbers.push(value);
      }
    }
    return numbers;
  };

  const rounding = (value, digits, mode) => {
    const bad = problem(value, digits);
    if (bad) return bad;
    const factor = 10 ** Math.trunc(digits);
    const scaled = Math.abs(value) * factor;
    const rounded = mode === 'up' ? Math.ceil(scaled - 1e-9) : mode === 'down' ? Math.floor(scaled + 1e-9) : Math.round(scaled + 1e-9);
    return (Math.sign(value) * rounded) / factor;
  };

  /** @param {string} formula @param {string} sheet */
  function evaluate(formula, sheet) {
    const tokens = tokenize(formula);
    if (!tokens) return UNKNOWN;
    let at = 0;
    const peek = () => tokens[at];
    const take = (type) => (tokens[at]?.type === type ? tokens[at++] : undefined);

    const args = () => {
      const list = [];
      if (take(')')) return list;
      do {
        const token = peek();
        // A bare range is a list of values; anything else is one value.
        if (token?.type === 'ref' && token.to && [',', ')'].includes(tokens[at + 1]?.type)) {
          at += 1;
          list.push({ range: rangeValues(token, sheet) });
        } else list.push({ value: comparison() });
      } while (take(','));
      return take(')') ? list : UNKNOWN;
    };

    const call = (name) => {
      const list = args();
      if (list === UNKNOWN) return UNKNOWN;
      if (AGGREGATES[name]) {
        const numbers = numbersOf(list);
        return Array.isArray(numbers) ? AGGREGATES[name](numbers) : numbers;
      }
      if (name === 'COUNT' || name === 'COUNTA') {
        let count = 0;
        for (const item of list) {
          if (item.range === UNKNOWN || item.value === UNKNOWN) return UNKNOWN;
          for (const value of item.range ?? [item.value]) {
            if (name === 'COUNT' ? typeof value === 'number' : value !== null && value !== '') count += 1;
          }
        }
        return count;
      }
      if (list.some((item) => item.range)) return UNKNOWN; // a range where one value belongs
      const values = list.map((item) => item.value);
      switch (name) {
        case 'ABS': return values.length === 1 ? (problem(toNumber(values[0])) ?? Math.abs(toNumber(values[0]))) : UNKNOWN;
        case 'ROUND': return values.length === 2 ? rounding(toNumber(values[0]), toNumber(values[1]), 'half') : UNKNOWN;
        case 'ROUNDUP': return values.length === 2 ? rounding(toNumber(values[0]), toNumber(values[1]), 'up') : UNKNOWN;
        case 'ROUNDDOWN': return values.length === 2 ? rounding(toNumber(values[0]), toNumber(values[1]), 'down') : UNKNOWN;
        case 'IF': {
          if (values.length < 2 || values.length > 3) return UNKNOWN;
          const condition = toBoolean(values[0]);
          if (condition === UNKNOWN || isError(condition)) return condition;
          // Only the branch taken matters, as in Excel.
          return condition ? values[1] : (values.length === 3 ? values[2] : false);
        }
        case 'IFERROR': {
          if (values.length !== 2 || values[0] === UNKNOWN) return UNKNOWN;
          return isError(values[0]) ? values[1] : values[0];
        }
        case 'AND':
        case 'OR': {
          if (!values.length) return UNKNOWN;
          const booleans = values.map(toBoolean);
          const bad = problem(...booleans);
          if (bad) return bad;
          return name === 'AND' ? booleans.every(Boolean) : booleans.some(Boolean);
        }
        case 'NOT': {
          if (values.length !== 1) return UNKNOWN;
          const value = toBoolean(values[0]);
          return problem(value) ?? !value;
        }
        default: return UNKNOWN;
      }
    };

    const primary = () => {
      const token = tokens[at++];
      if (!token) return UNKNOWN;
      if (token.type === 'num' || token.type === 'str' || token.type === 'bool') return token.value;
      if (token.type === 'ref') {
        if (token.to) return UNKNOWN;
        return valueOf(token.sheet ?? sheet, `${token.from[0]}${token.from[1]}`);
      }
      if (token.type === 'fn') return take('(') ? call(token.name) : UNKNOWN;
      if (token.type === '(') {
        const value = comparison();
        return take(')') ? value : UNKNOWN;
      }
      // As in Excel, a leading minus belongs to the number: -2^2 is 4.
      if (token.type === '-') { const value = toNumber(primary()); return problem(value) ?? -value; }
      if (token.type === '+') return primary();
      return UNKNOWN;
    };
    const arithmetic = (left, right, op) => {
      const a = toNumber(left);
      const b = toNumber(right);
      const bad = problem(a, b);
      if (bad) return bad;
      switch (op) {
        case '+': return a + b;
        case '-': return a - b;
        case '*': return a * b;
        case '/': return b === 0 ? DIV0 : a / b;
        case '^': return a ** b;
        default: return UNKNOWN;
      }
    };
    const power = () => {
      let value = primary();
      while (peek()?.type === '^') { at += 1; value = arithmetic(value, primary(), '^'); }
      return value;
    };
    const term = () => {
      let value = power();
      while (peek()?.type === '*' || peek()?.type === '/') { const op = tokens[at++].type; value = arithmetic(value, power(), op); }
      return value;
    };
    const sum = () => {
      let value = term();
      while (peek()?.type === '+' || peek()?.type === '-') { const op = tokens[at++].type; value = arithmetic(value, term(), op); }
      return value;
    };
    const joined = () => {
      let value = sum();
      while (take('&')) {
        const a = toText(value);
        const b = toText(sum());
        value = problem(a, b) ?? a + b;
      }
      return value;
    };
    /** Excel's order: numbers before text before TRUE/FALSE; text ignores case; empty is 0 or "". */
    const compare = (a, b) => {
      const rank = (value) => (typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2);
      if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
      if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
      if (rank(a) !== rank(b)) return rank(a) - rank(b);
      if (typeof a === 'string') return a.toLowerCase().localeCompare(b.toLowerCase());
      return Number(a) - Number(b);
    };
    function comparison() {
      let value = joined();
      while (['=', '<>', '<', '>', '<=', '>='].includes(peek()?.type)) {
        const op = tokens[at++].type;
        const right = joined();
        const bad = problem(value, right);
        if (bad) { value = bad; continue; }
        const order = compare(value, right);
        value = op === '=' ? order === 0 : op === '<>' ? order !== 0 : op === '<' ? order < 0 : op === '>' ? order > 0 : op === '<=' ? order <= 0 : order >= 0;
      }
      return value;
    }

    const value = comparison();
    if (at !== tokens.length || value === UNKNOWN) return UNKNOWN;
    if (value === null) return 0; // a formula pointing at an empty cell shows 0
    if (typeof value === 'number' && !Number.isFinite(value)) return error('#NUM!');
    return value;
  }

  return (sheet, address) => {
    const cell = cellAt(sheet, address);
    if (!cell?.formula) return undefined;
    const value = valueOf(sheet, address);
    return value === UNKNOWN ? undefined : value;
  };
}
