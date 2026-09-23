export function colNum(col) {
  let x = 0;
  for (const ch of col) x = x * 26 + (ch.charCodeAt(0) - 64);
  return x;
}
export function colName(k) {
  let s = '';
  while (k > 0) {
    const m = (k - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    k = Math.floor((k - 1) / 26);
  }
  return s;
}
