export function percent(value: string): number | null {
 const text = value.trim().replace(/%+$/, "").trim();
 if (!text || !/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(text)) return null;
 const n = Number(text); return Number.isFinite(n) ? n : null;
}
export function allowedExcel(name: string): boolean { return /\.(xlsx|xlsm|xlsb|xls)$/i.test(name); }
export function label(key: string): string { return key.replace(/([A-Z])/g, " $1").replace(/^./, x => x.toUpperCase()); }
