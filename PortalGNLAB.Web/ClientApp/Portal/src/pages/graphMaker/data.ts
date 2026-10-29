/** Convert records exports to the same delimited input consumed by GraphMakerParser. */
export function normalizeGraphData(text:string):string {
 if(!text.trimStart().startsWith("{"))return text;
 const data=JSON.parse(text) as {records?:unknown[]};
 if(!Array.isArray(data.records))throw Error("JSON must contain a records array.");
 const records=data.records.filter((r):r is Record<string,unknown>=>!!r&&typeof r==="object"&&!Array.isArray(r));
 const headers=[...new Set(records.flatMap(Object.keys))];
 const cell=(v:unknown)=>String(v??"").replace(/[\t\r\n]/g," ");
 return [headers.map(cell).join("\t"),...records.map(r=>headers.map(h=>cell(r[h])).join("\t"))].join("\n");
}
