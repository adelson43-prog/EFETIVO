/* Lê o "Personnel escala" (abas SORTER, MOLDAGEM, OFFSET, BLOWMOLDING) e devolve a base de pessoas. */
(function (root) {
  const SHEETS = { SORTER: "Sorter", MOLDAGEM: "Moldagem", OFFSET: "Offset", BLOWMOLDING: "Blow Molding" };
  const FUNC_KEYS = [["líder", "Líder"], ["lider", "Líder"], ["auxiliar", "Auxiliar Operacional"], ["scrap man", "Scrap man"],
    ["abastecimento", "Abastecimento"], ["conferência scrap", "Conferência Scrap"], ["conferencia scrap", "Conferência Scrap"],
    ["balança", "Balança / Transf. Sorter"], ["operadores", "Operador"], ["embaladores", "Embalador"],
    ["afastados", "Afastado"], ["supervisor", "Supervisor"], ["impressor", "Impressor"], ["administrativo", "Administrativo"],
    ["grupo a", "Operador"], ["grupo b", "Operador"], ["grupo d", "Operador"]];
  const TURNOS = ["A/D", "B/E", "C/F"];

  function isMat(v) {
    if (typeof v === "number" && v >= 100 && v <= 999999 && Number.isInteger(v)) return String(v);
    if (typeof v === "string" && /^\s*\d{2}-\d{6}\s*$/.test(v)) return v.trim();
    return null;
  }
  function isName(v) {
    if (typeof v !== "string") return false;
    const low = v.toLowerCase();
    return v.trim().split(/\s+/).length >= 2 && /[A-Za-zÀ-ú]{3}/.test(v) && !["total", "pessoa", "turno", "grupo"].some(k => low.includes(k));
  }
  const empty = v => v === null || v === undefined || v === "";

  function sheetToGrid(ws, XLSX) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }

  function parseWorkbook(wb, XLSX) {
    const people = {};
    const missing = [];
    for (const [sh, setor] of Object.entries(SHEETS)) {
      const name = wb.SheetNames.find(n => n.trim().toUpperCase() === sh);
      if (!name) { missing.push(sh); continue; }
      const grid = sheetToGrid(wb.Sheets[name], XLSX);
      const found = [];
      grid.forEach((row, r) => {
        if (!row) return;
        for (let c = 0; c < row.length - 1; c++) {
          const m = isMat(row[c]);
          if (!m) continue;
          if (isName(row[c + 1])) found.push([r, c, m, c + 1, false]);
          else if (c + 2 < row.length && (empty(row[c + 1]) || row[c + 1] === "E") && isName(row[c + 2])) found.push([r, c, m, c + 2, row[c + 1] === "E"]);
        }
      });
      const cols = [...new Set(found.map(f => f[1]))].sort((a, b) => a - b);
      const blocks = [];
      cols.forEach(c => { const last = blocks[blocks.length - 1]; if (last && c - last[last.length - 1] <= 3) last.push(c); else blocks.push([c]); });
      const blockOf = c => blocks.findIndex(b => b.includes(c));
      for (const [r, c, m, nc, isE] of found) {
        const row = grid[r];
        const nome = String(row[nc]).replace(/\s+/g, " ").trim();
        const bi = blockOf(c);
        let func = "";
        for (let rr = r - 1; rr >= 0; rr--) {
          const g = grid[rr] || [];
          const txt = [];
          for (let k = Math.max(0, c - 4); k < Math.min(g.length, c + 3); k++) if (typeof g[k] === "string") txt.push(g[k]);
          const t = txt.join(" ").toLowerCase();
          const hit = FUNC_KEYS.find(([k]) => t.includes(k));
          if (hit) { func = hit[1]; break; }
        }
        const after = [], before = [];
        for (let k = nc + 1; k < Math.min(row.length, nc + 5); k++) if (!empty(row[k])) after.push(String(row[k]).trim());
        for (let k = Math.max(0, c - 3); k < c; k++) if (!empty(row[k])) before.push(String(row[k]).trim());
        if (isE) func = "Afastado";
        const tipo = after.some(x => x.toUpperCase().startsWith("TEMP")) ? "Temporário" : "Efetivo";
        const pcd = after.some(x => x.toUpperCase() === "PCD") || nome.toUpperCase().includes("PCD");
        const ferias = before.includes("F");
        let codigos = after.filter(x => !["PCD", "EFETIVO", "TEMP", "TEMPORÁRIO"].includes(x.toUpperCase()) && x.length <= 3);
        if (func === "Afastado") codigos = [];
        const turno = bi >= 0 && bi < 3 && !["Afastado", "Supervisor", "Administrativo", "Impressor"].includes(func) ? TURNOS[bi] : "ADM";
        if (people[m]) continue;
        people[m] = { mat: m, nome, setor, turno, funcao: func || "Operador", tipo, pcd, ferias, afastado: func === "Afastado", codigos, ativo: true };
      }
    }
    return { pessoas: people, faltando: missing };
  }

  const api = { parseWorkbook };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.PersonnelParser = api;
})(typeof window !== "undefined" ? window : globalThis);
