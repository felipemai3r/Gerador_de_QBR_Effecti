/* Gerador de QBR — Effecti
 * App 100% no navegador. Lê o CSV de disputas, calcula os indicadores,
 * monta as páginas do relatório e exporta PDF ou PowerPoint.
 */
(function () {
  "use strict";

  const { jsPDF } = window.jspdf;

  // ---------- Estado ----------
  const state = {
    disputeRows: null,   // linhas do CSV de disputas
    csRows: null,        // linhas do CSV de Customer Success
    healthImg: null,     // dataURL do print do health score
    metrics: null,
    exportCfg: null,     // cfg da última prévia (para PDF/PPTX)
    candidatePages: null,// páginas que podem entrar no PDF
    selectedPages: {},   // páginas marcadas pelo operador
  };

  // ---------- Helpers ----------
  const $ = (id) => document.getElementById(id);

  const brl = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
  const numFmt = new Intl.NumberFormat("pt-BR");

  function normalize(s) {
    return String(s || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  }

  // Converte texto numérico (BR "1.234,56" ou US "1234.56") para Number
  function toNumber(v) {
    if (v === null || v === undefined) return 0;
    if (typeof v === "number") return v;
    let s = String(v).trim();
    if (!s) return 0;
    s = s.replace(/[R$\s]/g, "");
    const hasComma = s.includes(",");
    const hasDot = s.includes(".");
    if (hasComma && hasDot) {
      // formato BR: ponto = milhar, vírgula = decimal
      s = s.replace(/\./g, "").replace(",", ".");
    } else if (hasComma) {
      s = s.replace(",", ".");
    }
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  function numOrNull(v) {
    const s = String(v == null ? "" : v).trim();
    if (s === "") return null;
    const n = parseFloat(s.replace(",", "."));
    return isNaN(n) ? null : n;
  }

  function toast(msg, isError) {
    const t = $("toast");
    t.textContent = msg;
    t.className = "toast" + (isError ? " err" : "");
    t.hidden = false;
    clearTimeout(t._timer);
    t._timer = setTimeout(() => (t.hidden = true), 4200);
  }

  // Encontra o valor de uma coluna pela lista de nomes possíveis (normalizados/contains)
  function pick(row, keys, candidates) {
    for (const c of candidates) {
      const cn = normalize(c);
      for (const k of keys) {
        if (normalize(k) === cn) return row[k];
      }
    }
    // fallback: contains
    for (const c of candidates) {
      const cn = normalize(c);
      for (const k of keys) {
        if (normalize(k).includes(cn)) return row[k];
      }
    }
    return undefined;
  }

  // ---------- Leitura de CSV ----------
  function readCsv(file) {
    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: "greedy",
        transformHeader: (h) => h.trim(),
        complete: (res) => resolve(res.data.filter((r) => Object.values(r).some((v) => String(v).trim() !== ""))),
        error: reject,
      });
    });
  }

  // ---------- Parsing de dados colados do sistema interno ----------
  // Formato: campos em linhas sequenciais (19 por disputa) OU linhas com colunas
  // separadas por TAB. A âncora de cada disputa é a data (dd/mm/aaaa) na 3ª posição.
  const CAMPOS_COLADOS = [
    "Cod. Empresa", "Nome Empresa", "Data", "Nome Filial", "CNPJ Filial",
    "Portal", "Cod. Único", "Numero", "Numero UASG", "Nome UASG",
    "Total Items", "Items Vencidos", "Items Fracassados", "Items Fracassados com Margem",
    "Valor Total Vencidos", "Diferença Valor Mínimo", "Variação Inicial", "Variação Final", "Tipo Redução",
  ];
  const DATE_RE = /^\d{2}\/\d{2}\/\d{4}$/;

  function parsePasted(text) {
    const rawLines = String(text || "").split(/\r?\n/);

    // Variante 1: linhas com colunas separadas por TAB
    const tabRows = rawLines.filter((l) => l.split("\t").length >= 10);
    if (tabRows.length > 0) {
      return tabRows.map((l) => {
        const parts = l.split("\t").map((p) => p.trim());
        const row = {};
        CAMPOS_COLADOS.forEach((c, i) => (row[c] = parts[i] !== undefined ? parts[i] : ""));
        return row;
      });
    }

    // Variante 2: um campo por linha, âncora = data na 3ª posição de cada registro
    const lines = rawLines.map((l) => l.trim()).filter((l) => l !== "");
    const dateIdx = [];
    for (let i = 0; i < lines.length; i++) {
      if (DATE_RE.test(lines[i])) dateIdx.push(i);
    }
    if (!dateIdx.length) throw new Error("Não reconheci o formato colado (nenhuma data dd/mm/aaaa encontrada).");

    const rows = [];
    for (let k = 0; k < dateIdx.length; k++) {
      const start = dateIdx[k] - 2; // data é o 3º campo do registro
      const end = k + 1 < dateIdx.length ? dateIdx[k + 1] - 2 : lines.length;
      if (start < 0) continue;
      const chunk = lines.slice(start, end);
      const row = {};
      CAMPOS_COLADOS.forEach((c, i) => (row[c] = chunk[i] !== undefined ? chunk[i] : ""));
      rows.push(row);
    }
    return rows;
  }

  // ---------- Filtro por período (quando as linhas têm data) ----------
  function parseBrDate(s) {
    const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(s || "").trim());
    if (!m) return null;
    return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
  }

  function filterByPeriod(rows) {
    if (!rows.length) return { rows, filtered: false };
    const keys = Object.keys(rows[0]);
    const hasData = rows.some((r) => parseBrDate(pick(r, keys, ["Data", "Data de homologação"])));
    const ini = $("periodoInicio").value; // yyyy-mm
    const fim = $("periodoFim").value;
    if (!hasData || (!ini && !fim)) return { rows, filtered: false };

    let from = null, to = null;
    if (ini) { const [y, m] = ini.split("-").map(Number); from = new Date(y, m - 1, 1); }
    if (fim) { const [y, m] = fim.split("-").map(Number); to = new Date(y, m, 0, 23, 59, 59); }

    const out = rows.filter((r) => {
      const d = parseBrDate(pick(r, keys, ["Data", "Data de homologação"]));
      if (!d) return false;
      if (from && d < from) return false;
      if (to && d > to) return false;
      return true;
    });
    return { rows: out, filtered: true, total: rows.length };
  }

  // ---------- Cálculo dos indicadores ----------
  function computeMetrics(rows) {
    if (!rows || !rows.length) throw new Error("CSV de disputas vazio.");
    const keys = Object.keys(rows[0]);

    let valorArrematado = 0;
    let retorno = 0;
    let itensVencidosTotal = 0;
    let pregoesAcimaMinimo = 0;
    let licitacoesVencidas = 0;
    const orgaos = new Set();

    for (const r of rows) {
      const valVenc = toNumber(pick(r, keys, ["Valor Total Vencidos", "Valor Vencidos", "Valor Total"]));
      const dif = toNumber(pick(r, keys, ["Diferença Valor Mínimo", "Diferenca Valor Minimo", "Diferença Valor Mín"]));
      const itVenc = toNumber(pick(r, keys, ["Items Vencidos", "Itens Vencidos"]));
      const org = pick(r, keys, ["Nome UASG", "Órgão", "Orgao", "Nome Filial"]);

      valorArrematado += valVenc;
      retorno += dif;
      itensVencidosTotal += itVenc;
      if (dif > 0) pregoesAcimaMinimo += 1;
      if (itVenc > 0) licitacoesVencidas += 1;
      if (org) orgaos.add(String(org).trim());
    }

    const disputasRealizadas = rows.length;

    return {
      valorArrematado,
      retorno,
      itensVencidosTotal,
      pregoesAcimaMinimo,
      licitacoesVencidas,
      disputasRealizadas,
      orgaosUnicos: orgaos.size,
    };
  }

  // Abaixo deste ROI a página de ROI vem desmarcada por padrão (operador pode re-marcar)
  const ROI_MIN_PCT = 20;
  const ROI_MESES_OPTS = [3, 4, 5, 6];

  function getRoiMeses() {
    const v = parseInt($("roiMeses").value, 10);
    return ROI_MESES_OPTS.includes(v) ? v : 3;
  }

  function roiInvestimentoLabel(meses, ano) {
    const a = ano || new Date().getFullYear();
    if (meses === 3) return "Investimento Trimestral no Plano em " + a;
    if (meses === 6) return "Investimento Semestral no Plano em " + a;
    return "Investimento de " + meses + " Meses no Plano em " + a;
  }

  function roiInvestimentoResumo(meses) {
    if (meses === 3) return "Investimento trimestral (×3)";
    if (meses === 6) return "Investimento semestral (×6)";
    return "Investimento (" + meses + " meses)";
  }

  // ROI no formato do modelo: baseado no Valor Total Arrematado em 1º lugar
  // vs. o investimento no período (mensalidade × meses).
  function computeRoi(valorArrematado, mensalidade, meses, ano) {
    if (!mensalidade || mensalidade <= 0) return null;
    const m = ROI_MESES_OPTS.includes(meses) ? meses : 3;
    const investimento = mensalidade * m;
    const roiPct = ((valorArrematado - investimento) / investimento) * 100;
    const porReal = valorArrematado / investimento; // retorno bruto por R$1 investido
    const porRealNum = Number(porReal).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return {
      investimento,
      meses: m,
      investimentoLabel: roiInvestimentoLabel(m, ano),
      roiPct,
      porReal,
      // "658,716%" — inteiro com separador de milhar estilo do modelo
      roiStr: Math.round(roiPct).toLocaleString("en-US") + "%",
      porRealNum,                       // "6.588,15" (sem R$)
      porRealStr: "R$ " + porRealNum,   // "R$ 6.588,15"
    };
  }

  // ---------- Período (rótulo) ----------
  const MESES = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
  const MESES_ABREV = ["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
  function periodLabel() {
    const ini = $("periodoInicio").value; // yyyy-mm
    const fim = $("periodoFim").value;
    const fmt = (v) => {
      if (!v) return "";
      const [y, m] = v.split("-");
      return `${MESES[parseInt(m, 10)]} de ${y}`;
    };
    const a = fmt(ini), b = fmt(fim);
    if (a && b) {
      const sameYear = ini.split("-")[0] === fim.split("-")[0];
      if (sameYear) return `${MESES[parseInt(ini.split("-")[1], 10)]} a ${b}`;
      return `${a} a ${b}`;
    }
    return a || b || "";
  }

  // Período abreviado para nome de arquivo, ex.: "Mar a Jul" (ou "Mar 25 a Jul 26")
  function periodShort() {
    const ini = $("periodoInicio").value; // yyyy-mm
    const fim = $("periodoFim").value;
    const fmt = (v, withYear) => {
      if (!v) return "";
      const [y, m] = v.split("-");
      const mes = MESES_ABREV[parseInt(m, 10)] || "";
      return withYear ? `${mes} ${y.slice(-2)}` : mes;
    };
    if (ini && fim) {
      const sameYear = ini.split("-")[0] === fim.split("-")[0];
      if (sameYear) {
        const a = fmt(ini, false), b = fmt(fim, false);
        return a === b ? a : `${a} a ${b}`;
      }
      return `${fmt(ini, true)} a ${fmt(fim, true)}`;
    }
    return fmt(ini || fim, true);
  }

  // ---------- Construção das páginas ----------
  // Escala pt(modelo 720x405) -> px(slide 1280x720)
  const S = 1280 / 720;
  // Distância do topo da caixa de linha até a baseline, em frações da fonte,
  // para line-height:1 com métricas da Poppins (asc~1.05 / desc~0.35em → ~0.85em).
  const ASC = 0.84;

  // Cores extraídas do modelo
  const COR = {
    navy: "#002352",   // valores grandes (págs 2 e 3)
    blue: "#4285f4",   // destaques/azul
    cyan: "#b7eef4",   // fundo ciano chapado (pág 4)
    cyanCard: "#b2edf3", // card da pág 15
    navyBg: "#081128", // fundo azul-escuro (pág 6)
    gray: "#585752",   // texto do período
    ink: "#141414",    // texto preto do corpo (pág 4)
    ringOn: "#1a49ad", // bolinha preenchida
    ringOff: "#e7f4f7",// bolinha vazia
    ringBg: "#b3edf3", // disco de fundo (esconde dots do modelo)
    numLight: "#eef2f7", // números da pág 6
    black: "#080a0a",  // e-mail pág 15
  };
  // ---------- Limpeza do fundo (inpainting) ----------
  // Em vez de cobrir o texto do modelo com um retângulo chapado (que deixava
  // uma "emenda" sobre o fundo texturizado/degradê), reconstruímos o fundo real
  // por baixo do texto interpolando, coluna a coluna, os pixels logo acima e
  // logo abaixo da região. Assim o degradê/textura atravessa sem emenda e os
  // valores dinâmicos são desenhados por cima de um fundo limpo.
  // Regiões (em pt, base 720×405) que contêm texto do modelo a ser removido:
  const HOLES = {
    2: [[54, 138, 392, 52], [58, 318, 126, 15], [416, 355, 244, 24]],
    3: [[55, 232, 326, 46], [16, 335, 350, 46]],
    4: [
      [153, 172, 470, 21], [153, 243, 470, 21],
      [153, 281, 470, 24], [153, 303, 470, 24], [153, 325, 470, 24],
    ],
    5: [[96, 146, 122, 122], [231, 146, 122, 122], [367, 146, 122, 122], [502, 146, 122, 122]],
    6: [[358, 101, 144, 40], [358, 168, 144, 40], [358, 235, 144, 40], [358, 323, 144, 40]],
    15: [[43, 255, 214, 15]],
  };
  const CLEANBG = {}; // cache pageNum -> dataURL do fundo limpo
  const INPAINT_REV = 4;
  let _cleanRev = 0;

  function loadImg(src) {
    return new Promise((res, rej) => {
      const i = new Image();
      i.onload = () => res(i);
      i.onerror = rej;
      i.src = src;
    });
  }

  function inpaintRect(ctx, imgW, imgH, x0, y0, w, h) {
    const leftX = x0 - 1, rightX = x0 + w, topY = y0 - 1, botY = y0 + h;
    const hasL = leftX >= 0, hasR = rightX < imgW, hasT = topY >= 0, hasB = botY < imgH;
    const left = hasL ? ctx.getImageData(leftX, y0, 1, h).data : null;
    const right = hasR ? ctx.getImageData(rightX, y0, 1, h).data : null;
    const top = hasT ? ctx.getImageData(x0, topY, w, 1).data : null;
    const bot = hasB ? ctx.getImageData(x0, botY, w, 1).data : null;
    const fill = ctx.createImageData(w, h);
    for (let yy = 0; yy < h; yy++) {
      const ty = hasT && hasB ? (yy + 1) / (h + 1) : 0.5;
      for (let xx = 0; xx < w; xx++) {
        const tx = hasL && hasR ? (xx + 1) / (w + 1) : 0.5;
        let r = 0, g = 0, b = 0, n = 0;
        if (top) { const i = xx * 4; r += top[i] * (1 - ty); g += top[i + 1] * (1 - ty); b += top[i + 2] * (1 - ty); n += 1 - ty; }
        if (bot) { const i = xx * 4; r += bot[i] * ty; g += bot[i + 1] * ty; b += bot[i + 2] * ty; n += ty; }
        if (left) { const i = yy * 4; r += left[i] * (1 - tx); g += left[i + 1] * (1 - tx); b += left[i + 2] * (1 - tx); n += 1 - tx; }
        if (right) { const i = yy * 4; r += right[i] * tx; g += right[i + 1] * tx; b += right[i + 2] * tx; n += tx; }
        const di = (yy * w + xx) * 4;
        fill.data[di] = Math.round(r / Math.max(n, 1));
        fill.data[di + 1] = Math.round(g / Math.max(n, 1));
        fill.data[di + 2] = Math.round(b / Math.max(n, 1));
        fill.data[di + 3] = 255;
      }
    }
    ctx.putImageData(fill, x0, y0);
  }

  function inpaintPage(img, holes) {
    const c = document.createElement("canvas");
    c.width = img.width; c.height = img.height;
    const ctx = c.getContext("2d");
    ctx.drawImage(img, 0, 0);
    const sc = img.width / 720;
    for (const [xPt, yPt, wPt, hPt] of holes) {
      const x0 = Math.max(0, Math.round(xPt * sc));
      const y0 = Math.max(0, Math.round(yPt * sc));
      const w = Math.min(img.width - x0, Math.round(wPt * sc));
      const h = Math.min(img.height - y0, Math.round(hPt * sc));
      if (w <= 0 || h <= 0) continue;
      inpaintRect(ctx, img.width, img.height, x0, y0, w, h);
    }
    return c.toDataURL("image/png");
  }

  async function ensureCleanBackgrounds() {
    if (!window.MODELO_PAGES) return;
    if (_cleanRev === INPAINT_REV && Object.keys(CLEANBG).length) return;
    for (const k of Object.keys(CLEANBG)) delete CLEANBG[k];
    for (const n of Object.keys(HOLES)) {
      const key = String(n).padStart(2, "0");
      const src = window.MODELO_PAGES[key];
      if (!src) continue;
      try {
        const img = await loadImg(src);
        CLEANBG[n] = inpaintPage(img, HOLES[n]);
      } catch (e) { /* mantém o fundo original se falhar */ }
    }
    _cleanRev = INPAINT_REV;
  }

  // ---------- Leitura do print do Health Score (OCR) ----------
  // Layout padrão: 4 colunas — % em cima, rótulo embaixo (Encontrar…Monitorar).
  // Lê pela posição horizontal (coluna), não pela distância ao rótulo.
  const HS_ORDER = ["encontrar", "cadastrar", "disputar", "monitorar"];
  const HS_LABELS = {
    encontrar: ["encontrar", "encontar", "buscar"],
    cadastrar: ["cadastrar", "cadastro"],
    disputar: ["disputar", "disputa"],
    monitorar: ["monitorar", "monitoramento", "monitorial"],
  };
  const stripAccents = (s) => s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  const cxy = (b) => ({ x: (b.x0 + b.x1) / 2, y: (b.y0 + b.y1) / 2, w: b.x1 - b.x0, h: b.y1 - b.y0 });

  function clusterByX(items, gap) {
    const sorted = [...items].sort((a, b) => a.x - b.x);
    const clusters = [];
    for (const it of sorted) {
      const c = clusters.find((cl) => Math.abs(cl.x - it.x) < gap);
      if (!c) clusters.push({ ...it });
      else if (it.w > c.w || (it.hasPct && !c.hasPct)) Object.assign(c, it);
    }
    return clusters.sort((a, b) => a.x - b.x);
  }

  function assignPctColumns(pctCands, labelCands) {
    let pool = [...pctCands];
    if (labelCands.length >= 2) {
      const labelY = labelCands.reduce((s, l) => s + l.y, 0) / labelCands.length;
      const upper = pool.filter((p) => p.y < labelY + 8);
      if (upper.length >= 3) pool = upper;
    }
    const clusters = clusterByX(pool, 80);
    if (clusters.length >= 4) {
      const out = {};
      HS_ORDER.forEach((k, i) => { out[k] = clusters[i].val; });
      return { values: out, found: 4 };
    }
    if (clusters.length === 3 && pool.length >= 4) {
      // tenta separar colunas muito próximas
      const tight = clusterByX(pool, 45);
      if (tight.length >= 4) {
        const out = {};
        HS_ORDER.forEach((k, i) => { out[k] = tight[i].val; });
        return { values: out, found: 4 };
      }
    }
    // Fallback: casa cada rótulo com o % mais próximo na horizontal, acima do rótulo
    const labels = clusterByX(labelCands, 60);
    labels.sort((a, b) => a.x - b.x);
    const out = {};
    for (const lab of labels) {
      let best = null, bestDx = Infinity;
      for (const p of pctCands) {
        if (p.y > lab.y + 5) continue;
        const dx = Math.abs(p.x - lab.x);
        if (dx < bestDx) { bestDx = dx; best = p; }
      }
      if (best && bestDx < 130) out[lab.key] = best.val;
    }
    const found = Object.keys(out).length;
    return found ? { values: out, found } : null;
  }

  async function extractHealthPercents(dataUrl) {
    if (typeof Tesseract === "undefined") return { error: "OCR indisponível (Tesseract não carregado)." };
    const res = await Tesseract.recognize(dataUrl, "por", { tessedit_pageseg_mode: "6" });
    const words = (res && res.data && res.data.words) || [];
    const pctCands = [];
    const labelCands = [];
    for (const w of words) {
      const raw = (w.text || "").trim();
      if (!raw) continue;
      const norm = stripAccents(raw);
      const bbox = cxy(w.bbox);
      const mPct = raw.match(/(\d{1,3})\s*%/) || (/^(\d{1,3})%?$/.test(raw) && raw.match(/^(\d{1,3})/));
      if (mPct) {
        const v = parseInt(mPct[1], 10);
        if (v >= 0 && v <= 100) pctCands.push({ val: v, ...bbox, hasPct: /%/.test(raw) });
      }
      for (const key of HS_ORDER) {
        if (HS_LABELS[key].some((t) => norm.includes(t))) labelCands.push({ key, ...bbox });
      }
    }
    if (!pctCands.length) {
      for (const w of words) {
        const m = (w.text || "").trim().match(/^(\d{1,3})$/);
        if (m) {
          const v = +m[1];
          if (v >= 0 && v <= 100) pctCands.push({ val: v, ...cxy(w.bbox), hasPct: false });
        }
      }
    }
    const assigned = assignPctColumns(pctCands, labelCands);
    if (assigned) return assigned;
    return { error: "Não consegui ler os percentuais no print.", raw: (res.data && res.data.text) || "" };
  }

  function fmtBRL(v) {
    return "R$" + Number(v).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function tplSlide(n) {
    const s = document.createElement("div");
    s.className = "slide tpl";
    s.dataset.page = String(n);
    let src;
    if (isFunPage(n)) {
      src = funSrc(n);
    } else {
      const key = String(n).padStart(2, "0");
      src = CLEANBG[n]
        || (window.MODELO_PAGES && window.MODELO_PAGES[key])
        || ("assets/modelo-p" + key + ".png");
    }
    s.style.backgroundImage = 'url("' + src + '")';
    return s;
  }

  function addPatch(slide, xPt, yPt, wPt, hPt, color) {
    const p = document.createElement("div");
    p.className = "patch";
    p.style.left = xPt * S + "px";
    p.style.top = yPt * S + "px";
    p.style.width = wPt * S + "px";
    p.style.height = hPt * S + "px";
    p.style.background = color;
    slide.appendChild(p);
    return p;
  }

  // opts: {xPt, baselinePt, sizePt, weight, color, text|html, align:'left'|'right', spacing}
  function addOv(slide, opts) {
    const d = document.createElement("div");
    d.className = "ov";
    const fs = opts.sizePt * S;
    d.style.fontSize = fs + "px";
    d.style.fontWeight = opts.weight || 400;
    d.style.color = opts.color || "#000";
    d.style.top = opts.baselinePt * S - fs * ASC + "px";
    if (opts.spacing) d.style.letterSpacing = opts.spacing;
    if (opts.align === "right") {
      d.style.right = 1280 - opts.xPt * S + "px";
      d.style.textAlign = "right";
    } else {
      d.style.left = opts.xPt * S + "px";
    }
    if (opts.html !== undefined) d.innerHTML = opts.html;
    else d.textContent = opts.text;
    slide.appendChild(d);
    return d;
  }

  // Anéis da pág 5 — medidos no modelo DROGAFONTE (1280×720)
  const RING = {
    cols: [157, 292, 428, 563],
    cy: 203,
    r: 46,
    dot: 13.5,   // diâmetro em pt (modelo ≈ 13,8 pt; antes era 7 pt)
    n: 16,
    pctSize: 21,
  };

  function drawRingSvg(svg, cxPt, cyPt, pct) {
    const cx = cxPt * S, cy = RING.cy * S, r = RING.r * S;
    const dotR = (RING.dot * S) / 2;
    const bgR = r + dotR + 2 * S;
    const p = Math.max(0, Math.min(100, Number(pct) || 0));
    const filled = p >= 100 ? RING.n : Math.max(0, Math.min(RING.n - 1, Math.round((p / 100) * RING.n)));

    const bg = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    bg.setAttribute("cx", cx); bg.setAttribute("cy", cy); bg.setAttribute("r", bgR);
    bg.setAttribute("fill", COR.ringBg);
    svg.appendChild(bg);

    for (let i = 0; i < RING.n; i++) {
      const ang = (-90 + i * (360 / RING.n)) * Math.PI / 180;
      const dot = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      dot.setAttribute("cx", cx + r * Math.cos(ang));
      dot.setAttribute("cy", cy + r * Math.sin(ang));
      dot.setAttribute("r", dotR);
      dot.setAttribute("fill", i < filled ? COR.ringOn : COR.ringOff);
      svg.appendChild(dot);
    }

    const t = document.createElementNS("http://www.w3.org/2000/svg", "text");
    t.setAttribute("x", cx);
    t.setAttribute("y", cy);
    t.setAttribute("text-anchor", "middle");
    t.setAttribute("dominant-baseline", "central");
    t.setAttribute("fill", COR.ringOn);
    t.setAttribute("font-size", RING.pctSize * S);
    t.setAttribute("font-weight", "700");
    t.setAttribute("font-family", "Poppins, Segoe UI, Arial, sans-serif");
    t.textContent = Math.round(p) + "%";
    svg.appendChild(t);
  }

  function drawRingsPage5(slide, pct) {
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 1280 720");
    svg.setAttribute("width", "1280");
    svg.setAttribute("height", "720");
    svg.style.cssText = "position:absolute;left:0;top:0;pointer-events:none;overflow:visible";
    const vals = [pct.encontrar, pct.cadastrar, pct.disputar, pct.monitorar];
    for (let i = 0; i < 4; i++) drawRingSvg(svg, RING.cols[i], RING.cy, vals[i]);
    slide.appendChild(svg);
  }

  // Aplica os overlays (patches + textos + anéis) sobre cada página do modelo
  function applyOverlays(n, s, cfg) {
    if (isFunPage(n)) return; // slides de diversão são só imagem
    const m = cfg.metrics;
    const roi = cfg.roi;
    if (n === 2) {
      addOv(s, { xPt: 60, baselinePt: 178.7, sizePt: 43, weight: 700, color: COR.navy, text: fmtBRL(m.valorArrematado) });
      if (cfg.periodo) {
        addOv(s, { xPt: 61.4, baselinePt: 327.9, sizePt: 9, weight: 400, color: COR.gray, text: cfg.periodo });
      }
      addOv(s, { xPt: 424.7, baselinePt: 371.3, sizePt: 16, weight: 700, color: COR.blue, text: numFmt.format(m.disputasRealizadas) + " Disputas Realizadas via Robô" });
    } else if (n === 3) {
      addOv(s, { xPt: 60, baselinePt: 269.8, sizePt: 40, weight: 700, color: COR.navy, text: fmtBRL(m.retorno) });
      addOv(s, { xPt: 20, baselinePt: 351.5, sizePt: 16, weight: 700, color: COR.blue, text: numFmt.format(m.pregoesAcimaMinimo) + " pregões vencidos acima do valor limite, via" });
      addOv(s, { xPt: 20, baselinePt: 370.7, sizePt: 16, weight: 700, color: COR.blue, text: "estratégias do robô." });
    } else if (n === 4 && roi) {
      // Título estático fica no fundo do modelo — só sobrepomos linhas com valores do cliente
      addOv(s, { xPt: 166, baselinePt: 185.3, sizePt: 16, weight: 400, color: COR.ink, html: roi.investimentoLabel + ": <u>" + fmtBRL(roi.investimento) + "</u>" });
      addOv(s, { xPt: 166, baselinePt: 256.1, sizePt: 16, weight: 400, color: COR.ink, html: "Valor Total Arrematado em 1º Lugar: <u>" + fmtBRL(m.valorArrematado) + "</u>" });
      const di = '<span style="color:' + COR.blue + ';font-size:.68em">\u25C6</span>&nbsp;';
      addOv(s, { xPt: 166, baselinePt: 296.4, sizePt: 18, weight: 600, color: COR.ink, html: di + "ROI aproximado: <span style=\"color:" + COR.blue + '">' + roi.roiStr + "</span>" });
      addOv(s, { xPt: 166, baselinePt: 318, sizePt: 18, weight: 400, color: COR.ink, html: di + 'Para cada <b style="color:' + COR.blue + '">R$ 1,00 investido</b>, foi gerado cerca de <b style="color:' + COR.blue + '">R$</b>' });
      addOv(s, { xPt: 166, baselinePt: 339.6, sizePt: 18, weight: 400, color: COR.ink, html: '<b style="color:' + COR.blue + '">' + roi.porRealNum + " em retorno bruto</b>." });
    } else if (n === 5 && cfg.pct) {
      drawRingsPage5(s, cfg.pct);
    } else if (n === 6 && cfg.nums) {
      const rows = [
        { v: cfg.nums.capturadas, base: 129.4, py: 103 },
        { v: cfg.nums.cadastradas, base: 195.6, py: 170 },
        { v: cfg.nums.disputadas, base: 262.8, py: 237 },
        { v: cfg.nums.monitoradas, base: 351.6, py: 325 },
      ];
      for (const r of rows) {
        if (!r.v) continue;
        addOv(s, { xPt: 491, baselinePt: r.base, sizePt: 30, weight: 700, color: COR.numLight, align: "right", text: r.v });
      }
    } else if (n === 15 && cfg.emailGerente) {
      addOv(s, { xPt: 46.5, baselinePt: 265, sizePt: 9, weight: 400, color: COR.black, text: "E-mail: " + cfg.emailGerente });
    }
  }

  // Rótulos das páginas para o seletor
  const PAGE_LABELS = {
    1: "Capa", 2: "Valor arrematado", 3: "Economia / margem", 4: "ROI",
    5: "Health Score", 6: "Como a Effecti ajudou", 7: "Últimas implantações",
    8: "Disputar Web", 9: "Monitoramento de diligências", 11: "Funcionalidades adicionais",
    12: "Serviços adicionais", 13: "Cursos", 14: "Pesquisa de satisfação",
    15: "Contatos", 16: "Encerramento",
  };

  // Slides de encerramento (versão diversão) — ids >= 101 para não colidir com o modelo.
  // A ordem aqui é a ordem em que entram no fim da apresentação.
  const FUN_SLIDES = [
    { id: 101, key: "f1", label: "FIM — Está tudo acabado" },
    { id: 102, key: "f2", label: "Absolute Cinema" },
    { id: 103, key: "f3", label: "Obrigado — evento canônico" },
    { id: 104, key: "f4", label: "Nosso trabalho acabou" },
    { id: 105, key: "f5", label: "A culpa é do presidente" },
    { id: 106, key: "f6", label: "Não aceitamos difamações" },
    { id: 107, key: "f7", label: "Dúvidas, manda no Pix" },
  ];
  const FUN_BY_ID = {};
  FUN_SLIDES.forEach((s) => { FUN_BY_ID[s.id] = s; PAGE_LABELS[s.id] = s.label; });
  function isFunPage(n) { return n >= 101; }
  function funSrc(n) {
    const s = FUN_BY_ID[n];
    return (s && window.FUN_PAGES && window.FUN_PAGES[s.key]) || "";
  }

  function pageLabel(n) { return PAGE_LABELS[n] || ("Página " + n); }

  // Páginas que PODEM entrar (têm dados/são renderizáveis). A seleção fina fica
  // por conta do operador; algumas vêm desmarcadas por padrão (ver defaults).
  function computeCandidatePages(cfg) {
    const usaRobo = cfg.usaRobo !== false;
    const base = cfg.institucional
      ? [1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16]
      : [2, 3, 4, 5, 6];
    const pages = base.filter((n) => {
      if (!usaRobo && (n === 2 || n === 3 || n === 4)) return false;
      // Sem mensalidade não há como renderizar a página de ROI
      if (n === 4 && !cfg.roi) return false;
      return true;
    });
    // Slides de encerramento (diversão) entram no fim, se marcado
    if (cfg.fun && window.FUN_PAGES) {
      FUN_SLIDES.forEach((s) => { if (window.FUN_PAGES[s.key]) pages.push(s.id); });
    }
    return pages;
  }

  // Aviso de resultado baixo (a página continua marcada; o operador decide remover)
  function pageWarnReason(n, cfg) {
    if (n === 4 && cfg.roi && cfg.roi.roiPct < ROI_MIN_PCT) {
      return `ROI ${cfg.roi.roiStr} abaixo de ${ROI_MIN_PCT}% — considere remover`;
    }
    if (n === 3 && cfg.metrics && Number(cfg.metrics.retorno) === 0) {
      return "Sem economia (Diferença R$ 0,00) — considere remover";
    }
    return null;
  }

  // Todas as páginas do relatório entram marcadas; o operador remove pelo checkbox
  function isPageDefaultSelected(_n, _cfg) {
    return true;
  }

  function getSelectedPages() {
    const cands = state.candidatePages || [];
    return cands.filter((n) => state.selectedPages && state.selectedPages[n]);
  }

  function updateSelectionInfo() {
    const info = $("selInfo");
    const total = (state.candidatePages || []).length;
    const sel = getSelectedPages().length;
    if (info) info.textContent = `${sel} de ${total} páginas selecionadas`;
    const noSel = sel === 0;
    if ($("btnPdf")) $("btnPdf").disabled = noSel;
    if ($("btnPptx")) $("btnPptx").disabled = noSel;
  }

  function setAllPages(on) {
    if (!state.candidatePages) return;
    for (const n of state.candidatePages) state.selectedPages[n] = on;
    document.querySelectorAll("#deck .slide-item").forEach((item) => {
      const n = parseInt(item.dataset.page, 10);
      const cb = item.querySelector("input[type=checkbox]");
      if (cb) cb.checked = !!state.selectedPages[n];
      item.classList.toggle("off", !state.selectedPages[n]);
    });
    updateSelectionInfo();
  }

  function getPageBgSrc(n) {
    if (isFunPage(n)) return funSrc(n);
    return CLEANBG[n]
      || (window.MODELO_PAGES && window.MODELO_PAGES[String(n).padStart(2, "0")])
      || ("assets/modelo-p" + String(n).padStart(2, "0") + ".png");
  }

  function setCanvasFont(ctx, sizePt, weight) {
    ctx.font = `${weight || 400} ${sizePt * S}px Poppins, "Segoe UI", Arial, sans-serif`;
  }

  function drawCanvasText(ctx, xPt, baselinePt, text, opts) {
    setCanvasFont(ctx, opts.sizePt, opts.weight);
    ctx.fillStyle = opts.color || COR.ink;
    ctx.textBaseline = "alphabetic";
    ctx.textAlign = opts.align === "right" ? "right" : "left";
    const x = xPt * S;
    const y = baselinePt * S;
    ctx.fillText(text, x, y);
    return ctx.measureText(text).width;
  }

  function drawCanvasUnderline(ctx, x, y, w, color) {
    ctx.strokeStyle = color || COR.ink;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(x, y + 2);
    ctx.lineTo(x + w, y + 2);
    ctx.stroke();
  }

  function drawCanvasLabelValue(ctx, xPt, baselinePt, label, value, opts) {
    setCanvasFont(ctx, opts.sizePt, opts.weight);
    ctx.fillStyle = opts.color || COR.ink;
    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
    let x = xPt * S;
    const y = baselinePt * S;
    ctx.fillText(label, x, y);
    x += ctx.measureText(label).width;
    ctx.fillText(value, x, y);
    drawCanvasUnderline(ctx, x, y, ctx.measureText(value).width, opts.color || COR.ink);
  }

  function drawCanvasDiamondLine(ctx, xPt, baselinePt, parts, sizePt, weight) {
    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
    let x = xPt * S;
    const y = baselinePt * S;
    for (const p of parts) {
      setCanvasFont(ctx, p.sizePt || sizePt, p.weight || weight || 400);
      ctx.fillStyle = p.color || COR.ink;
      ctx.fillText(p.text, x, y);
      x += ctx.measureText(p.text).width;
    }
  }

  function drawRingCanvas(ctx, cxPt, pct) {
    const cx = cxPt * S;
    const cy = RING.cy * S;
    const r = RING.r * S;
    const dotR = (RING.dot * S) / 2;
    const bgR = r + dotR + 2 * S;
    const p = Math.max(0, Math.min(100, Number(pct) || 0));
    const filled = p >= 100 ? RING.n : Math.max(0, Math.min(RING.n - 1, Math.round((p / 100) * RING.n)));

    ctx.beginPath();
    ctx.arc(cx, cy, bgR, 0, Math.PI * 2);
    ctx.fillStyle = COR.ringBg;
    ctx.fill();

    for (let i = 0; i < RING.n; i++) {
      const ang = (-90 + i * (360 / RING.n)) * Math.PI / 180;
      ctx.beginPath();
      ctx.arc(cx + r * Math.cos(ang), cy + r * Math.sin(ang), dotR, 0, Math.PI * 2);
      ctx.fillStyle = i < filled ? COR.ringOn : COR.ringOff;
      ctx.fill();
    }

    setCanvasFont(ctx, RING.pctSize, 700);
    ctx.fillStyle = COR.ringOn;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(Math.round(p) + "%", cx, cy);
    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
  }

  function applyCanvasOverlays(ctx, n, cfg) {
    if (isFunPage(n)) return; // slides de diversão são só imagem
    const m = cfg.metrics;
    const roi = cfg.roi;
    if (n === 2) {
      drawCanvasText(ctx, 60, 178.7, fmtBRL(m.valorArrematado), { sizePt: 43, weight: 700, color: COR.navy });
      if (cfg.periodo) {
        drawCanvasText(ctx, 61.4, 327.9, cfg.periodo, { sizePt: 9, weight: 400, color: COR.gray });
      }
      drawCanvasText(ctx, 424.7, 371.3, numFmt.format(m.disputasRealizadas) + " Disputas Realizadas via Robô", { sizePt: 16, weight: 700, color: COR.blue });
    } else if (n === 3) {
      drawCanvasText(ctx, 60, 269.8, fmtBRL(m.retorno), { sizePt: 40, weight: 700, color: COR.navy });
      drawCanvasText(ctx, 20, 351.5, numFmt.format(m.pregoesAcimaMinimo) + " pregões vencidos acima do valor limite, via", { sizePt: 16, weight: 700, color: COR.blue });
      drawCanvasText(ctx, 20, 370.7, "estratégias do robô.", { sizePt: 16, weight: 700, color: COR.blue });
    } else if (n === 4 && roi) {
      drawCanvasLabelValue(ctx, 166, 185.3, roi.investimentoLabel + ": ", fmtBRL(roi.investimento), { sizePt: 16, weight: 400, color: COR.ink });
      drawCanvasLabelValue(ctx, 166, 256.1, "Valor Total Arrematado em 1º Lugar: ", fmtBRL(m.valorArrematado), { sizePt: 16, weight: 400, color: COR.ink });
      drawCanvasDiamondLine(ctx, 166, 296.4, [
        { text: "\u25C6 ", sizePt: 18 * 0.68, color: COR.blue },
        { text: "ROI aproximado: ", sizePt: 18, weight: 600, color: COR.ink },
        { text: roi.roiStr, sizePt: 18, weight: 600, color: COR.blue },
      ], 18, 600);
      drawCanvasDiamondLine(ctx, 166, 318, [
        { text: "\u25C6 ", sizePt: 18 * 0.68, color: COR.blue },
        { text: "Para cada ", sizePt: 18, weight: 400, color: COR.ink },
        { text: "R$ 1,00 investido", sizePt: 18, weight: 700, color: COR.blue },
        { text: ", foi gerado cerca de ", sizePt: 18, weight: 400, color: COR.ink },
        { text: "R$", sizePt: 18, weight: 700, color: COR.blue },
      ], 18, 400);
      drawCanvasDiamondLine(ctx, 166, 339.6, [
        { text: roi.porRealNum + " em retorno bruto.", sizePt: 18, weight: 700, color: COR.blue },
      ], 18, 700);
    } else if (n === 5 && cfg.pct) {
      const vals = [cfg.pct.encontrar, cfg.pct.cadastrar, cfg.pct.disputar, cfg.pct.monitorar];
      for (let i = 0; i < 4; i++) drawRingCanvas(ctx, RING.cols[i], vals[i]);
    } else if (n === 6 && cfg.nums) {
      const rows = [
        { v: cfg.nums.capturadas, base: 129.4 },
        { v: cfg.nums.cadastradas, base: 195.6 },
        { v: cfg.nums.disputadas, base: 262.8 },
        { v: cfg.nums.monitoradas, base: 351.6 },
      ];
      for (const r of rows) {
        if (!r.v) continue;
        drawCanvasText(ctx, 491, r.base, r.v, { sizePt: 30, weight: 700, color: COR.numLight, align: "right" });
      }
    } else if (n === 15 && cfg.emailGerente) {
      drawCanvasText(ctx, 46.5, 265, "E-mail: " + cfg.emailGerente, { sizePt: 9, weight: 400, color: COR.black });
    }
  }

  function renderPageToCanvas(n, cfg, bgImg) {
    const canvas = document.createElement("canvas");
    canvas.width = 1280;
    canvas.height = 720;
    const ctx = canvas.getContext("2d");
    if (bgImg) ctx.drawImage(bgImg, 0, 0, 1280, 720);
    applyCanvasOverlays(ctx, n, cfg);
    return canvas;
  }

  function buildDeck(cfg) {
    const deck = $("deck");
    deck.innerHTML = "";
    const slides = [];
    const pages = computeCandidatePages(cfg);
    state.candidatePages = pages;
    state.selectedPages = {};

    // Barra de seleção (contador + marcar/desmarcar tudo)
    const bar = document.createElement("div");
    bar.className = "sel-bar";
    bar.innerHTML =
      '<span id="selInfo" class="sel-info"></span>' +
      '<span class="sel-actions">' +
      '<button type="button" class="btn ghost small" id="btnSelAll">Marcar todas</button>' +
      '<button type="button" class="btn ghost small" id="btnSelNone">Desmarcar todas</button>' +
      "</span>";
    deck.appendChild(bar);

    const scale = computePreviewScale();
    for (const n of pages) {
      const s = tplSlide(n);
      applyOverlays(n, s, cfg);
      slides.push(s);

      const on = isPageDefaultSelected(n, cfg);
      state.selectedPages[n] = on;

      const item = document.createElement("div");
      item.className = "slide-item" + (on ? "" : " off");
      item.dataset.page = n;

      const head = document.createElement("label");
      head.className = "slide-check";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = on;
      cb.addEventListener("change", () => {
        state.selectedPages[n] = cb.checked;
        item.classList.toggle("off", !cb.checked);
        updateSelectionInfo();
      });
      const txt = document.createElement("span");
      txt.className = "slide-name";
      txt.textContent = `Pág. ${n} — ${pageLabel(n)}`;
      head.appendChild(cb);
      head.appendChild(txt);
      const reason = pageWarnReason(n, cfg);
      if (reason) {
        const why = document.createElement("em");
        why.className = "slide-why";
        why.textContent = reason;
        head.appendChild(why);
      }

      const frame = document.createElement("div");
      frame.className = "slide-frame";
      frame.style.width = 1280 * scale + "px";
      frame.style.height = 720 * scale + "px";
      frame.style.overflow = "hidden";
      s.style.transform = `scale(${scale})`;
      s.style.transformOrigin = "top left";
      frame.appendChild(s);

      item.appendChild(head);
      item.appendChild(frame);
      deck.appendChild(item);
    }

    $("btnSelAll").addEventListener("click", () => setAllPages(true));
    $("btnSelNone").addEventListener("click", () => setAllPages(false));
    updateSelectionInfo();

    return slides;
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  }

  function computePreviewScale() {
    const col = document.querySelector(".preview");
    const w = (col ? col.clientWidth : 900) - 8;
    return Math.min(1, Math.max(0.3, w / 1280));
  }

  function setExportProgress(cur, total, title) {
    const overlay = $("exportOverlay");
    const status = $("exportStatus");
    const fill = $("exportBarFill");
    const hint = $("exportHint");
    const ttl = $("exportTitle");
    if (ttl && title) ttl.textContent = title;
    if (status) status.textContent = `Página ${cur} de ${total}`;
    if (fill) fill.style.width = Math.round((cur / total) * 100) + "%";
    if (hint) {
      hint.textContent = document.hidden
        ? "A exportação continua em segundo plano — você pode usar outros programas."
        : "Pode minimizar ou trocar de janela; a exportação continua.";
    }
    if (overlay) overlay.hidden = false;
  }

  function hideExportProgress() {
    const overlay = $("exportOverlay");
    if (overlay) overlay.hidden = true;
    const fill = $("exportBarFill");
    if (fill) fill.style.width = "0%";
  }

  function pause(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  // Exportação via canvas (não usa html2canvas — continua com aba minimizada)
  async function loadExportBackgrounds(pages) {
    await ensureCleanBackgrounds();
    if (document.fonts && document.fonts.ready) {
      try { await document.fonts.ready; } catch (e) { /* ignore */ }
    }
    const bgs = {};
    for (const n of pages) {
      try {
        bgs[n] = await loadImg(getPageBgSrc(n));
      } catch (e) {
        throw new Error("Falha ao carregar página " + n + ".");
      }
      await pause(0);
    }
    return bgs;
  }

  async function exportPdf(plan, filename) {
    const pages = plan.pages;
    const cfg = plan.cfg;
    const bgs = await loadExportBackgrounds(pages);

    try {
      const pdf = new jsPDF({ orientation: "landscape", unit: "px", format: [1280, 720], compress: true });
      for (let i = 0; i < pages.length; i++) {
        setExportProgress(i + 1, pages.length, "Gerando PDF…");
        const canvas = renderPageToCanvas(pages[i], cfg, bgs[pages[i]]);
        const img = canvas.toDataURL("image/png");
        if (i > 0) pdf.addPage([1280, 720], "landscape");
        pdf.addImage(img, "PNG", 0, 0, 1280, 720, undefined, "FAST");
        await pause(0);
      }
      pdf.save(filename);
    } finally {
      hideExportProgress();
    }
  }

  async function exportPptx(plan, filename) {
    const PptxGenJS = window.PptxGenJS;
    if (!PptxGenJS) throw new Error("Biblioteca PowerPoint não carregou.");
    const pages = plan.pages;
    const cfg = plan.cfg;
    const bgs = await loadExportBackgrounds(pages);

    try {
      const pptx = new PptxGenJS();
      pptx.defineLayout({ name: "QBR_16x9", width: 13.333, height: 7.5 });
      pptx.layout = "QBR_16x9";
      pptx.author = "Effecti";
      pptx.title = filename.replace(/\.pptx$/i, "");
      for (let i = 0; i < pages.length; i++) {
        setExportProgress(i + 1, pages.length, "Gerando PowerPoint…");
        const canvas = renderPageToCanvas(pages[i], cfg, bgs[pages[i]]);
        const img = canvas.toDataURL("image/jpeg", 0.92);
        const slide = pptx.addSlide();
        slide.addImage({ data: img, x: 0, y: 0, w: 13.333, h: 7.5 });
        await pause(0);
      }
      await pptx.writeFile({ fileName: filename });
    } finally {
      hideExportProgress();
    }
  }

  function slugify(s) {
    return normalize(s).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "cliente";
  }

  // Remove caracteres inválidos p/ nome de arquivo, mantendo acentos e espaços
  function sanitizeFileName(s) {
    return String(s || "")
      .replace(/[\\/:*?"<>|]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  // Monta "Resultados Mar a Jul - EMPRESA.pdf" (ou .pptx)
  function buildDownloadFilename(ext) {
    const empresa = sanitizeFileName($("cliente").value) || "Cliente";
    const periodo = sanitizeFileName(periodShort());
    const partes = ["Resultados"];
    if (periodo) partes.push(periodo);
    partes.push("-", empresa);
    return partes.join(" ").replace(/\s+/g, " ").trim() + "." + ext;
  }

  // ---------- Fluxo principal ----------
  async function gerar() {
    try {
      let usaRobo = false;
      let metrics = null;

      if (state.disputeRows && state.disputeRows.length) {
        const { rows: rowsFiltradas, filtered, total } = filterByPeriod(state.disputeRows);
        if (rowsFiltradas.length) {
          usaRobo = true;
          metrics = computeMetrics(rowsFiltradas);
          state.metrics = metrics;
          if (filtered && rowsFiltradas.length < total) {
            toast(`Filtro por período: ${rowsFiltradas.length} de ${total} disputas consideradas.`);
          }
        } else {
          state.metrics = null;
          toast("Nenhuma disputa no período — páginas de arremate, economia e ROI não serão incluídas.");
        }
      } else {
        state.metrics = null;
        toast("Sem dados de disputa — páginas de arremate, economia e ROI omitidas (cliente sem robô).");
      }

      // mensalidade: se vazia e tiver CS, tenta buscar
      let mensalidade = toNumber($("mensalidade").value);
      if ((!mensalidade || mensalidade <= 0) && state.csRows) {
        const found = lookupMensalidade($("cliente").value);
        if (found) { mensalidade = found; $("mensalidade").value = numFmt.format(found); }
      }

      const roiMeses = getRoiMeses();
      const anoRoi = ($("periodoFim").value || $("periodoInicio").value || "").split("-")[0] || String(new Date().getFullYear());
      const roi = usaRobo && metrics ? computeRoi(metrics.valorArrematado, mensalidade, roiMeses, anoRoi) : null;

      const pctVals = {
        encontrar: numOrNull($("pctEncontrar").value),
        cadastrar: numOrNull($("pctCadastrar").value),
        disputar: numOrNull($("pctDisputar").value),
        monitorar: numOrNull($("pctMonitorar").value),
      };
      const temPct = Object.values(pctVals).some((v) => v !== null);
      const pct = temPct
        ? {
            encontrar: pctVals.encontrar || 0,
            cadastrar: pctVals.cadastrar || 0,
            disputar: pctVals.disputar || 0,
            monitorar: pctVals.monitorar || 0,
          }
        : null;

      const fmtInt = (s) => {
        const v = String(s || "").trim();
        if (!v) return "";
        const n = toNumber(v);
        return n ? numFmt.format(Math.round(n)) : v;
      };
      const nums = {
        capturadas: fmtInt($("numCapturadas").value),
        cadastradas: fmtInt($("numCadastradas").value),
        disputadas: fmtInt($("numDisputadas").value),
        monitoradas: fmtInt($("numMonitoradas").value),
      };

      const cfg = {
        cliente: $("cliente").value.trim(),
        periodo: periodLabel(),
        metrics: metrics || {},
        usaRobo,
        roi,
        pct,
        nums,
        emailGerente: $("emailGerente").value.trim(),
        healthImg: state.healthImg,
        institucional: $("incluirInstitucional").checked,
        fun: $("incluirFun").checked,
      };

      renderMetricsPanel({ metrics, usaRobo, roi, mensalidade, roiMeses });

      await ensureCleanBackgrounds();
      const slides = buildDeck(cfg);
      state.slides = slides;
      state.exportCfg = cfg;

      $("emptyState").hidden = true;
      updateSelectionInfo();
      toast(`Prévia gerada. ${getSelectedPages().length} de ${slides.length} páginas selecionadas.`);
    } catch (e) {
      console.error(e);
      toast("Erro ao gerar: " + e.message, true);
    }
  }

  function renderMetricsPanel({ metrics: m, usaRobo, roi, mensalidade, roiMeses }) {
    const box = $("metrics");
    const list = $("metricsList");
    if (!usaRobo) {
      list.innerHTML = `<li class="warn"><span>Robô de disputas</span><b>não utilizado — págs. 2, 3 e 4 omitidas</b></li>`;
      box.hidden = false;
      return;
    }
    const meses = roi ? roi.meses : roiMeses;
    const invLine = mensalidade
      ? `<li><span>${roiInvestimentoResumo(meses)}</span><b>${brl.format(roi ? roi.investimento : mensalidade * meses)}</b></li>`
      : "";
    const roiLine = roi
      ? `<li class="${roi.roiPct < ROI_MIN_PCT ? "warn" : ""}"><span>ROI aproximado</span><b>${roi.roiStr}${roi.roiPct < ROI_MIN_PCT ? " (considere remover)" : ""}</b></li>
         <li><span>Retorno por R$1 investido</span><b>${roi.porRealStr}</b></li>`
      : `<li class="warn"><span>ROI</span><b>informe a mensalidade</b></li>`;
    list.innerHTML = `
      <li><span>Valor arrematado</span><b>${brl.format(m.valorArrematado)}</b></li>
      <li><span>Retorno financeiro (margem)</span><b>${brl.format(m.retorno)}</b></li>
      <li><span>Disputas realizadas</span><b>${numFmt.format(m.disputasRealizadas)}</b></li>
      <li><span>Pregões acima do mínimo</span><b>${numFmt.format(m.pregoesAcimaMinimo)}</b></li>
      <li><span>Licitações vencidas</span><b>${numFmt.format(m.licitacoesVencidas)}</b></li>
      <li><span>Itens vencidos</span><b>${numFmt.format(m.itensVencidosTotal)}</b></li>
      ${invLine}
      ${roiLine}`;
    box.hidden = false;
  }

  function lookupMensalidade(nome) {
    if (!state.csRows || !state.csRows.length) return null;
    const alvo = normalize(nome);
    if (!alvo) return null;
    const keys = Object.keys(state.csRows[0]);
    let best = null;
    for (const r of state.csRows) {
      const nm = normalize(pick(r, keys, ["Nome", "Cliente", "Razão Social"]));
      if (!nm) continue;
      if (nm === alvo || nm.includes(alvo) || alvo.includes(nm)) {
        const men = toNumber(pick(r, keys, ["R$ Mensalidade", "Mensalidade", "Valor Mensalidade"]));
        if (men > 0) { best = men; break; }
      }
    }
    return best;
  }

  // ---------- Eventos ----------
  function bind() {
    hideExportProgress();
    $("csvDisputa").addEventListener("change", async (e) => {
      const f = e.target.files[0];
      if (!f) return;
      try {
        state.disputeRows = await readCsv(f);
        $("csvDisputaInfo").textContent = `✓ ${f.name} — ${state.disputeRows.length} linhas`;
      } catch (err) {
        toast("Falha ao ler CSV de disputas.", true);
      }
    });

    $("btnProcessarColado").addEventListener("click", () => {
      const text = $("dadosColados").value;
      if (!text.trim()) { toast("Cole os dados na caixa de texto primeiro.", true); return; }
      try {
        const rows = parsePasted(text);
        if (!rows.length) throw new Error("nenhuma disputa reconhecida");
        state.disputeRows = rows;
        $("csvDisputa").value = "";
        $("csvDisputaInfo").textContent = "";
        const comData = rows.filter((r) => parseBrDate(r["Data"])).length;
        $("coladoInfo").textContent = `✓ ${rows.length} disputas reconhecidas (${comData} com data)`;
        toast(`${rows.length} disputas processadas dos dados colados.`);
      } catch (err) {
        toast("Falha ao processar: " + err.message, true);
      }
    });

    $("csvCs").addEventListener("change", async (e) => {
      const f = e.target.files[0];
      if (!f) return;
      try {
        state.csRows = await readCsv(f);
        $("csvCsInfo").textContent = `✓ ${f.name} — ${state.csRows.length} clientes`;
      } catch (err) {
        toast("Falha ao ler CSV de Customer Success.", true);
      }
    });

    // ---- Health Score: colar (Ctrl+V) / arrastar / clicar ----
    function loadHealthFile(f) {
      if (!f || f.type.indexOf("image") !== 0) { toast("Cole ou selecione uma imagem.", true); return; }
      const reader = new FileReader();
      reader.onload = async () => {
        state.healthImg = reader.result;
        $("healthPreview").innerHTML = `<img src="${state.healthImg}" alt="preview" />`;
        const z = $("healthPaste");
        z.classList.add("has-img");
        $("pasteMsg").innerHTML = "Lendo os percentuais do print…";
        try {
          const r = await extractHealthPercents(state.healthImg);
          if (r && r.values) {
            const map = { encontrar: "pctEncontrar", cadastrar: "pctCadastrar", disputar: "pctDisputar", monitorar: "pctMonitorar" };
            Object.keys(r.values).forEach((k) => { $(map[k]).value = r.values[k]; });
            $("pasteMsg").innerHTML = `✓ Print lido — ${r.found}/4 percentuais preenchidos automaticamente`;
            toast(`Health Score: ${r.found} de 4 percentuais lidos do print.`);
          } else {
            $("pasteMsg").innerHTML = "✓ Print carregado — preencha os % manualmente (não consegui ler)";
            if (r && r.error) toast(r.error, true);
          }
        } catch (err) {
          $("pasteMsg").innerHTML = "✓ Print carregado — preencha os % manualmente";
          toast("Falha no OCR do print: " + err.message, true);
        }
      };
      reader.readAsDataURL(f);
    }
    function handlePaste(e) {
      const items = (e.clipboardData || window.clipboardData).items || [];
      for (const it of items) {
        if (it.type && it.type.indexOf("image") === 0) {
          loadHealthFile(it.getAsFile());
          e.preventDefault();
          return true;
        }
      }
      return false;
    }
    const zone = $("healthPaste");
    zone.addEventListener("click", () => $("healthPrint").click());
    zone.addEventListener("paste", handlePaste);
    zone.addEventListener("dragover", (e) => { e.preventDefault(); zone.classList.add("has-img"); });
    zone.addEventListener("dragleave", () => { if (!state.healthImg) zone.classList.remove("has-img"); });
    zone.addEventListener("drop", (e) => {
      e.preventDefault();
      const f = e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) loadHealthFile(f);
    });
    document.addEventListener("paste", (e) => {
      if (document.activeElement === zone || zone.contains(document.activeElement)) handlePaste(e);
    });
    $("healthPrint").addEventListener("change", (e) => {
      const f = e.target.files[0];
      if (f) loadHealthFile(f);
    });

    $("btnBuscarMensalidade").addEventListener("click", () => {
      if (!state.csRows) { toast("Carregue o CSV de Customer Success primeiro.", true); return; }
      const found = lookupMensalidade($("cliente").value);
      if (found) { $("mensalidade").value = numFmt.format(found); toast("Mensalidade preenchida: " + brl.format(found)); }
      else toast("Cliente não encontrado no CSV de CS.", true);
    });

    $("btnGerar").addEventListener("click", gerar);

    async function runDownload(kind) {
      if (!state.exportCfg) { toast("Gere a prévia primeiro.", true); return; }
      const pages = getSelectedPages();
      if (!pages.length) { toast("Selecione ao menos uma página para baixar.", true); return; }
      const btnPdf = $("btnPdf");
      const btnPptx = $("btnPptx");
      const isPdf = kind === "pdf";
      const btn = isPdf ? btnPdf : btnPptx;
      const old = btn.textContent;
      btnPdf.disabled = true;
      btnPptx.disabled = true;
      btn.textContent = "Gerando…";
      try {
        const plan = { pages, cfg: state.exportCfg };
        if (isPdf) {
          await exportPdf(plan, buildDownloadFilename("pdf"));
          toast("PDF gerado com sucesso.");
        } else {
          await exportPptx(plan, buildDownloadFilename("pptx"));
          toast("PowerPoint gerado com sucesso.");
        }
      } catch (e) {
        console.error(e);
        toast("Erro ao gerar " + (isPdf ? "PDF" : "PowerPoint") + ": " + e.message, true);
      } finally {
        btn.textContent = old;
        updateSelectionInfo();
      }
    }
    $("btnPdf").addEventListener("click", () => runDownload("pdf"));
    $("btnPptx").addEventListener("click", () => runDownload("pptx"));

    $("btnLimpar").addEventListener("click", () => {
      ["cliente", "mensalidade", "emailGerente",
       "pctEncontrar", "pctCadastrar", "pctDisputar", "pctMonitorar",
       "numCapturadas", "numCadastradas", "numDisputadas", "numMonitoradas"].forEach((id) => ($(id).value = ""));
      ["csvDisputa", "csvCs", "healthPrint", "periodoInicio", "periodoFim", "dadosColados"].forEach((id) => ($(id).value = ""));
      $("roiMeses").value = "3";
      $("incluirFun").checked = false;
      state.disputeRows = state.csRows = state.healthImg = state.slides = state.metrics = null;
      state.exportCfg = null;
      state.candidatePages = null;
      state.selectedPages = {};
      $("csvDisputaInfo").textContent = "";
      $("csvCsInfo").textContent = "";
      $("coladoInfo").textContent = "";
      $("healthPreview").innerHTML = "";
      $("healthPaste").classList.remove("has-img");
      $("pasteMsg").innerHTML = "Clique aqui e cole (Ctrl+V) o print do Health Score,<br />ou arraste a imagem / clique para escolher";
      $("deck").innerHTML = "";
      $("metrics").hidden = true;
      $("emptyState").hidden = false;
      $("btnPdf").disabled = true;
      $("btnPdf").textContent = "Baixar PDF";
      $("btnPptx").disabled = true;
      $("btnPptx").textContent = "Baixar PowerPoint";
      hideExportProgress();
    });
  }

  document.addEventListener("DOMContentLoaded", bind);

  // Hook para inspeção/testes (não interfere no uso normal)
  window.QBR = {
    state, gerar, buildDeck, computeMetrics, computeRoi, toNumber, readCsv, parsePasted, filterByPeriod, ROI_MIN_PCT, ROI_MESES_OPTS, getRoiMeses,
    ensureCleanBackgrounds, extractHealthPercents, renderPageToCanvas, exportPdf, exportPptx,
    setDisputeRows: (rows) => { state.disputeRows = rows; },
  };
})();
