import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

const T = {
  page: "#f5f6f2",
  panel: "#ffffff",
  panel2: "#f9faf7",
  text: "#1f2933",
  soft: "#5f6b7a",
  muted: "#8a94a3",
  line: "#dfe4dc",
  line2: "#cdd6cc",
  ink: "#101828",
  teal: "#0f766e",
  tealBg: "#e3f4f1",
  blue: "#2563eb",
  blueBg: "#e8efff",
  amber: "#b45309",
  amberBg: "#fff4d6",
  red: "#b42318",
  redBg: "#ffe8e5",
  green: "#087443",
  greenBg: "#def8ea",
  purple: "#6d28d9",
  purpleBg: "#efe7ff",
};

const DEC = {
  buy: { label: "COMPRAR", bg: T.greenBg, fg: T.green, bd: "#a8e8c2" },
  caution: { label: "REVISAR", bg: T.amberBg, fg: T.amber, bd: "#f3d47b" },
  avoid: { label: "NO CONV.", bg: T.redBg, fg: T.red, bd: "#f6b8b0" },
};

const STOP_WORDS = new Set([
  "alimento", "balanceado", "bolsa", "bolson", "perros", "perro", "dog",
  "gatos", "gato", "cat", "para", "con", "sin", "por", "the", "de", "del",
  "y", "x", "kg", "kgs", "kilo", "kilos", "un", "una",
]);

const $n = n => (n != null && !isNaN(n) && n !== "")
  ? `$${Math.round(n).toLocaleString("es-AR")}`
  : "-";

const LOGISTICS_FIXED_UNDER = 33000;
const LOGISTICS_FIXED_COST = 3000;

function parseNum(raw) {
  if (typeof raw === "number") return Number.isFinite(raw) ? raw : 0;
  const s = String(raw ?? "").trim();
  if (!s) return 0;
  const lp = s.lastIndexOf(".");
  const lc = s.lastIndexOf(",");
  let clean;
  if (lp >= 0 && lc >= 0) {
    clean = lp > lc ? s.replace(/,/g, "") : s.replace(/\./g, "").replace(",", ".");
  } else if (lp >= 0) {
    clean = /^\d{1,3}(?:\.\d{3})+$/.test(s.replace(/[^\d.]/g, ""))
      ? s.replace(/\./g, "")
      : s;
  } else if (lc >= 0) {
    clean = /^\d{1,3}(?:,\d{3})+$/.test(s.replace(/[^\d,]/g, ""))
      ? s.replace(/,/g, "")
      : s.replace(",", ".");
  } else {
    clean = s;
  }
  return parseFloat(clean.replace(/[^\d.-]/g, "")) || 0;
}

function excelColName(index) {
  let n = index + 1;
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function findHeaderRow(rows) {
  const productRx = /descrip|nombre|product|item|presenta|articulo|artículo/i;
  const priceRx = /precio|costo|neto|importe|valor|meli/i;
  const limit = Math.min(rows.length, 40);
  for (let i = 0; i < limit; i += 1) {
    const labels = (rows[i] || []).map(v => String(v ?? "").trim()).filter(Boolean);
    if (labels.length < 2) continue;
    if (labels.some(v => productRx.test(v)) && labels.some(v => priceRx.test(v))) return i;
  }
  return 0;
}

function costIncludesVatFromLabel(label) {
  const s = String(label || "");
  if (/sin\s*iva|s\/\s*iva/i.test(s)) return 0;
  if (/c\/\s*iva|con\s*iva/i.test(s)) return 1;
  return null;
}

const norm = v => String(v ?? "")
  .toLowerCase()
  .normalize("NFD")
  .replace(/[\u0300-\u036f]/g, "")
  .replace(/[^a-z0-9.,\s]/g, " ")
  .replace(/\s+/g, " ")
  .trim();

function detectBrand(text) {
  const s = norm(text);
  if (/\broyal\s+canin\b/.test(s)) return "royal canin";
  if (/\bpro\s*plan\b|\bproplan\b/.test(s)) return "pro plan";
  if (/\bcat\s+chow\b|\bcatchow\b|\bdog\s+chow\b|\bdogchow\b|\bexcellent\b/.test(s)) return "purina";
  if (/\bpurina\b/.test(s)) return "purina";
  return "";
}

function extractWeightKg(text) {
  const s = norm(text);
  const packKg = s.match(/\b(\d{1,2})\s*(?:x|unidades?|bolsas?)\s*(?:de\s*)?(\d+(?:[.,]\d+)?)\s*k(?:g|ilo|ilos)?\b/);
  if (packKg) return +(parseNum(packKg[1]) * parseNum(packKg[2])).toFixed(3);
  const kg = s.match(/\b(\d+(?:[.,]\d+)?)\s*k(?:g|ilo|ilos)?\b/);
  if (kg) return +parseNum(kg[1]).toFixed(3);
  const grams = s.match(/\b(\d{3,5})\s*g(?:r|ramos)?\b/);
  if (grams) return +(parseNum(grams[1]) / 1000).toFixed(3);
  return null;
}

function productTokens(text) {
  return norm(text)
    .replace(/\bcat\s+chow\b/g, "catchow")
    .replace(/\bdog\s+chow\b/g, "dogchow")
    .replace(/\bpro\s+plan\b/g, "proplan")
    .split(" ")
    .map(t => t.replace(",", "."))
    .filter(t => t.length > 2 && !STOP_WORDS.has(t) && !/^\d+(?:\.\d+)?(?:kg|kgs|kilo|kilos|g|gr)?$/.test(t));
}

function buildMlQuery(name) {
  const brand = detectBrand(name);
  const weight = extractWeightKg(name);
  const tokens = productTokens(name)
    .filter(t => t !== "royal" && t !== "canin" && t !== "purina")
    .slice(0, 7);
  return [brand, ...tokens, weight ? `${weight} kg` : ""].filter(Boolean).join(" ").trim();
}

function buildMlUrl(name) {
  const query = String(name || "").trim().replace(/\s+/g, " ");
  const slug = query
    .trim()
    .replace(/\s+/g, "-")
    .toLowerCase()
    .split("")
    .map(ch => /[a-zA-Z0-9-]/.test(ch) ? ch : encodeURIComponent(ch))
    .join("");
  const fragment = query.toUpperCase().replace(/\s+/g, "%20");
  return `https://listado.mercadolibre.com.ar/${slug}#D[A:${fragment}]`;
}

function logisticsCost(price, cfg) {
  if (price < LOGISTICS_FIXED_UNDER) return Number(cfg.fixedLogistics) || LOGISTICS_FIXED_COST;
  return price * ((Number(cfg.fullPct) || 0) / 100);
}

function costNet(cost, cfg) {
  const purchaseVat = (Number(cfg.purchaseVatPct) || 0) / 100;
  return Number(cfg.costIncludesVat) ? cost / (1 + purchaseVat) : cost;
}

function meliParts(price, cost, cfg) {
  const ivaFactor = 1 + ((Number(cfg.ivaSalePct) || 0) / 100);
  const baseCost = costNet(cost, cfg);
  const saleVat = price * ((Number(cfg.ivaSalePct) || 0) / 100) / ivaFactor;
  const netSale = price - saleVat;
  const commission = price * ((Number(cfg.commission) || 0) / 100);
  const commissionNet = commission / ivaFactor;
  const installments = price * ((Number(cfg.installmentsPct) || 0) / 100);
  const installmentsNet = installments / ivaFactor;
  const iibb = price * ((Number(cfg.iibbPct) || 0) / 100);
  const ads = price * ((Number(cfg.adsPct) || 0) / 100);
  const logistics = logisticsCost(price, cfg);
  const logisticsNet = logistics / ivaFactor;
  const packaging = Number(cfg.packaging) || 0;
  const packagingNet = packaging / ivaFactor;
  const profit = price - saleVat - packagingNet - commissionNet - installmentsNet - logisticsNet - iibb - ads - baseCost;
  const incomeTax = Math.max(profit, 0) * ((Number(cfg.incomeTaxPct) || 0) / 100);
  return {
    ads,
    baseCost,
    commission,
    commissionNet,
    grossSale: price,
    installments,
    installmentsNet,
    iibb,
    incomeTax,
    logistics,
    logisticsNet,
    net: profit,
    netAfterTax: profit - incomeTax,
    netSale,
    packaging,
    packagingNet,
    saleVat,
  };
}

function calcNet(price, cost, cfg) {
  return meliParts(price, cost, cfg);
}

function meliFormulaPrice(cost, cfg, profitPct) {
  const ivaFactor = 1 + ((Number(cfg.ivaSalePct) || 0) / 100);
  const ivaPart = ((Number(cfg.ivaSalePct) || 0) / 100) / ivaFactor;
  const commissionNetPct = ((Number(cfg.commission) || 0) / 100) / ivaFactor;
  const installmentsNetPct = ((Number(cfg.installmentsPct) || 0) / 100) / ivaFactor;
  const iibbPct = (Number(cfg.iibbPct) || 0) / 100;
  const adsPct = (Number(cfg.adsPct) || 0) / 100;
  const fullNetPct = ((Number(cfg.fullPct) || 0) / 100) / ivaFactor;
  const baseCost = costNet(cost, cfg);
  const packagingNet = (Number(cfg.packaging) || 0) / ivaFactor;
  const targetCostAndProfit = baseCost * (1 + (profitPct || 0) / 100);
  const commonDen = 1 - ivaPart - commissionNetPct - installmentsNetPct - iibbPct - adsPct;
  const fixedLogistics = Number(cfg.fixedLogistics) || LOGISTICS_FIXED_COST;
  const lowPrice = (targetCostAndProfit + packagingNet + fixedLogistics / ivaFactor) / commonDen;
  if (lowPrice < LOGISTICS_FIXED_UNDER) return Math.ceil(lowPrice / 10) * 10;
  const highDen = commonDen - fullNetPct;
  if (highDen <= 0) return null;
  return Math.ceil((targetCostAndProfit + packagingNet) / highDen / 10) * 10;
}

function calcPrices(cost, cfg) {
  if (!cost || cost <= 0) return null;
  const baseCost = costNet(cost, cfg);
  const spMin = meliFormulaPrice(cost, cfg, 0);
  const spRec = meliFormulaPrice(cost, cfg, Number(cfg.markup) || 0);
  if (!spMin || !spRec) return null;
  const result = calcNet(spRec, cost, cfg);
  const discount = Math.min(Math.max(Number(cfg.discountPct) || 0, 0), 95);
  const spPublished = discount > 0 ? Math.ceil((spRec / (1 - discount / 100)) / 10) * 10 : spRec;
  return {
    ads: Math.round(result.ads),
    baseCost: Math.round(result.baseCost),
    commission: Math.round(result.commission),
    installments: Math.round(result.installments),
    iibb: Math.round(result.iibb),
    spRec,
    spMin,
    spPublished,
    saleVat: Math.round(result.saleVat),
    netSale: Math.round(result.netSale),
    packaging: Math.round(result.packaging),
    preTaxNet: Math.round(result.net),
    net: Math.round(result.netAfterTax),
    netPct: +((result.netAfterTax / baseCost) * 100).toFixed(1),
    incomeTax: Math.round(result.incomeTax),
    logistics: Math.round(result.logistics),
  };
}

function mlFromManual(record) {
  const prices = (record?.prices || []).map(parseNum).filter(n => n > 0).sort((a, b) => a - b);
  if (!prices.length) return { found: false, prices: record?.prices || ["", "", ""] };
  return {
    found: true,
    prices: record?.prices || ["", "", ""],
    min: prices[0],
    max: prices[prices.length - 1],
    representative: prices[Math.floor(prices.length / 2)],
    count: prices.length,
  };
}

function getDecision(c, ml) {
  if (!c || !ml?.found) return null;
  const rep = Number(ml.representative);
  if (!rep || rep <= 0) return null;
  if (rep >= c.spRec) return "buy";
  if (rep > c.spMin) return "caution";
  return "avoid";
}

export default function App() {
  const [cfg, setCfg] = useState({
    commission: 16.5,
    ivaSalePct: 21,
    purchaseVatPct: 21,
    costIncludesVat: 1,
    fixedLogistics: 3000,
    fullPct: 8,
    installmentsPct: 0,
    adsPct: 5,
    iibbPct: 5.48,
    incomeTaxPct: 35,
    packaging: 0,
    discountPct: 0,
    markup: 180,
  });
  const [cfgOpen, setCfgOpen] = useState(true);
  const [products, setProducts] = useState([]);
  const [mlData, setMlData] = useState({});
  const [cols, setCols] = useState(null);
  const [rawRows, setRawRows] = useState(null);
  const [colMap, setColMap] = useState({ name: "", price: "" });
  const [drag, setDrag] = useState(false);
  const [mName, setMName] = useState("");
  const [mCost, setMCost] = useState("");
  const fileRef = useRef();

  const ivaFactor = 1 + (Number(cfg.ivaSalePct) || 0) / 100;
  const totalFees = ((Number(cfg.ivaSalePct) || 0) / 100) / ivaFactor
    + ((Number(cfg.commission) || 0) / 100) / ivaFactor
    + ((Number(cfg.installmentsPct) || 0) / 100) / ivaFactor
    + (Number(cfg.adsPct) || 0) / 100
    + (Number(cfg.iibbPct) || 0) / 100;

  const stats = useMemo(() => products.reduce((a, p) => {
    const ml = mlFromManual(mlData[p.id]);
    const dec = getDecision(calcPrices(p.cost, cfg), ml);
    a.total += 1;
    if (ml.found) a.withPrices += 1;
    if (dec) a[dec] += 1;
    return a;
  }, { total: 0, withPrices: 0, buy: 0, caution: 0, avoid: 0 }), [products, mlData, cfg]);

  const input = {
    width: "100%",
    boxSizing: "border-box",
    height: "36px",
    border: `1px solid ${T.line2}`,
    borderRadius: "7px",
    background: "#fff",
    color: T.text,
    padding: "0 10px",
    outline: "none",
    font: "inherit",
    fontSize: "13px",
  };

  const button = (variant = "soft") => ({
    alignItems: "center",
    display: "inline-flex",
    justifyContent: "center",
    height: "34px",
    border: "none",
    borderRadius: "7px",
    padding: "0 12px",
    font: "inherit",
    fontSize: "12px",
    fontWeight: 750,
    cursor: "pointer",
    textDecoration: "none",
    whiteSpace: "nowrap",
    color: variant === "primary" ? "#fff" : variant === "danger" ? T.red : T.ink,
    background: variant === "primary" ? T.teal : variant === "danger" ? T.redBg : T.panel2,
    borderColor: variant === "primary" ? T.teal : T.line2,
    boxShadow: variant === "primary" ? "0 8px 18px rgba(15, 118, 110, .18)" : "none",
  });

  const parseFile = (file) => {
    if (!file) return;
    const r = new FileReader();
    r.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheetName = wb.SheetNames.find(n => /^meli$/i.test(String(n).trim())) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const all = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        const headerRow = findHeaderRow(all);
        const hdr = (all[headerRow] || [])
          .map((h, index) => ({ index, label: String(h).trim() }))
          .filter(c => c.label);
        const dat = all.slice(headerRow + 1).filter(row => row.some(c => c !== "" && c != null));
        if (!hdr.length) {
          alert("El archivo no tiene encabezados en la primera fila");
          return;
        }
        const findCol = (patterns) => {
          for (const rx of patterns) {
            const found = hdr.find(c => rx.test(c.label));
            if (found) return found.index;
          }
          return null;
        };
        const nameCol = findCol([/descrip/i, /nombre/i, /product/i, /item/i, /presenta/i, /articulo|artículo/i]);
        const priceCol = findCol([/costo.*compra.*sin\s*iva/i, /precio\s*final\s*c\/?iva/i, /final.*iva/i, /costo/i, /neto/i, /precio\s*lista/i, /importe/i, /valor/i]);
        const defaultCostMode = costIncludesVatFromLabel(hdr.find(c => c.index === priceCol)?.label);
        setCols(hdr);
        setRawRows(dat);
        setColMap({
          name: nameCol ?? hdr[0]?.index ?? "",
          price: priceCol ?? hdr[1]?.index ?? "",
        });
        if (defaultCostMode != null) {
          setCfg(c => ({ ...c, costIncludesVat: defaultCostMode }));
        }
      } catch (err) {
        alert(`Error al leer el archivo: ${err.message}`);
      }
    };
    r.readAsArrayBuffer(file);
  };

  const applyMap = () => {
    const ni = colMap.name === "" ? -1 : Number(colMap.name);
    const pi = colMap.price === "" ? -1 : Number(colMap.price);
    if (ni < 0 || pi < 0) {
      alert("Seleccioná columnas válidas");
      return;
    }
    const selectedCostMode = costIncludesVatFromLabel(cols?.find(c => c.index === pi)?.label);
    if (selectedCostMode != null) {
      setCfg(c => ({ ...c, costIncludesVat: selectedCostMode }));
    }
    const parsed = rawRows
      .map((row, i) => ({
        id: `f-${Date.now()}-${i}`,
        name: String(row[ni] ?? "").trim(),
        cost: parseNum(row[pi]),
      }))
      .filter(p => p.name && p.cost > 0);
    setProducts(prev => {
      const ex = new Set(prev.map(p => p.name.toLowerCase()));
      return [...prev, ...parsed.filter(p => !ex.has(p.name.toLowerCase()))];
    });
    setCols(null);
    setRawRows(null);
  };

  const addManual = () => {
    const cost = parseNum(mCost);
    if (!mName.trim() || cost <= 0) return;
    setProducts(prev => [...prev, { id: `m-${Date.now()}`, name: mName.trim(), cost }]);
    setMName("");
    setMCost("");
  };

  const removeProd = (id) => {
    setProducts(prev => prev.filter(p => p.id !== id));
    setMlData(prev => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
  };

  const updateMlPrice = (id, idx, value) => {
    setMlData(prev => {
      const prices = [...(prev[id]?.prices || ["", "", ""])];
      prices[idx] = value;
      return { ...prev, [id]: { prices } };
    });
  };

  const clearMlPrices = (id) => {
    setMlData(prev => ({ ...prev, [id]: { prices: ["", "", ""] } }));
  };

  const copyQuery = async (product) => {
    const url = buildMlUrl(product.name);
    try {
      await navigator.clipboard.writeText(url);
    } catch {
      window.prompt("Copiar link de Mercado Libre", url);
    }
  };

  const mapPreview = cols && rawRows
    ? rawRows.slice(0, 5).map((row) => {
      const ni = colMap.name === "" ? -1 : Number(colMap.name);
      const pi = colMap.price === "" ? -1 : Number(colMap.price);
      const name = ni >= 0 ? String(row[ni] ?? "").trim() : "";
      return {
        name,
        searchName: name,
        rawCost: pi >= 0 ? row[pi] : "",
        parsedCost: pi >= 0 ? parseNum(row[pi]) : 0,
      };
    })
    : [];

  return (
    <main style={{ minHeight: "100vh", background: T.page, color: T.text, fontFamily: "Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" }}>
      <div style={{ maxWidth: "1440px", margin: "0 auto", padding: "22px 18px 36px" }}>
        <header style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: "16px", marginBottom: "18px", flexWrap: "wrap" }}>
          <div>
            <div style={{ display: "inline-flex", alignItems: "center", gap: "8px", padding: "5px 9px", border: `1px solid ${T.line}`, background: T.panel, borderRadius: "999px", color: T.soft, fontSize: "12px", fontWeight: 750, marginBottom: "9px" }}>
              Royal Canin / Purina · Mercado Libre
            </div>
            <h1 style={{ margin: 0, color: T.ink, fontSize: "30px", lineHeight: 1.08, letterSpacing: 0, fontWeight: 850 }}>Decisor de compra de balanceado</h1>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, minmax(96px, 1fr))", gap: "8px", width: "min(520px, 100%)" }}>
            {[
              ["Productos", stats.total, T.blueBg, T.blue],
              ["Con ML", stats.withPrices, T.purpleBg, T.purple],
              ["Comprar", stats.buy, T.greenBg, T.green],
              ["No conv.", stats.avoid, T.redBg, T.red],
            ].map(([label, value, bg, fg]) => (
              <div key={label} style={{ background: bg, border: `1px solid ${T.line}`, borderRadius: "8px", padding: "10px 11px" }}>
                <div style={{ color: fg, fontSize: "22px", fontWeight: 850, lineHeight: 1 }}>{value}</div>
                <div style={{ color: T.soft, fontSize: "11px", fontWeight: 750, marginTop: "3px" }}>{label}</div>
              </div>
            ))}
          </div>
        </header>

        <section style={{ background: T.panel, border: `1px solid ${T.line}`, borderRadius: "8px", overflow: "hidden", marginBottom: "12px" }}>
          <button onClick={() => setCfgOpen(o => !o)} style={{ width: "100%", height: "44px", border: 0, background: T.panel, display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 14px", cursor: "pointer", color: T.ink, fontSize: "13px", fontWeight: 800 }}>
            <span>Estructura de costos</span>
            <span style={{ color: T.muted, fontSize: "12px" }}>{cfgOpen ? "Ocultar" : "Mostrar"}</span>
          </button>
          {cfgOpen && (
            <div style={{ borderTop: `1px solid ${T.line}`, padding: "14px" }}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: "10px", marginBottom: "12px" }}>
                {[
                  { k: "commission", l: "Costo MELI" },
                  { k: "ivaSalePct", l: "IVA venta" },
                  { k: "purchaseVatPct", l: "IVA compra" },
                  { k: "fixedLogistics", l: "Logística fija <33k $", money: true },
                  { k: "fullPct", l: "Full logística" },
                  { k: "installmentsPct", l: "Cuotas" },
                  { k: "adsPct", l: "ADS pauta" },
                  { k: "iibbPct", l: "IIBB" },
                  { k: "incomeTaxPct", l: "Ganancias" },
                  { k: "packaging", l: "Packaging $", money: true },
                  { k: "discountPct", l: "Descuento publicado" },
                  { k: "markup", l: "Rentabilidad hoja MELI" },
                ].map(({ k, l, money }) => (
                  <label key={k} style={{ display: "grid", gap: "5px", color: T.soft, fontSize: "11px", fontWeight: 800 }}>
                    {l}{money ? "" : " %"}
                    <input type="number" step="0.5" value={cfg[k] ?? ""}
                      onChange={e => setCfg(c => ({ ...c, [k]: parseFloat(e.target.value) || 0 }))}
                      style={{ ...input, color: T.teal, fontWeight: 800, fontSize: "15px" }} />
                  </label>
                ))}
              </div>
              <label style={{ display: "inline-flex", alignItems: "center", gap: "8px", marginBottom: "12px", color: T.soft, fontSize: "12px", fontWeight: 800 }}>
                <input type="checkbox" checked={!!Number(cfg.costIncludesVat)}
                  onChange={e => setCfg(c => ({ ...c, costIncludesVat: e.target.checked ? 1 : 0 }))}
                  style={{ width: "16px", height: "16px" }} />
                El costo cargado incluye IVA compra
              </label>
              <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", padding: "10px 12px", background: T.panel2, border: `1px solid ${T.line}`, borderRadius: "8px", color: T.soft, fontSize: "12px", fontWeight: 700 }}>
                <span>Denominador MELI descuenta <strong style={{ color: T.ink }}>{(totalFees * 100).toFixed(2)}%</strong></span>
                <span>Venta neta <strong style={{ color: T.teal }}>PV - IVA venta</strong></span>
                <span>Logística <strong style={{ color: T.teal }}>{cfg.fullPct}%</strong></span>
                <span>Si PV &lt; {$n(LOGISTICS_FIXED_UNDER)} <strong style={{ color: T.blue }}>logística fija {$n(cfg.fixedLogistics || LOGISTICS_FIXED_COST)}</strong></span>
                <span>Ganancias <strong style={{ color: T.purple }}>{cfg.incomeTaxPct}% sobre utilidad</strong></span>
              </div>
            </div>
          )}
        </section>

        {cols && (
          <section style={{ background: T.panel, border: `1px solid ${T.line}`, borderRadius: "8px", padding: "14px", marginBottom: "12px" }}>
            <div style={{ color: T.ink, fontSize: "14px", fontWeight: 850, marginBottom: "10px" }}>Mapeo del archivo</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: "10px", marginBottom: "12px" }}>
              {[{ k: "name", l: "Producto" }, { k: "price", l: "Costo" }].map(({ k, l }) => (
                <label key={k} style={{ display: "grid", gap: "5px", color: T.soft, fontSize: "11px", fontWeight: 800 }}>
                  {l}
                  <select value={colMap[k]} onChange={e => setColMap(m => ({ ...m, [k]: e.target.value }))} style={input}>
                    {cols.map(c => (
                      <option key={`${c.index}-${c.label}`} value={c.index}>
                        {excelColName(c.index)} · {c.label}
                      </option>
                    ))}
                  </select>
                </label>
              ))}
            </div>
            {mapPreview.length > 0 && (
              <div style={{ marginBottom: "12px", border: `1px solid ${T.line}`, borderRadius: "8px", overflow: "hidden" }}>
                <div style={{ padding: "8px 10px", background: T.panel2, color: T.soft, fontSize: "11px", fontWeight: 850 }}>Previsualización de importación</div>
                {mapPreview.map((row, idx) => (
                  <div key={`${row.name}-${idx}`} style={{ display: "grid", gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr) 140px 140px", gap: "8px", padding: "8px 10px", borderTop: `1px solid ${T.line}`, alignItems: "center", fontSize: "12px" }}>
                    <div style={{ color: T.ink, fontWeight: 750, overflow: "hidden", whiteSpace: "nowrap", textOverflow: "ellipsis" }}>{row.name || "-"}</div>
                    <div style={{ color: T.blue, fontWeight: 750, overflow: "hidden", whiteSpace: "nowrap", textOverflow: "ellipsis" }}>ML: {row.searchName || "-"}</div>
                    <div style={{ color: T.soft }}>Excel: {String(row.rawCost || "-")}</div>
                    <div style={{ color: T.teal, fontWeight: 850 }}>Costo: {$n(row.parsedCost)}</div>
                  </div>
                ))}
              </div>
            )}
            <div style={{ display: "flex", gap: "8px", flexWrap: "wrap" }}>
              <button onClick={applyMap} style={button("primary")}>Confirmar</button>
              <button onClick={() => { setCols(null); setRawRows(null); }} style={button()}>Cancelar</button>
            </div>
          </section>
        )}

        {products.length === 0 && !cols && (
          <section style={{ display: "grid", gridTemplateColumns: "minmax(280px, 1.2fr) minmax(280px, .8fr)", gap: "12px", marginBottom: "12px" }}>
            <div
              onDrop={e => { e.preventDefault(); setDrag(false); parseFile(e.dataTransfer.files[0]); }}
              onDragOver={e => { e.preventDefault(); setDrag(true); }}
              onDragLeave={() => setDrag(false)}
              onClick={() => fileRef.current?.click()}
              style={{ minHeight: "150px", display: "grid", placeItems: "center", textAlign: "center", cursor: "pointer", background: drag ? T.tealBg : T.panel, border: `2px dashed ${drag ? T.teal : T.line2}`, borderRadius: "8px", padding: "24px" }}
            >
              <div>
                <div style={{ fontSize: "15px", fontWeight: 850, color: T.ink, marginBottom: "4px" }}>Subir listado de precios</div>
                <div style={{ fontSize: "13px", color: T.soft }}>Excel o CSV</div>
              </div>
            </div>
            <div style={{ background: T.panel, border: `1px solid ${T.line}`, borderRadius: "8px", padding: "14px" }}>
              <div style={{ color: T.ink, fontSize: "14px", fontWeight: 850, marginBottom: "10px" }}>Producto puntual</div>
              <div style={{ display: "grid", gap: "8px" }}>
                <input placeholder="Royal Canin Maxi Adult 15kg" value={mName} onChange={e => setMName(e.target.value)} onKeyDown={e => e.key === "Enter" && addManual()} style={input} />
                <input placeholder="Costo $" value={mCost} onChange={e => setMCost(e.target.value)} onKeyDown={e => e.key === "Enter" && addManual()} style={input} />
                <button onClick={addManual} style={button("primary")}>Agregar producto</button>
              </div>
            </div>
          </section>
        )}

        <input ref={fileRef} type="file" accept=".xlsx,.xls,.xlsm,.csv" style={{ display: "none" }}
          onChange={e => { parseFile(e.target.files[0]); if (e.target) e.target.value = ""; }} />

        {products.length > 0 && (
          <section style={{ background: T.panel, border: `1px solid ${T.line}`, borderRadius: "8px", overflow: "hidden" }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 150px auto auto", gap: "8px", alignItems: "center", padding: "12px", borderBottom: `1px solid ${T.line}`, background: T.panel }}>
              <input placeholder="Producto" value={mName} onChange={e => setMName(e.target.value)} onKeyDown={e => e.key === "Enter" && addManual()} style={input} />
              <input placeholder="Costo $" value={mCost} onChange={e => setMCost(e.target.value)} onKeyDown={e => e.key === "Enter" && addManual()} style={input} />
              <button onClick={addManual} style={button("primary")}>Agregar</button>
              <div style={{ display: "flex", gap: "7px", justifyContent: "flex-end" }}>
                <button onClick={() => fileRef.current?.click()} style={button()}>Cargar lista</button>
                <button onClick={() => { setProducts([]); setMlData({}); }} style={button("danger")}>Limpiar</button>
              </div>
            </div>

            <div>
              {products.map((p) => {
                const c = calcPrices(p.cost, cfg);
                const mlRecord = mlData[p.id] || { prices: ["", "", ""] };
                const ml = mlFromManual(mlRecord);
                const dec = getDecision(c, ml);
                const gap = c && ml.found ? ml.representative - c.spRec : null;
                const prices = mlRecord.prices || ["", "", ""];
                const metric = (label, value, color = T.ink, detail = "") => (
                  <div style={{ minHeight: "58px", padding: "9px 10px", border: `1px solid ${T.line}`, borderRadius: "8px", background: T.panel2 }}>
                    <div style={{ color: T.muted, fontSize: "10px", fontWeight: 850, marginBottom: "3px" }}>{label}</div>
                    <div style={{ color, fontSize: "15px", fontWeight: 900, lineHeight: 1.15 }}>{value}</div>
                    {detail && <div style={{ color: T.muted, fontSize: "11px", marginTop: "2px", fontWeight: 700 }}>{detail}</div>}
                  </div>
                );
                return (
                  <article key={p.id} style={{ padding: "14px", borderBottom: `1px solid ${T.line}` }}>
                    <div style={{ display: "grid", gridTemplateColumns: "minmax(0, 1fr) auto", gap: "10px", alignItems: "start", marginBottom: "12px" }}>
                      <div style={{ minWidth: 0 }}>
                        <div title={p.name} style={{ color: T.ink, fontWeight: 900, fontSize: "15px", lineHeight: 1.25, overflowWrap: "anywhere" }}>{p.name}</div>
                      </div>
                      <div style={{ display: "flex", gap: "7px", flexWrap: "wrap", justifyContent: "flex-end" }}>
                        <a href={buildMlUrl(p.name)} style={button("primary")}>Abrir ML</a>
                        <button onClick={() => copyQuery(p)} style={button()}>Copiar link</button>
                      </div>
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(116px, 1fr))", gap: "8px", marginBottom: "12px" }}>
                      {metric("Costo bruto", $n(p.cost), T.soft, Number(cfg.costIncludesVat) ? "c/IVA" : "sin IVA")}
                      {metric("Costo neto", c ? $n(c.baseCost) : "-", T.soft)}
                      {metric("PV mínimo", c ? $n(c.spMin) : "-", T.amber)}
                      {metric("PV MELI", c ? $n(c.spRec) : "-", T.teal, "calculado")}
                      {Number(cfg.discountPct) > 0 && metric("Publicado", c ? $n(c.spPublished) : "-", T.blue, `${cfg.discountPct}% desc.`)}
                      {metric("Venta neta", c ? $n(c.netSale) : "-", T.ink, "sin IVA venta")}
                      {metric("IVA venta", c ? $n(c.saleVat) : "-", T.soft)}
                      {metric("Gcia contable", c ? $n(c.preTaxNet) : "-", T.purple, "antes Ganancias")}
                      {metric("Gcia final", c ? $n(c.net) : "-", T.green, c ? `${c.netPct}% s/costo` : "")}
                      {metric("Cargo vender", c ? $n(c.commission) : "-", T.soft, `${cfg.commission}% bruto`)}
                      {metric("Cuotas", c ? $n(c.installments) : "-", T.soft, `${cfg.installmentsPct}%`)}
                      {metric("IIBB", c ? $n(c.iibb) : "-", T.soft, `${cfg.iibbPct}%`)}
                      {metric("ADS", c ? $n(c.ads) : "-", T.soft, `${cfg.adsPct}%`)}
                      {metric("Logística", c ? $n(c.logistics) : "-", c?.spRec < LOGISTICS_FIXED_UNDER ? T.blue : T.soft, c?.spRec < LOGISTICS_FIXED_UNDER ? "fija" : `${cfg.fullPct}% PV`)}
                      {metric("Ganancias", c ? $n(c.incomeTax) : "-", T.purple, "info")}
                      {metric("ML típico", ml.found ? $n(ml.representative) : "-", T.ink, ml.found ? `${ml.count} precio${ml.count > 1 ? "s" : ""}` : "")}
                      {metric("Brecha", gap != null ? `${gap >= 0 ? "+" : ""}${$n(gap)}` : "-", gap != null && gap < 0 ? T.red : T.green)}
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(110px, 1fr))", gap: "8px", alignItems: "end" }}>
                      {[0, 1, 2].map(idx => (
                        <label key={idx} style={{ display: "grid", gap: "5px", color: T.soft, fontSize: "11px", fontWeight: 850 }}>
                          Precio ML {idx + 1}
                          <input inputMode="numeric" placeholder="$" value={prices[idx] || ""} onChange={e => updateMlPrice(p.id, idx, e.target.value)} style={{ ...input, height: "34px", fontWeight: 850 }} />
                        </label>
                      ))}
                      <div style={{ gridColumn: "1 / -1", display: "flex", gap: "7px", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap" }}>
                        {dec ? (
                          <span style={{ display: "inline-flex", alignItems: "center", height: "30px", padding: "0 10px", borderRadius: "999px", background: DEC[dec].bg, border: `1px solid ${DEC[dec].bd}`, color: DEC[dec].fg, fontSize: "11px", fontWeight: 900 }}>
                            {DEC[dec].label}
                          </span>
                        ) : <span style={{ color: T.muted, fontSize: "12px", fontWeight: 800 }}>Pendiente</span>}
                        <div style={{ display: "flex", gap: "6px", marginLeft: "auto" }}>
                          <button onClick={() => clearMlPrices(p.id)} style={button()}>Borrar ML</button>
                          <button onClick={() => removeProd(p.id)} style={button("danger")}>Quitar</button>
                        </div>
                      </div>
                    </div>
                  </article>
                );
              })}
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", gap: "10px", flexWrap: "wrap", padding: "12px", background: T.panel2, borderTop: `1px solid ${T.line}`, color: T.soft, fontSize: "12px", fontWeight: 700 }}>
              <span>PV MELI se calcula desde el costo: costo neto + {cfg.markup}% de rentabilidad, dividido por el denominador de la hoja MELI.</span>
              <span>No toma un precio ya calculado: lo reproduce con la fórmula usando el listado cargado.</span>
              <span>Si el PV es menor a {$n(LOGISTICS_FIXED_UNDER)}, logística usa {$n(cfg.fixedLogistics || LOGISTICS_FIXED_COST)} fijo.</span>
              <span>ML típico usa la mediana de los precios cargados.</span>
            </div>
          </section>
        )}
      </div>
    </main>
  );
}
