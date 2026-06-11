[antibiotic-consumption-tool (2).tsx](https://github.com/user-attachments/files/28817795/antibiotic-consumption-tool.2.tsx)
import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { Upload, Database, FlaskConical, LayoutDashboard, Copy, Trash2, Plus, Download, AlertCircle, CheckCircle2, Search, Settings } from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, ComposedChart, Line, LabelList } from "recharts";

// ALURA BRAND COLORS
const ALURA = {
  vinotinto: "#993935",
  coral: "#EB5852",
  grayLight: "#CCCCCC",
  gray: "#8B8B8D",
  grayDark: "#4A4A4D",
  white: "#FFFFFF",
  bgSoft: "#F5F5F5",
  bgTint: "#FBE9E7",
};

const CHART_COLORS = [ALURA.vinotinto, ALURA.coral, ALURA.gray, ALURA.grayLight, "#D4756F", "#6F2D2A", "#A8A8AB"];

// EMA categories (more concerning -> less concerning)
const EMA_LEVELS = {
  A: { label: "A – Avoid", color: "#7B1F1B" },
  B: { label: "B – Restriction", color: "#993935" },
  C: { label: "C – Caution", color: "#EB5852" },
  D: { label: "D – Prudence", color: "#8B8B8D" },
  NC: { label: "No clasificado", color: "#CCCCCC" },
};

// FDA categories
const FDA_LEVELS = {
  C: { label: "C - Critically Important", color: "#7B1F1B" },
  H: { label: "H - Highly Important", color: "#EB5852" },
  NO: { label: "No MI", color: "#8B8B8D" },
  NC: { label: "No clasificado", color: "#CCCCCC" },
};

// Pre-loaded EMA/FDA reference table for common active ingredients
const ACTIVE_INGREDIENTS_REF = {
  "halquinol": { ema: "D", fda: "NO" },
  "virginiamicina": { ema: "NC", fda: "NC" },
  "tilosina": { ema: "C", fda: "H" },
  "tiamulina": { ema: "D", fda: "H" },
  "clortetraciclina": { ema: "D", fda: "H" },
  "tilmicosina": { ema: "C", fda: "H" },
  "ceftiofur": { ema: "B", fda: "C" },
  "tulatromicina": { ema: "C", fda: "H" },
  "penicilina": { ema: "D", fda: "H" },
  "enrofloxacina": { ema: "B", fda: "C" },
  "enrofloxaxina": { ema: "B", fda: "C" },
  "oxitetraciclina": { ema: "D", fda: "H" },
  "amoxicilina": { ema: "D", fda: "H" },
  "lincomicina": { ema: "C", fda: "H" },
  "colistina": { ema: "B", fda: "C" },
  "florfenicol": { ema: "C", fda: "H" },
  "gentamicina": { ema: "C", fda: "C" },
  "neomicina": { ema: "D", fda: "H" },
  "espectinomicina": { ema: "D", fda: "H" },
  "sulfadiazina": { ema: "D", fda: "H" },
  "sulfametoxazol": { ema: "D", fda: "H" },
  "trimetoprima": { ema: "D", fda: "H" },
  "doxiciclina": { ema: "D", fda: "H" },
  "lincosamida": { ema: "C", fda: "H" },
};

const VIAS = ["Alimento", "Inyectable", "Agua"];

// Helpers
const normalizeFeedName = (s) => {
  if (!s) return "";
  return String(s)
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
};

const num = (v) => {
  if (v === null || v === undefined || v === "") return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(",", "."));
  return isNaN(n) ? 0 : n;
};

const findCol = (row, candidates) => {
  const keys = Object.keys(row);
  const norm = (s) => normalizeFeedName(s);
  for (const c of candidates) {
    const nc = norm(c);
    const found = keys.find((k) => norm(k) === nc);
    if (found) return found;
    const partial = keys.find((k) => norm(k).includes(nc));
    if (partial) return partial;
  }
  return null;
};

const inferEMA_FDA = (activeName) => {
  const n = normalizeFeedName(activeName);
  for (const [key, val] of Object.entries(ACTIVE_INGREDIENTS_REF)) {
    if (n.includes(key) || key.includes(n)) return val;
  }
  return { ema: "NC", fda: "NC" };
};

const parseExcelDate = (v) => {
  if (!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === "number") {
    // Excel serial date
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return isNaN(d) ? null : d;
  }
  const d = new Date(v);
  return isNaN(d) ? null : d;
};

const monthLabel = (d) => {
  if (!d) return "Sin fecha";
  const months = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  return `${months[d.getMonth()]} ${d.getFullYear()}`;
};

const monthKey = (d) => d ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}` : "0000-00";

const TABS = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "upload", label: "Cargar datos", icon: Upload },
  { id: "products", label: "Productos", icon: Database },
  { id: "recipes", label: "Recetas por lote", icon: FlaskConical },
  { id: "classification", label: "Clasificaciones", icon: Settings },
];

// ===== LOGO COMPONENTS =====
function AluraLogo({ size = 32, withText = true }) {
  return (
    <div className="flex items-center gap-2">
      <svg viewBox="0 0 100 100" width={size} height={size} fill="none">
        <path
          d="M 28 80 L 50 22 L 72 80 M 38 80 L 50 50 L 62 80"
          stroke={ALURA.vinotinto}
          strokeWidth="8"
          strokeLinecap="round"
          strokeLinejoin="round"
          fill="none"
        />
      </svg>
      {withText && (
        <span className="font-semibold text-2xl tracking-tight" style={{ color: ALURA.gray, fontFamily: "Nunito, Quicksand, sans-serif" }}>
          Alura
        </span>
      )}
    </div>
  );
}

function AluraLogoCircle({ size = 36 }) {
  return (
    <div
      className="flex items-center justify-center rounded-full"
      style={{ width: size, height: size, background: ALURA.vinotinto }}
    >
      <svg viewBox="0 0 100 100" width={size * 0.6} height={size * 0.6} fill="none">
        <path
          d="M 28 80 L 50 22 L 72 80 M 38 80 L 50 50 L 62 80"
          stroke="white"
          strokeWidth="9"
          strokeLinecap="round"
          strokeLinejoin="round"
          fill="none"
        />
      </svg>
    </div>
  );
}

// ===== MAIN APP =====
export default function App() {
  const [tab, setTab] = useState("dashboard");
  const [consumos, setConsumos] = useState([]);
  const [productos, setProductos] = useState([]);
  const [recetas, setRecetas] = useState({});
  const [feedAliases, setFeedAliases] = useState({});
  const [activosOverride, setActivosOverride] = useState({}); // {activoNormalized: {ema, fda}}
  const [loading, setLoading] = useState(true);
  const [msg, setMsg] = useState(null);

  useEffect(() => {
    (async () => {
      try {
        const k = ["consumos", "productos", "recetas", "feedAliases", "activosOverride"];
        const results = await Promise.allSettled(k.map((key) => window.storage.get(key)));
        const [c, p, r, f, a] = results.map((res) => (res.status === "fulfilled" && res.value ? res.value.value : null));
        if (c) setConsumos(JSON.parse(c));
        if (p) setProductos(JSON.parse(p));
        if (r) setRecetas(JSON.parse(r));
        if (f) setFeedAliases(JSON.parse(f));
        if (a) setActivosOverride(JSON.parse(a));
      } catch (e) {
        console.log("No previous data");
      }
      setLoading(false);
    })();
  }, []);

  useEffect(() => { if (!loading) window.storage.set("consumos", JSON.stringify(consumos)).catch(() => {}); }, [consumos, loading]);
  useEffect(() => { if (!loading) window.storage.set("productos", JSON.stringify(productos)).catch(() => {}); }, [productos, loading]);
  useEffect(() => { if (!loading) window.storage.set("recetas", JSON.stringify(recetas)).catch(() => {}); }, [recetas, loading]);
  useEffect(() => { if (!loading) window.storage.set("feedAliases", JSON.stringify(feedAliases)).catch(() => {}); }, [feedAliases, loading]);
  useEffect(() => { if (!loading) window.storage.set("activosOverride", JSON.stringify(activosOverride)).catch(() => {}); }, [activosOverride, loading]);

  const showMsg = (text, type = "info") => {
    setMsg({ text, type });
    setTimeout(() => setMsg(null), 4000);
  };

  const handleFile = async (file, kind) => {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!rows.length) {
        showMsg("El archivo está vacío", "error");
        return;
      }
      if (kind === "consumos") {
        const sample = rows[0];
        const cols = {
          fechaIni: findCol(sample, ["fecha inicial del consumo", "fecha inicial", "fecha ini"]),
          fechaFin: findCol(sample, ["fecha final del consumo", "fecha final", "fecha fin"]),
          granja: findCol(sample, ["granja"]),
          etapa: findCol(sample, ["etapa"]),
          lote: findCol(sample, ["lote"]),
          dias: findCol(sample, ["dias de vida", "días de vida"]),
          semana: findCol(sample, ["semana"]),
          alimento: findCol(sample, ["alimento"]),
          consumo: findCol(sample, ["consumo de alimento (kg)", "consumo de alimento", "consumo kg", "consumo"]),
        };
        const missing = Object.entries(cols).filter(([k, v]) => !v && k !== "dias" && k !== "semana").map(([k]) => k);
        if (missing.length) {
          showMsg(`Faltan columnas: ${missing.join(", ")}`, "error");
          return;
        }
        const parsed = rows.map((r, i) => ({
          id: i,
          fechaIni: r[cols.fechaIni],
          fechaFin: r[cols.fechaFin],
          granja: String(r[cols.granja] || ""),
          etapa: String(r[cols.etapa] || ""),
          lote: String(r[cols.lote] || ""),
          dias: cols.dias ? num(r[cols.dias]) : null,
          semana: cols.semana ? num(r[cols.semana]) : null,
          alimento: String(r[cols.alimento] || ""),
          consumoKg: num(r[cols.consumo]),
        })).filter((r) => r.lote && r.alimento);
        setConsumos(parsed);
        const aliases = { ...feedAliases };
        parsed.forEach((r) => {
          const k = normalizeFeedName(r.alimento);
          if (!aliases[k]) aliases[k] = r.alimento.trim();
        });
        setFeedAliases(aliases);
        showMsg(`Cargados ${parsed.length} consumos`, "success");
      } else {
        const sample = rows[0];
        const cols = {
          producto: findCol(sample, ["producto comercial", "producto"]),
          empresa: findCol(sample, ["empresa"]),
          activo: findCol(sample, ["ingrediente activo", "activo"]),
          conc: findCol(sample, ["concentracion", "concentración"]),
          precio: findCol(sample, ["precio"]),
          via: findCol(sample, ["via", "vía", "via de administracion"]),
        };
        const missing = Object.entries(cols).filter(([k, v]) => !v && k !== "via").map(([k]) => k);
        if (missing.length) {
          showMsg(`Faltan columnas: ${missing.join(", ")}`, "error");
          return;
        }
        const parsed = rows.map((r, i) => ({
          id: `p${i}`,
          producto: String(r[cols.producto] || ""),
          empresa: String(r[cols.empresa] || ""),
          activo: String(r[cols.activo] || ""),
          concentracion: num(r[cols.conc]),
          precio: num(r[cols.precio]),
          via: cols.via ? String(r[cols.via] || "Alimento") : "Alimento",
        })).filter((r) => r.producto);
        setProductos(parsed);
        showMsg(`Cargados ${parsed.length} productos`, "success");
      }
    } catch (e) {
      showMsg(`Error: ${e.message}`, "error");
    }
  };

  // Derived: unique (lote, feedKey)
  const loteAlimentos = useMemo(() => {
    const map = new Map();
    consumos.forEach((c) => {
      const fk = normalizeFeedName(c.alimento);
      const key = `${c.lote}||${fk}`;
      if (!map.has(key)) {
        map.set(key, { key, lote: c.lote, feedKey: fk, alimentoDisplay: feedAliases[fk] || c.alimento, totalKg: 0, granjas: new Set(), etapas: new Set() });
      }
      const o = map.get(key);
      o.totalKg += c.consumoKg;
      o.granjas.add(c.granja);
      o.etapas.add(c.etapa);
    });
    return Array.from(map.values()).map((o) => ({ ...o, granjas: Array.from(o.granjas).join(", "), etapas: Array.from(o.etapas).join(", ") }));
  }, [consumos, feedAliases]);

  // Derived: results
  const resultados = useMemo(() => {
    const rows = [];
    consumos.forEach((c) => {
      const fk = normalizeFeedName(c.alimento);
      const rkey = `${c.lote}||${fk}`;
      const receta = recetas[rkey] || [];
      const fecha = parseExcelDate(c.fechaIni);
      if (!receta.length) {
        rows.push({
          ...c,
          fechaParsed: fecha,
          monthKey: monthKey(fecha),
          monthLabel: monthLabel(fecha),
          alimentoDisplay: feedAliases[fk] || c.alimento,
          producto: "(sin receta)",
          empresa: "", activo: "", via: "",
          kgProducto: 0, kgActivo: 0, mgActivo: 0, costo: 0, sinReceta: true,
          ema: "NC", fda: "NC",
        });
        return;
      }
      receta.forEach((item) => {
        const p = productos.find((pp) => pp.id === item.productId);
        if (!p) return;
        const kgProducto = (c.consumoKg * num(item.dosis_g_ton)) / 1_000_000;
        const kgActivo = kgProducto * (p.concentracion / 100);
        const mgActivo = kgActivo * 1_000_000;
        const costo = kgProducto * p.precio;
        const ovKey = normalizeFeedName(p.activo);
        const ov = activosOverride[ovKey] || inferEMA_FDA(p.activo);
        rows.push({
          ...c,
          fechaParsed: fecha,
          monthKey: monthKey(fecha),
          monthLabel: monthLabel(fecha),
          alimentoDisplay: feedAliases[fk] || c.alimento,
          producto: p.producto, empresa: p.empresa, activo: p.activo,
          via: p.via || "Alimento",
          kgProducto, kgActivo, mgActivo, costo,
          sinReceta: false,
          ema: ov.ema, fda: ov.fda,
        });
      });
    });
    return rows;
  }, [consumos, productos, recetas, feedAliases, activosOverride]);

  const resetAll = async () => {
    if (!confirm("¿Borrar TODOS los datos?")) return;
    setConsumos([]); setProductos([]); setRecetas({}); setFeedAliases({}); setActivosOverride({});
    await Promise.all([
      window.storage.delete("consumos").catch(() => {}),
      window.storage.delete("productos").catch(() => {}),
      window.storage.delete("recetas").catch(() => {}),
      window.storage.delete("feedAliases").catch(() => {}),
      window.storage.delete("activosOverride").catch(() => {}),
    ]);
    showMsg("Datos reiniciados", "success");
  };

  if (loading) return <div className="p-8 text-center" style={{ color: ALURA.gray }}>Cargando...</div>;

  return (
    <div className="min-h-screen" style={{ background: ALURA.bgSoft, fontFamily: "Inter, Manrope, system-ui, sans-serif" }}>
      <div className="max-w-[1400px] mx-auto p-4">
        {/* HEADER */}
        <header className="flex items-center justify-between mb-4 pb-3 border-b-2" style={{ borderColor: ALURA.vinotinto }}>
          <div className="flex items-center gap-3">
            <AluraLogo size={42} />
            <div className="border-l-2 pl-3" style={{ borderColor: ALURA.grayLight }}>
              <h1 className="font-bold text-lg leading-tight" style={{ color: ALURA.grayDark, fontFamily: "Nunito, sans-serif" }}>
                Consumo de Antibióticos
              </h1>
              <p className="text-xs italic" style={{ color: ALURA.gray }}>
                Calculadora de mg de ingrediente activo y costo de medicación
              </p>
            </div>
          </div>
          <div className="text-right text-[10px]" style={{ color: ALURA.gray }}>
            <div>Última actualización</div>
            <div className="font-medium">{new Date().toLocaleString("es-CO", { dateStyle: "short", timeStyle: "short" })}</div>
          </div>
        </header>

        {/* MESSAGES */}
        {msg && (
          <div className={`mb-3 p-3 rounded flex items-center gap-2 text-sm`} style={{
            background: msg.type === "error" ? "#FEE2E2" : msg.type === "success" ? "#D1FAE5" : "#DBEAFE",
            color: msg.type === "error" ? "#991B1B" : msg.type === "success" ? "#065F46" : "#1E40AF",
          }}>
            {msg.type === "error" ? <AlertCircle size={16} /> : <CheckCircle2 size={16} />}
            {msg.text}
          </div>
        )}

        {/* NAV */}
        <nav className="flex gap-1 mb-4 border-b overflow-x-auto" style={{ borderColor: ALURA.grayLight }}>
          {TABS.map((t) => {
            const Icon = t.icon;
            const active = tab === t.id;
            return (
              <button
                key={t.id}
                onClick={() => setTab(t.id)}
                className="px-4 py-2 text-sm font-medium border-b-2 flex items-center gap-2 whitespace-nowrap transition-colors"
                style={{
                  borderColor: active ? ALURA.vinotinto : "transparent",
                  color: active ? ALURA.vinotinto : ALURA.gray,
                  fontFamily: "Nunito, sans-serif",
                }}
              >
                <Icon size={16} />
                {t.label}
              </button>
            );
          })}
        </nav>

        {tab === "dashboard" && <Dashboard resultados={resultados} consumos={consumos} />}
        {tab === "upload" && <UploadTab consumos={consumos} productos={productos} onFile={handleFile} onReset={resetAll} />}
        {tab === "products" && <ProductsTab productos={productos} setProductos={setProductos} />}
        {tab === "recipes" && <RecipesTab loteAlimentos={loteAlimentos} productos={productos} recetas={recetas} setRecetas={setRecetas} feedAliases={feedAliases} setFeedAliases={setFeedAliases} />}
        {tab === "classification" && <ClassificationTab productos={productos} activosOverride={activosOverride} setActivosOverride={setActivosOverride} />}

        {/* FOOTER */}
        <footer className="mt-6 pt-3 border-t flex items-center justify-end" style={{ borderColor: ALURA.grayLight }}>
          <AluraLogo size={28} />
        </footer>
      </div>
    </div>
  );
}

// ===== DASHBOARD TAB =====
function Dashboard({ resultados, consumos }) {
  const [fActivo, setFActivo] = useState("");
  const [fMonth, setFMonth] = useState("");
  const [fGranja, setFGranja] = useState("");

  // Placeholder: kg producidos by lote (until events file integration)
  // For now: sum of food consumption (kg) per granja as proxy - WILL BE REPLACED IN STEP 2
  const kgProducidosTotal = useMemo(() => {
    // Until events file is integrated, allow user to see ALL metrics by using a placeholder
    // We use a conservative estimate: total food / 2.5 (typical feed conversion ratio)
    // BUT we mark it as "estimado" — proper value comes from events file
    const totalFood = consumos.reduce((s, c) => s + c.consumoKg, 0);
    return totalFood > 0 ? totalFood / 2.5 : 0;
  }, [consumos]);

  const filtered = resultados.filter((r) => {
    if (r.sinReceta) return false;
    if (fActivo && r.activo !== fActivo) return false;
    if (fMonth && r.monthKey !== fMonth) return false;
    if (fGranja && r.granja !== fGranja) return false;
    return true;
  });

  const uniq = (arr, field) => Array.from(new Set(arr.map((r) => r[field]).filter(Boolean))).sort();
  const activos = uniq(resultados.filter((r) => !r.sinReceta), "activo");
  const granjas = uniq(resultados, "granja");
  const months = Array.from(new Set(resultados.map((r) => r.monthKey).filter((k) => k !== "0000-00")))
    .sort()
    .map((k) => ({ key: k, label: resultados.find((r) => r.monthKey === k)?.monthLabel || k }));

  // KPIs
  const totalMg = filtered.reduce((s, r) => s + r.mgActivo, 0);
  const totalCosto = filtered.reduce((s, r) => s + r.costo, 0);
  const mgPorKg = kgProducidosTotal > 0 ? totalMg / kgProducidosTotal : 0;
  const costoPorKg = kgProducidosTotal > 0 ? totalCosto / kgProducidosTotal : 0;

  // Chart: mg + costo por mes
  const mgCostoMes = useMemo(() => {
    const map = new Map();
    filtered.forEach((r) => {
      if (!map.has(r.monthKey)) map.set(r.monthKey, { key: r.monthKey, mes: r.monthLabel, mg: 0, costo: 0 });
      const o = map.get(r.monthKey);
      o.mg += r.mgActivo;
      o.costo += r.costo;
    });
    return Array.from(map.values()).sort((a, b) => a.key.localeCompare(b.key));
  }, [filtered]);

  // Chart: mg/kg + $/kg por mes
  const mgKgMes = useMemo(() => {
    // group by month: mg / kgProducidos (proxy split by month proportion)
    return mgCostoMes.map((m) => ({
      mes: m.mes,
      mgKg: kgProducidosTotal > 0 ? (m.mg / kgProducidosTotal) : 0,
      costoKg: kgProducidosTotal > 0 ? (m.costo / kgProducidosTotal) : 0,
    }));
  }, [mgCostoMes, kgProducidosTotal]);

  // Donut: vía de administración (costo)
  const viaData = useMemo(() => {
    const map = new Map();
    filtered.forEach((r) => {
      const v = r.via || "Alimento";
      map.set(v, (map.get(v) || 0) + r.costo);
    });
    return Array.from(map.entries()).map(([name, value]) => ({ name, value }));
  }, [filtered]);

  // Heatmap granja x mes (mg/kg)
  const heatmap = useMemo(() => {
    const map = new Map();
    filtered.forEach((r) => {
      const k = `${r.granja}||${r.monthKey}`;
      if (!map.has(k)) map.set(k, { granja: r.granja, monthKey: r.monthKey, mes: r.monthLabel, mg: 0 });
      map.get(k).mg += r.mgActivo;
    });
    const rows = granjas.map((g) => {
      const obj = { granja: g };
      months.forEach((m) => {
        const cell = map.get(`${g}||${m.key}`);
        obj[m.key] = cell ? (cell.mg / kgProducidosTotal) : null;
      });
      return obj;
    });
    return rows;
  }, [filtered, granjas, months, kgProducidosTotal]);

  // Bar: mg/kg por molecula por granja (stacked)
  const moleculaPorGranja = useMemo(() => {
    const topActivos = (() => {
      const map = new Map();
      filtered.forEach((r) => map.set(r.activo, (map.get(r.activo) || 0) + r.mgActivo));
      return Array.from(map.entries()).sort((a, b) => b[1] - a[1]).slice(0, 4).map((x) => x[0]);
    })();
    const rows = granjas.map((g) => {
      const obj = { granja: g };
      let otros = 0;
      filtered.filter((r) => r.granja === g).forEach((r) => {
        const v = kgProducidosTotal > 0 ? r.mgActivo / kgProducidosTotal : 0;
        if (topActivos.includes(r.activo)) {
          obj[r.activo] = (obj[r.activo] || 0) + v;
        } else {
          otros += v;
        }
      });
      if (otros > 0) obj["Otros"] = otros;
      return obj;
    });
    return { rows, keys: [...topActivos, "Otros"] };
  }, [filtered, granjas, kgProducidosTotal]);

  // EMA stacked bars by month
  const emaMes = useMemo(() => {
    const map = new Map();
    filtered.forEach((r) => {
      if (!map.has(r.monthKey)) map.set(r.monthKey, { mes: r.monthLabel, key: r.monthKey });
      const o = map.get(r.monthKey);
      const mgKg = kgProducidosTotal > 0 ? r.mgActivo / kgProducidosTotal : 0;
      o[r.ema] = (o[r.ema] || 0) + mgKg;
    });
    return Array.from(map.values()).sort((a, b) => a.key.localeCompare(b.key));
  }, [filtered, kgProducidosTotal]);

  // FDA stacked bars by month
  const fdaMes = useMemo(() => {
    const map = new Map();
    filtered.forEach((r) => {
      if (!map.has(r.monthKey)) map.set(r.monthKey, { mes: r.monthLabel, key: r.monthKey });
      const o = map.get(r.monthKey);
      const mgKg = kgProducidosTotal > 0 ? r.mgActivo / kgProducidosTotal : 0;
      o[r.fda] = (o[r.fda] || 0) + mgKg;
    });
    return Array.from(map.values()).sort((a, b) => a.key.localeCompare(b.key));
  }, [filtered, kgProducidosTotal]);

  // Tabla cumplimiento por activo
  const cumplimientoTabla = useMemo(() => {
    const map = new Map();
    filtered.forEach((r) => {
      if (!map.has(r.activo)) map.set(r.activo, { activo: r.activo, mg: 0, ema: r.ema, fda: r.fda });
      map.get(r.activo).mg += r.mgActivo;
    });
    return Array.from(map.values()).sort((a, b) => b.mg - a.mg);
  }, [filtered]);

  if (!resultados.length || resultados.every((r) => r.sinReceta)) {
    return (
      <div className="bg-white border rounded-lg p-10 text-center" style={{ borderColor: ALURA.grayLight }}>
        <FlaskConical size={40} className="mx-auto mb-3" style={{ color: ALURA.gray }} />
        <h3 className="text-lg font-semibold mb-1" style={{ color: ALURA.grayDark, fontFamily: "Nunito, sans-serif" }}>
          Aún no hay datos para mostrar
        </h3>
        <p className="text-sm" style={{ color: ALURA.gray }}>
          Carga los archivos de consumos y productos, y captura las recetas por lote.
        </p>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {/* Filters row */}
      <div className="bg-white rounded-lg p-3 flex flex-wrap gap-2 items-center border" style={{ borderColor: ALURA.grayLight }}>
        <span className="text-xs font-semibold uppercase tracking-wide" style={{ color: ALURA.gray }}>Filtros:</span>
        <select value={fActivo} onChange={(e) => setFActivo(e.target.value)} className="text-sm border rounded px-2 py-1" style={{ borderColor: ALURA.grayLight }}>
          <option value="">Todos los principios activos</option>
          {activos.map((a) => <option key={a} value={a}>{a}</option>)}
        </select>
        <select value={fMonth} onChange={(e) => setFMonth(e.target.value)} className="text-sm border rounded px-2 py-1" style={{ borderColor: ALURA.grayLight }}>
          <option value="">Todas las fechas</option>
          {months.map((m) => <option key={m.key} value={m.key}>{m.label}</option>)}
        </select>
        <select value={fGranja} onChange={(e) => setFGranja(e.target.value)} className="text-sm border rounded px-2 py-1" style={{ borderColor: ALURA.grayLight }}>
          <option value="">Todas las granjas</option>
          {granjas.map((g) => <option key={g} value={g}>{g}</option>)}
        </select>
      </div>

      {/* KPI row */}
      <div>
        <div className="text-[10px] uppercase tracking-wider mb-1 font-semibold" style={{ color: ALURA.gray }}>Indicadores Consolidados del Periodo</div>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          <KpiCard icon="🐷" label="Kg Producidos" value={kgProducidosTotal.toLocaleString(undefined, { maximumFractionDigits: 0 })} hint="(estimado · pendiente archivo de eventos)" />
          <KpiCard icon="💊" label="mg AB/kg producido" value={mgPorKg.toLocaleString(undefined, { maximumFractionDigits: 1 })} highlighted />
          <KpiCard icon="💰" label="$/Kg producido" value={costoPorKg.toLocaleString(undefined, { maximumFractionDigits: 1 })} highlighted />
          <KpiCard icon="🧪" label="Consumo total mg" value={totalMg.toLocaleString(undefined, { maximumFractionDigits: 0 })} />
        </div>
      </div>

      {/* Indicadores mensuales */}
      <div>
        <div className="text-[10px] uppercase tracking-wider mb-1 font-semibold" style={{ color: ALURA.gray }}>Indicadores mensuales</div>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
          <ChartCard title="Consumo total mg AB por mes – costo mes">
            {mgCostoMes.length === 0 ? <EmptyChart /> : (
              <ResponsiveContainer width="100%" height={220}>
                <ComposedChart data={mgCostoMes}>
                  <CartesianGrid strokeDasharray="3 3" stroke={ALURA.grayLight} />
                  <XAxis dataKey="mes" tick={{ fontSize: 10 }} />
                  <YAxis yAxisId="left" tick={{ fontSize: 10 }} tickFormatter={(v) => `${(v / 1e6).toFixed(0)}M`} />
                  <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 10 }} tickFormatter={(v) => `$${(v / 1000).toFixed(0)}k`} />
                  <Tooltip formatter={(v, n) => n === "costo" ? `$${v.toLocaleString(undefined, { maximumFractionDigits: 0 })}` : v.toLocaleString(undefined, { maximumFractionDigits: 0 })} />
                  <Legend wrapperStyle={{ fontSize: 10 }} />
                  <Bar yAxisId="left" dataKey="mg" name="mg ingrediente activo" fill={ALURA.vinotinto} />
                  <Line yAxisId="right" type="monotone" dataKey="costo" name="Costo" stroke={ALURA.grayDark} strokeWidth={2} dot={{ fill: ALURA.grayDark }} />
                </ComposedChart>
              </ResponsiveContainer>
            )}
          </ChartCard>
          <ChartCard title="Consumo mg AB/kg – $/kg por Mes">
            {mgKgMes.length === 0 ? <EmptyChart /> : (
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={mgKgMes}>
                  <CartesianGrid strokeDasharray="3 3" stroke={ALURA.grayLight} />
                  <XAxis dataKey="mes" tick={{ fontSize: 10 }} />
                  <YAxis tick={{ fontSize: 10 }} />
                  <Tooltip formatter={(v) => v.toLocaleString(undefined, { maximumFractionDigits: 1 })} />
                  <Legend wrapperStyle={{ fontSize: 10 }} />
                  <Bar dataKey="mgKg" name="mg activo / kg" fill={ALURA.vinotinto} />
                  <Bar dataKey="costoKg" name="$ / kg" fill={ALURA.gray} />
                </BarChart>
              </ResponsiveContainer>
            )}
          </ChartCard>
          <ChartCard title="$/kg por Vía de administración">
            {viaData.length === 0 ? <EmptyChart /> : (
              <ResponsiveContainer width="100%" height={220}>
                <PieChart>
                  <Pie data={viaData} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={50} outerRadius={80} label={(e) => `${((e.value / viaData.reduce((s, d) => s + d.value, 0)) * 100).toFixed(1)}%`}>
                    {viaData.map((d, i) => <Cell key={i} fill={d.name === "Alimento" ? ALURA.gray : ALURA.vinotinto} />)}
                  </Pie>
                  <Tooltip formatter={(v) => `$${v.toLocaleString(undefined, { maximumFractionDigits: 0 })}`} />
                  <Legend wrapperStyle={{ fontSize: 10 }} />
                </PieChart>
              </ResponsiveContainer>
            )}
          </ChartCard>
        </div>
      </div>

      {/* Indicadores por granja */}
      <div>
        <div className="text-[10px] uppercase tracking-wider mb-1 font-semibold" style={{ color: ALURA.gray }}>Indicadores por granja</div>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          <ChartCard title="Consumo mg AB/kg por granja">
            {heatmap.length === 0 ? <EmptyChart /> : (
              <div className="overflow-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr style={{ background: ALURA.bgSoft }}>
                      <th className="text-left p-2 font-semibold" style={{ color: ALURA.grayDark }}>Granja</th>
                      {months.map((m) => <th key={m.key} className="text-right p-2 font-semibold" style={{ color: ALURA.grayDark }}>{m.label}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {heatmap.map((row) => (
                      <tr key={row.granja} className="border-t" style={{ borderColor: ALURA.grayLight }}>
                        <td className="p-2 font-medium" style={{ color: ALURA.grayDark }}>{row.granja}</td>
                        {months.map((m) => {
                          const v = row[m.key];
                          const maxV = Math.max(...heatmap.flatMap((r) => months.map((mm) => r[mm.key] || 0)));
                          const intensity = v && maxV ? v / maxV : 0;
                          return (
                            <td key={m.key} className="p-2 text-right font-medium" style={{
                              color: intensity > 0.5 ? "white" : ALURA.grayDark,
                              background: v == null ? "transparent" : `rgba(153, 57, 53, ${0.1 + intensity * 0.8})`,
                            }}>
                              {v == null ? "-" : v.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </ChartCard>
          <ChartCard title="Consumo de ingrediente activo por molécula por Granja">
            {moleculaPorGranja.rows.length === 0 ? <EmptyChart /> : (
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={moleculaPorGranja.rows}>
                  <CartesianGrid strokeDasharray="3 3" stroke={ALURA.grayLight} />
                  <XAxis dataKey="granja" tick={{ fontSize: 10 }} />
                  <YAxis tick={{ fontSize: 10 }} />
                  <Tooltip formatter={(v) => v.toLocaleString(undefined, { maximumFractionDigits: 0 })} />
                  <Legend wrapperStyle={{ fontSize: 10 }} />
                  {moleculaPorGranja.keys.map((k, i) => (
                    <Bar key={k} dataKey={k} stackId="a" fill={CHART_COLORS[i % CHART_COLORS.length]} />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            )}
          </ChartCard>
        </div>
      </div>

      {/* Perfil de Riesgo */}
      <div>
        <div className="text-[10px] uppercase tracking-wider mb-1 font-semibold" style={{ color: ALURA.gray }}>Perfil de Riesgo y Cumplimiento Antimicrobiano</div>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
          <ChartCard title="Consumo de ingrediente activo por Clasificación EMA">
            {emaMes.length === 0 ? <EmptyChart /> : (
              <ResponsiveContainer width="100%" height={240}>
                <BarChart data={emaMes}>
                  <CartesianGrid strokeDasharray="3 3" stroke={ALURA.grayLight} />
                  <XAxis dataKey="mes" tick={{ fontSize: 10 }} />
                  <YAxis tick={{ fontSize: 10 }} />
                  <Tooltip formatter={(v) => v.toLocaleString(undefined, { maximumFractionDigits: 0 })} />
                  <Legend wrapperStyle={{ fontSize: 9 }} />
                  {Object.entries(EMA_LEVELS).map(([k, v]) => (
                    <Bar key={k} dataKey={k} stackId="ema" fill={v.color} name={v.label} />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            )}
          </ChartCard>
          <ChartCard title="Consumo de ingrediente activo por Clasificación FDA">
            {fdaMes.length === 0 ? <EmptyChart /> : (
              <ResponsiveContainer width="100%" height={240}>
                <BarChart data={fdaMes}>
                  <CartesianGrid strokeDasharray="3 3" stroke={ALURA.grayLight} />
                  <XAxis dataKey="mes" tick={{ fontSize: 10 }} />
                  <YAxis tick={{ fontSize: 10 }} />
                  <Tooltip formatter={(v) => v.toLocaleString(undefined, { maximumFractionDigits: 0 })} />
                  <Legend wrapperStyle={{ fontSize: 9 }} />
                  {Object.entries(FDA_LEVELS).map(([k, v]) => (
                    <Bar key={k} dataKey={k} stackId="fda" fill={v.color} name={v.label} />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            )}
          </ChartCard>
          <ChartCard title="Cumplimiento Consumo AB – Alura">
            {cumplimientoTabla.length === 0 ? <EmptyChart /> : (
              <div className="overflow-auto max-h-[240px]">
                <table className="w-full text-xs">
                  <thead className="sticky top-0" style={{ background: ALURA.bgSoft }}>
                    <tr>
                      <th className="text-left p-1.5 font-semibold" style={{ color: ALURA.grayDark }}>Activo</th>
                      <th className="text-center p-1.5 font-semibold" style={{ color: ALURA.grayDark }}>EMA</th>
                      <th className="text-center p-1.5 font-semibold" style={{ color: ALURA.grayDark }}>FDA</th>
                      <th className="text-right p-1.5 font-semibold" style={{ color: ALURA.grayDark }}>mg</th>
                    </tr>
                  </thead>
                  <tbody>
                    {cumplimientoTabla.map((r) => (
                      <tr key={r.activo} className="border-t" style={{ borderColor: ALURA.grayLight }}>
                        <td className="p-1.5">{r.activo}</td>
                        <td className="p-1.5 text-center">{r.ema === "A" || r.ema === "B" ? "❌" : r.ema === "C" ? "⚠️" : r.ema === "D" ? "✅" : "—"}</td>
                        <td className="p-1.5 text-center">{r.fda === "C" ? "❌" : r.fda === "H" ? "⚠️" : r.fda === "NO" ? "✅" : "—"}</td>
                        <td className="p-1.5 text-right font-medium">{r.mg.toLocaleString(undefined, { maximumFractionDigits: 0 })}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </ChartCard>
        </div>
      </div>
    </div>
  );
}

function KpiCard({ icon, label, value, hint, highlighted }) {
  return (
    <div
      className="rounded-lg p-3 border"
      style={{
        background: highlighted ? ALURA.bgTint : "white",
        borderColor: highlighted ? ALURA.coral : ALURA.grayLight,
        borderWidth: highlighted ? 2 : 1,
      }}
    >
      <div className="flex items-start justify-between mb-1">
        <span className="text-xs" style={{ color: ALURA.gray }}>{label}</span>
        <span className="text-lg">{icon}</span>
      </div>
      <div className="text-2xl font-bold" style={{ color: highlighted ? ALURA.vinotinto : ALURA.grayDark, fontFamily: "Nunito, sans-serif" }}>
        {value}
      </div>
      {hint && <div className="text-[9px] italic mt-1" style={{ color: ALURA.gray }}>{hint}</div>}
    </div>
  );
}

function ChartCard({ title, children }) {
  return (
    <div className="bg-white rounded-lg border overflow-hidden" style={{ borderColor: ALURA.grayLight }}>
      <div className="px-3 py-1.5 text-xs font-semibold text-white text-center" style={{ background: ALURA.vinotinto, fontFamily: "Nunito, sans-serif" }}>
        {title}
      </div>
      <div className="p-2">{children}</div>
    </div>
  );
}

function EmptyChart() {
  return <div className="h-[200px] flex items-center justify-center text-xs" style={{ color: ALURA.gray }}>Sin datos</div>;
}

// ===== UPLOAD TAB =====
function UploadTab({ consumos, productos, onFile, onReset }) {
  return (
    <div className="grid md:grid-cols-2 gap-4">
      <FileBox title="Consumos de alimento" desc="Excel con columnas: Fecha inicial, Fecha final, Granja, Etapa, Lote, Días de vida, Semana, Alimento, Consumo de alimento (Kg)" loaded={consumos.length} onFile={(f) => onFile(f, "consumos")} />
      <FileBox title="Catálogo de medicamentos" desc="Excel con columnas: Producto Comercial, Empresa, Ingrediente Activo, Concentración (%), Precio ($/kg). Opcional: Vía." loaded={productos.length} onFile={(f) => onFile(f, "productos")} />
      <div className="md:col-span-2 rounded-lg border p-4 bg-white" style={{ borderColor: ALURA.grayLight }}>
        <h3 className="font-semibold mb-2" style={{ color: ALURA.grayDark, fontFamily: "Nunito, sans-serif" }}>Estado actual</h3>
        <div className="text-sm space-y-1" style={{ color: ALURA.grayDark }}>
          <div>Consumos cargados: <span className="font-medium">{consumos.length}</span></div>
          <div>Productos en catálogo: <span className="font-medium">{productos.length}</span></div>
          <div style={{ color: ALURA.gray }} className="text-xs">Tus datos se guardan automáticamente entre sesiones.</div>
        </div>
        <button onClick={onReset} className="mt-3 px-3 py-1.5 text-sm rounded hover:opacity-90" style={{ background: "#FEE2E2", color: "#991B1B" }}>
          Borrar todo
        </button>
      </div>
    </div>
  );
}

function FileBox({ title, desc, loaded, onFile }) {
  return (
    <div className="rounded-lg border p-4 bg-white" style={{ borderColor: ALURA.grayLight }}>
      <h3 className="font-semibold mb-1" style={{ color: ALURA.grayDark, fontFamily: "Nunito, sans-serif" }}>{title}</h3>
      <p className="text-xs mb-3" style={{ color: ALURA.gray }}>{desc}</p>
      <label className="block">
        <span className="inline-flex items-center gap-2 px-3 py-2 text-white rounded text-sm cursor-pointer hover:opacity-90" style={{ background: ALURA.vinotinto }}>
          <Upload size={14} />
          {loaded > 0 ? "Reemplazar archivo" : "Subir archivo"}
        </span>
        <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => e.target.files[0] && onFile(e.target.files[0])} />
      </label>
      {loaded > 0 && <p className="text-xs mt-2" style={{ color: ALURA.vinotinto }}>{loaded} registros cargados</p>}
    </div>
  );
}

// ===== PRODUCTS TAB =====
function ProductsTab({ productos, setProductos }) {
  const [q, setQ] = useState("");
  const filtered = productos.filter((p) =>
    [p.producto, p.empresa, p.activo].some((s) => normalizeFeedName(s).includes(normalizeFeedName(q)))
  );
  const update = (id, field, value) => {
    setProductos(productos.map((p) => p.id === id ? { ...p, [field]: field === "concentracion" || field === "precio" ? num(value) : value } : p));
  };
  if (!productos.length) {
    return <div className="bg-white border rounded-lg p-8 text-center" style={{ borderColor: ALURA.grayLight, color: ALURA.gray }}>Sube el archivo de productos primero.</div>;
  }
  return (
    <div className="bg-white rounded-lg border" style={{ borderColor: ALURA.grayLight }}>
      <div className="p-3 border-b flex items-center gap-2" style={{ borderColor: ALURA.grayLight }}>
        <Search size={16} style={{ color: ALURA.gray }} />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar producto, empresa o activo..." className="flex-1 text-sm outline-none" />
        <span className="text-xs" style={{ color: ALURA.gray }}>{filtered.length} de {productos.length}</span>
      </div>
      <div className="overflow-auto max-h-[600px]">
        <table className="w-full text-sm">
          <thead className="sticky top-0" style={{ background: ALURA.bgSoft }}>
            <tr>
              <th className="text-left p-2 font-semibold" style={{ color: ALURA.grayDark }}>Producto</th>
              <th className="text-left p-2 font-semibold" style={{ color: ALURA.grayDark }}>Empresa</th>
              <th className="text-left p-2 font-semibold" style={{ color: ALURA.grayDark }}>Ing. activo</th>
              <th className="text-right p-2 font-semibold" style={{ color: ALURA.grayDark }}>Conc. (%)</th>
              <th className="text-right p-2 font-semibold" style={{ color: ALURA.grayDark }}>Precio ($/kg)</th>
              <th className="text-left p-2 font-semibold" style={{ color: ALURA.grayDark }}>Vía</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((p) => (
              <tr key={p.id} className="border-t" style={{ borderColor: ALURA.grayLight }}>
                <td className="p-2">{p.producto}</td>
                <td className="p-2" style={{ color: ALURA.gray }}>{p.empresa}</td>
                <td className="p-2" style={{ color: ALURA.gray }}>{p.activo}</td>
                <td className="p-2">
                  <input type="number" step="0.01" value={p.concentracion} onChange={(e) => update(p.id, "concentracion", e.target.value)} className="w-20 text-right border rounded px-1 py-0.5" style={{ borderColor: ALURA.grayLight }} />
                </td>
                <td className="p-2">
                  <input type="number" step="0.01" value={p.precio} onChange={(e) => update(p.id, "precio", e.target.value)} className="w-24 text-right border rounded px-1 py-0.5" style={{ borderColor: ALURA.grayLight }} />
                </td>
                <td className="p-2">
                  <select value={p.via || "Alimento"} onChange={(e) => update(p.id, "via", e.target.value)} className="text-sm border rounded px-1 py-0.5" style={{ borderColor: ALURA.grayLight }}>
                    {VIAS.map((v) => <option key={v} value={v}>{v}</option>)}
                  </select>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ===== RECIPES TAB =====
function RecipesTab({ loteAlimentos, productos, recetas, setRecetas, feedAliases, setFeedAliases }) {
  const [selectedKey, setSelectedKey] = useState(null);
  const [q, setQ] = useState("");
  const [editingAlias, setEditingAlias] = useState(null);

  const filtered = loteAlimentos.filter((la) =>
    normalizeFeedName(`${la.lote} ${la.alimentoDisplay} ${la.granjas} ${la.etapas}`).includes(normalizeFeedName(q))
  );

  const selected = loteAlimentos.find((la) => la.key === selectedKey);
  const currentRecipe = selected ? (recetas[selected.key] || []) : [];

  const addItem = () => {
    if (!selected || !productos.length) return;
    setRecetas({ ...recetas, [selected.key]: [...currentRecipe, { productId: productos[0].id, dosis_g_ton: 0 }] });
  };

  const updateItem = (idx, field, value) => {
    const updated = currentRecipe.map((it, i) => i === idx ? { ...it, [field]: field === "dosis_g_ton" ? num(value) : value } : it);
    setRecetas({ ...recetas, [selected.key]: updated });
  };

  const removeItem = (idx) => {
    const updated = currentRecipe.filter((_, i) => i !== idx);
    setRecetas({ ...recetas, [selected.key]: updated });
  };

  const copyFrom = (sourceKey) => {
    const source = recetas[sourceKey];
    if (!source || !selected) return;
    setRecetas({ ...recetas, [selected.key]: source.map((i) => ({ ...i })) });
  };

  const recipeCount = (key) => (recetas[key] || []).length;
  const saveAlias = (oldKey, newDisplay) => { setFeedAliases({ ...feedAliases, [oldKey]: newDisplay }); setEditingAlias(null); };

  if (!loteAlimentos.length) return <div className="bg-white border rounded-lg p-8 text-center" style={{ borderColor: ALURA.grayLight, color: ALURA.gray }}>Sube los consumos primero.</div>;
  if (!productos.length) return <div className="bg-white border rounded-lg p-8 text-center" style={{ borderColor: ALURA.grayLight, color: ALURA.gray }}>Sube el catálogo de productos primero.</div>;

  const sourcesAvailable = loteAlimentos.filter((la) => la.key !== selectedKey && recipeCount(la.key) > 0);

  return (
    <div className="grid md:grid-cols-5 gap-4">
      <div className="md:col-span-2 bg-white rounded-lg border" style={{ borderColor: ALURA.grayLight }}>
        <div className="p-3 border-b flex items-center gap-2" style={{ borderColor: ALURA.grayLight }}>
          <Search size={16} style={{ color: ALURA.gray }} />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar lote o alimento..." className="flex-1 text-sm outline-none" />
        </div>
        <div className="overflow-auto max-h-[600px]">
          {filtered.map((la) => {
            const count = recipeCount(la.key);
            const active = la.key === selectedKey;
            return (
              <button key={la.key} onClick={() => setSelectedKey(la.key)} className="w-full text-left p-3 border-b hover:bg-gray-50" style={{ borderColor: ALURA.grayLight, background: active ? ALURA.bgTint : "white" }}>
                <div className="flex justify-between items-start gap-2">
                  <div className="flex-1 min-w-0">
                    <div className="text-xs" style={{ color: ALURA.gray }}>Lote {la.lote} · {la.etapas}</div>
                    <div className="text-sm font-medium truncate" style={{ color: ALURA.grayDark }}>{la.alimentoDisplay}</div>
                    <div className="text-xs" style={{ color: ALURA.gray }}>{la.totalKg.toLocaleString()} kg · {la.granjas}</div>
                  </div>
                  <span className="text-xs px-2 py-0.5 rounded" style={{ background: count > 0 ? "#D1FAE5" : ALURA.bgSoft, color: count > 0 ? "#065F46" : ALURA.gray }}>
                    {count > 0 ? `${count} producto${count > 1 ? "s" : ""}` : "sin receta"}
                  </span>
                </div>
              </button>
            );
          })}
        </div>
      </div>

      <div className="md:col-span-3 bg-white rounded-lg border p-4" style={{ borderColor: ALURA.grayLight }}>
        {!selected ? (
          <div className="text-center py-12" style={{ color: ALURA.gray }}>Selecciona un lote/alimento de la izquierda.</div>
        ) : (
          <>
            <div className="mb-3">
              <div className="text-xs" style={{ color: ALURA.gray }}>Lote {selected.lote} · {selected.etapas} · {selected.granjas}</div>
              {editingAlias === selected.feedKey ? (
                <input defaultValue={selected.alimentoDisplay} onBlur={(e) => saveAlias(selected.feedKey, e.target.value)} onKeyDown={(e) => e.key === "Enter" && saveAlias(selected.feedKey, e.target.value)} className="w-full border rounded px-2 py-1 text-sm mt-1" style={{ borderColor: ALURA.grayLight }} autoFocus />
              ) : (
                <h3 className="text-lg font-semibold cursor-pointer hover:bg-yellow-50 inline-block px-1" style={{ color: ALURA.grayDark, fontFamily: "Nunito, sans-serif" }} onClick={() => setEditingAlias(selected.feedKey)} title="Click para renombrar">
                  {selected.alimentoDisplay}
                </h3>
              )}
              <div className="text-sm mt-1" style={{ color: ALURA.grayDark }}>Total consumido: {selected.totalKg.toLocaleString()} kg</div>
            </div>

            <div className="border-t pt-3" style={{ borderColor: ALURA.grayLight }}>
              <div className="flex items-center justify-between mb-2">
                <h4 className="font-medium text-sm" style={{ color: ALURA.grayDark }}>Productos en esta receta</h4>
                <div className="flex gap-2">
                  {sourcesAvailable.length > 0 && (
                    <select onChange={(e) => e.target.value && copyFrom(e.target.value)} value="" className="text-xs border rounded px-2 py-1" style={{ borderColor: ALURA.grayLight }}>
                      <option value="">📋 Copiar receta de...</option>
                      {sourcesAvailable.map((la) => (<option key={la.key} value={la.key}>Lote {la.lote} · {la.alimentoDisplay}</option>))}
                    </select>
                  )}
                  <button onClick={addItem} className="text-xs px-2 py-1 text-white rounded flex items-center gap-1 hover:opacity-90" style={{ background: ALURA.vinotinto }}>
                    <Plus size={12} /> Agregar producto
                  </button>
                </div>
              </div>

              {currentRecipe.length === 0 ? (
                <div className="text-sm text-center py-6 rounded" style={{ color: ALURA.gray, background: ALURA.bgSoft }}>
                  No hay productos en esta receta aún.
                </div>
              ) : (
                <div className="space-y-2">
                  {currentRecipe.map((item, idx) => {
                    const p = productos.find((pp) => pp.id === item.productId);
                    const kgProductoTotal = (selected.totalKg * num(item.dosis_g_ton)) / 1_000_000;
                    const kgActivoTotal = p ? kgProductoTotal * (p.concentracion / 100) : 0;
                    const costoTotal = p ? kgProductoTotal * p.precio : 0;
                    return (
                      <div key={idx} className="border rounded p-2" style={{ borderColor: ALURA.grayLight, background: ALURA.bgSoft }}>
                        <div className="flex gap-2 items-start">
                          <select value={item.productId} onChange={(e) => updateItem(idx, "productId", e.target.value)} className="flex-1 text-sm border rounded px-2 py-1 bg-white" style={{ borderColor: ALURA.grayLight }}>
                            {productos.map((p) => (<option key={p.id} value={p.id}>{p.producto} — {p.activo} ({p.concentracion}%)</option>))}
                          </select>
                          <div className="flex items-center gap-1">
                            <input type="number" step="1" value={item.dosis_g_ton} onChange={(e) => updateItem(idx, "dosis_g_ton", e.target.value)} className="w-20 text-right text-sm border rounded px-2 py-1" style={{ borderColor: ALURA.grayLight }} placeholder="0" />
                            <span className="text-xs" style={{ color: ALURA.gray }}>g/ton</span>
                          </div>
                          <button onClick={() => removeItem(idx)} className="hover:bg-red-50 p-1 rounded" style={{ color: ALURA.vinotinto }}>
                            <Trash2 size={14} />
                          </button>
                        </div>
                        {p && num(item.dosis_g_ton) > 0 && (
                          <div className="text-xs mt-1 grid grid-cols-3 gap-1" style={{ color: ALURA.gray }}>
                            <span>Producto: <b style={{ color: ALURA.grayDark }}>{kgProductoTotal.toFixed(3)} kg</b></span>
                            <span>Activo: <b style={{ color: ALURA.grayDark }}>{kgActivoTotal.toFixed(3)} kg</b></span>
                            <span>Costo: <b style={{ color: ALURA.grayDark }}>${costoTotal.toLocaleString(undefined, { maximumFractionDigits: 0 })}</b></span>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ===== CLASSIFICATION TAB =====
function ClassificationTab({ productos, activosOverride, setActivosOverride }) {
  const activos = useMemo(() => {
    const set = new Set();
    productos.forEach((p) => set.add(p.activo));
    return Array.from(set).filter(Boolean).sort();
  }, [productos]);

  if (!productos.length) {
    return <div className="bg-white border rounded-lg p-8 text-center" style={{ borderColor: ALURA.grayLight, color: ALURA.gray }}>Sube el catálogo de productos primero.</div>;
  }

  const updateOverride = (activo, field, value) => {
    const k = normalizeFeedName(activo);
    setActivosOverride({ ...activosOverride, [k]: { ...(activosOverride[k] || inferEMA_FDA(activo)), [field]: value } });
  };

  return (
    <div className="bg-white rounded-lg border" style={{ borderColor: ALURA.grayLight }}>
      <div className="p-4 border-b" style={{ borderColor: ALURA.grayLight }}>
        <h3 className="font-semibold" style={{ color: ALURA.grayDark, fontFamily: "Nunito, sans-serif" }}>Clasificación EMA y FDA por ingrediente activo</h3>
        <p className="text-xs mt-1" style={{ color: ALURA.gray }}>
          Hay valores pre-cargados para los activos más comunes. Puedes editarlos según las referencias más recientes.
        </p>
      </div>
      <div className="overflow-auto max-h-[600px]">
        <table className="w-full text-sm">
          <thead className="sticky top-0" style={{ background: ALURA.bgSoft }}>
            <tr>
              <th className="text-left p-2 font-semibold" style={{ color: ALURA.grayDark }}>Ingrediente activo</th>
              <th className="text-left p-2 font-semibold" style={{ color: ALURA.grayDark }}>EMA</th>
              <th className="text-left p-2 font-semibold" style={{ color: ALURA.grayDark }}>FDA</th>
            </tr>
          </thead>
          <tbody>
            {activos.map((a) => {
              const k = normalizeFeedName(a);
              const current = activosOverride[k] || inferEMA_FDA(a);
              return (
                <tr key={a} className="border-t" style={{ borderColor: ALURA.grayLight }}>
                  <td className="p-2">{a}</td>
                  <td className="p-2">
                    <select value={current.ema} onChange={(e) => updateOverride(a, "ema", e.target.value)} className="text-sm border rounded px-2 py-1" style={{ borderColor: ALURA.grayLight }}>
                      {Object.entries(EMA_LEVELS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                    </select>
                  </td>
                  <td className="p-2">
                    <select value={current.fda} onChange={(e) => updateOverride(a, "fda", e.target.value)} className="text-sm border rounded px-2 py-1" style={{ borderColor: ALURA.grayLight }}>
                      {Object.entries(FDA_LEVELS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                    </select>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
