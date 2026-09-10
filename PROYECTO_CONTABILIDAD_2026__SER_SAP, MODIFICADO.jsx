import React, { useState, useEffect, useMemo } from "react";
import {
  LayoutDashboard, ListTree, BookOpen, Scale, FileBarChart2,
  Users, Package, Plus, Trash2, ChevronDown, ChevronRight,
  AlertCircle, RefreshCw, Receipt, FileDown, Printer, UserCog
} from "lucide-react";
import * as XLSX from "xlsx";

const uid = () => (crypto.randomUUID ? crypto.randomUUID() : String(Math.random()).slice(2));

const DEFAULT_COMPANY = { name: "Mi Empresa", currency: "CRC", ivaRate: 13 };

const DEFAULT_ACCOUNTS = [
  { id: uid(), code: "1010", name: "Caja", type: "Activo" },
  { id: uid(), code: "1020", name: "Bancos", type: "Activo" },
  { id: uid(), code: "1030", name: "Cuentas por Cobrar", type: "Activo" },
  { id: uid(), code: "1040", name: "Inventario", type: "Activo" },
  { id: uid(), code: "1050", name: "Mobiliario y Equipo", type: "Activo" },
  { id: uid(), code: "2010", name: "Cuentas por Pagar", type: "Pasivo" },
  { id: uid(), code: "2020", name: "IVA por Pagar", type: "Pasivo" },
  { id: uid(), code: "2030", name: "Préstamos por Pagar", type: "Pasivo" },
  { id: uid(), code: "3010", name: "Capital Social", type: "Patrimonio" },
  { id: uid(), code: "3020", name: "Utilidades Retenidas", type: "Patrimonio" },
  { id: uid(), code: "4010", name: "Ventas", type: "Ingreso" },
  { id: uid(), code: "4020", name: "Otros Ingresos", type: "Ingreso" },
  { id: uid(), code: "5010", name: "Costo de Ventas", type: "Gasto" },
  { id: uid(), code: "5020", name: "Gastos Operativos", type: "Gasto" },
  { id: uid(), code: "5030", name: "Gastos Administrativos", type: "Gasto" },
];

const TYPES = ["Activo", "Pasivo", "Patrimonio", "Ingreso", "Gasto"];
const isDebitNormal = (type) => type === "Activo" || type === "Gasto";

function fmt(n, currency) {
  const v = Number.isFinite(n) ? n : 0;
  try {
    return new Intl.NumberFormat("es-CR", { style: "currency", currency: currency || "CRC", maximumFractionDigits: 2 }).format(v);
  } catch {
    return v.toFixed(2);
  }
}

async function loadKey(key, fallback) {
  try {
    const res = await window.storage.get(key, false);
    return res && res.value ? JSON.parse(res.value) : fallback;
  } catch {
    return fallback;
  }
}
async function saveKey(key, value) {
  try {
    const res = await window.storage.set(key, JSON.stringify(value), false);
    if (!res) console.error("No se pudo guardar", key);
  } catch (e) {
    console.error("Error guardando", key, e);
  }
}

/* ---------- Períodos (informes) ---------- */
function toISO(d) { return d.toISOString().slice(0, 10); }
function monthRange(d = new Date()) {
  return { start: toISO(new Date(d.getFullYear(), d.getMonth(), 1)), end: toISO(new Date(d.getFullYear(), d.getMonth() + 1, 0)) };
}
function yearRange(d = new Date()) {
  return { start: toISO(new Date(d.getFullYear(), 0, 1)), end: toISO(new Date(d.getFullYear(), 11, 31)) };
}
function inPeriod(dateStr, period) {
  if (!period) return true;
  if (period.start && dateStr < period.start) return false;
  if (period.end && dateStr > period.end) return false;
  return true;
}
function computeBalances(accounts, journalEntries) {
  const map = {};
  accounts.forEach((a) => { map[a.id] = { debit: 0, credit: 0 }; });
  journalEntries.forEach((e) => e.lines.forEach((l) => {
    if (!map[l.accountId]) return;
    map[l.accountId].debit += Number(l.debit) || 0;
    map[l.accountId].credit += Number(l.credit) || 0;
  }));
  const out = {};
  accounts.forEach((a) => {
    const { debit, credit } = map[a.id];
    out[a.id] = { debit, credit, balance: isDebitNormal(a.type) ? debit - credit : credit - debit };
  });
  return out;
}
function computeTotalsByType(accounts, balances) {
  const t = { Activo: 0, Pasivo: 0, Patrimonio: 0, Ingreso: 0, Gasto: 0 };
  accounts.forEach((a) => { t[a.type] += balances[a.id]?.balance || 0; });
  return t;
}

/* ---------- Exportación ---------- */
function exportExcel(filename, sheetName, rows) {
  try {
    if (!rows || rows.length === 0) { alert("No hay datos para exportar en el período seleccionado."); return; }
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, (sheetName || "Hoja1").slice(0, 31));
    XLSX.writeFile(wb, filename);
  } catch (e) {
    console.error("Error exportando a Excel", e);
    alert("No se pudo exportar a Excel.");
  }
}

function PeriodBar({ period, onChange }) {
  return (
    <div className="row period-bar no-print" style={{ flexWrap: "wrap", gap: 8 }}>
      <button className="btn" onClick={() => onChange({ start: "", end: "" })}>Todo</button>
      <button className="btn" onClick={() => onChange(monthRange())}>Este mes</button>
      <button className="btn" onClick={() => onChange(yearRange())}>Este año</button>
      <input className="input" type="date" style={{ maxWidth: 148 }} value={period.start}
        onChange={(e) => onChange({ ...period, start: e.target.value })} />
      <span style={{ color: "var(--muted)", fontSize: 12.5 }}>a</span>
      <input className="input" type="date" style={{ maxWidth: 148 }} value={period.end}
        onChange={(e) => onChange({ ...period, end: e.target.value })} />
    </div>
  );
}

function ExportButtons({ onExcel }) {
  return (
    <div className="row no-print" style={{ gap: 8 }}>
      <button className="btn" onClick={onExcel}><FileDown size={14} /> Excel</button>
      <button className="btn" onClick={() => window.print()}><Printer size={14} /> PDF</button>
    </div>
  );
}

function PrintHeader({ company, title, period }) {
  return (
    <div className="print-only">
      <div style={{ fontFamily: "'IBM Plex Serif',serif", fontWeight: 600, fontSize: 16 }}>{company?.name}</div>
      <div style={{ fontSize: 12.5, color: "#555", marginBottom: 10 }}>
        {title} · {period && (period.start || period.end) ? `Período: ${period.start || "inicio"} a ${period.end || "hoy"}` : "Todo el historial"}
      </div>
    </div>
  );
}

/* ---------- Roles y permisos ---------- */
const ROLES = {
  admin: {
    label: "Administrador",
    nav: ["dashboard", "accounts", "journal", "ledger", "trial", "statements", "billing", "inventory"],
    canEditAccounts: true, canDeleteAccounts: true, canEditJournal: true, canDeleteJournal: true,
    canEditCompany: true, canEditBilling: true, canEditInventory: true, canReset: true,
  },
  contador: {
    label: "Contador",
    nav: ["dashboard", "accounts", "journal", "ledger", "trial", "statements"],
    canEditAccounts: true, canDeleteAccounts: false, canEditJournal: true, canDeleteJournal: true,
    canEditCompany: false, canEditBilling: false, canEditInventory: false, canReset: false,
  },
  ventas: {
    label: "Facturación / Ventas",
    nav: ["billing", "inventory"],
    canEditAccounts: false, canDeleteAccounts: false, canEditJournal: false, canDeleteJournal: false,
    canEditCompany: false, canEditBilling: true, canEditInventory: true, canReset: false,
  },
  readonly: {
    label: "Solo lectura",
    nav: ["dashboard", "ledger", "trial", "statements"],
    canEditAccounts: false, canDeleteAccounts: false, canEditJournal: false, canDeleteJournal: false,
    canEditCompany: false, canEditBilling: false, canEditInventory: false, canReset: false,
  },
};

const NAV = [
  { id: "dashboard", label: "Resumen", icon: LayoutDashboard },
  { id: "accounts", label: "Plan de cuentas", icon: ListTree },
  { id: "journal", label: "Diario", icon: BookOpen },
  { id: "ledger", label: "Mayor", icon: Scale },
  { id: "trial", label: "Balance de comprobación", icon: Scale },
  { id: "statements", label: "Estados financieros", icon: FileBarChart2 },
  { id: "billing", label: "Clientes y facturación", icon: Receipt },
  { id: "inventory", label: "Inventario", icon: Package },
];

export default function App() {
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("dashboard");
  const [company, setCompany] = useState(DEFAULT_COMPANY);
  const [accounts, setAccounts] = useState([]);
  const [journal, setJournal] = useState([]);
  const [customers, setCustomers] = useState([]);
  const [invoices, setInvoices] = useState([]);
  const [inventory, setInventory] = useState([]);
  const [role, setRole] = useState("admin");

  useEffect(() => {
    (async () => {
      const [comp, accs, jrn, custs, invs, inv, rl] = await Promise.all([
        loadKey("company", DEFAULT_COMPANY),
        loadKey("accounts", DEFAULT_ACCOUNTS),
        loadKey("journal", []),
        loadKey("customers", []),
        loadKey("invoices", []),
        loadKey("inventory", []),
        loadKey("role", "admin"),
      ]);
      setCompany(comp); setAccounts(accs); setJournal(jrn);
      setCustomers(custs); setInvoices(invs); setInventory(inv);
      setRole(ROLES[rl] ? rl : "admin");
      setLoading(false);
    })();
  }, []);

  const update = {
    company: (v) => { setCompany(v); saveKey("company", v); },
    accounts: (v) => { setAccounts(v); saveKey("accounts", v); },
    journal: (v) => { setJournal(v); saveKey("journal", v); },
    customers: (v) => { setCustomers(v); saveKey("customers", v); },
    invoices: (v) => { setInvoices(v); saveKey("invoices", v); },
    inventory: (v) => { setInventory(v); saveKey("inventory", v); },
    role: (v) => { setRole(v); saveKey("role", v); },
  };

  const perm = ROLES[role] || ROLES.admin;

  useEffect(() => {
    if (!loading && !perm.nav.includes(tab)) setTab(perm.nav[0]);
  }, [role, loading]); // eslint-disable-line react-hooks/exhaustive-deps

  async function resetAll() {
    if (!confirm("Esto borrará todos los datos guardados en este dispositivo. ¿Continuar?")) return;
    update.company(DEFAULT_COMPANY);
    update.accounts(DEFAULT_ACCOUNTS);
    update.journal([]);
    update.customers([]);
    update.invoices([]);
    update.inventory([]);
  }

  const balances = useMemo(() => {
    const map = {};
    accounts.forEach((a) => { map[a.id] = { debit: 0, credit: 0 }; });
    journal.forEach((e) => e.lines.forEach((l) => {
      if (!map[l.accountId]) return;
      map[l.accountId].debit += Number(l.debit) || 0;
      map[l.accountId].credit += Number(l.credit) || 0;
    }));
    const out = {};
    accounts.forEach((a) => {
      const { debit, credit } = map[a.id];
      const balance = isDebitNormal(a.type) ? debit - credit : credit - debit;
      out[a.id] = { debit, credit, balance };
    });
    return out;
  }, [accounts, journal]);

  const totalsByType = useMemo(() => {
    const t = { Activo: 0, Pasivo: 0, Patrimonio: 0, Ingreso: 0, Gasto: 0 };
    accounts.forEach((a) => { t[a.type] += balances[a.id]?.balance || 0; });
    return t;
  }, [accounts, balances]);

  const netIncome = totalsByType.Ingreso - totalsByType.Gasto;

  if (loading) {
    return (
      <div style={{ padding: "3rem", textAlign: "center", color: "var(--muted)" }}>
        Cargando datos contables…
      </div>
    );
  }

  return (
    <div className="app-root">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Serif:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
        .app-root {
          --bg:#F6F7F5; --surface:#FFFFFF; --ink:#1B2430; --muted:#6B7688;
          --accent:#2F6F5E; --accent-bg:#E7F0ED; --negative:#B5453A; --negative-bg:#F6E9E7;
          --violet:#5F5480; --amber:#9A6B24; --border:#DEE2E7;
          font-family:'IBM Plex Sans',sans-serif; color:var(--ink); background:var(--bg);
          display:flex; min-height:600px; border:1px solid var(--border); border-radius:12px; overflow:hidden;
        }
        .app-root * { box-sizing:border-box; }
        .serif { font-family:'IBM Plex Serif',serif; }
        .mono { font-family:'IBM Plex Mono',monospace; font-variant-numeric:tabular-nums; }
        .sidebar { width:210px; background:var(--surface); border-right:1px solid var(--border); padding:20px 12px; flex-shrink:0; }
        .brand { font-family:'IBM Plex Serif',serif; font-weight:600; font-size:16px; padding:0 8px 16px; border-bottom:1px solid var(--border); margin-bottom:12px; }
        .navitem { display:flex; align-items:center; gap:10px; width:100%; text-align:left; padding:9px 10px; border-radius:6px; border:none; background:transparent; color:var(--ink); font-size:13.5px; cursor:pointer; margin-bottom:2px; }
        .navitem:hover { background:var(--bg); }
        .navitem.active { background:var(--accent-bg); color:var(--accent); font-weight:500; }
        .main { flex:1; padding:24px 28px; overflow-x:auto; }
        .h1 { font-family:'IBM Plex Serif',serif; font-weight:600; font-size:22px; margin:0 0 4px; }
        .sub { color:var(--muted); font-size:13px; margin:0 0 20px; }
        .card { background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:16px 18px; }
        .grid { display:grid; gap:12px; }
        .metric-label { color:var(--muted); font-size:12.5px; margin:0 0 6px; }
        .metric-value { font-family:'IBM Plex Mono',monospace; font-size:21px; font-weight:500; margin:0; }
        table.tbl { width:100%; border-collapse:collapse; font-size:13.5px; }
        table.tbl th { text-align:left; font-weight:500; color:var(--muted); font-size:12px; padding:6px 8px; border-bottom:1px solid var(--border); }
        table.tbl td { padding:7px 8px; border-bottom:1px solid var(--border); }
        table.tbl tr:last-child td { border-bottom:none; }
        .num { text-align:right; font-family:'IBM Plex Mono',monospace; }
        .btn { border:1px solid var(--border); background:var(--surface); color:var(--ink); padding:7px 12px; border-radius:6px; font-size:13px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; }
        .btn:hover { background:var(--bg); }
        .btn-primary { background:var(--accent); border-color:var(--accent); color:#fff; }
        .btn-primary:hover { opacity:0.92; }
        .btn-danger { color:var(--negative); }
        .btn-icon { padding:6px; }
        .input, select.input { width:100%; padding:7px 9px; border:1px solid var(--border); border-radius:6px; font-size:13.5px; font-family:inherit; background:var(--surface); color:var(--ink); }
        .row { display:flex; gap:8px; align-items:center; }
        .badge { display:inline-block; padding:2px 8px; border-radius:20px; font-size:11.5px; font-weight:500; }
        .badge-Activo { background:var(--accent-bg); color:var(--accent); }
        .badge-Pasivo { background:var(--negative-bg); color:var(--negative); }
        .badge-Patrimonio { background:#EEEBF3; color:var(--violet); }
        .badge-Ingreso { background:#E4F0EC; color:#1F4D40; }
        .badge-Gasto { background:#F5EEE0; color:var(--amber); }
        .section-title { font-size:14.5px; font-weight:600; margin:24px 0 10px; }
        .section-title:first-child { margin-top:0; }
        .warn { display:flex; gap:8px; align-items:flex-start; background:#FBF3EC; border:1px solid #E9D3B6; color:#7A4D14; padding:10px 12px; border-radius:8px; font-size:13px; margin-bottom:14px; }
        .empty { color:var(--muted); font-size:13.5px; padding:12px 4px; }
        .role-switch { margin-bottom:14px; padding-bottom:14px; border-bottom:1px solid var(--border); }
        .print-only { display:none; }

        @media print {
          .no-print, .sidebar { display:none !important; }
          .app-root { border:none; min-height:0; display:block; background:#fff; }
          .main { padding:0; overflow:visible; }
          .print-only { display:block; }
        }

        @media (max-width: 720px) {
          .app-root { flex-direction:column; min-height:0; }
          .sidebar { width:100%; display:flex; flex-wrap:wrap; align-items:flex-start; gap:4px; padding:12px; border-right:none; border-bottom:1px solid var(--border); }
          .brand { width:100%; border-bottom:none; margin-bottom:8px; padding-bottom:0; }
          .role-switch { width:100%; flex:1 1 100%; border-bottom:none; margin-bottom:8px; padding-bottom:0; }
          .navitem { width:auto; flex:1 1 auto; justify-content:center; }
          .main { padding:16px; }
          .row { flex-wrap:wrap; }
          table.tbl { display:block; overflow-x:auto; white-space:nowrap; }
        }
      `}</style>

      <div className="sidebar">
        <div className="brand">{company.name}</div>

        <div className="role-switch no-print">
          <label className="metric-label" style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <UserCog size={13} /> Rol activo
          </label>
          <select className="input" value={role} onChange={(e) => update.role(e.target.value)}>
            {Object.entries(ROLES).map(([key, r]) => <option key={key} value={key}>{r.label}</option>)}
          </select>
        </div>

        {NAV.filter((n) => perm.nav.includes(n.id)).map((n) => (
          <button key={n.id} className={`navitem ${tab === n.id ? "active" : ""}`} onClick={() => setTab(n.id)}>
            <n.icon size={16} /> {n.label}
          </button>
        ))}
        {perm.canReset && (
          <button className="navitem no-print" style={{ marginTop: 16, color: "var(--muted)" }} onClick={resetAll}>
            <RefreshCw size={15} /> Reiniciar datos
          </button>
        )}
      </div>

      <div className="main">
        {tab === "dashboard" && (
          <Dashboard company={company} totalsByType={totalsByType} netIncome={netIncome}
            invoices={invoices} inventory={inventory} onCompany={update.company} perm={perm} />
        )}
        {tab === "accounts" && (
          <AccountsTab accounts={accounts} onChange={update.accounts} balances={balances} currency={company.currency} perm={perm} />
        )}
        {tab === "journal" && (
          <JournalTab accounts={accounts} journal={journal} onChange={update.journal} currency={company.currency} company={company} perm={perm} />
        )}
        {tab === "ledger" && (
          <LedgerTab accounts={accounts} journal={journal} currency={company.currency} company={company} />
        )}
        {tab === "trial" && (
          <TrialBalanceTab accounts={accounts} journal={journal} currency={company.currency} company={company} />
        )}
        {tab === "statements" && (
          <StatementsTab accounts={accounts} journal={journal} currency={company.currency} company={company} />
        )}
        {tab === "billing" && (
          <BillingTab company={company} accounts={accounts} customers={customers} invoices={invoices}
            journal={journal} inventory={inventory} onCustomers={update.customers} onInvoices={update.invoices}
            onJournal={update.journal} onInventory={update.inventory} perm={perm} />
        )}
        {tab === "inventory" && (
          <InventoryTab inventory={inventory} onChange={update.inventory} currency={company.currency} perm={perm} />
        )}
      </div>
    </div>
  );
}

function Dashboard({ company, totalsByType, netIncome, invoices, inventory, onCompany, perm }) {
  const pending = invoices.filter((i) => i.status !== "Pagada").length;
  const invValue = inventory.reduce((s, i) => s + (Number(i.qty) || 0) * (Number(i.cost) || 0), 0);
  return (
    <div>
      <h1 className="h1">Resumen</h1>
      <p className="sub">Vista general del estado financiero de {company.name}.</p>

      <div className="grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(160px,1fr))", marginBottom: 24 }}>
        <Metric label="Activos" value={fmt(totalsByType.Activo, company.currency)} />
        <Metric label="Pasivos" value={fmt(totalsByType.Pasivo, company.currency)} />
        <Metric label="Patrimonio" value={fmt(totalsByType.Patrimonio, company.currency)} />
        <Metric label="Utilidad del período" value={fmt(netIncome, company.currency)} />
        <Metric label="Facturas pendientes" value={String(pending)} />
        <Metric label="Valor de inventario" value={fmt(invValue, company.currency)} />
      </div>

      <div className="section-title">Datos de la empresa</div>
      <div className="card" style={{ maxWidth: 420 }}>
        <div style={{ marginBottom: 10 }}>
          <label className="metric-label">Nombre</label>
          <input className="input" disabled={!perm.canEditCompany} value={company.name} onChange={(e) => onCompany({ ...company, name: e.target.value })} />
        </div>
        <div className="row">
          <div style={{ flex: 1 }}>
            <label className="metric-label">Moneda (código)</label>
            <input className="input" disabled={!perm.canEditCompany} value={company.currency} onChange={(e) => onCompany({ ...company, currency: e.target.value.toUpperCase() })} />
          </div>
          <div style={{ flex: 1 }}>
            <label className="metric-label">IVA por defecto (%)</label>
            <input className="input" type="number" disabled={!perm.canEditCompany} value={company.ivaRate}
              onChange={(e) => onCompany({ ...company, ivaRate: Number(e.target.value) })} />
          </div>
        </div>
        {!perm.canEditCompany && <p className="empty" style={{ padding: "8px 0 0" }}>Tu rol no permite editar los datos de la empresa.</p>}
      </div>
    </div>
  );
}

function Metric({ label, value }) {
  return (
    <div className="card">
      <p className="metric-label">{label}</p>
      <p className="metric-value">{value}</p>
    </div>
  );
}

function AccountsTab({ accounts, onChange, balances, currency, perm }) {
  const [form, setForm] = useState({ code: "", name: "", type: "Activo" });
  function add() {
    if (!form.code.trim() || !form.name.trim()) return;
    onChange([...accounts, { id: uid(), code: form.code.trim(), name: form.name.trim(), type: form.type }]);
    setForm({ code: "", name: "", type: "Activo" });
  }
  function remove(id) {
    onChange(accounts.filter((a) => a.id !== id));
  }
  return (
    <div>
      <h1 className="h1">Plan de cuentas</h1>
      <p className="sub">Catálogo de cuentas contables agrupadas por tipo.</p>

      {perm.canEditAccounts && (
        <div className="card no-print" style={{ marginBottom: 20 }}>
          <div className="row">
            <input className="input" placeholder="Código, p. ej. 1060" style={{ maxWidth: 120 }}
              value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} />
            <input className="input" placeholder="Nombre de la cuenta"
              value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
            <select className="input" style={{ maxWidth: 150 }} value={form.type}
              onChange={(e) => setForm({ ...form, type: e.target.value })}>
              {TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
            <button className="btn btn-primary" onClick={add}><Plus size={14} /> Agregar</button>
          </div>
        </div>
      )}

      <table className="tbl">
        <thead><tr><th>Código</th><th>Nombre</th><th>Tipo</th><th className="num">Saldo</th><th className="no-print"></th></tr></thead>
        <tbody>
          {[...accounts].sort((a, b) => a.code.localeCompare(b.code)).map((a) => (
            <tr key={a.id}>
              <td className="mono">{a.code}</td>
              <td>{a.name}</td>
              <td><span className={`badge badge-${a.type}`}>{a.type}</span></td>
              <td className="num">{fmt(balances[a.id]?.balance, currency)}</td>
              <td className="no-print">
                {perm.canDeleteAccounts && (
                  <button className="btn btn-icon btn-danger" onClick={() => remove(a.id)}><Trash2 size={14} /></button>
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function emptyLine() { return { accountId: "", debit: "", credit: "" }; }

function JournalTab({ accounts, journal, onChange, currency, company, perm }) {
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [desc, setDesc] = useState("");
  const [lines, setLines] = useState([emptyLine(), emptyLine()]);
  const [error, setError] = useState("");
  const [expanded, setExpanded] = useState(null);
  const [period, setPeriod] = useState({ start: "", end: "" });

  const totalDebit = lines.reduce((s, l) => s + (Number(l.debit) || 0), 0);
  const totalCredit = lines.reduce((s, l) => s + (Number(l.credit) || 0), 0);

  function updateLine(i, patch) {
    const next = [...lines];
    next[i] = { ...next[i], ...patch };
    setLines(next);
  }
  function addLine() { setLines([...lines, emptyLine()]); }
  function removeLine(i) { setLines(lines.filter((_, idx) => idx !== i)); }

  function submit() {
    const valid = lines.filter((l) => l.accountId && ((Number(l.debit) || 0) > 0 || (Number(l.credit) || 0) > 0));
    if (!desc.trim()) return setError("Agrega una descripción para el asiento.");
    if (valid.length < 2) return setError("Necesitas al menos dos líneas con cuenta y monto.");
    if (Math.abs(totalDebit - totalCredit) > 0.005) return setError("El total del debe debe ser igual al total del haber.");
    if (totalDebit === 0) return setError("El asiento no puede estar en cero.");
    setError("");
    const entry = {
      id: uid(), date, description: desc.trim(),
      lines: valid.map((l) => ({ accountId: l.accountId, debit: Number(l.debit) || 0, credit: Number(l.credit) || 0 })),
    };
    onChange([entry, ...journal]);
    setDesc(""); setLines([emptyLine(), emptyLine()]);
  }

  function removeEntry(id) { onChange(journal.filter((e) => e.id !== id)); }
  const accName = (id) => accounts.find((a) => a.id === id)?.name || "—";

  const filtered = useMemo(() => journal.filter((e) => inPeriod(e.date, period)), [journal, period]);

  function handleExportExcel() {
    const rows = [];
    filtered.forEach((e) => e.lines.forEach((l) => {
      rows.push({ Fecha: e.date, Descripción: e.description, Cuenta: accName(l.accountId), Debe: l.debit || 0, Haber: l.credit || 0 });
    }));
    exportExcel(`diario_${period.start || "inicio"}_${period.end || "hoy"}.xlsx`, "Diario", rows);
  }

  return (
    <div>
      <h1 className="h1">Diario general</h1>
      <p className="sub">Registra asientos de partida doble. El debe y el haber deben coincidir.</p>
      <PrintHeader company={company} title="Diario general" period={period} />

      {perm.canEditJournal && (
        <div className="card no-print" style={{ marginBottom: 20 }}>
          <div className="row" style={{ marginBottom: 10 }}>
            <input className="input" type="date" style={{ maxWidth: 160 }} value={date} onChange={(e) => setDate(e.target.value)} />
            <input className="input" placeholder="Descripción del asiento" value={desc} onChange={(e) => setDesc(e.target.value)} />
          </div>

          {lines.map((l, i) => (
            <div className="row" key={i} style={{ marginBottom: 6 }}>
              <select className="input" value={l.accountId} onChange={(e) => updateLine(i, { accountId: e.target.value })}>
                <option value="">Selecciona cuenta…</option>
                {[...accounts].sort((a, b) => a.code.localeCompare(b.code)).map((a) => (
                  <option key={a.id} value={a.id}>{a.code} · {a.name}</option>
                ))}
              </select>
              <input className="input" type="number" step="0.01" placeholder="Debe" style={{ maxWidth: 130 }}
                value={l.debit} onChange={(e) => updateLine(i, { debit: e.target.value, credit: e.target.value ? "" : l.credit })} />
              <input className="input" type="number" step="0.01" placeholder="Haber" style={{ maxWidth: 130 }}
                value={l.credit} onChange={(e) => updateLine(i, { credit: e.target.value, debit: e.target.value ? "" : l.debit })} />
              <button className="btn btn-icon btn-danger" onClick={() => removeLine(i)}><Trash2 size={14} /></button>
            </div>
          ))}

          <div className="row" style={{ marginTop: 8, justifyContent: "space-between" }}>
            <button className="btn" onClick={addLine}><Plus size={14} /> Línea</button>
            <div className="mono" style={{ fontSize: 13 }}>
              Debe: {fmt(totalDebit, currency)} · Haber: {fmt(totalCredit, currency)}
            </div>
          </div>

          {error && <div className="warn" style={{ marginTop: 10 }}><AlertCircle size={16} /> {error}</div>}
          <div style={{ marginTop: 10 }}>
            <button className="btn btn-primary" onClick={submit}>Registrar asiento</button>
          </div>
        </div>
      )}

      <div className="row" style={{ justifyContent: "space-between", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
        <PeriodBar period={period} onChange={setPeriod} />
        <ExportButtons onExcel={handleExportExcel} />
      </div>

      {filtered.length === 0 ? (
        <p className="empty">No hay asientos en el período seleccionado.</p>
      ) : (
        filtered.map((e) => (
          <div className="card" key={e.id} style={{ marginBottom: 10 }}>
            <div className="row" style={{ justifyContent: "space-between", cursor: "pointer" }}
              onClick={() => setExpanded(expanded === e.id ? null : e.id)}>
              <div className="row">
                {expanded === e.id ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
                <strong style={{ fontSize: 13.5 }}>{e.description}</strong>
                <span style={{ color: "var(--muted)", fontSize: 12.5 }}>{e.date}</span>
              </div>
              {perm.canDeleteJournal && (
                <button className="btn btn-icon btn-danger no-print" onClick={(ev) => { ev.stopPropagation(); removeEntry(e.id); }}>
                  <Trash2 size={14} />
                </button>
              )}
            </div>
            {expanded === e.id && (
              <table className="tbl" style={{ marginTop: 10 }}>
                <thead><tr><th>Cuenta</th><th className="num">Debe</th><th className="num">Haber</th></tr></thead>
                <tbody>
                  {e.lines.map((l, i) => (
                    <tr key={i}>
                      <td>{accName(l.accountId)}</td>
                      <td className="num">{l.debit ? fmt(l.debit, currency) : ""}</td>
                      <td className="num">{l.credit ? fmt(l.credit, currency) : ""}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        ))
      )}
    </div>
  );
}

function LedgerTab({ accounts, journal, currency, company }) {
  const sorted = [...accounts].sort((a, b) => a.code.localeCompare(b.code));
  const [accountId, setAccountId] = useState(sorted[0]?.id || "");
  const [period, setPeriod] = useState({ start: "", end: "" });
  const acc = accounts.find((a) => a.id === accountId);

  const rows = useMemo(() => {
    if (!acc) return [];
    const entries = [];
    journal.forEach((e) => {
      if (!inPeriod(e.date, period)) return;
      e.lines.forEach((l) => {
        if (l.accountId === accountId) entries.push({ date: e.date, desc: e.description, debit: l.debit, credit: l.credit });
      });
    });
    entries.sort((a, b) => a.date.localeCompare(b.date));
    let running = 0;
    return entries.map((r) => {
      running += isDebitNormal(acc.type) ? (r.debit - r.credit) : (r.credit - r.debit);
      return { ...r, running };
    });
  }, [acc, accountId, journal, period]);

  function handleExportExcel() {
    const rows2 = rows.map((r) => ({ Fecha: r.date, Descripción: r.desc, Debe: r.debit || 0, Haber: r.credit || 0, Saldo: r.running }));
    exportExcel(`mayor_${acc?.code || "cuenta"}_${period.start || "inicio"}_${period.end || "hoy"}.xlsx`, "Mayor", rows2);
  }

  return (
    <div>
      <h1 className="h1">Libro mayor</h1>
      <p className="sub">Movimientos y saldo acumulado por cuenta, en el período seleccionado.</p>
      <PrintHeader company={company} title={`Mayor · ${acc ? acc.code + " " + acc.name : ""}`} period={period} />

      <select className="input no-print" style={{ maxWidth: 320, marginBottom: 12 }} value={accountId} onChange={(e) => setAccountId(e.target.value)}>
        {sorted.map((a) => <option key={a.id} value={a.id}>{a.code} · {a.name}</option>)}
      </select>

      <div className="row" style={{ justifyContent: "space-between", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
        <PeriodBar period={period} onChange={setPeriod} />
        <ExportButtons onExcel={handleExportExcel} />
      </div>

      {rows.length === 0 ? (
        <p className="empty">Esta cuenta no tiene movimientos en el período seleccionado.</p>
      ) : (
        <table className="tbl">
          <thead><tr><th>Fecha</th><th>Descripción</th><th className="num">Debe</th><th className="num">Haber</th><th className="num">Saldo</th></tr></thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={i}>
                <td className="mono">{r.date}</td>
                <td>{r.desc}</td>
                <td className="num">{r.debit ? fmt(r.debit, currency) : ""}</td>
                <td className="num">{r.credit ? fmt(r.credit, currency) : ""}</td>
                <td className="num">{fmt(r.running, currency)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

function TrialBalanceTab({ accounts, journal, currency, company }) {
  const [period, setPeriod] = useState({ start: "", end: "" });
  const filteredJournal = useMemo(() => journal.filter((e) => inPeriod(e.date, period)), [journal, period]);
  const balances = useMemo(() => computeBalances(accounts, filteredJournal), [accounts, filteredJournal]);
  const sorted = [...accounts].sort((a, b) => a.code.localeCompare(b.code));
  const totalDebit = sorted.reduce((s, a) => s + (balances[a.id]?.debit || 0), 0);
  const totalCredit = sorted.reduce((s, a) => s + (balances[a.id]?.credit || 0), 0);
  const diff = totalDebit - totalCredit;

  function handleExportExcel() {
    const rows = sorted.map((a) => ({ Código: a.code, Cuenta: a.name, Debe: balances[a.id]?.debit || 0, Haber: balances[a.id]?.credit || 0 }));
    rows.push({ Código: "", Cuenta: "TOTALES", Debe: totalDebit, Haber: totalCredit });
    exportExcel(`balance_comprobacion_${period.start || "inicio"}_${period.end || "hoy"}.xlsx`, "Balance", rows);
  }

  return (
    <div>
      <h1 className="h1">Balance de comprobación</h1>
      <p className="sub">Suma de movimientos por cuenta en el período seleccionado. El total debe y haber deben coincidir.</p>
      <PrintHeader company={company} title="Balance de comprobación" period={period} />

      <div className="row" style={{ justifyContent: "space-between", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
        <PeriodBar period={period} onChange={setPeriod} />
        <ExportButtons onExcel={handleExportExcel} />
      </div>

      <table className="tbl">
        <thead><tr><th>Código</th><th>Cuenta</th><th className="num">Debe</th><th className="num">Haber</th></tr></thead>
        <tbody>
          {sorted.map((a) => (
            <tr key={a.id}>
              <td className="mono">{a.code}</td>
              <td>{a.name}</td>
              <td className="num">{fmt(balances[a.id]?.debit, currency)}</td>
              <td className="num">{fmt(balances[a.id]?.credit, currency)}</td>
            </tr>
          ))}
        </tbody>
        <tfoot>
          <tr>
            <td colSpan={2} style={{ fontWeight: 500 }}>Totales</td>
            <td className="num" style={{ fontWeight: 500 }}>{fmt(totalDebit, currency)}</td>
            <td className="num" style={{ fontWeight: 500 }}>{fmt(totalCredit, currency)}</td>
          </tr>
        </tfoot>
      </table>
      <p className="sub" style={{ marginTop: 10 }}>
        Diferencia: <span className="mono">{fmt(diff, currency)}</span> {Math.abs(diff) < 0.005 ? "· cuadrado" : "· revisa los asientos"}
      </p>
    </div>
  );
}

function StatementsTab({ accounts, journal, currency, company }) {
  const [period, setPeriod] = useState({ start: "", end: "" });
  const filteredJournal = useMemo(() => journal.filter((e) => inPeriod(e.date, period)), [journal, period]);
  const balances = useMemo(() => computeBalances(accounts, filteredJournal), [accounts, filteredJournal]);
  const totalsByType = useMemo(() => computeTotalsByType(accounts, balances), [accounts, balances]);
  const netIncome = totalsByType.Ingreso - totalsByType.Gasto;
  const byType = (t) => accounts.filter((a) => a.type === t);

  function handleExportExcel() {
    const rows = [
      { Rubro: "Ingresos", Monto: totalsByType.Ingreso },
      { Rubro: "Gastos", Monto: totalsByType.Gasto },
      { Rubro: "Utilidad neta", Monto: netIncome },
      { Rubro: "Total activos", Monto: totalsByType.Activo },
      { Rubro: "Pasivos", Monto: totalsByType.Pasivo },
      { Rubro: "Patrimonio", Monto: totalsByType.Patrimonio },
      { Rubro: "Total pasivo + patrimonio", Monto: totalsByType.Pasivo + totalsByType.Patrimonio + netIncome },
    ];
    exportExcel(`estados_financieros_${period.start || "inicio"}_${period.end || "hoy"}.xlsx`, "Estados", rows);
  }

  return (
    <div>
      <h1 className="h1">Estados financieros</h1>
      <p className="sub">Estado de resultados y balance general, calculados a partir del diario en el período seleccionado.</p>
      <PrintHeader company={company} title="Estados financieros" period={period} />

      <div className="row" style={{ justifyContent: "space-between", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
        <PeriodBar period={period} onChange={setPeriod} />
        <ExportButtons onExcel={handleExportExcel} />
      </div>

      <div className="section-title">Estado de resultados</div>
      <div className="card" style={{ maxWidth: 480 }}>
        <Line label="Ingresos" value={fmt(totalsByType.Ingreso, currency)} bold />
        <Line label="Gastos" value={fmt(totalsByType.Gasto, currency)} bold />
        <Line label="Utilidad neta" value={fmt(netIncome, currency)} bold accent />
      </div>

      <div className="section-title">Balance general</div>
      <div className="grid" style={{ gridTemplateColumns: "1fr 1fr", maxWidth: 700 }}>
        <div className="card">
          <p style={{ fontWeight: 500, fontSize: 13.5, marginBottom: 8 }}>Activos</p>
          {byType("Activo").map((a) => <Line key={a.id} label={a.name} value={fmt(0, currency)} skip />)}
          <Line label="Total activos" value={fmt(totalsByType.Activo, currency)} bold accent />
        </div>
        <div className="card">
          <p style={{ fontWeight: 500, fontSize: 13.5, marginBottom: 8 }}>Pasivo y patrimonio</p>
          <Line label="Pasivos" value={fmt(totalsByType.Pasivo, currency)} />
          <Line label="Patrimonio" value={fmt(totalsByType.Patrimonio, currency)} />
          <Line label="Utilidad del período" value={fmt(netIncome, currency)} />
          <Line label="Total pasivo + patrimonio" value={fmt(totalsByType.Pasivo + totalsByType.Patrimonio + netIncome, currency)} bold accent />
        </div>
      </div>
    </div>
  );
}

function Line({ label, value, bold, accent, skip }) {
  if (skip) return null;
  return (
    <div className="row" style={{ justifyContent: "space-between", padding: "4px 0" }}>
      <span style={{ fontSize: 13.5, fontWeight: bold ? 500 : 400 }}>{label}</span>
      <span className="mono" style={{ fontSize: 13.5, fontWeight: bold ? 500 : 400, color: accent ? "var(--accent)" : "var(--ink)" }}>{value}</span>
    </div>
  );
}

function BillingTab({ company, accounts, customers, invoices, journal, inventory, onCustomers, onInvoices, onJournal, onInventory, perm }) {
  const [sub, setSub] = useState("invoices");
  const [custForm, setCustForm] = useState({ name: "", taxId: "" });
  const [form, setForm] = useState({ customerId: "", date: new Date().toISOString().slice(0, 10), items: [{ description: "", qty: 1, price: 0, inventoryId: "" }] });
  const [msg, setMsg] = useState("");
  const [period, setPeriod] = useState({ start: "", end: "" });

  function addCustomer() {
    if (!custForm.name.trim()) return;
    onCustomers([...customers, { id: uid(), name: custForm.name.trim(), taxId: custForm.taxId.trim() }]);
    setCustForm({ name: "", taxId: "" });
  }
  function removeCustomer(id) { onCustomers(customers.filter((c) => c.id !== id)); }

  function updateItem(i, patch) {
    const items = [...form.items]; items[i] = { ...items[i], ...patch };
    setForm({ ...form, items });
  }
  function addItem() { setForm({ ...form, items: [...form.items, { description: "", qty: 1, price: 0, inventoryId: "" }] }); }
  function removeItem(i) { setForm({ ...form, items: form.items.filter((_, idx) => idx !== i) }); }

  const subtotal = form.items.reduce((s, it) => s + (Number(it.qty) || 0) * (Number(it.price) || 0), 0);
  const iva = subtotal * (company.ivaRate / 100);
  const total = subtotal + iva;

  function createInvoice() {
    if (!form.customerId) return setMsg("Selecciona un cliente.");
    const validItems = form.items.filter((it) => it.description.trim() && Number(it.qty) > 0);
    if (validItems.length === 0) return setMsg("Agrega al menos un artículo o servicio.");
    setMsg("");
    const number = `F-${String(invoices.length + 1).padStart(4, "0")}`;
    const invoice = { id: uid(), number, customerId: form.customerId, date: form.date, items: validItems, ivaRate: company.ivaRate, subtotal, iva, total, status: "Pendiente" };
    onInvoices([invoice, ...invoices]);

    const newEntries = [];
    const ar = accounts.find((a) => a.code === "1030");
    const sales = accounts.find((a) => a.code === "4010");
    const ivaAcc = accounts.find((a) => a.code === "2020");
    if (ar && sales && ivaAcc) {
      const custName = customers.find((c) => c.id === form.customerId)?.name || "cliente";
      const lines = [{ accountId: ar.id, debit: total, credit: 0 }, { accountId: sales.id, debit: 0, credit: subtotal }];
      if (iva > 0) lines.push({ accountId: ivaAcc.id, debit: 0, credit: iva });
      newEntries.push({ id: uid(), date: form.date, description: `Factura ${number} — ${custName}`, lines });
    } else {
      setMsg("Factura creada, pero no se generó el asiento automático: faltan cuentas 1030, 4010 o 2020 en el plan de cuentas.");
    }

    // Descontar inventario de los artículos vinculados y registrar el costo de venta
    const linkedItems = validItems.filter((it) => it.inventoryId);
    if (linkedItems.length > 0 && inventory) {
      let updatedInventory = inventory;
      let totalCost = 0;
      let stockWarning = false;
      linkedItems.forEach((it) => {
        const qtySold = Number(it.qty) || 0;
        const invItem = updatedInventory.find((x) => x.id === it.inventoryId);
        if (!invItem) return;
        if (qtySold > invItem.qty) stockWarning = true;
        totalCost += qtySold * (Number(invItem.cost) || 0);
        updatedInventory = updatedInventory.map((x) => x.id === it.inventoryId ? { ...x, qty: Math.max(0, x.qty - qtySold) } : x);
      });
      onInventory(updatedInventory);

      const cogs = accounts.find((a) => a.code === "5010");
      const invAcc = accounts.find((a) => a.code === "1040");
      if (cogs && invAcc && totalCost > 0) {
        newEntries.push({ id: uid(), date: form.date, description: `Costo de venta — Factura ${number}`, lines: [{ accountId: cogs.id, debit: totalCost, credit: 0 }, { accountId: invAcc.id, debit: 0, credit: totalCost }] });
      }
      if (stockWarning) setMsg((m) => (m ? m + " " : "") + "Aviso: alguna cantidad facturada superaba el stock disponible; el inventario se ajustó a 0.");
    }

    if (newEntries.length > 0) onJournal([...newEntries, ...journal]);
    setForm({ customerId: form.customerId, date: form.date, items: [{ description: "", qty: 1, price: 0, inventoryId: "" }] });
  }

  function markPaid(inv) {
    onInvoices(invoices.map((i) => i.id === inv.id ? { ...i, status: "Pagada" } : i));
    const bank = accounts.find((a) => a.code === "1020");
    const ar = accounts.find((a) => a.code === "1030");
    if (bank && ar) {
      onJournal([{ id: uid(), date: new Date().toISOString().slice(0, 10), description: `Cobro factura ${inv.number}`, lines: [{ accountId: bank.id, debit: inv.total, credit: 0 }, { accountId: ar.id, debit: 0, credit: inv.total }] }, ...journal]);
    }
  }

  const custName = (id) => customers.find((c) => c.id === id)?.name || "—";
  const filteredInvoices = useMemo(() => invoices.filter((inv) => inPeriod(inv.date, period)), [invoices, period]);

  function handleExportExcel() {
    const rows = filteredInvoices.map((inv) => ({
      Número: inv.number, Cliente: custName(inv.customerId), Fecha: inv.date,
      Subtotal: inv.subtotal, IVA: inv.iva, Total: inv.total, Estado: inv.status,
    }));
    exportExcel(`facturas_${period.start || "inicio"}_${period.end || "hoy"}.xlsx`, "Facturas", rows);
  }

  return (
    <div>
      <h1 className="h1">Clientes y facturación</h1>
      <p className="sub">Cada factura registrada genera automáticamente su asiento contable{inventory?.length ? " y descuenta el inventario vinculado" : ""}.</p>

      <div className="row no-print" style={{ marginBottom: 18, gap: 6 }}>
        <button className={`btn ${sub === "invoices" ? "btn-primary" : ""}`} onClick={() => setSub("invoices")}>Facturas</button>
        <button className={`btn ${sub === "customers" ? "btn-primary" : ""}`} onClick={() => setSub("customers")}>Clientes</button>
      </div>

      {sub === "customers" ? (
        <div>
          {perm.canEditBilling && (
            <div className="card no-print" style={{ marginBottom: 16 }}>
              <div className="row">
                <input className="input" placeholder="Nombre del cliente" value={custForm.name} onChange={(e) => setCustForm({ ...custForm, name: e.target.value })} />
                <input className="input" placeholder="Cédula / ID fiscal" style={{ maxWidth: 200 }} value={custForm.taxId} onChange={(e) => setCustForm({ ...custForm, taxId: e.target.value })} />
                <button className="btn btn-primary" onClick={addCustomer}><Plus size={14} /> Agregar</button>
              </div>
            </div>
          )}
          {customers.length === 0 ? <p className="empty">Todavía no hay clientes.</p> : (
            <table className="tbl">
              <thead><tr><th>Nombre</th><th>Cédula / ID</th><th className="no-print"></th></tr></thead>
              <tbody>
                {customers.map((c) => (
                  <tr key={c.id}><td>{c.name}</td><td className="mono">{c.taxId}</td>
                    <td className="no-print">{perm.canEditBilling && <button className="btn btn-icon btn-danger" onClick={() => removeCustomer(c.id)}><Trash2 size={14} /></button>}</td></tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      ) : (
        <div>
          {perm.canEditBilling && (
            <div className="card no-print" style={{ marginBottom: 20 }}>
              <div className="row" style={{ marginBottom: 10 }}>
                <select className="input" value={form.customerId} onChange={(e) => setForm({ ...form, customerId: e.target.value })}>
                  <option value="">Selecciona cliente…</option>
                  {customers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
                </select>
                <input className="input" type="date" style={{ maxWidth: 160 }} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
              </div>
              {form.items.map((it, i) => (
                <div className="row" key={i} style={{ marginBottom: 6 }}>
                  <select className="input" style={{ maxWidth: 210 }} value={it.inventoryId || ""} onChange={(e) => {
                    const invItem = (inventory || []).find((x) => x.id === e.target.value);
                    if (invItem) updateItem(i, { inventoryId: invItem.id, description: invItem.name, price: invItem.price });
                    else updateItem(i, { inventoryId: "" });
                  }}>
                    <option value="">Artículo libre (sin inventario)</option>
                    {(inventory || []).map((invIt) => (
                      <option key={invIt.id} value={invIt.id}>{invIt.sku ? `${invIt.sku} · ` : ""}{invIt.name} (disp. {invIt.qty})</option>
                    ))}
                  </select>
                  <input className="input" placeholder="Descripción" value={it.description} onChange={(e) => updateItem(i, { description: e.target.value })} />
                  <input className="input" type="number" placeholder="Cant." style={{ maxWidth: 90 }} value={it.qty} onChange={(e) => updateItem(i, { qty: e.target.value })} />
                  <input className="input" type="number" step="0.01" placeholder="Precio unit." style={{ maxWidth: 130 }} value={it.price} onChange={(e) => updateItem(i, { price: e.target.value })} />
                  <button className="btn btn-icon btn-danger" onClick={() => removeItem(i)}><Trash2 size={14} /></button>
                </div>
              ))}
              <div className="row" style={{ justifyContent: "space-between", marginTop: 8 }}>
                <button className="btn" onClick={addItem}><Plus size={14} /> Artículo</button>
                <div className="mono" style={{ fontSize: 13 }}>
                  Subtotal {fmt(subtotal, company.currency)} · IVA {fmt(iva, company.currency)} · Total {fmt(total, company.currency)}
                </div>
              </div>
              {msg && <div className="warn" style={{ marginTop: 10 }}><AlertCircle size={16} /> {msg}</div>}
              <div style={{ marginTop: 10 }}><button className="btn btn-primary" onClick={createInvoice}>Crear factura</button></div>
            </div>
          )}

          <PrintHeader company={company} title="Facturas" period={period} />
          <div className="row" style={{ justifyContent: "space-between", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
            <PeriodBar period={period} onChange={setPeriod} />
            <ExportButtons onExcel={handleExportExcel} />
          </div>

          {filteredInvoices.length === 0 ? <p className="empty">No hay facturas en el período seleccionado.</p> : (
            <table className="tbl">
              <thead><tr><th>N.°</th><th>Cliente</th><th>Fecha</th><th className="num">Total</th><th>Estado</th><th className="no-print"></th></tr></thead>
              <tbody>
                {filteredInvoices.map((inv) => (
                  <tr key={inv.id}>
                    <td className="mono">{inv.number}</td>
                    <td>{custName(inv.customerId)}</td>
                    <td className="mono">{inv.date}</td>
                    <td className="num">{fmt(inv.total, company.currency)}</td>
                    <td><span className={`badge ${inv.status === "Pagada" ? "badge-Activo" : "badge-Gasto"}`}>{inv.status}</span></td>
                    <td className="no-print">{perm.canEditBilling && inv.status !== "Pagada" && <button className="btn" onClick={() => markPaid(inv)}>Marcar pagada</button>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}
    </div>
  );
}

function InventoryTab({ inventory, onChange, currency, perm }) {
  const [form, setForm] = useState({ sku: "", name: "", qty: 0, cost: 0, price: 0 });
  function add() {
    if (!form.name.trim()) return;
    onChange([...inventory, { id: uid(), ...form, qty: Number(form.qty), cost: Number(form.cost), price: Number(form.price) }]);
    setForm({ sku: "", name: "", qty: 0, cost: 0, price: 0 });
  }
  function remove(id) { onChange(inventory.filter((i) => i.id !== id)); }

  function handleExportExcel() {
    const rows = inventory.map((i) => ({ SKU: i.sku, Artículo: i.name, Cantidad: i.qty, Costo: i.cost, Precio: i.price, "Valor total": i.qty * i.cost }));
    exportExcel("inventario.xlsx", "Inventario", rows);
  }

  return (
    <div>
      <h1 className="h1">Inventario</h1>
      <p className="sub">Existencias, costo y precio de venta. Se descuenta automáticamente al facturar artículos vinculados.</p>

      {perm.canEditInventory && (
        <div className="card no-print" style={{ marginBottom: 20 }}>
          <div className="row">
            <input className="input" placeholder="SKU" style={{ maxWidth: 110 }} value={form.sku} onChange={(e) => setForm({ ...form, sku: e.target.value })} />
            <input className="input" placeholder="Nombre del artículo" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
            <input className="input" type="number" placeholder="Cant." style={{ maxWidth: 90 }} value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} />
            <input className="input" type="number" step="0.01" placeholder="Costo" style={{ maxWidth: 110 }} value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} />
            <input className="input" type="number" step="0.01" placeholder="Precio" style={{ maxWidth: 110 }} value={form.price} onChange={(e) => setForm({ ...form, price: e.target.value })} />
            <button className="btn btn-primary" onClick={add}><Plus size={14} /></button>
          </div>
        </div>
      )}

      <div className="row no-print" style={{ justifyContent: "flex-end", marginBottom: 12 }}>
        <ExportButtons onExcel={handleExportExcel} />
      </div>

      {inventory.length === 0 ? <p className="empty">Todavía no hay artículos en inventario.</p> : (
        <table className="tbl">
          <thead><tr><th>SKU</th><th>Artículo</th><th className="num">Cantidad</th><th className="num">Costo</th><th className="num">Precio</th><th className="num">Valor total</th><th className="no-print"></th></tr></thead>
          <tbody>
            {inventory.map((i) => (
              <tr key={i.id}>
                <td className="mono">{i.sku}</td>
                <td>{i.name}</td>
                <td className="num">{i.qty}</td>
                <td className="num">{fmt(i.cost, currency)}</td>
                <td className="num">{fmt(i.price, currency)}</td>
                <td className="num">{fmt(i.qty * i.cost, currency)}</td>
                <td className="no-print">{perm.canEditInventory && <button className="btn btn-icon btn-danger" onClick={() => remove(i.id)}><Trash2 size={14} /></button>}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}
