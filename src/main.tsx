import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { motion, AnimatePresence } from 'framer-motion';
import {
  BarChart3,
  Bike,
  CalendarDays,
  Car,
  ChevronLeft,
  ClipboardList,
  Download,
  Droplets,
  FileSpreadsheet,
  FileText,
  FilterX,
  Gauge,
  LogOut,
  Plus,
  Search,
  ShieldCheck,
  Sparkles,
  Wrench,
  X,
} from 'lucide-react';
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import './styles.css';

type ServiceName = 'Mantenimiento' | 'Lavado' | 'Engrasada' | 'Cambio de aceite';

type VehicleRecord = {
  item: string;
  fecha: string;
  lugar: string;
  maestro: string;
  taller: string;
  chofer: string;
  placa: string;
  cantidadPago: number;
  descripcion: string;
  tipoServicio: ServiceName;
};

type Filters = {
  placa: string;
  desde: string;
  hasta: string;
};

type FormState = Omit<VehicleRecord, 'item' | 'cantidadPago'> & { cantidadPago: string };

const SHEET_ID = import.meta.env.VITE_GOOGLE_SHEET_ID || '1Evr1lpNSwLYXWgcRf5D5NSPUD5nR-YS_EczVK4PVAHI';
const SCRIPT_URL = import.meta.env.VITE_GOOGLE_SCRIPT_URL || '';

const SERVICES: ServiceName[] = ['Mantenimiento', 'Lavado', 'Engrasada', 'Cambio de aceite'];
const COLORS = ['#ff0000', '#25ff00', '#111827', '#9ca3af'];
const HEADERS = ['ITEM', 'FECHA', 'LUGAR', 'MAESTRO', 'TALLER', 'CHOFER', 'PLACA', 'CANTIDAD DE PAGO', 'DESCRIPCION'];
const SHEET_GIDS: Record<ServiceName, string> = {
  Mantenimiento: '0',
  Lavado: '819388144',
  Engrasada: '2024356449',
  'Cambio de aceite': '464443967',
};

const serviceMeta: Record<ServiceName, { icon: React.ElementType; accent: string; description: string }> = {
  Mantenimiento: { icon: Wrench, accent: 'red', description: 'Servicios correctivos y preventivos' },
  Lavado: { icon: Droplets, accent: 'green', description: 'Control de limpiezas vehiculares' },
  Engrasada: { icon: Gauge, accent: 'dark', description: 'Lubricacion y proteccion mecanica' },
  'Cambio de aceite': { icon: Bike, accent: 'gray', description: 'Registros de aceite y filtros' },
};

const mockData: VehicleRecord[] = [
  { item: '1', fecha: '2026-06-02', lugar: 'Santo Domingo', maestro: 'Carlos Diaz', taller: 'Taller Central', chofer: 'Luis Perez', placa: 'PAS-104', cantidadPago: 4200, descripcion: 'Revision general de frenos', tipoServicio: 'Mantenimiento' },
  { item: '2', fecha: '2026-06-05', lugar: 'La Vega', maestro: 'Ramon Cruz', taller: 'PAS Norte', chofer: 'Ana Gomez', placa: 'PAS-220', cantidadPago: 850, descripcion: 'Lavado completo', tipoServicio: 'Lavado' },
  { item: '3', fecha: '2026-06-08', lugar: 'Santiago', maestro: 'Miguel Soto', taller: 'LubriExpress', chofer: 'Jose Rivas', placa: 'PAS-104', cantidadPago: 1450, descripcion: 'Engrasada tren delantero', tipoServicio: 'Engrasada' },
  { item: '4', fecha: '2026-06-12', lugar: 'Distrito Nacional', maestro: 'Pedro Leon', taller: 'AutoServicio PAS', chofer: 'Marta Gil', placa: 'PAS-308', cantidadPago: 3650, descripcion: 'Cambio aceite 15W40', tipoServicio: 'Cambio de aceite' },
];

const emptyFilters: Filters = { placa: '', desde: '', hasta: '' };
const emptyForm: FormState = {
  fecha: new Date().toISOString().slice(0, 10),
  lugar: '',
  maestro: '',
  taller: '',
  chofer: '',
  placa: '',
  cantidadPago: '',
  descripcion: '',
  tipoServicio: 'Mantenimiento',
};

function money(value: number) {
  return new Intl.NumberFormat('es-DO', { style: 'currency', currency: 'DOP', maximumFractionDigits: 0 }).format(value || 0);
}

function normalizeDate(value: string) {
  if (!value) return '';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toISOString().slice(0, 10);
}

function parseNumber(value: unknown) {
  return Number(String(value ?? '').replace(/[^0-9.-]/g, '')) || 0;
}

function parseGviz(text: string): unknown[][] {
  const json = JSON.parse(text.substring(text.indexOf('{'), text.lastIndexOf('}') + 1));
  return (json.table.rows || []).map((row: { c: Array<{ v?: unknown; f?: string } | null> }) =>
    (row.c || []).map((cell) => cell?.f ?? cell?.v ?? '')
  );
}

async function fetchSheet(service: ServiceName): Promise<VehicleRecord[]> {
  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&gid=${SHEET_GIDS[service]}`;
  const response = await fetch(url);
  if (!response.ok) throw new Error(`No se pudo consultar ${service}`);
  const rows = parseGviz(await response.text());
  const dataRows = rows.filter((row) => String(row[0]).toUpperCase() !== 'ITEM' && row.some(Boolean));
  return dataRows.map((row) => ({
    item: String(row[0] ?? ''),
    fecha: normalizeDate(String(row[1] ?? '')),
    lugar: String(row[2] ?? ''),
    maestro: String(row[3] ?? ''),
    taller: String(row[4] ?? ''),
    chofer: String(row[5] ?? ''),
    placa: String(row[6] ?? '').toUpperCase(),
    cantidadPago: parseNumber(row[7]),
    descripcion: String(row[8] ?? ''),
    tipoServicio: service,
  }));
}

async function addRecord(record: FormState) {
  if (!SCRIPT_URL) {
    throw new Error('El formulario esta listo, pero falta configurar VITE_GOOGLE_SCRIPT_URL para guardar en Google Sheets. Revisa README.md.');
  }
  const response = await fetch(SCRIPT_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ ...record, cantidadPago: Number(record.cantidadPago) || 0 }),
  });
  if (!response.ok) throw new Error('Google Sheets no acepto el registro. Revisa permisos del Apps Script.');
  return response;
}

function applyFilters(records: VehicleRecord[], filters: Filters) {
  const plate = filters.placa.trim().toUpperCase();
  return records.filter((record) => {
    const byPlate = !plate || record.placa.toUpperCase().includes(plate);
    const byStart = !filters.desde || record.fecha >= filters.desde;
    const byEnd = !filters.hasta || record.fecha <= filters.hasta;
    return byPlate && byStart && byEnd;
  });
}

function Logo({ compact = false }: { compact?: boolean }) {
  return (
    <div className="logo" aria-label="Distribuidor Punto PAS">
      <img className="logo-image" src="/logo-punto-pas.svg" alt="Punto PAS" />
      {!compact && <div><strong>Distribuidor Punto PAS</strong><small>Gestion vehicular</small></div>}
    </div>
  );
}

function Login({ onLogin }: { onLogin: () => void }) {
  const [user, setUser] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');

  function submit(event: React.FormEvent) {
    event.preventDefault();
    if (user === 'Maria21' && password === 'PuntoPas2026*') {
      sessionStorage.setItem('punto-pas-session', 'Maria21');
      onLogin();
      return;
    }
    setError('Usuario o contrasena incorrectos. Verifique los datos e intente nuevamente.');
  }

  return (
    <main className="login-shell">
      <motion.section className="login-card" initial={{ opacity: 0, y: 24 }} animate={{ opacity: 1, y: 0 }}>
        <Logo />
        <div className="login-copy">
          <span className="eyebrow"><ShieldCheck size={16} /> Acceso seguro</span>
          <h1>Panel ejecutivo de mantenimientos vehiculares</h1>
          <p>Controle servicios, gastos, placas y reportes operativos desde una interfaz empresarial.</p>
        </div>
        <form onSubmit={submit} className="login-form">
          <label>Usuario<input value={user} onChange={(e) => setUser(e.target.value)} placeholder="Maria21" autoComplete="username" /></label>
          <label>Contrasena<input type="password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="********" autoComplete="current-password" /></label>
          {error && <div className="alert error">{error}</div>}
          <button className="primary full" type="submit">Ingresar al sistema</button>
        </form>
      </motion.section>
    </main>
  );
}

function FilterBar({ filters, setFilters, onClear }: { filters: Filters; setFilters: (filters: Filters) => void; onClear: () => void }) {
  return (
    <section className="filter-bar">
      <label><Search size={16} /> Placa<input value={filters.placa} onChange={(e) => setFilters({ ...filters, placa: e.target.value })} placeholder="Ej. PAS-104" /></label>
      <label><CalendarDays size={16} /> Desde<input type="date" value={filters.desde} onChange={(e) => setFilters({ ...filters, desde: e.target.value })} /></label>
      <label><CalendarDays size={16} /> Hasta<input type="date" value={filters.hasta} onChange={(e) => setFilters({ ...filters, hasta: e.target.value })} /></label>
      <button className="ghost" onClick={onClear}><FilterX size={16} /> Limpiar</button>
    </section>
  );
}

function SummaryCard({ title, value, hint, icon: Icon }: { title: string; value: string; hint: string; icon: React.ElementType }) {
  return <div className="summary-card"><div><span>{title}</span><strong>{value}</strong><small>{hint}</small></div><Icon size={24} /></div>;
}

function ServiceCard({ service, records, onOpen }: { service: ServiceName; records: VehicleRecord[]; onOpen: () => void }) {
  const meta = serviceMeta[service];
  const Icon = meta.icon;
  const total = records.reduce((sum, record) => sum + record.cantidadPago, 0);
  const last = [...records].sort((a, b) => b.fecha.localeCompare(a.fecha))[0];
  return (
    <motion.article whileHover={{ y: -5 }} className={`service-card ${meta.accent}`}>
      <div className="service-head"><div className="service-icon"><Icon size={28} /></div><span>{records.length} registros</span></div>
      <h3>{service.toUpperCase()}</h3>
      <p>{meta.description}</p>
      <div className="service-stats"><strong>{money(total)}</strong><small>Ultimo: {last ? `${last.fecha} - ${last.placa}` : 'Sin registros'}</small></div>
      <button className="secondary" onClick={onOpen}>Ingresar al modulo</button>
    </motion.article>
  );
}

function DataTable({ records, title, filters, service }: { records: VehicleRecord[]; title: string; filters: Filters; service?: ServiceName }) {
  const [sortKey, setSortKey] = useState<keyof VehicleRecord>('fecha');
  const [direction, setDirection] = useState<'asc' | 'desc'>('desc');
  const [page, setPage] = useState(1);
  const pageSize = 8;
  const sorted = useMemo(() => [...records].sort((a, b) => {
    const left = a[sortKey];
    const right = b[sortKey];
    const result = typeof left === 'number' && typeof right === 'number' ? left - right : String(left).localeCompare(String(right));
    return direction === 'asc' ? result : -result;
  }), [records, sortKey, direction]);
  const pages = Math.max(1, Math.ceil(sorted.length / pageSize));
  const visible = sorted.slice((page - 1) * pageSize, page * pageSize);
  const total = records.reduce((sum, record) => sum + record.cantidadPago, 0);

  useEffect(() => setPage(1), [records.length]);

  function setSort(key: keyof VehicleRecord) {
    if (sortKey === key) setDirection(direction === 'asc' ? 'desc' : 'asc');
    else { setSortKey(key); setDirection('asc'); }
  }

  return (
    <section className="table-card">
      <div className="table-toolbar"><div><h2>{title}</h2><p>{records.length} registros filtrados · Total {money(total)}</p></div><ExportButtons records={records} title={title} filters={filters} service={service} /></div>
      <div className="table-wrap">
        <table>
          <thead><tr>{['fecha', 'tipoServicio', 'placa', 'chofer', 'taller', 'lugar', 'cantidadPago', 'descripcion'].map((key) => <th key={key} onClick={() => setSort(key as keyof VehicleRecord)}>{labelFor(key)} {sortKey === key ? (direction === 'asc' ? '↑' : '↓') : ''}</th>)}</tr></thead>
          <tbody>
            {visible.map((record, index) => (
              <tr key={`${record.tipoServicio}-${record.item}-${index}`}>
                <td data-label="Fecha">{record.fecha}</td><td data-label="Servicio"><span className="pill">{record.tipoServicio}</span></td><td data-label="Placa"><strong>{record.placa}</strong></td><td data-label="Chofer">{record.chofer}</td><td data-label="Taller">{record.taller}</td><td data-label="Lugar">{record.lugar}</td><td data-label="Pago">{money(record.cantidadPago)}</td><td data-label="Descripcion">{record.descripcion}</td>
              </tr>
            ))}
            {!visible.length && <tr><td colSpan={8} className="empty">No hay registros para los filtros seleccionados.</td></tr>}
          </tbody>
        </table>
      </div>
      <div className="pagination"><button className="ghost" disabled={page === 1} onClick={() => setPage(page - 1)}>Anterior</button><span>Pagina {page} de {pages}</span><button className="ghost" disabled={page === pages} onClick={() => setPage(page + 1)}>Siguiente</button></div>
    </section>
  );
}

function labelFor(key: string) {
  return ({ fecha: 'Fecha', tipoServicio: 'Servicio', placa: 'Placa', chofer: 'Chofer', taller: 'Taller', lugar: 'Lugar', cantidadPago: 'Pago', descripcion: 'Descripcion' } as Record<string, string>)[key] || key;
}

function ExportButtons({ records, title, filters, service }: { records: VehicleRecord[]; title: string; filters: Filters; service?: ServiceName }) {
  function rows() {
    return records.map((record) => ({ Fecha: record.fecha, Servicio: record.tipoServicio, Placa: record.placa, Chofer: record.chofer, Taller: record.taller, Lugar: record.lugar, Maestro: record.maestro, Pago: record.cantidadPago, Descripcion: record.descripcion }));
  }
  function excel() {
    const ws = XLSX.utils.json_to_sheet(rows());
    XLSX.utils.sheet_add_aoa(ws, [[`Distribuidor Punto PAS - ${title}`], [`Generado: ${new Date().toLocaleString('es-DO')}`], [`Filtros: placa ${filters.placa || 'Todas'}, desde ${filters.desde || 'inicio'}, hasta ${filters.hasta || 'actual'}, servicio ${service || 'Todos'}`], []], { origin: 'A1' });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Reporte');
    XLSX.writeFile(wb, `${title.replace(/\s+/g, '_')}.xlsx`);
  }
  function pdf() {
    const doc = new jsPDF({ orientation: 'landscape' });
    doc.setFillColor(255, 0, 0); doc.rect(10, 10, 12, 12, 'F');
    doc.setFillColor(37, 255, 0); doc.rect(22, 10, 12, 12, 'F');
    doc.setFontSize(16); doc.text('Distribuidor Punto PAS', 40, 16);
    doc.setFontSize(11); doc.text(title, 40, 23);
    doc.text(`Generado: ${new Date().toLocaleString('es-DO')} | Placa: ${filters.placa || 'Todas'} | Servicio: ${service || 'Todos'}`, 10, 34);
    autoTable(doc, { startY: 40, head: [['Fecha', 'Servicio', 'Placa', 'Chofer', 'Taller', 'Lugar', 'Pago', 'Descripcion']], body: records.map((r) => [r.fecha, r.tipoServicio, r.placa, r.chofer, r.taller, r.lugar, money(r.cantidadPago), r.descripcion]), styles: { fontSize: 8 }, headStyles: { fillColor: [255, 0, 0] } });
    doc.save(`${title.replace(/\s+/g, '_')}.pdf`);
  }
  return <div className="export-actions"><button className="secondary" onClick={excel}><FileSpreadsheet size={16} /> Excel</button><button className="secondary" onClick={pdf}><FileText size={16} /> PDF</button></div>;
}

function Charts({ records }: { records: VehicleRecord[] }) {
  const byPlate = Object.values(records.reduce((acc, r) => { acc[r.placa] = acc[r.placa] || { name: r.placa, total: 0 }; acc[r.placa].total += r.cantidadPago; return acc; }, {} as Record<string, { name: string; total: number }>)).sort((a, b) => b.total - a.total).slice(0, 8);
  const byService = SERVICES.map((service) => ({ name: service, value: records.filter((r) => r.tipoServicio === service).length }));
  const byMonth = Object.values(records.reduce((acc, r) => { const key = r.fecha.slice(0, 7) || 'Sin fecha'; acc[key] = acc[key] || { name: key, total: 0 }; acc[key].total += r.cantidadPago; return acc; }, {} as Record<string, { name: string; total: number }>)).sort((a, b) => a.name.localeCompare(b.name));
  return <section className="charts-grid"><ChartCard title="Gasto por placa"><ResponsiveContainer width="100%" height={260}><BarChart data={byPlate}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis /><Tooltip formatter={(v) => money(Number(v))} /><Bar dataKey="total" fill="#ff0000" radius={[8, 8, 0, 0]} /></BarChart></ResponsiveContainer></ChartCard><ChartCard title="Servicios por tipo"><ResponsiveContainer width="100%" height={260}><PieChart><Pie data={byService} dataKey="value" nameKey="name" outerRadius={90} label>{byService.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}</Pie><Tooltip /></PieChart></ResponsiveContainer></ChartCard><ChartCard title="Tendencia mensual"><ResponsiveContainer width="100%" height={260}><LineChart data={byMonth}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis /><Tooltip formatter={(v) => money(Number(v))} /><Line dataKey="total" stroke="#25ff00" strokeWidth={3} dot={{ r: 5 }} /></LineChart></ResponsiveContainer></ChartCard></section>;
}

function ChartCard({ title, children }: { title: string; children: React.ReactNode }) {
  return <div className="chart-card"><h3>{title}</h3>{children}</div>;
}

function RecordModal({ open, onClose, onSaved }: { open: boolean; onClose: () => void; onSaved: (record: VehicleRecord) => void }) {
  const [form, setForm] = useState<FormState>(emptyForm);
  const [status, setStatus] = useState('');
  const [saving, setSaving] = useState(false);
  async function submit(event: React.FormEvent) {
    event.preventDefault();
    setSaving(true); setStatus('');
    try {
      await addRecord(form);
      const record: VehicleRecord = { ...form, item: String(Date.now()), cantidadPago: Number(form.cantidadPago) || 0, placa: form.placa.toUpperCase() };
      onSaved(record);
      setStatus('Registro guardado correctamente.');
      setForm(emptyForm);
      setTimeout(onClose, 700);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'No se pudo guardar el registro.');
    } finally { setSaving(false); }
  }
  return <AnimatePresence>{open && <motion.div className="modal-backdrop" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}><motion.form className="modal" onSubmit={submit} initial={{ scale: .96, y: 20 }} animate={{ scale: 1, y: 0 }} exit={{ scale: .96, y: 20 }}><div className="modal-head"><div><span className="eyebrow"><Plus size={14} /> Nuevo registro</span><h2>Agregar servicio vehicular</h2></div><button type="button" className="icon-btn" onClick={onClose}><X size={20} /></button></div><div className="form-grid"><Field label="Fecha" type="date" value={form.fecha} onChange={(v) => setForm({ ...form, fecha: v })} /><Field label="Lugar" value={form.lugar} onChange={(v) => setForm({ ...form, lugar: v })} /><Field label="Maestro" value={form.maestro} onChange={(v) => setForm({ ...form, maestro: v })} /><Field label="Taller" value={form.taller} onChange={(v) => setForm({ ...form, taller: v })} /><Field label="Chofer" value={form.chofer} onChange={(v) => setForm({ ...form, chofer: v })} /><Field label="Placa" value={form.placa} onChange={(v) => setForm({ ...form, placa: v.toUpperCase() })} /><Field label="Cantidad de pago" type="number" value={form.cantidadPago} onChange={(v) => setForm({ ...form, cantidadPago: v })} /><label>Tipo de servicio<select value={form.tipoServicio} onChange={(e) => setForm({ ...form, tipoServicio: e.target.value as ServiceName })}>{SERVICES.map((service) => <option key={service}>{service}</option>)}</select></label><label className="span-2">Descripcion<textarea value={form.descripcion} onChange={(e) => setForm({ ...form, descripcion: e.target.value })} required /></label></div>{status && <div className={`alert ${status.includes('correctamente') ? 'success' : 'error'}`}>{status}</div>}<button className="primary full" disabled={saving}>{saving ? 'Guardando...' : 'Guardar en Google Sheets'}</button></motion.form></motion.div>}</AnimatePresence>;
}

function Field({ label, value, onChange, type = 'text' }: { label: string; value: string; onChange: (value: string) => void; type?: string }) {
  return <label>{label}<input required type={type} value={value} onChange={(e) => onChange(e.target.value)} /></label>;
}

function ModuleView({ service, records, filters, setFilters, onBack, onAdd }: { service: ServiceName; records: VehicleRecord[]; filters: Filters; setFilters: (filters: Filters) => void; onBack: () => void; onAdd: () => void }) {
  const filtered = applyFilters(records.filter((r) => r.tipoServicio === service), filters);
  const Icon = serviceMeta[service].icon;
  return <><header className="topbar"><Logo /><div className="top-actions"><button className="ghost" onClick={onBack}><ChevronLeft size={16} /> Dashboard</button></div></header><main className="page"><section className="hero compact"><div><span className="eyebrow"><Icon size={16} /> Modulo individual</span><h1>{service}</h1><p>Consulta, ordena, filtra, exporta y registra servicios de {service.toLowerCase()}.</p></div><button className="primary" onClick={onAdd}><Plus size={18} /> Agregar registro</button></section><FilterBar filters={filters} setFilters={setFilters} onClear={() => setFilters(emptyFilters)} /><DataTable records={filtered} title={`Reporte de ${service}`} filters={filters} service={service} /></main><button className="fab" onClick={onAdd}><Plus /></button></>;
}

function Dashboard() {
  const [records, setRecords] = useState<VehicleRecord[]>([]);
  const [filters, setFilters] = useState<Filters>(emptyFilters);
  const [activeService, setActiveService] = useState<ServiceName | null>(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [loading, setLoading] = useState(true);
  const [notice, setNotice] = useState('');

  useEffect(() => {
    Promise.all(SERVICES.map(fetchSheet)).then((results) => setRecords(results.flat())).catch(() => { setRecords(mockData); setNotice('No se pudo leer Google Sheets. Se muestran datos de demostracion hasta publicar/compartir correctamente el archivo.'); }).finally(() => setLoading(false));
  }, []);

  if (activeService) {
    return <><ModuleView service={activeService} records={records} filters={filters} setFilters={setFilters} onBack={() => setActiveService(null)} onAdd={() => setModalOpen(true)} /><RecordModal open={modalOpen} onClose={() => setModalOpen(false)} onSaved={(record) => setRecords((current) => [record, ...current])} /></>;
  }

  const filtered = applyFilters(records, filters);
  const totalPaid = filtered.reduce((sum, record) => sum + record.cantidadPago, 0);
  const byPlate = Object.entries(filtered.reduce((acc, r) => { acc[r.placa] = (acc[r.placa] || 0) + r.cantidadPago; return acc; }, {} as Record<string, number>)).sort((a, b) => b[1] - a[1]);
  const last = [...filtered].sort((a, b) => b.fecha.localeCompare(a.fecha))[0];
  const avgByVehicle = byPlate.length ? totalPaid / byPlate.length : 0;

  function logout() { sessionStorage.removeItem('punto-pas-session'); window.location.reload(); }

  return <><header className="topbar"><Logo /><div className="top-meta"><span>Maria21</span><span>{new Date().toLocaleDateString('es-DO', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' })}</span><button className="ghost" onClick={logout}><LogOut size={16} /> Cerrar sesion</button></div></header><main className="page"><section className="hero"><div><span className="eyebrow"><Sparkles size={16} /> Dashboard ejecutivo</span><h1>Gestion de mantenimientos vehiculares</h1><p>Vista centralizada de servicios, costos, placas, tendencias y reportes para Distribuidor Punto PAS.</p></div><button className="primary" onClick={() => setModalOpen(true)}><Plus size={18} /> Agregar registro</button></section>{notice && <div className="alert error">{notice}</div>}{loading && <div className="alert success">Cargando informacion desde Google Sheets...</div>}<FilterBar filters={filters} setFilters={setFilters} onClear={() => setFilters(emptyFilters)} /><section className="summary-grid"><SummaryCard title="Total registros" value={String(filtered.length)} hint="Servicios encontrados" icon={ClipboardList} /><SummaryCard title="Total pagado" value={money(totalPaid)} hint="Segun filtros aplicados" icon={Download} /><SummaryCard title="Vehiculo mayor gasto" value={byPlate[0]?.[0] || 'N/A'} hint={byPlate[0] ? money(byPlate[0][1]) : 'Sin datos'} icon={Car} /><SummaryCard title="Promedio por vehiculo" value={money(avgByVehicle)} hint="Costo operacional medio" icon={BarChart3} /></section><section className="service-grid">{SERVICES.map((service) => <ServiceCard key={service} service={service} records={filtered.filter((r) => r.tipoServicio === service)} onOpen={() => setActiveService(service)} />)}</section><section className="summary-grid narrow">{SERVICES.map((service) => <SummaryCard key={service} title={`Total ${service.toLowerCase()}`} value={String(filtered.filter((r) => r.tipoServicio === service).length)} hint={money(filtered.filter((r) => r.tipoServicio === service).reduce((s, r) => s + r.cantidadPago, 0))} icon={serviceMeta[service].icon} />)}<SummaryCard title="Ultimo servicio" value={last?.placa || 'N/A'} hint={last ? `${last.tipoServicio} · ${last.fecha}` : 'Sin registros'} icon={CalendarDays} /></section><Charts records={filtered} />{filters.placa && <DataTable records={filtered} title={`Consulta general por placa ${filters.placa.toUpperCase()}`} filters={filters} />}{!filters.placa && <DataTable records={filtered} title="Tabla general de servicios" filters={filters} />}</main><RecordModal open={modalOpen} onClose={() => setModalOpen(false)} onSaved={(record) => setRecords((current) => [record, ...current])} /><button className="fab" onClick={() => setModalOpen(true)}><Plus /></button></>;
}

function App() {
  const [logged, setLogged] = useState(sessionStorage.getItem('punto-pas-session') === 'Maria21');
  return logged ? <Dashboard /> : <Login onLogin={() => setLogged(true)} />;
}

createRoot(document.getElementById('root')!).render(<React.StrictMode><App /></React.StrictMode>);
