import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  Plus, Minus, Send, ChefHat, ClipboardList, Check, Flame, Settings2, Trash2, Clock,
  Receipt, Pencil, Ban, AlertTriangle, Lock, Wallet, Users, Banknote, CreditCard, Smartphone,
  KeyRound, UserCircle2, PauseCircle, ArrowRightLeft, Download, Eye, EyeOff,
  Gift, Volume2, MessageSquarePlus, PlayCircle, StopCircle,
} from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";
import { saveShared, subscribeShared } from "./firebase.js";

const NUM_MESAS = 4;

const DEFAULT_MENU = [
  { id: "m1", name: "Limonada Natural", price: 7000, cat: "Limonadas y Jugos" },
  { id: "m2", name: "Limonada de Coco", price: 9000, cat: "Limonadas y Jugos" },
  { id: "m3", name: "Limonada de Cereza", price: 9000, cat: "Limonadas y Jugos" },
  { id: "m4", name: "Jugos en Agua", price: 9000, cat: "Limonadas y Jugos" },
  { id: "m5", name: "Jugos en Leche", price: 10000, cat: "Limonadas y Jugos" },
  { id: "m6", name: "Tinto", price: 1500, cat: "Bebidas Calientes" },
  { id: "m7", name: "Americano", price: 3000, cat: "Bebidas Calientes" },
  { id: "m8", name: "Espresso Sencillo", price: 4000, cat: "Bebidas Calientes" },
  { id: "m9", name: "Espresso Doble", price: 6000, cat: "Bebidas Calientes" },
  { id: "m10", name: "Café con Leche", price: 4000, cat: "Bebidas Calientes" },
  { id: "m11", name: "Capuchino", price: 6000, cat: "Bebidas Calientes" },
  { id: "m12", name: "Capuchino Brownie", price: 7000, cat: "Bebidas Calientes" },
  { id: "m13", name: "Capuchino Caramelo", price: 7000, cat: "Bebidas Calientes" },
  { id: "m14", name: "Mocaccino", price: 7000, cat: "Bebidas Calientes" },
  { id: "m15", name: "Latte Machiato", price: 7000, cat: "Bebidas Calientes" },
  { id: "m16", name: "Colada de Café", price: 7000, cat: "Bebidas Calientes" },
  { id: "m17", name: "Colada Tradicional", price: 6000, cat: "Bebidas Calientes" },
  { id: "m18", name: "Affogato ⭐", price: 9000, cat: "Bebidas Calientes" },
  { id: "m19", name: "Migote (con licor)", price: 14000, cat: "Bebidas Calientes" },
  { id: "m20", name: "Milo", price: 5000, cat: "Bebidas Calientes" },
  { id: "m21", name: "Milo con Masmelo", price: 6000, cat: "Bebidas Calientes" },
  { id: "m22", name: "Chocolate con Masmelo", price: 6000, cat: "Bebidas Calientes" },
  { id: "m23", name: "Aromática frutas deshidratadas", price: 6000, cat: "Bebidas Calientes" },
  { id: "m24", name: "Aromática bolsa", price: 3000, cat: "Bebidas Calientes" },
  { id: "m25", name: "Carajillo (aguardiente, ron o brandy)", price: 8000, cat: "Bebidas Calientes" },
  { id: "m26", name: "Capuchino con Licor (whisky, amaretto o ron)", price: 10000, cat: "Bebidas Calientes" },
  { id: "m27", name: "Malteada pequeña", price: 10000, cat: "Bebidas Frías" },
  { id: "m28", name: "Malteada grande", price: 16000, cat: "Bebidas Frías" },
  { id: "m29", name: "Granizado (café, milo, mango biche, maracuyá)", price: 12000, cat: "Bebidas Frías" },
  { id: "m30", name: "Nevado fantasía de café", price: 13000, cat: "Bebidas Frías" },
  { id: "m31", name: "Nevado de arequipe (whisky o piña colada)", price: 14000, cat: "Bebidas Frías" },
  { id: "m32", name: "Milo frío", price: 9000, cat: "Bebidas Frías" },
  { id: "m33", name: "Capuchino frío", price: 11000, cat: "Bebidas Frías" },
  { id: "m34", name: "Soda michelada saborizada", price: 13000, cat: "Bebidas Frías" },
  { id: "m35", name: "Tamarindo michelado", price: 7000, cat: "Bebidas Frías" },
  { id: "m36", name: "Tamarindo envenenado", price: 12000, cat: "Bebidas Frías" },
  { id: "m37", name: "Gaseosa Tamarindo", price: 5000, cat: "Bebidas Frías" },
  { id: "m38", name: "Coca-Cola", price: 5000, cat: "Bebidas Frías" },
  { id: "m39", name: "Jugos Hit", price: 4000, cat: "Bebidas Frías" },
  { id: "m40", name: "Agua", price: 2000, cat: "Bebidas Frías" },
  { id: "m41", name: "Pony Malta", price: 4000, cat: "Bebidas Frías" },
  { id: "m42", name: "Gatorade", price: 5000, cat: "Bebidas Frías" },
  { id: "m43", name: "Bretaña", price: 4000, cat: "Bebidas Frías" },
  { id: "m44", name: "Águila Light", price: 5000, cat: "Cervezas" },
  { id: "m45", name: "Corona", price: 7000, cat: "Cervezas" },
  { id: "m46", name: "Coronita", price: 5000, cat: "Cervezas" },
  { id: "m47", name: "Poker", price: 5000, cat: "Cervezas" },
  { id: "m48", name: "Pilsen", price: 5000, cat: "Cervezas" },
  { id: "m49", name: "Club Colombia roja/dorada (consultar precio)", price: 0, cat: "Cervezas" },
  { id: "m50", name: "Cerveza Saborizada", price: 14000, cat: "Cervezas" },
  { id: "m51", name: "Cerveza Mango Biche", price: 15000, cat: "Cervezas" },
  { id: "m52", name: "Michelada Pilsen · Águila · Poker", price: 6000, cat: "Micheladas" },
  { id: "m53", name: "Michelada Corona · Club Colombia", price: 8000, cat: "Micheladas" },
  { id: "m54", name: "Michelada Bretaña", price: 6000, cat: "Micheladas" },
  { id: "m55", name: "Copa de Helado", price: 11000, cat: "Helados y Postres" },
  { id: "m56", name: "Cono de Helado", price: 4000, cat: "Helados y Postres" },
  { id: "m57", name: "Brownie con Helado", price: 10000, cat: "Helados y Postres" },
  { id: "m58", name: "Fresas con Crema", price: 16000, cat: "Helados y Postres" },
  { id: "m59", name: "Ensalada de Frutas", price: 15000, cat: "Helados y Postres" },
  { id: "m60", name: "Copa de Queso", price: 14000, cat: "Helados y Postres" },
  { id: "m61", name: "Payaso Plim Plim", price: 11000, cat: "Menú Infantil" },
  { id: "m62", name: "Pulpo", price: 11000, cat: "Menú Infantil" },
  { id: "m63", name: "Osito Vetta", price: 12000, cat: "Menú Infantil" },
  { id: "m64", name: "Granizado para Niños", price: 10000, cat: "Menú Infantil" },
];

const STATE_META = {
  pendiente: { label: "En cocina", color: "#C1442D" },
  preparando: { label: "Preparando", color: "#B98A2E" },
  listo: { label: "Listo", color: "#5B7553" },
  entregado: { label: "Entregado", color: "#8A9B87" },
  cancelado: { label: "Cancelado", color: "#9A9382" },
};

const METODO_META = {
  efectivo: { label: "Efectivo", icon: Banknote, color: "#5B7553" },
  tarjeta: { label: "Tarjeta", icon: CreditCard, color: "#2F6690" },
  transferencia: { label: "Transferencia", icon: Smartphone, color: "#8A611A" },
};

const PROPINA_OPCIONES = [0, 10, 15, 20];

/* ---------------- helpers ---------------- */

function uid() {
  return Math.random().toString(36).slice(2, 10);
}
function money(n) {
  return `$${Math.round(n || 0).toLocaleString("es-CO")}`;
}
function orderTotal(order) {
  return order.items.reduce((sum, it) => sum + it.price * it.qty, 0);
}
function cuentaOrders(orders, cuentaId) {
  return orders.filter((o) => o.cuentaId === cuentaId);
}
function cuentaSubtotal(orders, cuentaId) {
  return cuentaOrders(orders, cuentaId)
    .filter((o) => o.estado !== "cancelado")
    .reduce((sum, o) => sum + orderTotal(o), 0);
}
function cuentaItemCount(orders, cuentaId) {
  return cuentaOrders(orders, cuentaId)
    .filter((o) => o.estado !== "cancelado")
    .reduce((sum, o) => sum + o.items.reduce((a, i) => a + i.qty, 0), 0);
}
// Total a cobrar de una cuenta = subtotal - descuento + propina
function cuentaGranTotal(orders, cuenta) {
  const subtotal = cuentaSubtotal(orders, cuenta.id);
  const descuento = (cuenta.descuento && cuenta.descuento.monto) || 0;
  const propina = (cuenta.propina && cuenta.propina.monto) || 0;
  return Math.max(0, subtotal - descuento + propina);
}
function itemsAgregados(orders, cuentaId) {
  const map = {};
  cuentaOrders(orders, cuentaId)
    .filter((o) => o.estado !== "cancelado")
    .forEach((o) =>
      o.items.forEach((it) => {
        const key = `${it.id}@${it.price}`;
        if (!map[key]) map[key] = { key, id: it.id, name: it.name, price: it.price, qty: 0 };
        map[key].qty += it.qty;
      })
    );
  return Object.values(map);
}
function pagosDe(cuenta) {
  return cuenta.pagos || [];
}
function pagosTotal(cuenta) {
  return pagosDe(cuenta).reduce((s, p) => s + p.monto, 0);
}
function qtyAsignadaPorProducto(cuenta) {
  const map = {};
  pagosDe(cuenta)
    .filter((p) => p.tipo === "por_producto")
    .forEach((p) => (p.items || []).forEach((it) => (map[it.key] = (map[it.key] || 0) + it.qty)));
  return map;
}

function elapsedLabel(ts) {
  const secs = Math.floor((Date.now() - ts) / 1000);
  const m = Math.floor(secs / 60);
  const s = secs % 60;
  return `${m}:${String(s).padStart(2, "0")}`;
}
function urgencyColor(ts) {
  const mins = (Date.now() - ts) / 60000;
  if (mins >= 10) return "#C1442D";
  if (mins >= 5) return "#B98A2E";
  return "#5B7553";
}
function timeLabel(ts) {
  return new Date(ts).toLocaleTimeString("es-CO", { hour: "2-digit", minute: "2-digit" });
}
function dateTimeLabel(ts) {
  const d = new Date(ts);
  const date = d.toLocaleDateString("es-CO", { day: "2-digit", month: "2-digit", year: "numeric" });
  return `${date} · ${timeLabel(ts)}`;
}
function dayKey(ts) {
  const d = new Date(ts);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function monthKey(ts) {
  const d = new Date(ts);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}
function yearKey(ts) {
  return `${new Date(ts).getFullYear()}`;
}
function dayLabel(ts) {
  const d = new Date(ts);
  const s = d.toLocaleDateString("es-ES", { weekday: "long", day: "numeric", month: "long" });
  return s.charAt(0).toUpperCase() + s.slice(1);
}
function monthLabel(ts) {
  const d = new Date(ts);
  const s = d.toLocaleDateString("es-ES", { month: "long", year: "numeric" });
  return s.charAt(0).toUpperCase() + s.slice(1);
}
function hourLabel(h) {
  return `${String(h).padStart(2, "0")}:00`;
}

function buildHistory(cuentas, orders) {
  const pagadas = cuentas.filter((c) => c.estado === "pagada");
  const days = {};
  pagadas.forEach((c) => {
    const ordsC = cuentaOrders(orders, c.id).filter((o) => o.estado !== "cancelado");
    const total = cuentaGranTotal(orders, c);
    const itemCount = ordsC.reduce((s, o) => s + o.items.reduce((a, i) => a + i.qty, 0), 0);
    const pagos = pagosDe(c);
    const dk = dayKey(c.pagadaTs);
    if (!days[dk]) {
      days[dk] = {
        key: dk, ts: c.pagadaTs, total: 0, cuentasCount: 0, productos: {}, cuentas: [],
        porMetodo: { efectivo: 0, tarjeta: 0, transferencia: 0 }, porHora: {},
      };
    }
    const day = days[dk];
    day.total += total;
    day.cuentasCount += 1;
    if (c.pagadaTs < day.ts) day.ts = c.pagadaTs;
    pagos.forEach((p) => { day.porMetodo[p.metodoPago] = (day.porMetodo[p.metodoPago] || 0) + p.monto; });
    const hora = new Date(c.pagadaTs).getHours();
    day.porHora[hora] = (day.porHora[hora] || 0) + total;
    day.cuentas.push({ id: c.id, mesa: c.mesa, ts: c.ts, pagadaTs: c.pagadaTs, total, itemCount, pagos, mesero: c.mesero, propina: c.propina, descuento: c.descuento });
    ordsC.forEach((o) =>
      o.items.forEach((it) => {
        if (!day.productos[it.name]) day.productos[it.name] = { qty: 0, subtotal: 0 };
        day.productos[it.name].qty += it.qty;
        day.productos[it.name].subtotal += it.qty * it.price;
      })
    );
  });

  const months = {};
  Object.values(days).forEach((day) => {
    const mk = monthKey(day.ts);
    if (!months[mk]) months[mk] = { key: mk, ts: day.ts, total: 0, dayKeys: [] };
    months[mk].total += day.total;
    months[mk].dayKeys.push(day.key);
    if (day.ts < months[mk].ts) months[mk].ts = day.ts;
  });
  const years = {};
  Object.values(months).forEach((m) => {
    const yk = yearKey(m.ts);
    if (!years[yk]) years[yk] = { key: yk, total: 0, monthKeys: [] };
    years[yk].total += m.total;
    years[yk].monthKeys.push(m.key);
  });
  return { days, months, years };
}

function playBeep() {
  try {
    const Ctx = window.AudioContext || window.webkitAudioContext;
    const ctx = new Ctx();
    [880, 1180].forEach((freq, i) => {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = "sine";
      osc.frequency.value = freq;
      gain.gain.setValueAtTime(0.001, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.25, ctx.currentTime + i * 0.16 + 0.02);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + i * 0.16 + 0.22);
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start(ctx.currentTime + i * 0.16);
      osc.stop(ctx.currentTime + i * 0.16 + 0.24);
    });
    setTimeout(() => ctx.close(), 900);
  } catch (e) {
    // silencioso si el navegador bloquea audio
  }
}

function descargarRespaldo({ menu, orders, cuentas, config, turnos }) {
  const data = { version: 1, exportadoTs: Date.now(), menu, orders, cuentas, config, turnos };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const fecha = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = `${(config.businessName || "comandas").replace(/\s+/g, "_")}_respaldo_${fecha}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function exportarExcel(cuentas, orders, config) {
  const pagadas = cuentas.filter((c) => c.estado === "pagada").sort((a, b) => a.pagadaTs - b.pagadaTs);
  const filasPagos = [];
  pagadas.forEach((c) => {
    pagosDe(c).forEach((p) => {
      filasPagos.push({
        Fecha: new Date(c.pagadaTs).toLocaleDateString("es-CO"),
        Hora: timeLabel(p.ts),
        Mesa: c.mesa,
        Mesero: c.mesero || "",
        Etiqueta: p.etiqueta || "Cuenta completa",
        Método: METODO_META[p.metodoPago]?.label || p.metodoPago,
        Monto: p.monto,
        "Cobrado por": p.procesadoPor || "",
        Ítems: p.items ? p.items.map((it) => `${it.qty}x ${it.name}`).join(", ") : "",
      });
    });
  });

  const productoMap = {};
  pagadas.forEach((c) => {
    cuentaOrders(orders, c.id).filter((o) => o.estado !== "cancelado").forEach((o) =>
      o.items.forEach((it) => {
        if (!productoMap[it.name]) productoMap[it.name] = { qty: 0, subtotal: 0 };
        productoMap[it.name].qty += it.qty;
        productoMap[it.name].subtotal += it.qty * it.price;
      })
    );
  });
  const filasProductos = Object.entries(productoMap)
    .sort((a, b) => b[1].subtotal - a[1].subtotal)
    .map(([name, d]) => ({ Producto: name, Cantidad: d.qty, Subtotal: d.subtotal }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filasPagos), "Pagos");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filasProductos), "Productos");
  XLSX.writeFile(wb, `${(config.businessName || "comandas").replace(/\s+/g, "_")}_historial.xlsx`);
}

/* ---------------- App ---------------- */

const DEFAULT_CONFIG = { businessName: "Dulce & Café", waiters: [], usuarios: [], auditLog: [] };

export default function App() {
  const [role, setRole] = useState("mesero");
  const [menu, setMenu] = useState(DEFAULT_MENU);
  const [orders, setOrders] = useState([]);
  const [cuentas, setCuentas] = useState([]);
  const [config, setConfig] = useState(DEFAULT_CONFIG);
  const [turnos, setTurnos] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [, forceTick] = useState(0);
  const [meseroActual, setMeseroActual] = useState(null);
  const [cajaUsuario, setCajaUsuario] = useState(null); // { nombre, rol }
  const [lastSync, setLastSync] = useState(Date.now());

  // Los navegadores bloquean audio automático hasta que hay una interacción real
  // del usuario. Con el primer toque en cualquier parte de la app "despertamos"
  // el audio para que la campanita de cocina sí pueda sonar más adelante.
  useEffect(() => {
    const unlock = () => {
      try {
        const Ctx = window.AudioContext || window.webkitAudioContext;
        const ctx = new Ctx();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        gain.gain.value = 0;
        osc.connect(gain); gain.connect(ctx.destination);
        osc.start(); osc.stop(ctx.currentTime + 0.01);
        setTimeout(() => ctx.close(), 200);
      } catch (e) {}
      window.removeEventListener("touchstart", unlock);
      window.removeEventListener("click", unlock);
    };
    window.addEventListener("touchstart", unlock, { once: true });
    window.addEventListener("click", unlock, { once: true });
    return () => { window.removeEventListener("touchstart", unlock); window.removeEventListener("click", unlock); };
  }, []);

  // Sincronización en tiempo real con Firebase: cada vez que cualquier
  // dispositivo conectado cambia algo, todos los demás lo reciben al instante
  // (no hace falta esperar un sondeo cada 3 segundos como antes).
  useEffect(() => {
    const loadedKeys = new Set();
    const markLoaded = (key) => {
      loadedKeys.add(key);
      if (loadedKeys.size === 5) setLoaded(true);
    };

    const unsubs = [
      subscribeShared("menu", DEFAULT_MENU, (v) => { setMenu(v); setLastSync(Date.now()); markLoaded("menu"); }),
      subscribeShared("orders", [], (v) => { setOrders(v); setLastSync(Date.now()); markLoaded("orders"); }),
      subscribeShared("cuentas", [], (v) => { setCuentas(v); setLastSync(Date.now()); markLoaded("cuentas"); }),
      subscribeShared("config", DEFAULT_CONFIG, (v) => { setConfig({ ...DEFAULT_CONFIG, ...v }); setLastSync(Date.now()); markLoaded("config"); }),
      subscribeShared("turnos", [], (v) => { setTurnos(v); setLastSync(Date.now()); markLoaded("turnos"); }),
    ];
    const clock = setInterval(() => forceTick((t) => t + 1), 1000);
    return () => { unsubs.forEach((u) => u()); clearInterval(clock); };
  }, []);

  const persistOrders = useCallback(async (next) => { setOrders(next); await saveShared("orders", next); }, []);
  const persistMenu = useCallback(async (next) => { setMenu(next); await saveShared("menu", next); }, []);
  const persistCuentas = useCallback(async (next) => { setCuentas(next); await saveShared("cuentas", next); }, []);
  const persistConfig = useCallback(async (next) => { setConfig(next); await saveShared("config", next); }, []);
  const persistTurnos = useCallback(async (next) => { setTurnos(next); await saveShared("turnos", next); }, []);

  const turnoAbierto = turnos.find((t) => t.estado === "abierto") || null;
  const staleMs = Date.now() - lastSync;

  if (!loaded) {
    return (
      <div style={styles.loadingScreen}>
        <div style={styles.loadingStamp}>CARGANDO COMANDAS…</div>
      </div>
    );
  }

  return (
    <div style={styles.app}>
      <style>{fontImports}</style>
      <TopBar
        role={role} setRole={setRole} config={config} meseroActual={meseroActual}
        onCambiarMesero={() => setMeseroActual(null)} staleMs={staleMs}
        cajaUsuario={cajaUsuario} onCambiarCajaUsuario={() => setCajaUsuario(null)}
      />
      {role === "mesero" && (
        !meseroActual ? (
          <MeseroLogin config={config} persistConfig={persistConfig} onEntrar={setMeseroActual} />
        ) : (
          <MeseroView
            menu={menu} orders={orders} cuentas={cuentas} meseroActual={meseroActual} config={config}
            persistOrders={persistOrders} persistMenu={persistMenu} persistCuentas={persistCuentas} persistConfig={persistConfig}
          />
        )
      )}
      {role === "cocina" && <CocinaView orders={orders} persistOrders={persistOrders} />}
      {role === "caja" && (
        !cajaUsuario ? (
          <CajaLogin config={config} persistConfig={persistConfig} onLogin={setCajaUsuario} />
        ) : (
          <CajaView
            menu={menu} orders={orders} cuentas={cuentas} persistCuentas={persistCuentas}
            turnos={turnos} turnoAbierto={turnoAbierto} persistTurnos={persistTurnos}
            config={config} persistConfig={persistConfig} usuarioActual={cajaUsuario}
          />
        )
      )}
    </div>
  );
}

function TopBar({ role, setRole, config, meseroActual, onCambiarMesero, staleMs, cajaUsuario, onCambiarCajaUsuario }) {
  const desync = staleMs > 15000;
  return (
    <div style={styles.topBar}>
      <div style={styles.brand}>
        <span style={styles.brandMark}>◆</span>
        <span style={styles.brandText}>{(config.businessName || "LA COMANDA").toUpperCase()}</span>
        <span style={{ ...styles.syncDot, background: desync ? "#C1442D" : "#5B7553" }} title={desync ? "Sin sincronizar" : "Sincronizado"} />
      </div>
      {role === "mesero" && meseroActual && (
        <button style={styles.meseroChip} onClick={onCambiarMesero}>
          <UserCircle2 size={13} /> {meseroActual}
        </button>
      )}
      {role === "caja" && cajaUsuario && (
        <button style={styles.meseroChip} onClick={onCambiarCajaUsuario}>
          <UserCircle2 size={13} /> {cajaUsuario.nombre} · {cajaUsuario.rol === "admin" ? "Admin" : "Cajero"}
        </button>
      )}
      <div style={styles.roleSwitch}>
        <button onClick={() => setRole("mesero")} style={{ ...styles.roleBtn, ...(role === "mesero" ? styles.roleBtnActive : {}) }}>
          <ClipboardList size={15} strokeWidth={2.2} /> Mesero
        </button>
        <button onClick={() => setRole("cocina")} style={{ ...styles.roleBtn, ...(role === "cocina" ? styles.roleBtnActiveDark : {}) }}>
          <ChefHat size={15} strokeWidth={2.2} /> Cocina
        </button>
        <button onClick={() => setRole("caja")} style={{ ...styles.roleBtn, ...(role === "caja" ? styles.roleBtnActive : {}) }}>
          <Receipt size={15} strokeWidth={2.2} /> Caja
        </button>
      </div>
    </div>
  );
}

/* ---------------- Mesero: login e identificación ---------------- */

function MeseroLogin({ config, persistConfig, onEntrar }) {
  const waiters = config.waiters || [];
  const [seleccionado, setSeleccionado] = useState(null);
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [mostrarNuevo, setMostrarNuevo] = useState(false);
  const [nombreNuevo, setNombreNuevo] = useState("");
  const [pinNuevo, setPinNuevo] = useState("");
  const [pinNuevo2, setPinNuevo2] = useState("");

  const crearMesero = async () => {
    if (!nombreNuevo.trim()) { setError("Escribe tu nombre."); return; }
    if (waiters.some((w) => w.nombre.toLowerCase() === nombreNuevo.trim().toLowerCase())) { setError("Ya existe un mesero con ese nombre."); return; }
    if (pinNuevo.length < 4) { setError("El PIN debe tener al menos 4 dígitos."); return; }
    if (pinNuevo !== pinNuevo2) { setError("Los PIN no coinciden."); return; }
    const nuevo = { id: uid(), nombre: nombreNuevo.trim(), pin: pinNuevo };
    await persistConfig({ ...config, waiters: [...waiters, nuevo] });
    onEntrar(nuevo.nombre);
  };

  if (mostrarNuevo || waiters.length === 0) {
    return (
      <div style={styles.screen}>
        <div style={styles.loginWrap}>
          <div style={styles.pageEyebrow}>{waiters.length === 0 ? "PRIMER MESERO" : "MESERO NUEVO"}</div>
          <div style={styles.pageTitle}>Crea tu acceso</div>
          <div style={styles.cajaEmptyText}>Elige un PIN corto — lo usarás cada vez que entres, para que nadie más tome pedidos como si fuera tú.</div>
          <input style={styles.editInput} placeholder="Tu nombre" value={nombreNuevo} onChange={(e) => setNombreNuevo(e.target.value)} />
          <input style={{ ...styles.editInput, marginTop: 8 }} type="password" inputMode="numeric" placeholder="Nuevo PIN" value={pinNuevo} onChange={(e) => setPinNuevo(e.target.value.replace(/\D/g, ""))} />
          <input style={{ ...styles.editInput, marginTop: 8 }} type="password" inputMode="numeric" placeholder="Repite el PIN" value={pinNuevo2} onChange={(e) => setPinNuevo2(e.target.value.replace(/\D/g, ""))} />
          {error && <div style={styles.pinError}>{error}</div>}
          <button style={{ ...styles.addItemBtn, marginTop: 10 }} onClick={crearMesero}><KeyRound size={15} /> Crear y entrar</button>
          {waiters.length > 0 && (
            <button style={styles.backBtn} onClick={() => { setMostrarNuevo(false); setError(""); }}>← Ya tengo cuenta</button>
          )}
        </div>
      </div>
    );
  }

  if (!seleccionado) {
    return (
      <div style={styles.screen}>
        <div style={styles.loginWrap}>
          <div style={styles.pageEyebrow}>IDENTIFÍCATE</div>
          <div style={styles.pageTitle}>¿Quién eres?</div>
          <div style={styles.loginList}>
            {waiters.map((w) => (
              <button key={w.id} style={styles.loginBtn} onClick={() => { setSeleccionado(w); setError(""); }}>
                <UserCircle2 size={18} /> {w.nombre}
              </button>
            ))}
          </div>
          <button style={styles.inlineAddLink} onClick={() => setMostrarNuevo(true)}><Plus size={12} /> Soy nuevo, crear mi acceso</button>
        </div>
      </div>
    );
  }

  const intentar = () => {
    if (pin === seleccionado.pin) onEntrar(seleccionado.nombre);
    else { setError("PIN incorrecto."); setPin(""); }
  };

  return (
    <div style={styles.screen}>
      <div style={styles.loginWrap}>
        <button style={styles.backBtn} onClick={() => { setSeleccionado(null); setPin(""); setError(""); }}>← Elegir otro nombre</button>
        <div style={{ ...styles.pageEyebrow, marginTop: 10 }}>HOLA, {seleccionado.nombre.toUpperCase()}</div>
        <div style={styles.pageTitle}>Tu PIN</div>
        <input
          style={styles.editInput} type="password" inputMode="numeric" placeholder="PIN"
          value={pin} onChange={(e) => { setPin(e.target.value.replace(/\D/g, "")); setError(""); }}
          onKeyDown={(e) => e.key === "Enter" && intentar()}
        />
        {error && <div style={styles.pinError}>{error}</div>}
        <button style={{ ...styles.addItemBtn, marginTop: 10 }} onClick={intentar}><KeyRound size={15} /> Entrar</button>
      </div>
    </div>
  );
}

/* ---------------- Caja: usuarios individuales con rol ---------------- */

function CajaLogin({ config, persistConfig, onLogin }) {
  const usuarios = config.usuarios || [];
  const hayUsuarios = usuarios.length > 0;

  const [seleccionado, setSeleccionado] = useState(null);
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");

  // solo se usan cuando NO hay usuarios todavía (creación del administrador)
  const [nombreAdmin, setNombreAdmin] = useState("");
  const [pinAdmin, setPinAdmin] = useState("");
  const [pinAdmin2, setPinAdmin2] = useState("");

  if (!hayUsuarios) {
    const crear = async () => {
      if (!nombreAdmin.trim()) { setError("Escribe tu nombre."); return; }
      if (pinAdmin.length < 4) { setError("El PIN debe tener al menos 4 dígitos."); return; }
      if (pinAdmin !== pinAdmin2) { setError("Los PIN no coinciden."); return; }
      const admin = { id: uid(), nombre: nombreAdmin.trim(), pin: pinAdmin, rol: "admin" };
      const registro = { ts: Date.now(), accion: "Cuenta administradora creada", actor: nombreAdmin.trim() };
      await persistConfig({ ...config, usuarios: [admin], auditLog: [registro] });
      onLogin({ nombre: admin.nombre, rol: "admin" });
    };

    return (
      <div style={styles.screen}>
        <div style={styles.loginWrap}>
          <div style={styles.pageEyebrow}>CONFIGURAR CAJA</div>
          <div style={styles.pageTitle}>Crea el administrador</div>
          <div style={styles.cajaEmptyText}>Esta primera cuenta queda como Administrador — puede gestionar otros usuarios, aplicar descuentos y todo lo demás. Los siguientes usuarios los crea el administrador desde adentro.</div>
          <input style={styles.editInput} placeholder="Tu nombre" value={nombreAdmin} onChange={(e) => setNombreAdmin(e.target.value)} />
          <input style={{ ...styles.editInput, marginTop: 8 }} type="password" inputMode="numeric" placeholder="Nuevo PIN" value={pinAdmin} onChange={(e) => setPinAdmin(e.target.value.replace(/\D/g, ""))} />
          <input style={{ ...styles.editInput, marginTop: 8 }} type="password" inputMode="numeric" placeholder="Repite el PIN" value={pinAdmin2} onChange={(e) => setPinAdmin2(e.target.value.replace(/\D/g, ""))} />
          {error && <div style={styles.pinError}>{error}</div>}
          <button style={{ ...styles.addItemBtn, marginTop: 10 }} onClick={crear}><KeyRound size={15} /> Crear y entrar</button>
        </div>
      </div>
    );
  }

  // hay usuarios: elegir quién eres, luego pedir su PIN
  if (!seleccionado) {
    return (
      <div style={styles.screen}>
        <div style={styles.loginWrap}>
          <div style={styles.pageEyebrow}>ACCESO RESTRINGIDO</div>
          <div style={styles.pageTitle}>¿Quién eres?</div>
          <div style={styles.loginList}>
            {usuarios.map((u) => (
              <button key={u.id} style={styles.loginBtn} onClick={() => { setSeleccionado(u); setError(""); }}>
                <UserCircle2 size={18} /> {u.nombre}
                <span style={styles.rolBadge}>{u.rol === "admin" ? "Admin" : "Cajero"}</span>
              </button>
            ))}
          </div>
        </div>
      </div>
    );
  }

  const intentar = () => {
    if (pin === seleccionado.pin) onLogin({ nombre: seleccionado.nombre, rol: seleccionado.rol });
    else { setError("PIN incorrecto."); setPin(""); }
  };

  return (
    <div style={styles.screen}>
      <div style={styles.loginWrap}>
        <button style={styles.backBtn} onClick={() => { setSeleccionado(null); setPin(""); setError(""); }}>← Elegir otro usuario</button>
        <div style={{ ...styles.pageEyebrow, marginTop: 10 }}>HOLA, {seleccionado.nombre.toUpperCase()}</div>
        <div style={styles.pageTitle}>Tu PIN</div>
        <input
          style={styles.editInput} type="password" inputMode="numeric" placeholder="PIN"
          value={pin} onChange={(e) => { setPin(e.target.value.replace(/\D/g, "")); setError(""); }}
          onKeyDown={(e) => e.key === "Enter" && intentar()}
        />
        {error && <div style={styles.pinError}>{error}</div>}
        <button style={{ ...styles.addItemBtn, marginTop: 10 }} onClick={intentar}><KeyRound size={15} /> Entrar</button>
      </div>
    </div>
  );
}

/* ---------------- MESERO ---------------- */

function MeseroView({ menu, orders, cuentas, meseroActual, config, persistOrders, persistMenu, persistCuentas, persistConfig }) {
  const [selectedMesa, setSelectedMesa] = useState(null);
  const [draft, setDraft] = useState({}); // { itemId: { qty, nota } }
  const [note, setNote] = useState("");
  const [showMenuEditor, setShowMenuEditor] = useState(false);
  const [editingOrderId, setEditingOrderId] = useState(null);
  const [confirmCancelId, setConfirmCancelId] = useState(null);
  const [justSent, setJustSent] = useState(false);
  const [activeCat, setActiveCat] = useState(null);
  const [search, setSearch] = useState("");
  const [notingItemId, setNotingItemId] = useState(null);
  const [mesaAction, setMesaAction] = useState(null); // 'trasladar' | 'unir'
  const [avisosListos, setAvisosListos] = useState([]); // [{id, mesa, ts}]
  const listoIdsPrevRef = useRef(null); // null = todavía no inicializado

  // Aviso al mesero cuando cocina marca un pedido como "Listo"
  useEffect(() => {
    const listoAhora = orders.filter((o) => o.estado === "listo");
    const idsAhora = new Set(listoAhora.map((o) => o.id));

    if (listoIdsPrevRef.current === null) {
      // primera carga: solo establecemos la base, sin sonar (evita avisos falsos de pedidos viejos)
      listoIdsPrevRef.current = idsAhora;
      return;
    }

    const nuevos = listoAhora.filter((o) => !listoIdsPrevRef.current.has(o.id));
    if (nuevos.length > 0) {
      playBeep();
      const items = nuevos.map((o) => ({ id: o.id, mesa: o.mesa, ts: Date.now() }));
      setAvisosListos((prev) => [...items, ...prev]);
      items.forEach((it) => {
        setTimeout(() => setAvisosListos((prev) => prev.filter((a) => a.id !== it.id)), 15000);
      });
    }
    listoIdsPrevRef.current = idsAhora;
  }, [orders]);

  const cuentaAbiertaDe = (mesaNum) => cuentas.find((c) => c.mesa === mesaNum && c.estado === "abierta");

  const totalItems = Object.values(draft).reduce((a, b) => a + b.qty, 0);

  const resetDraft = () => { setDraft({}); setNote(""); setEditingOrderId(null); setNotingItemId(null); };
  const enterMesa = (n) => { resetDraft(); setActiveCat(null); setSearch(""); setConfirmCancelId(null); setMesaAction(null); setSelectedMesa(n); };
  const exitMesa = () => { resetDraft(); setActiveCat(null); setSearch(""); setMesaAction(null); setSelectedMesa(null); };

  const addQty = (id, d) => {
    setDraft((prev) => {
      const cur = prev[id] || { qty: 0, nota: "" };
      const nextQty = Math.max(0, cur.qty + d);
      const next = { ...prev };
      if (nextQty === 0) delete next[id];
      else next[id] = { ...cur, qty: nextQty };
      return next;
    });
  };
  const setItemNota = (id, text) => {
    setDraft((prev) => (prev[id] ? { ...prev, [id]: { ...prev[id], nota: text } } : prev));
  };

  const startEdit = (order) => {
    const d = {};
    order.items.forEach((it) => (d[it.id] = { qty: it.qty, nota: it.nota || "" }));
    setDraft(d);
    setNote(order.nota || "");
    setEditingOrderId(order.id);
  };

  const repetirUltimoPedido = (cuenta) => {
    const propios = cuentaOrders(orders, cuenta.id).sort((a, b) => b.ts - a.ts);
    if (propios.length === 0) return;
    const d = {};
    propios[0].items.forEach((it) => (d[it.id] = { qty: it.qty, nota: it.nota || "" }));
    setDraft(d); setNote(""); setEditingOrderId(null);
  };

  const cancelarPedido = async (orderId) => {
    const next = orders.map((o) => (o.id === orderId ? { ...o, estado: "cancelado" } : o));
    await persistOrders(next);
    setConfirmCancelId(null);
    if (editingOrderId === orderId) resetDraft();
  };

  const marcarEntregado = async (orderId) => {
    const next = orders.map((o) => (o.id === orderId ? { ...o, estado: "entregado" } : o));
    await persistOrders(next);
  };

  const enviarPedido = async (mesaNum) => {
    if (totalItems === 0) return;
    const items = Object.entries(draft).map(([id, { qty, nota }]) => {
      const m = menu.find((x) => x.id === id);
      return { id, name: m.name, price: m.price, qty, nota: (nota || "").trim() || undefined };
    });

    if (editingOrderId) {
      const next = orders.map((o) => (o.id === editingOrderId ? { ...o, items, nota: note.trim(), editado: true } : o));
      await persistOrders(next);
    } else {
      let cuenta = cuentaAbiertaDe(mesaNum);
      if (!cuenta) {
        cuenta = { id: uid(), mesa: mesaNum, ts: Date.now(), estado: "abierta", mesero: meseroActual };
        await persistCuentas([...cuentas, cuenta]);
      }
      const order = { id: uid(), mesa: mesaNum, cuentaId: cuenta.id, items, nota: note.trim(), estado: "pendiente", ts: Date.now() };
      await persistOrders([...orders, order]);
      setJustSent(true);
      setTimeout(() => setJustSent(false), 2200);
    }
    resetDraft();
  };

  const mesasLibres = (excluir) => Array.from({ length: NUM_MESAS }, (_, i) => i + 1).filter((n) => n !== excluir && !cuentaAbiertaDe(n));
  const mesasOcupadas = (excluir) => Array.from({ length: NUM_MESAS }, (_, i) => i + 1).filter((n) => n !== excluir && cuentaAbiertaDe(n));

  const trasladarA = async (cuenta, nuevaMesa) => {
    const nextOrders = orders.map((o) => (o.cuentaId === cuenta.id ? { ...o, mesa: nuevaMesa } : o));
    const nextCuentas = cuentas.map((c) => (c.id === cuenta.id ? { ...c, mesa: nuevaMesa } : c));
    await persistOrders(nextOrders);
    await persistCuentas(nextCuentas);
    setMesaAction(null);
    setSelectedMesa(nuevaMesa);
  };

  const unirCon = async (cuentaOrigen, cuentaDestino) => {
    const nextOrders = orders.map((o) =>
      o.cuentaId === cuentaOrigen.id ? { ...o, cuentaId: cuentaDestino.id, mesa: cuentaDestino.mesa } : o
    );
    const nextCuentas = cuentas.map((c) => (c.id === cuentaOrigen.id ? { ...c, estado: "fusionada", fusionadaEnId: cuentaDestino.id } : c));
    await persistOrders(nextOrders);
    await persistCuentas(nextCuentas);
    setMesaAction(null);
    setSelectedMesa(cuentaDestino.mesa);
  };

  if (showMenuEditor) {
    return <MenuEditor menu={menu} persistMenu={persistMenu} config={config} persistConfig={persistConfig} onClose={() => setShowMenuEditor(false)} />;
  }

  if (selectedMesa) {
    const categories = [...new Set(menu.map((m) => m.cat))];
    const cuenta = cuentaAbiertaDe(selectedMesa);
    const pending = cuenta ? cuentaOrders(orders, cuenta.id).filter((o) => o.estado !== "cancelado") : [];

    return (
      <div style={styles.screen}>
        <div style={styles.subHeader}>
          <button style={styles.backBtn} onClick={exitMesa}>← Mesas</button>
          <div style={styles.mesaTitleWrap}>
            <span style={styles.mesaEyebrow}>MESA</span>
            <span style={styles.mesaTitle}>{selectedMesa}</span>
            {cuenta && <span style={styles.cuentaOpenSub}>Cuenta abierta · {timeLabel(cuenta.ts)}</span>}
          </div>
          {cuenta ? (
            <button style={styles.mesaActionBtn} onClick={() => setMesaAction(mesaAction ? null : "menu")} title="Trasladar o unir esta mesa">
              <ArrowRightLeft size={15} />
            </button>
          ) : (
            <div style={{ width: 34 }} />
          )}
        </div>

        {mesaAction === "menu" && (
          <div style={styles.mesaActionPanel}>
            <button style={styles.modoChoiceBtn} onClick={() => setMesaAction("trasladar")}>
              <ArrowRightLeft size={15} /> Trasladar esta mesa a otra
            </button>
            <button style={styles.modoChoiceBtn} onClick={() => setMesaAction("unir")}>
              <Users size={15} /> Unir con otra mesa
            </button>
          </div>
        )}
        {mesaAction === "trasladar" && (
          <div style={styles.mesaActionPanel}>
            <div style={styles.closeConfirmText}>Elige la mesa libre destino:</div>
            <div style={styles.mesaPickerRow}>
              {mesasLibres(selectedMesa).length === 0 && <span style={styles.cajaEmptyText}>No hay mesas libres.</span>}
              {mesasLibres(selectedMesa).map((n) => (
                <button key={n} style={styles.mesaPickerBtn} onClick={() => trasladarA(cuenta, n)}>Mesa {n}</button>
              ))}
            </div>
            <button style={styles.cancelEditBtn} onClick={() => setMesaAction("menu")}>Volver</button>
          </div>
        )}
        {mesaAction === "unir" && (
          <div style={styles.mesaActionPanel}>
            <div style={styles.closeConfirmText}>Se unirá esta cuenta a la mesa que elijas (quedará todo en una sola cuenta):</div>
            <div style={styles.mesaPickerRow}>
              {mesasOcupadas(selectedMesa).length === 0 && <span style={styles.cajaEmptyText}>No hay otras mesas ocupadas.</span>}
              {mesasOcupadas(selectedMesa).map((n) => (
                <button key={n} style={styles.mesaPickerBtn} onClick={() => unirCon(cuenta, cuentaAbiertaDe(n))}>Mesa {n}</button>
              ))}
            </div>
            <button style={styles.cancelEditBtn} onClick={() => setMesaAction("menu")}>Volver</button>
          </div>
        )}

        {pending.length > 0 && (
          <div style={styles.ticketStripCol}>
            {pending.map((o) => {
              const puedeModificar = o.estado === "pendiente" || o.estado === "preparando";
              return (
                <div key={o.id} style={styles.miniTicketRow}>
                  <div style={styles.miniTicket}>
                    <span style={{ ...styles.stateDot, background: STATE_META[o.estado].color }} />
                    {o.items.reduce((a, i) => a + i.qty, 0)} ítems · {STATE_META[o.estado].label} · {money(orderTotal(o))}
                    <span style={styles.miniTicketTime}>{timeLabel(o.ts)}</span>
                    {o.editado && <span style={styles.editedTag}>editado</span>}
                  </div>
                  {puedeModificar && (
                    <div style={styles.miniTicketActions}>
                      {confirmCancelId === o.id ? (
                        <>
                          <span style={styles.confirmText}>¿Cancelar?</span>
                          <button style={styles.confirmYes} onClick={() => cancelarPedido(o.id)}>Sí</button>
                          <button style={styles.confirmNo} onClick={() => setConfirmCancelId(null)}>No</button>
                        </>
                      ) : (
                        <>
                          <button style={styles.iconBtn} onClick={() => startEdit(o)} title="Modificar"><Pencil size={13} /></button>
                          <button style={{ ...styles.iconBtn, color: "#C1442D" }} onClick={() => setConfirmCancelId(o.id)} title="Cancelar"><Ban size={13} /></button>
                        </>
                      )}
                    </div>
                  )}
                  {o.estado === "listo" && (
                    <button style={styles.entregadoBtn} onClick={() => marcarEntregado(o.id)}>
                      <Check size={12} /> Ya lo llevé
                    </button>
                  )}
                </div>
              );
            })}
            <div style={styles.miniTicketTotal}>Total cuenta hasta ahora: <b>{money(cuentaSubtotal(orders, cuenta.id))}</b></div>
          </div>
        )}

        {editingOrderId && (
          <div style={styles.editingBanner}><AlertTriangle size={14} /> Modificando pedido ya enviado — los cambios se reflejan en cocina.</div>
        )}
        {justSent && (
          <div style={styles.sentBanner}><Check size={14} /> Pedido enviado a cocina</div>
        )}
        {!editingOrderId && totalItems === 0 && cuenta && cuentaOrders(orders, cuenta.id).length > 0 && (
          <button style={styles.repeatBtn} onClick={() => repetirUltimoPedido(cuenta)}>↻ Repetir último pedido de esta cuenta</button>
        )}

        <div style={styles.catPillsWrap}>
          <input style={styles.searchInput} placeholder="🔍 Buscar producto…" value={search} onChange={(e) => setSearch(e.target.value)} />
          {!search.trim() && (
            <div style={styles.catPills}>
              {categories.map((cat) => {
                const qtyInCat = menu.filter((m) => m.cat === cat).reduce((s, m) => s + (draft[m.id]?.qty || 0), 0);
                const isActive = (activeCat || categories[0]) === cat;
                return (
                  <button key={cat} style={{ ...styles.catPill, ...(isActive ? styles.catPillActive : {}) }} onClick={() => setActiveCat(cat)}>
                    {cat}
                    {qtyInCat > 0 && <span style={styles.catPillBadge}>{qtyInCat}</span>}
                  </button>
                );
              })}
            </div>
          )}
        </div>

        <div style={styles.menuScroll}>
          {(search.trim()
            ? menu.filter((m) => m.name.toLowerCase().includes(search.trim().toLowerCase()))
            : menu.filter((m) => m.cat === (activeCat || categories[0]))
          ).map((item) => {
            const d = draft[item.id];
            const agotado = !!item.agotado;
            return (
              <div key={item.id} style={{ ...styles.menuRow, opacity: agotado ? 0.5 : 1 }}>
                <div>
                  <div style={styles.menuItemName}>{item.name} {agotado && <span style={styles.agotadoTag}>AGOTADO</span>}</div>
                  {search.trim() && <div style={styles.menuItemCat}>{item.cat}</div>}
                  <div style={styles.menuItemPrice}>{money(item.price)}</div>
                  {d && d.qty > 0 && (
                    notingItemId === item.id ? (
                      <input
                        autoFocus style={styles.itemNoteInput} placeholder="Nota para este producto…"
                        value={d.nota} onChange={(e) => setItemNota(item.id, e.target.value)}
                        onBlur={() => setNotingItemId(null)}
                      />
                    ) : (
                      <button style={styles.itemNoteBtn} onClick={() => setNotingItemId(item.id)}>
                        <MessageSquarePlus size={11} /> {d.nota ? d.nota : "agregar nota"}
                      </button>
                    )
                  )}
                </div>
                <div style={styles.stepper}>
                  <button style={styles.stepBtn} disabled={agotado} onClick={() => addQty(item.id, -1)}><Minus size={14} /></button>
                  <span style={styles.stepVal}>{d?.qty || 0}</span>
                  <button style={{ ...styles.stepBtn, ...styles.stepBtnPlus }} disabled={agotado} onClick={() => addQty(item.id, 1)}><Plus size={14} /></button>
                </div>
              </div>
            );
          })}
          {search.trim() && menu.filter((m) => m.name.toLowerCase().includes(search.trim().toLowerCase())).length === 0 && (
            <div style={styles.noResults}>Sin resultados para "{search}"</div>
          )}
        </div>

        <div style={styles.orderBar}>
          <input placeholder="Nota general para cocina (opcional)" value={note} onChange={(e) => setNote(e.target.value)} style={styles.noteInput} />
          <div style={{ display: "flex", gap: 8 }}>
            {editingOrderId && <button style={styles.cancelEditBtn} onClick={resetDraft}>Descartar</button>}
            <button style={{ ...styles.sendBtn, flex: 1, opacity: totalItems === 0 ? 0.4 : 1 }} disabled={totalItems === 0} onClick={() => enviarPedido(selectedMesa)}>
              {editingOrderId ? <Check size={16} /> : <Send size={16} />}
              {editingOrderId ? `Guardar cambios (${totalItems})` : `Enviar pedido ${totalItems > 0 ? `(${totalItems})` : ""}`}
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={styles.screen}>
      <div style={styles.mesaGridHeader}>
        <div>
          <div style={styles.pageEyebrow}>SALÓN</div>
          <div style={styles.pageTitle}>Mesas</div>
        </div>
        <button style={styles.editMenuBtn} onClick={() => setShowMenuEditor(true)}><Settings2 size={15} /> Editar menú</button>
      </div>

      {avisosListos.length > 0 && (
        <div style={styles.avisosListosWrap}>
          {avisosListos.map((a) => (
            <button key={a.id} style={styles.avisoListoBanner} onClick={() => { enterMesa(a.mesa); setAvisosListos((prev) => prev.filter((x) => x.id !== a.id)); }}>
              <Volume2 size={14} /> Mesa {a.mesa}: pedido listo para servir
              <span style={styles.avisoListoClose} onClick={(e) => { e.stopPropagation(); setAvisosListos((prev) => prev.filter((x) => x.id !== a.id)); }}>✕</span>
            </button>
          ))}
        </div>
      )}

      <div style={styles.mesaGrid}>
        {Array.from({ length: NUM_MESAS }, (_, i) => i + 1).map((n) => {
          const cuenta = cuentaAbiertaDe(n);
          const ordsMesa = cuenta ? cuentaOrders(orders, cuenta.id) : [];
          const hasActivePedido = ordsMesa.some((o) => o.estado === "pendiente" || o.estado === "preparando");
          const hasListoParaServir = !hasActivePedido && ordsMesa.some((o) => o.estado === "listo");
          const tagBg = hasActivePedido ? "#C1442D" : hasListoParaServir ? "#2F6690" : cuenta ? "#5B7553" : "#EDE7D8";
          const tagColor = cuenta ? "#F7F4EC" : "#8A8272";
          const tagText = hasActivePedido ? "En cocina" : hasListoParaServir ? "🔔 Listo para servir" : cuenta ? "Cuenta abierta" : "Libre";
          return (
            <button key={n} style={styles.mesaCard} onClick={() => enterMesa(n)}>
              <span style={styles.mesaCardNum}>{n}</span>
              <span style={{ ...styles.mesaCardTag, background: tagBg, color: tagColor }}>{tagText}</span>
              {cuenta && <span style={styles.mesaCardSince}>desde {timeLabel(cuenta.ts)}</span>}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function MenuEditor({ menu, persistMenu, config, persistConfig, onClose }) {
  const [name, setName] = useState("");
  const [price, setPrice] = useState("");
  const [cat, setCat] = useState("");
  const [businessName, setBusinessName] = useState(config.businessName || "");
  const [savedFlash, setSavedFlash] = useState(false);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);

  const addItem = () => {
    if (!name.trim() || !price) return;
    const item = { id: uid(), name: name.trim(), price: parseFloat(price) || 0, cat: cat.trim() || "Otros" };
    persistMenu([...menu, item]);
    setName(""); setPrice(""); setCat("");
  };
  const removeItem = (id) => { persistMenu(menu.filter((m) => m.id !== id)); setConfirmDeleteId(null); };
  const toggleAgotado = (id) => persistMenu(menu.map((m) => (m.id === id ? { ...m, agotado: !m.agotado } : m)));

  const guardarNombre = async () => {
    if (!businessName.trim()) return;
    await persistConfig({ ...config, businessName: businessName.trim() });
    setSavedFlash(true);
    setTimeout(() => setSavedFlash(false), 1800);
  };

  return (
    <div style={styles.screen}>
      <div style={styles.subHeader}>
        <button style={styles.backBtn} onClick={onClose}>← Mesas</button>
        <div style={styles.mesaTitleWrap}><span style={styles.mesaEyebrow}>AJUSTES</span><span style={styles.mesaTitle}>Menú</span></div>
        <div style={{ width: 70 }} />
      </div>
      <div style={styles.menuScroll}>
        <div style={styles.addItemCard}>
          <div style={styles.catLabel}>NOMBRE DEL NEGOCIO</div>
          <input style={styles.editInput} placeholder="Nombre del negocio" value={businessName} onChange={(e) => setBusinessName(e.target.value)} />
          <button style={styles.addItemBtn} onClick={guardarNombre}>
            {savedFlash ? <><Check size={15} /> Guardado</> : "Guardar nombre"}
          </button>
        </div>

        <div style={styles.addItemCard}>
          <input style={styles.editInput} placeholder="Nombre del plato" value={name} onChange={(e) => setName(e.target.value)} />
          <div style={{ display: "flex", gap: 8 }}>
            <input style={{ ...styles.editInput, flex: 1 }} placeholder="Precio" type="number" value={price} onChange={(e) => setPrice(e.target.value)} />
            <input style={{ ...styles.editInput, flex: 1 }} placeholder="Categoría" value={cat} onChange={(e) => setCat(e.target.value)} />
          </div>
          <button style={styles.addItemBtn} onClick={addItem}><Plus size={15} /> Agregar al menú</button>
        </div>

        {[...new Set(menu.map((m) => m.cat))].map((c) => (
          <div key={c} style={{ marginBottom: 18 }}>
            <div style={styles.catLabel}>{c}</div>
            {menu.filter((m) => m.cat === c).map((item) => (
              <div key={item.id} style={styles.editRow}>
                <button
                  style={{ ...styles.agotadoToggle, ...(item.agotado ? styles.agotadoToggleActive : {}) }}
                  onClick={() => toggleAgotado(item.id)}
                  title={item.agotado ? "Marcar disponible" : "Marcar agotado"}
                >
                  <PauseCircle size={14} />
                </button>
                <span style={{ ...styles.menuItemName, ...(item.agotado ? { textDecoration: "line-through", color: "#9A9382" } : {}) }}>{item.name}</span>
                <span style={styles.menuItemPrice}>{money(item.price)}</span>
                {confirmDeleteId === item.id ? (
                  <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    <button style={styles.confirmYes} onClick={() => removeItem(item.id)}>Sí</button>
                    <button style={styles.confirmNo} onClick={() => setConfirmDeleteId(null)}>No</button>
                  </div>
                ) : (
                  <button style={styles.deleteBtn} onClick={() => setConfirmDeleteId(item.id)}><Trash2 size={14} /></button>
                )}
              </div>
            ))}
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---------------- COCINA ---------------- */

function CocinaView({ orders, persistOrders }) {
  const active = orders.filter((o) => o.estado === "pendiente" || o.estado === "preparando").sort((a, b) => a.ts - b.ts);
  const prevCountRef = useRef(active.length);
  const [flash, setFlash] = useState(false);

  useEffect(() => {
    if (active.length > prevCountRef.current) {
      playBeep();
      setFlash(true);
      setTimeout(() => setFlash(false), 900);
    }
    prevCountRef.current = active.length;
  }, [active.length]);

  const advance = async (id) => {
    const next = orders.map((o) => (o.id !== id ? o : { ...o, estado: o.estado === "pendiente" ? "preparando" : "listo" }));
    await persistOrders(next);
  };

  return (
    <div style={{ ...styles.screenDark, ...(flash ? styles.screenFlash : {}) }}>
      <div style={styles.kdsHeader}>
        <div>
          <div style={styles.kdsEyebrow}>COCINA — PANTALLA EN VIVO <Volume2 size={11} style={{ verticalAlign: -2, marginLeft: 4 }} /></div>
          <div style={styles.kdsTitle}>Comandas activas</div>
        </div>
        <div style={styles.kdsCount}>{active.length}</div>
      </div>

      {active.length === 0 ? (
        <div style={styles.emptyKitchen}><Flame size={26} strokeWidth={1.5} color="#5C6670" /><div style={styles.emptyKitchenText}>Sin pedidos pendientes</div></div>
      ) : (
        <div style={styles.rail}>
          {active.map((o) => (
            <div key={o.id} style={{ ...styles.ticket, borderLeftColor: urgencyColor(o.ts) }}>
              <div style={styles.ticketHead}>
                <span style={styles.ticketMesa}>MESA {o.mesa}{o.editado && <span style={styles.editedBadge}>MODIFICADO</span>}</span>
                <span style={{ ...styles.ticketTime, color: urgencyColor(o.ts), fontWeight: 700 }}><Clock size={12} style={{ marginRight: 4, verticalAlign: -2 }} />{elapsedLabel(o.ts)}</span>
              </div>
              <div style={styles.ticketExactTime}>Pedido a las {timeLabel(o.ts)}</div>
              <div style={styles.ticketDivider} />
              {o.items.map((it, idx) => (
                <div key={idx}>
                  <div style={styles.ticketLine}><span style={styles.ticketQty}>{it.qty}×</span><span style={styles.ticketItemName}>{it.name}</span></div>
                  {it.nota && <div style={styles.ticketItemNote}>↳ {it.nota}</div>}
                </div>
              ))}
              {o.nota && <div style={styles.ticketNote}>"{o.nota}"</div>}
              <div style={styles.ticketDivider} />
              <button style={{ ...styles.ticketBtn, background: o.estado === "pendiente" ? "#B98A2E" : "#5B7553" }} onClick={() => advance(o.id)}>
                {o.estado === "pendiente" ? <><Flame size={14} /> Empezar preparación</> : <><Check size={14} /> Marcar listo</>}
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ---------------- CAJA ---------------- */

function CajaView({ menu, orders, cuentas, persistCuentas, turnos, turnoAbierto, persistTurnos, config, persistConfig, usuarioActual }) {
  const [tab, setTab] = useState("mesas");
  const [expandedId, setExpandedId] = useState(null);

  const abiertas = cuentas.filter((c) => c.estado === "abierta").sort((a, b) => a.mesa - b.mesa);
  const hoyKey = dayKey(Date.now());
  const pagadasHoy = cuentas.filter((c) => c.estado === "pagada" && dayKey(c.pagadaTs) === hoyKey);
  const vendidoHoy = pagadasHoy.reduce((s, c) => s + cuentaGranTotal(orders, c), 0);
  const porMetodoHoy = { efectivo: 0, tarjeta: 0, transferencia: 0 };
  pagadasHoy.forEach((c) => pagosDe(c).forEach((p) => { porMetodoHoy[p.metodoPago] = (porMetodoHoy[p.metodoPago] || 0) + p.monto; }));

  return (
    <div style={styles.screen}>
      <div style={styles.mesaGridHeader}>
        <div><div style={styles.pageEyebrow}>CAJA</div><div style={styles.pageTitle}>Control de consumo</div></div>
      </div>

      <div style={styles.cajaTabSwitch}>
        <button style={{ ...styles.cajaTabBtn, ...(tab === "turno" ? styles.cajaTabBtnActive : {}) }} onClick={() => setTab("turno")}>Turno</button>
        <button style={{ ...styles.cajaTabBtn, ...(tab === "mesas" ? styles.cajaTabBtnActive : {}) }} onClick={() => setTab("mesas")}>Mesas {abiertas.length > 0 ? `(${abiertas.length})` : ""}</button>
        <button style={{ ...styles.cajaTabBtn, ...(tab === "historial" ? styles.cajaTabBtnActive : {}) }} onClick={() => setTab("historial")}>Historial</button>
        <button style={{ ...styles.cajaTabBtn, ...(tab === "reportes" ? styles.cajaTabBtnActive : {}) }} onClick={() => setTab("reportes")}>Reportes</button>
      </div>

      {tab === "turno" && (
        <TurnoView turnos={turnos} turnoAbierto={turnoAbierto} persistTurnos={persistTurnos} orders={orders} cuentas={cuentas} config={config} persistConfig={persistConfig} usuarioActual={usuarioActual} menu={menu} />
      )}

      {tab === "mesas" && (
        <div style={styles.menuScroll}>
          <div style={styles.cajaTotalCard}>
            <span style={styles.cajaTotalLabel}>Vendido hoy ({pagadasHoy.length} {pagadasHoy.length === 1 ? "cuenta" : "cuentas"})</span>
            <span style={styles.cajaTotalValue}>{money(vendidoHoy)}</span>
            <div style={styles.metodoBreakdownRow}>
              {Object.entries(METODO_META).map(([key, meta]) => {
                const Icon = meta.icon;
                return (
                  <div key={key} style={styles.metodoBreakdownItem}>
                    <Icon size={13} color="#B3A891" /><span style={styles.metodoBreakdownLabel}>{meta.label}</span><span style={styles.metodoBreakdownValue}>{money(porMetodoHoy[key] || 0)}</span>
                  </div>
                );
              })}
            </div>
          </div>

          {!turnoAbierto && (
            <button style={{ ...styles.warnBanner, width: "100%", border: "none", cursor: "pointer", textAlign: "left" }} onClick={() => setTab("turno")}>
              <AlertTriangle size={14} /> Abre el turno para poder registrar cobros — toca aquí
            </button>
          )}

          {abiertas.length === 0 ? (
            <div style={styles.cajaEmptyText}>No hay mesas con cuenta abierta en este momento.</div>
          ) : (
            abiertas.map((c) => (
              <MesaCuentaCard
                key={c.id} cuenta={c} cuentasAll={cuentas} orders={orders} persistCuentas={persistCuentas}
                turnoAbierto={!!turnoAbierto} usuarioActual={usuarioActual}
                expanded={expandedId === c.id} onToggle={() => setExpandedId(expandedId === c.id ? null : c.id)}
              />
            ))
          )}
          {cuentas.some((c) => c.estado === "pagada") && (
            <div style={styles.goHistLink} onClick={() => setTab("historial")}>Ver historial completo por día, mes y año →</div>
          )}
        </div>
      )}

      {tab === "historial" && <HistorialView cuentas={cuentas} orders={orders} config={config} />}
      {tab === "reportes" && <ReportesView cuentas={cuentas} orders={orders} />}
    </div>
  );
}

/* ---------------- Turno / arqueo de caja ---------------- */

function TurnoView({ turnos, turnoAbierto, persistTurnos, orders, cuentas, config, persistConfig, usuarioActual, menu }) {
  const [base, setBase] = useState("");
  const [contado, setContado] = useState("");
  const [confirmandoCierre, setConfirmandoCierre] = useState(false);

  const mesasAbiertas = cuentas.filter((c) => c.estado === "abierta");
  const esAdmin = usuarioActual && usuarioActual.rol === "admin";

  const abrirTurno = async () => {
    const baseNum = parseFloat(base) || 0;
    const turno = { id: uid(), aperturaTs: Date.now(), baseEfectivo: baseNum, estado: "abierto" };
    await persistTurnos([...turnos, turno]);
    setBase("");
  };

  const efectivoEsperado = () => {
    if (!turnoAbierto) return 0;
    let total = turnoAbierto.baseEfectivo;
    cuentas.filter((c) => c.estado === "pagada" && c.pagadaTs >= turnoAbierto.aperturaTs).forEach((c) => {
      pagosDe(c).filter((p) => p.metodoPago === "efectivo").forEach((p) => (total += p.monto));
    });
    return total;
  };

  const esperado = efectivoEsperado();
  const contadoNum = parseFloat(contado) || 0;
  const diferencia = contadoNum - esperado;

  const cerrarTurno = async () => {
    const next = turnos.map((t) =>
      t.id === turnoAbierto.id
        ? { ...t, estado: "cerrado", cierreTs: Date.now(), esperado, contado: contadoNum, diferencia }
        : t
    );
    await persistTurnos(next);
    setContado(""); setConfirmandoCierre(false);
  };

  const cerrados = turnos.filter((t) => t.estado === "cerrado").sort((a, b) => b.cierreTs - a.cierreTs);

  return (
    <div style={styles.menuScroll}>
      {!turnoAbierto ? (
        <div style={styles.turnoBox}>
          <div style={styles.pageEyebrow}>TURNO CERRADO</div>
          <div style={styles.closeConfirmText}>Registra cuánto efectivo hay en caja para empezar el turno.</div>
          <input style={styles.editInput} type="number" placeholder="Base de efectivo inicial" value={base} onChange={(e) => setBase(e.target.value)} />
          <button style={{ ...styles.addItemBtn, marginTop: 8 }} onClick={abrirTurno}><PlayCircle size={15} /> Abrir turno</button>
        </div>
      ) : (
        <div style={styles.turnoBox}>
          <div style={styles.pageEyebrow}>TURNO ABIERTO</div>
          <div style={styles.closeConfirmText}>Abierto {dateTimeLabel(turnoAbierto.aperturaTs)} · base {money(turnoAbierto.baseEfectivo)}</div>
          <div style={styles.cajaTotalValue2}>{money(esperado)}<span style={styles.turnoEsperadoLabel}>efectivo esperado en caja ahora</span></div>

          {!confirmandoCierre ? (
            <>
              {mesasAbiertas.length > 0 && (
                <div style={styles.warnBanner}>
                  <AlertTriangle size={14} /> {mesasAbiertas.length} {mesasAbiertas.length === 1 ? "mesa tiene" : "mesas tienen"} cuenta abierta sin cobrar todavía.
                </div>
              )}
              <button style={styles.closeCajaBtn} onClick={() => setConfirmandoCierre(true)}><StopCircle size={15} /> Cerrar turno (arqueo)</button>
            </>
          ) : (
            <div style={styles.closeConfirmBox}>
              {mesasAbiertas.length > 0 && (
                <div style={styles.closeConfirmText}>
                  <b>Ojo:</b> {mesasAbiertas.length} {mesasAbiertas.length === 1 ? "mesa sigue" : "mesas siguen"} sin cobrar (mesa{mesasAbiertas.length > 1 ? "s" : ""} {mesasAbiertas.map((c) => c.mesa).join(", ")}). Si cierras ahora, ese dinero no quedará contado en este turno.
                </div>
              )}
              <div style={styles.closeConfirmText}>Cuenta el efectivo físico en caja y escríbelo aquí:</div>
              <input style={styles.editInput} type="number" placeholder="Efectivo contado" value={contado} onChange={(e) => setContado(e.target.value)} />
              {contado !== "" && (
                <div style={{ ...styles.diferenciaBox, background: diferencia === 0 ? "#E4EEE1" : diferencia > 0 ? "#FBEFD9" : "#FBE2DC" }}>
                  {diferencia === 0 ? "Cuadra exacto ✓" : diferencia > 0 ? `Sobran ${money(diferencia)}` : `Faltan ${money(Math.abs(diferencia))}`}
                </div>
              )}
              <div style={{ display: "flex", gap: 8 }}>
                <button style={styles.cancelEditBtn} onClick={() => setConfirmandoCierre(false)}>Cancelar</button>
                <button style={{ ...styles.sendBtn, flex: 1, opacity: contado !== "" ? 1 : 0.4 }} disabled={contado === ""} onClick={cerrarTurno}>
                  <Lock size={15} /> Confirmar cierre
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {cerrados.length > 0 && (
        <>
          <div style={{ ...styles.catLabel, marginTop: 20 }}>TURNOS ANTERIORES</div>
          {cerrados.map((t) => (
            <div key={t.id} style={styles.turnoRow}>
              <span style={styles.cierreDate}>{dateTimeLabel(t.aperturaTs)} → {timeLabel(t.cierreTs)}</span>
              <span style={{ ...styles.turnoDiff, color: t.diferencia === 0 ? "#5B7553" : "#C1442D" }}>
                {t.diferencia === 0 ? "Cuadró" : t.diferencia > 0 ? `+${money(t.diferencia)}` : money(t.diferencia)}
              </span>
            </div>
          ))}
        </>
      )}

      {esAdmin && (
        <RespaldoBox menu={menu} orders={orders} cuentas={cuentas} config={config} turnos={turnos} persistConfig={persistConfig} usuarioActual={usuarioActual} />
      )}

      <div style={{ ...styles.catLabel, marginTop: 24 }}>{esAdmin ? "USUARIOS Y SEGURIDAD" : "MI CUENTA"}</div>
      {esAdmin ? (
        <GestionUsuarios config={config} persistConfig={persistConfig} usuarioActual={usuarioActual} />
      ) : (
        <MiActividad config={config} usuarioActual={usuarioActual} />
      )}
    </div>
  );
}

function MiActividad({ config, usuarioActual }) {
  const auditLog = config.auditLog || [];
  const relevantes = auditLog.filter((r) => r.accion.includes(`"${usuarioActual.nombre}"`));
  return (
    <div>
      <div style={styles.lockedNote}><Lock size={11} /> Solo un administrador puede gestionar usuarios. Aquí ves las acciones registradas sobre tu propia cuenta.</div>
      {relevantes.length === 0 ? (
        <div style={styles.cajaEmptyText}>No hay acciones registradas sobre tu cuenta todavía.</div>
      ) : (
        relevantes.map((r, i) => (
          <div key={i} style={styles.turnoRow}>
            <span style={styles.cierreDate}>{dateTimeLabel(r.ts)}</span>
            <span style={styles.auditText}>{r.actor}: {r.accion}</span>
          </div>
        ))
      )}
    </div>
  );
}

function RespaldoBox({ menu, orders, cuentas, config, turnos, persistConfig, usuarioActual }) {
  const [restaurando, setRestaurando] = useState(false);
  const [confirmando, setConfirmando] = useState(false);
  const [archivo, setArchivo] = useState(null);
  const [error, setError] = useState("");
  const fileInputRef = useRef(null);

  const onArchivoElegido = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target.result);
        if (!data.menu || !data.orders || !data.cuentas || !data.config) throw new Error("Archivo incompleto");
        setArchivo(data);
        setConfirmando(true);
        setError("");
      } catch (err) {
        setError("Ese archivo no parece un respaldo válido.");
      }
    };
    reader.readAsText(file);
  };

  const confirmarRestauracion = async () => {
    if (!archivo) return;
    await saveShared("menu", archivo.menu);
    await saveShared("orders", archivo.orders);
    await saveShared("cuentas", archivo.cuentas);
    await saveShared("turnos", archivo.turnos || []);
    const nuevoConfig = {
      ...archivo.config,
      auditLog: [{ ts: Date.now(), accion: "Restauró un respaldo desde archivo", actor: usuarioActual.nombre }, ...(archivo.config.auditLog || [])].slice(0, 50),
    };
    await saveShared("config", nuevoConfig);
    window.location.reload();
  };

  return (
    <div style={{ ...styles.catLabel, marginTop: 24 }}>
      RESPALDO
      <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 8 }}>
        <button style={styles.modoChoiceBtn} onClick={() => descargarRespaldo({ menu, orders, cuentas, config, turnos })}>
          <Download size={15} /> Descargar copia de seguridad completa
        </button>
        {!restaurando ? (
          <button style={styles.inlineAddLink} onClick={() => setRestaurando(true)}>Restaurar desde un archivo de respaldo</button>
        ) : (
          <div style={styles.turnoBox}>
            <div style={styles.closeConfirmText}><b>Cuidado:</b> restaurar reemplaza todos los datos actuales (menú, pedidos, historial) por los del archivo. No se puede deshacer.</div>
            <input ref={fileInputRef} type="file" accept="application/json" onChange={onArchivoElegido} />
            {error && <div style={styles.pinError}>{error}</div>}
            {confirmando && archivo && (
              <div style={styles.closeConfirmBox}>
                <div style={styles.closeConfirmText}>
                  Este archivo es de "{archivo.config.businessName}", exportado el {dateTimeLabel(archivo.exportadoTs)}. ¿Restaurar y reemplazar todo lo actual?
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <button style={styles.cancelEditBtn} onClick={() => { setConfirmando(false); setArchivo(null); }}>Cancelar</button>
                  <button style={{ ...styles.sendBtn, flex: 1 }} onClick={confirmarRestauracion}><Lock size={15} /> Sí, restaurar</button>
                </div>
              </div>
            )}
            <button style={styles.cancelEditBtn} onClick={() => { setRestaurando(false); setArchivo(null); setConfirmando(false); setError(""); }}>Cerrar</button>
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------------- Cobro por mesa (con propina, descuento y división) ---------------- */

function GestionUsuarios({ config, persistConfig, usuarioActual }) {
  const usuarios = config.usuarios || [];
  const auditLog = config.auditLog || [];
  const [nombre, setNombre] = useState("");
  const [pin, setPin] = useState("");
  const [rol, setRol] = useState("cajero");
  const [error, setError] = useState("");
  const [resetId, setResetId] = useState(null);
  const [resetPin, setResetPin] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [pinesVisibles, setPinesVisibles] = useState({}); // { [usuarioId]: true }

  const registrar = (accion) => ({ ts: Date.now(), accion, actor: usuarioActual.nombre });

  const togglePinVisible = (u) => {
    const yaVisible = !!pinesVisibles[u.id];
    if (!yaVisible) {
      // queda registrado en el log de seguridad cada vez que alguien revela un PIN
      persistConfig({ ...config, auditLog: [registrar(`Vio el PIN de "${u.nombre}"`), ...auditLog].slice(0, 50) });
    }
    setPinesVisibles((prev) => ({ ...prev, [u.id]: !yaVisible }));
  };

  const agregarUsuario = async () => {
    setError("");
    if (!nombre.trim()) { setError("Escribe un nombre."); return; }
    if (usuarios.some((u) => u.nombre.toLowerCase() === nombre.trim().toLowerCase())) { setError("Ya existe un usuario con ese nombre."); return; }
    if (pin.length < 4) { setError("El PIN debe tener al menos 4 dígitos."); return; }
    const nuevo = { id: uid(), nombre: nombre.trim(), pin, rol };
    await persistConfig({
      ...config,
      usuarios: [...usuarios, nuevo],
      auditLog: [registrar(`Creó el usuario "${nuevo.nombre}" (${rol === "admin" ? "Admin" : "Cajero"})`), ...auditLog].slice(0, 50),
    });
    setNombre(""); setPin(""); setRol("cajero");
  };

  const cambiarRol = async (u) => {
    const nuevoRol = u.rol === "admin" ? "cajero" : "admin";
    const admins = usuarios.filter((x) => x.rol === "admin");
    if (u.rol === "admin" && admins.length <= 1) { setError("Debe quedar al menos un administrador."); return; }
    const next = usuarios.map((x) => (x.id === u.id ? { ...x, rol: nuevoRol } : x));
    await persistConfig({ ...config, usuarios: next, auditLog: [registrar(`Cambió el rol de "${u.nombre}" a ${nuevoRol === "admin" ? "Admin" : "Cajero"}`), ...auditLog].slice(0, 50) });
  };

  const resetearPin = async (u) => {
    if (resetPin.length < 4) { setError("El nuevo PIN debe tener al menos 4 dígitos."); return; }
    const next = usuarios.map((x) => (x.id === u.id ? { ...x, pin: resetPin } : x));
    await persistConfig({ ...config, usuarios: next, auditLog: [registrar(`Restableció el PIN de "${u.nombre}"`), ...auditLog].slice(0, 50) });
    setResetId(null); setResetPin(""); setError("");
  };

  const eliminarUsuario = async (u) => {
    const admins = usuarios.filter((x) => x.rol === "admin");
    if (u.rol === "admin" && admins.length <= 1) { setError("Debe quedar al menos un administrador."); return; }
    const next = usuarios.filter((x) => x.id !== u.id);
    await persistConfig({ ...config, usuarios: next, auditLog: [registrar(`Eliminó al usuario "${u.nombre}"`), ...auditLog].slice(0, 50) });
    setConfirmDeleteId(null);
  };

  return (
    <div>
      <div style={styles.catLabel}>EQUIPO</div>
      {usuarios.map((u) => (
        <div key={u.id} style={styles.usuarioRow}>
          <div style={styles.usuarioRowLeft}>
            <UserCircle2 size={16} color="#8A8272" />
            <span style={styles.pagoRowEtiqueta}>{u.nombre}</span>
            <span style={styles.rolBadge}>{u.rol === "admin" ? "Admin" : "Cajero"}</span>
            <button style={styles.pinRevealBtn} onClick={() => togglePinVisible(u)} title={pinesVisibles[u.id] ? "Ocultar PIN" : "Ver PIN"}>
              {pinesVisibles[u.id] ? (
                <><EyeOff size={12} /> {u.pin}</>
              ) : (
                <><Eye size={12} /> ••••</>
              )}
            </button>
          </div>
          <div style={styles.usuarioRowActions}>
            <button style={styles.inlineLink} onClick={() => cambiarRol(u)}>{u.rol === "admin" ? "hacer cajero" : "hacer admin"}</button>
            <button style={styles.inlineLink} onClick={() => { setResetId(resetId === u.id ? null : u.id); setResetPin(""); setError(""); }}>PIN</button>
            {u.nombre !== usuarioActual.nombre && (
              confirmDeleteId === u.id ? (
                <>
                  <button style={styles.confirmYes} onClick={() => eliminarUsuario(u)}>Sí</button>
                  <button style={styles.confirmNo} onClick={() => setConfirmDeleteId(null)}>No</button>
                </>
              ) : (
                <button style={{ ...styles.inlineLink, color: "#C1442D" }} onClick={() => setConfirmDeleteId(u.id)}>eliminar</button>
              )
            )}
          </div>
          {resetId === u.id && (
            <div style={styles.resetPinRow}>
              <input style={styles.editInput} type="password" inputMode="numeric" placeholder="Nuevo PIN para este usuario" value={resetPin} onChange={(e) => setResetPin(e.target.value.replace(/\D/g, ""))} />
              <button style={styles.addItemBtn} onClick={() => resetearPin(u)}><KeyRound size={14} /> Guardar</button>
            </div>
          )}
        </div>
      ))}

      <MeserosGestion config={config} persistConfig={persistConfig} registrar={registrar} auditLog={auditLog} />

      <div style={{ ...styles.catLabel, marginTop: 16 }}>AGREGAR USUARIO</div>
      <div style={styles.turnoBox}>
        <input style={styles.editInput} placeholder="Nombre" value={nombre} onChange={(e) => setNombre(e.target.value)} />
        <input style={styles.editInput} type="password" inputMode="numeric" placeholder="PIN (mínimo 4 dígitos)" value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} />
        <div style={styles.rolPillsRow}>
          <button style={{ ...styles.rolPill, ...(rol === "cajero" ? styles.rolPillActive : {}) }} onClick={() => setRol("cajero")}>Cajero</button>
          <button style={{ ...styles.rolPill, ...(rol === "admin" ? styles.rolPillActive : {}) }} onClick={() => setRol("admin")}>Admin</button>
        </div>
        {error && <div style={styles.pinError}>{error}</div>}
        <button style={styles.addItemBtn} onClick={agregarUsuario}><Plus size={15} /> Agregar usuario</button>
      </div>

      {auditLog.length > 0 && (
        <>
          <div style={{ ...styles.catLabel, marginTop: 16 }}>REGISTRO DE SEGURIDAD</div>
          {auditLog.map((r, i) => (
            <div key={i} style={styles.turnoRow}>
              <span style={styles.cierreDate}>{dateTimeLabel(r.ts)}</span>
              <span style={styles.auditText}>{r.actor}: {r.accion}</span>
            </div>
          ))}
        </>
      )}
    </div>
  );
}

function MeserosGestion({ config, persistConfig, registrar, auditLog }) {
  const waiters = config.waiters || [];
  const [resetId, setResetId] = useState(null);
  const [resetPin, setResetPin] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [pinesVisibles, setPinesVisibles] = useState({});
  const [error, setError] = useState("");

  const togglePinVisible = (w) => {
    const yaVisible = !!pinesVisibles[w.id];
    if (!yaVisible) {
      persistConfig({ ...config, auditLog: [registrar(`Vio el PIN del mesero "${w.nombre}"`), ...auditLog].slice(0, 50) });
    }
    setPinesVisibles((prev) => ({ ...prev, [w.id]: !yaVisible }));
  };

  const resetearPin = async (w) => {
    if (resetPin.length < 4) { setError("El nuevo PIN debe tener al menos 4 dígitos."); return; }
    const next = waiters.map((x) => (x.id === w.id ? { ...x, pin: resetPin } : x));
    await persistConfig({ ...config, waiters: next, auditLog: [registrar(`Restableció el PIN del mesero "${w.nombre}"`), ...auditLog].slice(0, 50) });
    setResetId(null); setResetPin(""); setError("");
  };

  const eliminarMesero = async (w) => {
    const next = waiters.filter((x) => x.id !== w.id);
    await persistConfig({ ...config, waiters: next, auditLog: [registrar(`Eliminó al mesero "${w.nombre}"`), ...auditLog].slice(0, 50) });
    setConfirmDeleteId(null);
  };

  if (waiters.length === 0) return null;

  return (
    <>
      <div style={{ ...styles.catLabel, marginTop: 16 }}>MESEROS</div>
      {waiters.map((w) => (
        <div key={w.id} style={styles.usuarioRow}>
          <div style={styles.usuarioRowLeft}>
            <UserCircle2 size={16} color="#8A8272" />
            <span style={styles.pagoRowEtiqueta}>{w.nombre}</span>
            <button style={styles.pinRevealBtn} onClick={() => togglePinVisible(w)} title={pinesVisibles[w.id] ? "Ocultar PIN" : "Ver PIN"}>
              {pinesVisibles[w.id] ? (<><EyeOff size={12} /> {w.pin}</>) : (<><Eye size={12} /> ••••</>)}
            </button>
          </div>
          <div style={styles.usuarioRowActions}>
            <button style={styles.inlineLink} onClick={() => { setResetId(resetId === w.id ? null : w.id); setResetPin(""); setError(""); }}>PIN</button>
            {confirmDeleteId === w.id ? (
              <>
                <button style={styles.confirmYes} onClick={() => eliminarMesero(w)}>Sí</button>
                <button style={styles.confirmNo} onClick={() => setConfirmDeleteId(null)}>No</button>
              </>
            ) : (
              <button style={{ ...styles.inlineLink, color: "#C1442D" }} onClick={() => setConfirmDeleteId(w.id)}>eliminar</button>
            )}
          </div>
          {resetId === w.id && (
            <div style={styles.resetPinRow}>
              <input style={styles.editInput} type="password" inputMode="numeric" placeholder="Nuevo PIN para este mesero" value={resetPin} onChange={(e) => setResetPin(e.target.value.replace(/\D/g, ""))} />
              <button style={styles.addItemBtn} onClick={() => resetearPin(w)}><KeyRound size={14} /> Guardar</button>
            </div>
          )}
          {error && resetId === w.id && <div style={styles.pinError}>{error}</div>}
        </div>
      ))}
    </>
  );
}

function MetodoPills({ selected, onSelect }) {
  return (
    <div style={styles.metodoPillsRow}>
      {Object.entries(METODO_META).map(([key, meta]) => {
        const Icon = meta.icon;
        const isSel = selected === key;
        return (
          <button key={key} style={{ ...styles.metodoPill, ...(isSel ? { background: meta.color, color: "#F7F4EC", border: `1px solid ${meta.color}` } : {}) }} onClick={() => onSelect(key)}>
            <Icon size={14} /> {meta.label}
          </button>
        );
      })}
    </div>
  );
}

function MesaCuentaCard({ cuenta, cuentasAll, orders, persistCuentas, turnoAbierto, usuarioActual, expanded, onToggle }) {
  const subtotal = cuentaSubtotal(orders, cuenta.id);
  const total = cuentaGranTotal(orders, cuenta);
  const items = cuentaItemCount(orders, cuenta.id);
  const ords = cuentaOrders(orders, cuenta.id).filter((o) => o.estado !== "cancelado");
  const pagos = pagosDe(cuenta);
  const cobrado = pagosTotal(cuenta);
  const restante = Math.max(0, total - cobrado);
  const enCurso = pagos.length > 0 && restante > 0;

  const [modoLocal, setModoLocal] = useState(null);
  const [metodoSel, setMetodoSel] = useState(null);
  const [numPartes, setNumPartes] = useState(2);
  const [cart, setCart] = useState({});
  const [etiqueta, setEtiqueta] = useState("");
  const [propinaCustom, setPropinaCustom] = useState("");
  const [mostrarDescuento, setMostrarDescuento] = useState(false);
  const [descMonto, setDescMonto] = useState("");
  const [descMotivo, setDescMotivo] = useState("");

  useEffect(() => { if (!expanded) { setModoLocal(null); setMetodoSel(null); setCart({}); setEtiqueta(""); setMostrarDescuento(false); } }, [expanded]);

  const effectiveModo = cuenta.splitMode || modoLocal;

  async function patchCuenta(patch) {
    const next = cuentasAll.map((c) => (c.id === cuenta.id ? { ...c, ...patch } : c));
    await persistCuentas(next);
  }

  async function guardarPago(pago) {
    const next = cuentasAll.map((c) => {
      if (c.id !== cuenta.id) return c;
      const nuevosPagos = [...pagosDe(c), pago];
      const totalPagado = nuevosPagos.reduce((s, p) => s + p.monto, 0);
      const completo = totalPagado >= total - 1;
      return { ...c, pagos: nuevosPagos, splitMode: pago.tipo === "unico" ? c.splitMode : effectiveModo, estado: completo ? "pagada" : "abierta", pagadaTs: completo ? Date.now() : c.pagadaTs };
    });
    await persistCuentas(next);
    setMetodoSel(null); setCart({}); setEtiqueta("");
  }

  const esAdmin = usuarioActual && usuarioActual.rol === "admin";

  const setPropina = async (pct) => {
    const monto = pct === "custom" ? Math.round(parseFloat(propinaCustom) || 0) : Math.round((subtotal * pct) / 100);
    await patchCuenta({ propina: { pct: pct === "custom" ? null : pct, monto } });
  };
  const aplicarDescuento = async () => {
    if (!esAdmin) return;
    const monto = Math.round(parseFloat(descMonto) || 0);
    if (monto <= 0) return;
    await patchCuenta({ descuento: { monto, motivo: descMotivo.trim() || "Sin motivo especificado", aplicadoPor: usuarioActual.nombre } });
    setMostrarDescuento(false); setDescMonto(""); setDescMotivo("");
  };
  const quitarDescuento = async () => { if (esAdmin) patchCuenta({ descuento: null }); };

  const confirmarUnico = () => { if (!metodoSel) return; guardarPago({ id: uid(), ts: Date.now(), metodoPago: metodoSel, monto: total, tipo: "unico", etiqueta: null, items: itemsAgregados(orders, cuenta.id), procesadoPor: usuarioActual.nombre }); };

  const parteMonto = numPartes > 0 ? Math.round(total / numPartes) : 0;
  const partesRegistradas = pagos.filter((p) => p.tipo === "equitativo").length;
  const parteActual = partesRegistradas + 1;
  const esUltimaParte = parteActual >= numPartes;
  const montoParteActual = esUltimaParte ? total - parteMonto * (numPartes - 1) : parteMonto;
  const confirmarParte = () => { if (!metodoSel) return; guardarPago({ id: uid(), ts: Date.now(), metodoPago: metodoSel, monto: montoParteActual, tipo: "equitativo", etiqueta: `Parte ${parteActual} de ${numPartes}`, procesadoPor: usuarioActual.nombre }); };

  const agregados = itemsAgregados(orders, cuenta.id);
  const asignada = qtyAsignadaPorProducto(cuenta);
  const pendientes = agregados.map((it) => ({ ...it, qtyPend: it.qty - (asignada[it.key] || 0) })).filter((it) => it.qtyPend > 0);
  const cartMonto = Object.entries(cart).reduce((s, [key, qty]) => { const it = agregados.find((a) => a.key === key); return s + (it ? it.price * qty : 0); }, 0);
  const numPersonaSugerido = pagos.filter((p) => p.tipo === "por_producto").length + 1;

  const addToCart = (key, d) => {
    setCart((prev) => {
      const it = agregados.find((a) => a.key === key);
      const pend = it.qty - (asignada[key] || 0);
      const cur = prev[key] || 0;
      const next = Math.max(0, Math.min(pend, cur + d));
      const copy = { ...prev, [key]: next };
      if (next === 0) delete copy[key];
      return copy;
    });
  };
  const confirmarProducto = () => {
    if (!metodoSel || Object.keys(cart).length === 0) return;
    const itemsSel = Object.entries(cart).map(([key, qty]) => { const it = agregados.find((a) => a.key === key); return { key, name: it.name, price: it.price, qty }; });
    guardarPago({ id: uid(), ts: Date.now(), metodoPago: metodoSel, monto: cartMonto, tipo: "por_producto", etiqueta: etiqueta.trim() || `Persona ${numPersonaSugerido}`, items: itemsSel, procesadoPor: usuarioActual.nombre });
  };
  const cobrarResto = () => {
    if (!metodoSel || restante <= 0) return;
    const itemsSel = pendientes.map((it) => ({ key: it.key, name: it.name, price: it.price, qty: it.qtyPend }));
    guardarPago({ id: uid(), ts: Date.now(), metodoPago: metodoSel, monto: restante, tipo: "por_producto", etiqueta: etiqueta.trim() || `Persona ${numPersonaSugerido} (resto)`, items: itemsSel, procesadoPor: usuarioActual.nombre });
  };

  return (
    <div style={styles.mesaCuentaCard}>
      <button style={styles.mesaCuentaHead} onClick={onToggle}>
        <div style={styles.mesaCuentaLeft}>
          <span style={styles.mesaCuentaNum}>Mesa {cuenta.mesa}</span>
          <span style={styles.mesaCuentaSince}>Abierta desde {dateTimeLabel(cuenta.ts)}{cuenta.mesero ? ` · ${cuenta.mesero}` : ""}</span>
          {enCurso && <span style={styles.enCursoTag}>Cobro en curso · {money(cobrado)} de {money(total)}</span>}
        </div>
        <div style={styles.mesaCuentaRight}><span style={styles.mesaCuentaItems}>{items} ítems</span><span style={styles.mesaCuentaTotal}>{money(total)}</span></div>
      </button>

      {expanded && (
        <div style={styles.mesaCuentaBody}>
          {ords.slice().sort((a, b) => a.ts - b.ts).map((o) => (
            <div key={o.id} style={styles.mesaCuentaOrderRow}>
              <span style={styles.mesaCuentaOrderTime}>{timeLabel(o.ts)}</span>
              <span style={styles.mesaCuentaOrderItems}>{o.items.map((it) => `${it.qty}× ${it.name}${it.nota ? ` (${it.nota})` : ""}`).join(", ")}</span>
              <span style={styles.mesaCuentaOrderTotal}>{money(orderTotal(o))}</span>
            </div>
          ))}

          <div style={styles.subtotalRow}>
            <span>Subtotal</span><span>{money(subtotal)}</span>
          </div>
          {cuenta.descuento ? (
            <div style={styles.subtotalRow}>
              <span>
                Descuento ({cuenta.descuento.motivo}{cuenta.descuento.aplicadoPor ? ` · ${cuenta.descuento.aplicadoPor}` : ""})
                {esAdmin && <button style={styles.inlineLink} onClick={quitarDescuento}>quitar</button>}
              </span>
              <span style={{ color: "#C1442D" }}>−{money(cuenta.descuento.monto)}</span>
            </div>
          ) : pagos.length === 0 ? (
            !esAdmin ? (
              <div style={styles.lockedNote}><Lock size={11} /> Solo un administrador puede aplicar descuentos o cortesías.</div>
            ) : mostrarDescuento ? (
              <div style={styles.descuentoBox}>
                <input style={styles.editInput} type="number" placeholder="Monto del descuento" value={descMonto} onChange={(e) => setDescMonto(e.target.value)} />
                <input style={{ ...styles.editInput, marginTop: 6 }} placeholder="Motivo (ej: cortesía, cliente frecuente)" value={descMotivo} onChange={(e) => setDescMotivo(e.target.value)} />
                <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
                  <button style={styles.cancelEditBtn} onClick={() => setMostrarDescuento(false)}>Cancelar</button>
                  <button style={{ ...styles.sendBtn, flex: 1 }} onClick={aplicarDescuento}><Gift size={14} /> Aplicar</button>
                </div>
              </div>
            ) : (
              <button style={styles.inlineAddLink} onClick={() => setMostrarDescuento(true)}><Gift size={12} /> Aplicar descuento o cortesía</button>
            )
          ) : null}

          {cuenta.propina ? (
            <div style={styles.subtotalRow}><span>Propina {cuenta.propina.pct != null ? `(${cuenta.propina.pct}%)` : ""}</span><span>{money(cuenta.propina.monto)}</span></div>
          ) : pagos.length === 0 && !mostrarDescuento ? (
            <div style={styles.propinaBox}>
              <span style={styles.closeConfirmText}>Propina:</span>
              <div style={styles.propinaPillsRow}>
                {PROPINA_OPCIONES.map((p) => (
                  <button key={p} style={styles.propinaPill} onClick={() => setPropina(p)}>{p}%</button>
                ))}
                <input style={styles.propinaCustomInput} type="number" placeholder="$" value={propinaCustom} onChange={(e) => setPropinaCustom(e.target.value)} />
                <button style={styles.propinaCustomBtn} onClick={() => setPropina("custom")} disabled={!propinaCustom}>Ok</button>
              </div>
            </div>
          ) : null}

          <div style={{ ...styles.subtotalRow, ...styles.subtotalRowFinal }}><span>Total a cobrar</span><span>{money(total)}</span></div>

          {pagos.length > 0 && (
            <div style={styles.pagosRegistradosBox}>
              <div style={styles.pagosRegistradosLabel}>PAGOS REGISTRADOS</div>
              {pagos.map((p) => {
                const meta = METODO_META[p.metodoPago]; const Icon = meta.icon;
                return (
                  <div key={p.id} style={styles.pagoRow}>
                    <Icon size={13} color={meta.color} />
                    <div style={styles.pagoRowMid}>
                      <span style={styles.pagoRowEtiqueta}>{p.etiqueta || "Cuenta completa"}</span>
                      {p.tipo === "por_producto" && <span style={styles.pagoRowItems}>{p.items.map((it) => `${it.qty}× ${it.name}`).join(", ")}</span>}
                      <span style={styles.pagoRowTime}>{timeLabel(p.ts)} · {meta.label}{p.procesadoPor ? ` · cobró ${p.procesadoPor}` : ""}</span>
                    </div>
                    <span style={styles.pagoRowMonto}>{money(p.monto)}</span>
                  </div>
                );
              })}
            </div>
          )}

          {!turnoAbierto && restante > 0 && (
            <div style={styles.warnBanner}><AlertTriangle size={14} /> Abre el turno para poder cobrar esta mesa.</div>
          )}

          {restante > 0 && turnoAbierto && (
            <div style={styles.cobroBox}>
              {!effectiveModo ? (
                <>
                  <div style={styles.closeConfirmText}>¿Cómo quieren pagar?</div>
                  <div style={styles.modoChoiceRow}>
                    <button style={styles.modoChoiceBtn} onClick={() => setModoLocal("unico")}><Wallet size={16} /> Pago único</button>
                    <button style={styles.modoChoiceBtn} onClick={() => setModoLocal("igual")}><Users size={16} /> Partes iguales</button>
                    <button style={styles.modoChoiceBtn} onClick={() => setModoLocal("producto")}><Receipt size={16} /> Por producto</button>
                  </div>
                </>
              ) : effectiveModo === "unico" ? (
                <>
                  <div style={styles.closeConfirmText}>Cobrar {money(total)}. Elige el método de pago:</div>
                  <MetodoPills selected={metodoSel} onSelect={setMetodoSel} />
                  <div style={{ display: "flex", gap: 8 }}>
                    <button style={styles.cancelEditBtn} onClick={() => setModoLocal(null)}>Volver</button>
                    <button style={{ ...styles.sendBtn, flex: 1, opacity: metodoSel ? 1 : 0.4 }} disabled={!metodoSel} onClick={confirmarUnico}><Lock size={15} /> Confirmar cobro</button>
                  </div>
                </>
              ) : effectiveModo === "igual" ? (
                <>
                  {partesRegistradas === 0 ? (
                    <div style={styles.partesStepperRow}>
                      <span style={styles.closeConfirmText}>Dividir entre:</span>
                      <div style={styles.stepper}>
                        <button style={styles.stepBtn} onClick={() => setNumPartes((n) => Math.max(2, n - 1))}><Minus size={14} /></button>
                        <span style={styles.stepVal}>{numPartes}</span>
                        <button style={{ ...styles.stepBtn, ...styles.stepBtnPlus }} onClick={() => setNumPartes((n) => Math.min(12, n + 1))}><Plus size={14} /></button>
                      </div>
                      <span style={styles.closeConfirmText}>personas</span>
                    </div>
                  ) : (
                    <div style={styles.closeConfirmText}>Dividido en {numPartes} partes de {money(parteMonto)} cada una.</div>
                  )}
                  <div style={styles.closeConfirmText}>Parte {parteActual} de {numPartes} — {money(montoParteActual)}. Método de pago:</div>
                  <MetodoPills selected={metodoSel} onSelect={setMetodoSel} />
                  <div style={{ display: "flex", gap: 8 }}>
                    {partesRegistradas === 0 && <button style={styles.cancelEditBtn} onClick={() => setModoLocal(null)}>Volver</button>}
                    <button style={{ ...styles.sendBtn, flex: 1, opacity: metodoSel ? 1 : 0.4 }} disabled={!metodoSel} onClick={confirmarParte}><Check size={15} /> Registrar parte {parteActual}</button>
                  </div>
                </>
              ) : (
                <>
                  <div style={styles.closeConfirmText}>Toca los productos de esta persona:</div>
                  {pendientes.map((it) => (
                    <div key={it.key} style={styles.productoSplitRow}>
                      <div><div style={styles.menuItemName}>{it.name}</div><div style={styles.menuItemPrice}>{money(it.price)} · quedan {it.qtyPend}</div></div>
                      <div style={styles.stepper}>
                        <button style={styles.stepBtn} onClick={() => addToCart(it.key, -1)}><Minus size={14} /></button>
                        <span style={styles.stepVal}>{cart[it.key] || 0}</span>
                        <button style={{ ...styles.stepBtn, ...styles.stepBtnPlus }} onClick={() => addToCart(it.key, 1)}><Plus size={14} /></button>
                      </div>
                    </div>
                  ))}
                  <input style={styles.editInput} placeholder={`Nombre de la persona (opcional, ej. "Persona ${numPersonaSugerido}")`} value={etiqueta} onChange={(e) => setEtiqueta(e.target.value)} />
                  <div style={styles.closeConfirmText}>Subtotal de esta persona: <b>{money(cartMonto)}</b></div>
                  <MetodoPills selected={metodoSel} onSelect={setMetodoSel} />
                  <div style={{ display: "flex", gap: 8 }}>
                    {pagos.length === 0 && <button style={styles.cancelEditBtn} onClick={() => setModoLocal(null)}>Volver</button>}
                    <button style={{ ...styles.sendBtn, flex: 1, opacity: metodoSel && cartMonto > 0 ? 1 : 0.4 }} disabled={!metodoSel || cartMonto === 0} onClick={confirmarProducto}><Check size={15} /> Registrar pago — {money(cartMonto)}</button>
                  </div>
                  {Object.keys(cart).length === 0 && (
                    <button style={{ ...styles.repeatBtn, margin: 0, opacity: metodoSel ? 1 : 0.5 }} disabled={!metodoSel} onClick={cobrarResto}>Cobrar todo lo que queda — {money(restante)}</button>
                  )}
                </>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ---------------- Historial ---------------- */

function HistorialView({ cuentas, orders, config }) {
  const [view, setView] = useState("years");
  const [year, setYear] = useState(null);
  const [month, setMonth] = useState(null);
  const [day, setDay] = useState(null);
  const [expandedCuentaId, setExpandedCuentaId] = useState(null);

  const { days, months, years } = buildHistory(cuentas, orders);

  if (Object.keys(days).length === 0) {
    return (
      <div style={styles.menuScroll}>
        <div style={styles.cajaEmptyText}>Todavía no hay cuentas cobradas. Cuando cobres una mesa, va a aparecer aquí organizada por día, con fecha y hora exactas.</div>
      </div>
    );
  }

  if (view === "years") {
    const list = Object.values(years).sort((a, b) => b.key - a.key);
    return (
      <div style={styles.menuScroll}>
        <div style={styles.rowBetween}>
          <div style={styles.catLabel}>AÑOS</div>
          <button style={styles.exportLink} onClick={() => exportarExcel(cuentas, orders, config)}><Download size={13} /> Exportar Excel</button>
        </div>
        {list.map((y) => (
          <button key={y.key} style={styles.histRow} onClick={() => { setYear(y.key); setView("months"); }}>
            <span style={styles.histRowTitle}>{y.key}</span>
            <span style={styles.histRowSub}>{y.monthKeys.length} {y.monthKeys.length === 1 ? "mes" : "meses"} con ventas</span>
            <span style={styles.histRowTotal}>{money(y.total)}</span>
          </button>
        ))}
      </div>
    );
  }
  if (view === "months") {
    const list = Object.values(months).filter((m) => m.key.startsWith(year)).sort((a, b) => b.key.localeCompare(a.key));
    return (
      <div style={styles.menuScroll}>
        <button style={styles.backBtn} onClick={() => setView("years")}>← Años</button>
        <div style={{ ...styles.catLabel, marginTop: 12 }}>{year}</div>
        {list.map((m) => (
          <button key={m.key} style={styles.histRow} onClick={() => { setMonth(m.key); setView("days"); }}>
            <span style={styles.histRowTitle}>{monthLabel(m.ts)}</span>
            <span style={styles.histRowSub}>{m.dayKeys.length} {m.dayKeys.length === 1 ? "día" : "días"} con ventas</span>
            <span style={styles.histRowTotal}>{money(m.total)}</span>
          </button>
        ))}
      </div>
    );
  }
  if (view === "days") {
    const list = Object.values(days).filter((d) => d.key.startsWith(month)).sort((a, b) => b.key.localeCompare(a.key));
    return (
      <div style={styles.menuScroll}>
        <button style={styles.backBtn} onClick={() => setView("months")}>← Meses</button>
        <div style={{ ...styles.catLabel, marginTop: 12 }}>{list[0] ? monthLabel(list[0].ts) : ""}</div>
        {list.map((d) => (
          <button key={d.key} style={styles.histRow} onClick={() => { setDay(d.key); setView("day"); }}>
            <span style={styles.histRowTitle}>{dayLabel(d.ts)}</span>
            <span style={styles.histRowSub}>{d.cuentasCount} {d.cuentasCount === 1 ? "cuenta cobrada" : "cuentas cobradas"}</span>
            <span style={styles.histRowTotal}>{money(d.total)}</span>
          </button>
        ))}
      </div>
    );
  }

  const d = days[day];
  if (!d) return null;
  const productos = Object.entries(d.productos).sort((a, b) => b[1].subtotal - a[1].subtotal);
  return (
    <div style={styles.menuScroll}>
      <button style={styles.backBtn} onClick={() => setView("days")}>← Días</button>
      <div style={{ marginTop: 12, marginBottom: 14 }}>
        <div style={styles.pageEyebrow}>{dayLabel(d.ts)}</div>
        <div style={{ ...styles.cajaTotalValue, color: "#242019" }}>{money(d.total)}</div>
      </div>

      <div style={styles.metodoBreakdownRowLight}>
        {Object.entries(METODO_META).map(([key, meta]) => {
          const Icon = meta.icon;
          return (
            <div key={key} style={styles.metodoBreakdownItemLight}>
              <Icon size={13} color={meta.color} /><span style={styles.metodoBreakdownLabelLight}>{meta.label}</span><span style={styles.metodoBreakdownValueLight}>{money(d.porMetodo[key] || 0)}</span>
            </div>
          );
        })}
      </div>

      <div style={{ ...styles.catLabel, marginTop: 18 }}>CUENTAS COBRADAS ESE DÍA</div>
      {d.cuentas.slice().sort((a, b) => a.pagadaTs - b.pagadaTs).map((c) => {
        const dividida = c.pagos && c.pagos.length > 1;
        const isExpanded = expandedCuentaId === c.id;
        const meta = !dividida && c.pagos && c.pagos[0] ? METODO_META[c.pagos[0].metodoPago] : null;
        return (
          <div key={c.id}>
            <button style={{ ...styles.cierreRow, width: "100%", background: "transparent", border: "none", cursor: dividida ? "pointer" : "default", textAlign: "left" }} onClick={() => dividida && setExpandedCuentaId(isExpanded ? null : c.id)}>
              <span style={styles.cierreDate}>
                {meta && <meta.icon size={12} color={meta.color} style={{ marginRight: 5, verticalAlign: -2 }} />}
                Mesa {c.mesa}{c.mesero ? ` · ${c.mesero}` : ""} · cobrada {timeLabel(c.pagadaTs)}
                {dividida && <span style={styles.dividedTag}> · dividida en {c.pagos.length} pagos {isExpanded ? "▲" : "▼"}</span>}
              </span>
              <span style={styles.cierrePedidos}>{c.itemCount} ítems</span>
              <span style={styles.cierreTotal}>{money(c.total)}</span>
            </button>
            {dividida && isExpanded && (
              <div style={styles.pagosRegistradosBox}>
                {c.pagos.map((p) => {
                  const pmeta = METODO_META[p.metodoPago]; const PIcon = pmeta.icon;
                  return (
                    <div key={p.id} style={styles.pagoRow}>
                      <PIcon size={13} color={pmeta.color} />
                      <div style={styles.pagoRowMid}>
                        <span style={styles.pagoRowEtiqueta}>{p.etiqueta || "Cuenta completa"}</span>
                        {p.tipo === "por_producto" && <span style={styles.pagoRowItems}>{p.items.map((it) => `${it.qty}× ${it.name}`).join(", ")}</span>}
                        <span style={styles.pagoRowTime}>{timeLabel(p.ts)} · {pmeta.label}{p.procesadoPor ? ` · cobró ${p.procesadoPor}` : ""}</span>
                      </div>
                      <span style={styles.pagoRowMonto}>{money(p.monto)}</span>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}

      <div style={{ ...styles.catLabel, marginTop: 18 }}>PRODUCTOS VENDIDOS</div>
      {productos.map(([name, p]) => (
        <div key={name} style={styles.cajaProductRow}><span style={styles.cajaProductQty}>{p.qty}×</span><span style={styles.cajaProductName}>{name}</span><span style={styles.cajaProductSubtotal}>{money(p.subtotal)}</span></div>
      ))}
    </div>
  );
}

/* ---------------- Reportes ---------------- */

function ReportesView({ cuentas, orders }) {
  const [rango, setRango] = useState("hoy");

  const inicioRango = () => {
    const d = new Date();
    if (rango === "hoy") { d.setHours(0, 0, 0, 0); return d.getTime(); }
    if (rango === "semana") { d.setDate(d.getDate() - 7); return d.getTime(); }
    d.setDate(d.getDate() - 30); return d.getTime();
  };
  const desde = inicioRango();
  const pagadasRango = cuentas.filter((c) => c.estado === "pagada" && c.pagadaTs >= desde);

  const porHora = {};
  pagadasRango.forEach((c) => { const h = new Date(c.pagadaTs).getHours(); porHora[h] = (porHora[h] || 0) + cuentaGranTotal(orders, c); });
  const dataHora = Array.from({ length: 24 }, (_, h) => ({ hora: hourLabel(h), total: porHora[h] || 0 })).filter((d, i) => {
    // solo mostrar el rango de horas de operación con datos, ampliado un poco
    const conDatos = Object.keys(porHora).map(Number);
    if (conDatos.length === 0) return i >= 6 && i <= 22;
    return i >= Math.max(0, Math.min(...conDatos) - 1) && i <= Math.min(23, Math.max(...conDatos) + 1);
  });

  const productoMap = {};
  pagadasRango.forEach((c) => cuentaOrders(orders, c.id).filter((o) => o.estado !== "cancelado").forEach((o) => o.items.forEach((it) => {
    if (!productoMap[it.name]) productoMap[it.name] = { qty: 0, subtotal: 0 };
    productoMap[it.name].qty += it.qty; productoMap[it.name].subtotal += it.qty * it.price;
  })));
  const topProductos = Object.entries(productoMap).sort((a, b) => b[1].qty - a[1].qty).slice(0, 5);

  const totalRango = pagadasRango.reduce((s, c) => s + cuentaGranTotal(orders, c), 0);
  const propinasRango = pagadasRango.reduce((s, c) => s + ((c.propina && c.propina.monto) || 0), 0);

  const meseroMap = {};
  pagadasRango.forEach((c) => {
    const nombre = c.mesero || "Sin asignar";
    if (!meseroMap[nombre]) meseroMap[nombre] = { cuentas: 0, total: 0, propinas: 0 };
    meseroMap[nombre].cuentas += 1;
    meseroMap[nombre].total += cuentaGranTotal(orders, c);
    meseroMap[nombre].propinas += (c.propina && c.propina.monto) || 0;
  });
  const meserosOrdenados = Object.entries(meseroMap).sort((a, b) => b[1].total - a[1].total);

  return (
    <div style={styles.menuScroll}>
      <div style={styles.cajaTabSwitch}>
        <button style={{ ...styles.cajaTabBtn, ...(rango === "hoy" ? styles.cajaTabBtnActive : {}) }} onClick={() => setRango("hoy")}>Hoy</button>
        <button style={{ ...styles.cajaTabBtn, ...(rango === "semana" ? styles.cajaTabBtnActive : {}) }} onClick={() => setRango("semana")}>7 días</button>
        <button style={{ ...styles.cajaTabBtn, ...(rango === "mes" ? styles.cajaTabBtnActive : {}) }} onClick={() => setRango("mes")}>30 días</button>
      </div>

      <div style={styles.cajaTotalCard}>
        <span style={styles.cajaTotalLabel}>Total vendido</span>
        <span style={styles.cajaTotalValue}>{money(totalRango)}</span>
        <span style={styles.turnoEsperadoLabel}>{pagadasRango.length} cuentas · {money(propinasRango)} en propinas</span>
      </div>

      <div style={styles.catLabel}>VENTAS POR HORA</div>
      <div style={styles.chartBox}>
        {dataHora.length === 0 || totalRango === 0 ? (
          <div style={styles.cajaEmptyText}>Sin ventas en este rango todavía.</div>
        ) : (
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={dataHora}>
              <CartesianGrid strokeDasharray="3 3" stroke="#E4DFCE" />
              <XAxis dataKey="hora" tick={{ fontSize: 10, fill: "#8A8272" }} interval={1} />
              <YAxis tick={{ fontSize: 10, fill: "#8A8272" }} tickFormatter={(v) => (v >= 1000 ? `${v / 1000}k` : v)} />
              <Tooltip formatter={(v) => money(v)} contentStyle={{ fontSize: 12, borderRadius: 8 }} />
              <Bar dataKey="total" fill="#C1442D" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        )}
      </div>

      <div style={{ ...styles.catLabel, marginTop: 20 }}>TOP 5 PRODUCTOS</div>
      {topProductos.length === 0 ? (
        <div style={styles.cajaEmptyText}>Sin datos todavía.</div>
      ) : (
        topProductos.map(([name, p]) => (
          <div key={name} style={styles.cajaProductRow}><span style={styles.cajaProductQty}>{p.qty}×</span><span style={styles.cajaProductName}>{name}</span><span style={styles.cajaProductSubtotal}>{money(p.subtotal)}</span></div>
        ))
      )}

      <div style={{ ...styles.catLabel, marginTop: 20 }}>VENTAS POR MESERO</div>
      {meserosOrdenados.length === 0 ? (
        <div style={styles.cajaEmptyText}>Sin datos todavía.</div>
      ) : (
        meserosOrdenados.map(([nombre, d]) => (
          <div key={nombre} style={styles.meseroStatRow}>
            <div style={styles.meseroStatLeft}>
              <UserCircle2 size={14} color="#8A8272" />
              <span style={styles.pagoRowEtiqueta}>{nombre}</span>
              <span style={styles.meseroStatSub}>{d.cuentas} {d.cuentas === 1 ? "cuenta" : "cuentas"}{d.propinas > 0 ? ` · ${money(d.propinas)} propina` : ""}</span>
            </div>
            <span style={styles.mesaCuentaTotal}>{money(d.total)}</span>
          </div>
        ))
      )}
    </div>
  );
}

/* ---------------- estilos ---------------- */

const fontImports = `
@import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500;600&display=swap');
`;

const styles = {
  app: { fontFamily: "'Inter', sans-serif", minHeight: "100vh", background: "#F7F4EC", display: "flex", flexDirection: "column" },
  loadingScreen: { minHeight: "100vh", background: "#1E2124", display: "flex", alignItems: "center", justifyContent: "center" },
  loadingStamp: { fontFamily: "'IBM Plex Mono', monospace", color: "#8A9B87", letterSpacing: 2, fontSize: 13 },
  topBar: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 18px", background: "#242019", flexShrink: 0, gap: 8, flexWrap: "wrap" },
  brand: { display: "flex", alignItems: "center", gap: 8 },
  brandMark: { color: "#C1442D", fontSize: 14 },
  brandText: { fontFamily: "'Oswald', sans-serif", color: "#F7F4EC", letterSpacing: 2, fontSize: 15, fontWeight: 600 },
  syncDot: { width: 7, height: 7, borderRadius: "50%", display: "inline-block", marginLeft: 2 },
  meseroChip: { display: "flex", alignItems: "center", gap: 5, background: "#33291f", color: "#D9CFAE", border: "none", borderRadius: 20, padding: "5px 10px", fontSize: 11.5, fontWeight: 600, cursor: "pointer" },
  rolBadge: { marginLeft: "auto", fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: "#8A611A", background: "#FBEFD9", padding: "2px 8px", borderRadius: 10, fontWeight: 700 },
  roleSwitch: { display: "flex", background: "#33291f", borderRadius: 10, padding: 3, gap: 2 },
  roleBtn: { display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", fontSize: 12.5, fontWeight: 600, fontFamily: "'Inter', sans-serif", border: "none", borderRadius: 8, background: "transparent", color: "#B3A891", cursor: "pointer" },
  roleBtnActive: { background: "#F7F4EC", color: "#242019" },
  roleBtnActiveDark: { background: "#5C6670", color: "#F7F4EC" },

  screen: { flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" },
  screenDark: { flex: 1, display: "flex", flexDirection: "column", background: "#1E2124", overflow: "hidden", transition: "background 0.2s" },
  screenFlash: { background: "#2b1f1c" },

  loginWrap: { padding: 24, display: "flex", flexDirection: "column", gap: 10, maxWidth: 420, margin: "0 auto", width: "100%" },
  loginList: { display: "flex", flexDirection: "column", gap: 8, margin: "14px 0" },
  loginBtn: { display: "flex", alignItems: "center", gap: 10, background: "#FFFEFA", border: "1px solid #E4DFCE", borderRadius: 10, padding: "13px 14px", fontSize: 14, fontWeight: 600, color: "#242019", cursor: "pointer" },
  pinError: { color: "#C1442D", fontSize: 12, fontWeight: 600, marginTop: 6 },

  mesaGridHeader: { display: "flex", justifyContent: "space-between", alignItems: "flex-end", padding: "22px 20px 10px" },
  pageEyebrow: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: "#8A8272", letterSpacing: 1.5 },
  pageTitle: { fontFamily: "'Oswald', sans-serif", fontSize: 26, color: "#242019", fontWeight: 600 },
  editMenuBtn: { display: "flex", alignItems: "center", gap: 6, background: "transparent", border: "1px solid #D9D2BE", color: "#5C5644", borderRadius: 8, padding: "8px 12px", fontSize: 12.5, fontWeight: 600, cursor: "pointer" },
  mesaGrid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, padding: 20 },
  mesaCard: { aspectRatio: "1", border: "1.5px solid #E4DFCE", borderRadius: 16, background: "#FFFEFA", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 10, cursor: "pointer", padding: 8 },
  mesaCardNum: { fontFamily: "'Oswald', sans-serif", fontSize: 44, color: "#242019", fontWeight: 600 },
  mesaCardTag: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, padding: "4px 10px", borderRadius: 20, fontWeight: 500 },
  mesaCardSince: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 9.5, color: "#9A9382" },

  subHeader: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "16px 18px 10px" },
  backBtn: { background: "none", border: "none", color: "#8A8272", fontSize: 13, fontWeight: 600, cursor: "pointer" },
  mesaTitleWrap: { display: "flex", flexDirection: "column", alignItems: "center" },
  mesaEyebrow: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: "#8A8272", letterSpacing: 1.5 },
  mesaTitle: { fontFamily: "'Oswald', sans-serif", fontSize: 20, color: "#242019", fontWeight: 600 },
  cuentaOpenSub: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: "#5B7553", fontWeight: 600, marginTop: 2 },
  mesaActionBtn: { width: 34, height: 34, borderRadius: 9, border: "1.5px solid #B98A2E", background: "#FBEFD9", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: "#8A611A" },
  mesaActionPanel: { margin: "0 18px 12px", background: "#FFFEFA", border: "1px solid #E4DFCE", borderRadius: 12, padding: 14, display: "flex", flexDirection: "column", gap: 8 },
  mesaPickerRow: { display: "flex", flexWrap: "wrap", gap: 8 },
  mesaPickerBtn: { background: "#242019", color: "#F7F4EC", border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer" },

  ticketStripCol: { display: "flex", flexDirection: "column", gap: 6, padding: "0 18px 10px" },
  miniTicketRow: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 },
  miniTicket: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, background: "#EFE9D9", color: "#5C5644", padding: "5px 10px", borderRadius: 20, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" },
  miniTicketTime: { color: "#9A9382" },
  miniTicketTotal: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5, color: "#5C5644", textAlign: "right", padding: "2px 4px 0" },
  entregadoBtn: { display: "flex", alignItems: "center", gap: 4, alignSelf: "flex-end", background: "#E4EEE1", color: "#3F5638", border: "none", borderRadius: 8, padding: "5px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" },
  avisosListosWrap: { display: "flex", flexDirection: "column", gap: 6, padding: "0 20px 10px" },
  avisoListoBanner: { display: "flex", alignItems: "center", gap: 8, background: "#2F6690", color: "#F7F4EC", border: "none", borderRadius: 10, padding: "11px 14px", fontSize: 13, fontWeight: 700, cursor: "pointer", textAlign: "left" },
  avisoListoClose: { marginLeft: "auto", opacity: 0.8, fontSize: 13, padding: "0 4px" },
  stateDot: { width: 7, height: 7, borderRadius: "50%", display: "inline-block", flexShrink: 0 },
  editedTag: { fontSize: 9.5, fontWeight: 700, color: "#B98A2E", marginLeft: 2 },
  miniTicketActions: { display: "flex", alignItems: "center", gap: 4, flexShrink: 0 },
  iconBtn: { width: 34, height: 34, borderRadius: 9, border: "1px solid #E4DFCE", background: "#FFFEFA", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: "#5C5644" },
  confirmText: { fontSize: 11, color: "#5C5644", fontWeight: 600 },
  confirmYes: { background: "#C1442D", color: "#F7F4EC", border: "none", borderRadius: 6, padding: "4px 9px", fontSize: 11, fontWeight: 700, cursor: "pointer" },
  confirmNo: { background: "#EDE7D8", color: "#5C5644", border: "none", borderRadius: 6, padding: "4px 9px", fontSize: 11, fontWeight: 700, cursor: "pointer" },
  editingBanner: { display: "flex", alignItems: "center", gap: 8, background: "#FBEFD9", color: "#8A611A", fontSize: 12, fontWeight: 600, padding: "9px 18px", margin: "0 18px 10px", borderRadius: 8 },
  sentBanner: { display: "flex", alignItems: "center", gap: 8, background: "#E4EEE1", color: "#3F5638", fontSize: 12.5, fontWeight: 700, padding: "10px 18px", margin: "0 18px 10px", borderRadius: 8 },
  warnBanner: { display: "flex", alignItems: "center", gap: 8, background: "#FBE2DC", color: "#8A2E1A", fontSize: 12, fontWeight: 600, padding: "9px 14px", margin: "0 0 12px", borderRadius: 8 },
  repeatBtn: { display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "#FFFEFA", border: "1px dashed #C7BE9F", color: "#8A611A", borderRadius: 10, padding: "11px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", margin: "0 18px 12px" },
  cancelEditBtn: { background: "#EDE7D8", color: "#5C5644", border: "none", borderRadius: 10, padding: "13px 16px", fontSize: 13, fontWeight: 600, cursor: "pointer" },

  catPillsWrap: { padding: "0 18px 8px", flexShrink: 0 },
  searchInput: { width: "100%", border: "1px solid #E4DFCE", borderRadius: 10, padding: "10px 12px", fontSize: 13.5, fontFamily: "'Inter', sans-serif", background: "#FFFEFA", marginBottom: 8, boxSizing: "border-box" },
  catPills: { display: "flex", gap: 6, overflowX: "auto", paddingBottom: 2, WebkitOverflowScrolling: "touch", scrollbarWidth: "none" },
  catPill: { flexShrink: 0, display: "flex", alignItems: "center", gap: 5, fontSize: 12, fontWeight: 600, letterSpacing: 0.2, color: "#5C5644", background: "#FFFEFA", border: "1px solid #E4DFCE", padding: "8px 13px", borderRadius: 18, whiteSpace: "nowrap", cursor: "pointer" },
  catPillActive: { background: "#242019", color: "#F7F4EC", border: "1px solid #242019" },
  catPillBadge: { background: "#C1442D", color: "#F7F4EC", fontSize: 10, fontWeight: 700, borderRadius: 10, padding: "1px 6px", minWidth: 16, textAlign: "center" },
  menuItemCat: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: "#B98A2E", marginTop: 1 },
  noResults: { textAlign: "center", color: "#9A9382", fontSize: 13, fontStyle: "italic", padding: "30px 0" },
  menuScroll: { flex: 1, overflowY: "auto", padding: "6px 18px 100px" },
  catLabel: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: "#B98A2E", letterSpacing: 1.5, marginBottom: 6, marginTop: 4 },
  rowBetween: { display: "flex", justifyContent: "space-between", alignItems: "center" },
  menuRow: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 0", borderBottom: "1px solid #EBE5D3" },
  menuItemName: { fontSize: 14.5, color: "#242019", fontWeight: 500 },
  menuItemPrice: { fontSize: 12.5, color: "#9A9382", fontFamily: "'IBM Plex Mono', monospace", marginTop: 2 },
  agotadoTag: { fontSize: 9, fontWeight: 700, color: "#C1442D", marginLeft: 4 },
  itemNoteBtn: { display: "flex", alignItems: "center", gap: 4, background: "none", border: "none", color: "#B98A2E", fontSize: 10.5, padding: "3px 0 0", cursor: "pointer" },
  itemNoteInput: { border: "1px solid #E4DFCE", borderRadius: 6, padding: "5px 8px", fontSize: 11, marginTop: 4, width: "90%" },
  stepper: { display: "flex", alignItems: "center", gap: 10 },
  stepBtn: { width: 38, height: 38, borderRadius: 9, border: "1px solid #E4DFCE", background: "#FFFEFA", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: "#5C5644" },
  stepBtnPlus: { background: "#242019", color: "#F7F4EC", border: "none" },
  stepVal: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 15, width: 18, textAlign: "center", fontWeight: 600 },

  orderBar: { position: "sticky", bottom: 0, background: "#F7F4EC", borderTop: "1px solid #E4DFCE", padding: 14, display: "flex", flexDirection: "column", gap: 10 },
  noteInput: { border: "1px solid #E4DFCE", borderRadius: 8, padding: "10px 12px", fontSize: 13, fontFamily: "'Inter', sans-serif", background: "#FFFEFA" },
  sendBtn: { display: "flex", alignItems: "center", justifyContent: "center", gap: 8, background: "#C1442D", color: "#F7F4EC", border: "none", borderRadius: 10, padding: "13px", fontSize: 14, fontWeight: 600, cursor: "pointer" },

  addItemCard: { background: "#FFFEFA", border: "1px solid #E4DFCE", borderRadius: 12, padding: 14, display: "flex", flexDirection: "column", gap: 8, marginBottom: 20 },
  editInput: { border: "1px solid #E4DFCE", borderRadius: 8, padding: "10px 12px", fontSize: 13, fontFamily: "'Inter', sans-serif", width: "100%", boxSizing: "border-box" },
  addItemBtn: { display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "#242019", color: "#F7F4EC", border: "none", borderRadius: 8, padding: "10px", fontSize: 13, fontWeight: 600, cursor: "pointer" },
  editRow: { display: "flex", alignItems: "center", gap: 10, padding: "10px 0", borderBottom: "1px solid #EBE5D3" },
  deleteBtn: { marginLeft: "auto", background: "none", border: "none", color: "#C1442D", cursor: "pointer", padding: 4 },
  agotadoToggle: { width: 26, height: 26, borderRadius: 7, border: "1px solid #E4DFCE", background: "#FFFEFA", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: "#9A9382", flexShrink: 0 },
  agotadoToggleActive: { background: "#C1442D", color: "#F7F4EC", border: "none" },

  kdsHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "20px 20px 14px" },
  kdsEyebrow: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: "#5C6670", letterSpacing: 1.5 },
  kdsTitle: { fontFamily: "'Oswald', sans-serif", fontSize: 22, color: "#F7F4EC", fontWeight: 600, marginTop: 2 },
  kdsCount: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 20, color: "#F7F4EC", background: "#33393F", borderRadius: 10, width: 40, height: 40, display: "flex", alignItems: "center", justifyContent: "center" },
  emptyKitchen: { flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 10 },
  emptyKitchenText: { fontFamily: "'IBM Plex Mono', monospace", color: "#5C6670", fontSize: 12.5 },
  rail: { flex: 1, overflowY: "auto", padding: "0 16px 30px", display: "flex", flexDirection: "column", gap: 14 },
  ticket: { background: "#F7F4EC", borderRadius: 4, padding: "16px 16px 14px", boxShadow: "0 4px 14px rgba(0,0,0,0.35)", borderLeft: "4px solid #C1442D" },
  ticketHead: { display: "flex", justifyContent: "space-between", alignItems: "center" },
  ticketMesa: { fontFamily: "'Oswald', sans-serif", fontSize: 22, fontWeight: 700, color: "#242019", display: "flex", alignItems: "center", gap: 8, letterSpacing: 0.5 },
  ticketExactTime: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: "#9A9382", marginTop: 2 },
  editedBadge: { fontFamily: "'Inter', sans-serif", fontSize: 9, fontWeight: 700, color: "#F7F4EC", background: "#B98A2E", padding: "2px 6px", borderRadius: 4, letterSpacing: 0.5 },
  ticketTime: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5, color: "#8A8272" },
  ticketDivider: { borderTop: "1.5px dashed #D9D2BE", margin: "10px 0" },
  ticketLine: { display: "flex", gap: 8, padding: "3px 0", fontFamily: "'IBM Plex Mono', monospace" },
  ticketQty: { color: "#C1442D", fontWeight: 600, fontSize: 13 },
  ticketItemName: { fontSize: 13, color: "#242019" },
  ticketItemNote: { fontSize: 11, color: "#B98A2E", fontStyle: "italic", paddingLeft: 20, marginBottom: 2 },
  ticketNote: { fontFamily: "'Inter', sans-serif", fontStyle: "italic", fontSize: 12.5, color: "#8A8272", marginTop: 8 },
  ticketBtn: { width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, border: "none", borderRadius: 10, padding: "16px", fontSize: 15, fontWeight: 700, color: "#F7F4EC", cursor: "pointer", minHeight: 52 },

  cajaTabSwitch: { display: "flex", gap: 6, padding: "0 20px 12px" },
  cajaTabBtn: { flex: 1, background: "#EFE9D9", border: "none", borderRadius: 9, padding: "9px", fontSize: 12, fontWeight: 600, color: "#8A8272", cursor: "pointer", fontFamily: "'Inter', sans-serif" },
  cajaTabBtnActive: { background: "#242019", color: "#F7F4EC" },

  cajaTotalCard: { background: "#242019", borderRadius: 14, padding: "20px 18px", display: "flex", flexDirection: "column", gap: 4, marginBottom: 16 },
  cajaTotalLabel: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: "#B3A891", letterSpacing: 1 },
  cajaTotalValue: { fontFamily: "'Oswald', sans-serif", fontSize: 34, color: "#F7F4EC", fontWeight: 600 },
  cajaTotalValue2: { fontFamily: "'Oswald', sans-serif", fontSize: 30, color: "#242019", fontWeight: 600, display: "flex", flexDirection: "column", margin: "8px 0" },
  turnoEsperadoLabel: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: "#B3A891", fontWeight: 400, marginTop: 2 },
  cajaEmptyText: { fontSize: 12.5, color: "#9A9382", fontStyle: "italic", padding: "4px 0 18px", lineHeight: 1.6 },

  metodoBreakdownRow: { display: "flex", gap: 14, marginTop: 12, paddingTop: 12, borderTop: "1px solid #3a332a" },
  metodoBreakdownItem: { display: "flex", flexDirection: "column", alignItems: "flex-start", gap: 2 },
  metodoBreakdownLabel: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 9.5, color: "#B3A891" },
  metodoBreakdownValue: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: "#F7F4EC", fontWeight: 700 },
  metodoBreakdownRowLight: { display: "flex", gap: 10, marginBottom: 16 },
  metodoBreakdownItemLight: { flex: 1, background: "#FFFEFA", border: "1px solid #E4DFCE", borderRadius: 10, padding: "10px 8px", display: "flex", flexDirection: "column", alignItems: "center", gap: 3 },
  metodoBreakdownLabelLight: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 9.5, color: "#8A8272" },
  metodoBreakdownValueLight: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: "#242019", fontWeight: 700 },
  metodoPillsRow: { display: "flex", gap: 6 },
  metodoPill: { flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 5, background: "#FFFEFA", border: "1px solid #D9CFAE", color: "#5C5644", borderRadius: 9, padding: "9px 6px", fontSize: 11.5, fontWeight: 600, cursor: "pointer" },

  turnoBox: { background: "#FFFEFA", border: "1px solid #E4DFCE", borderRadius: 14, padding: 16, display: "flex", flexDirection: "column", gap: 8, marginBottom: 16 },
  turnoRow: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: "1px solid #EBE5D3" },
  turnoDiff: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 700 },
  diferenciaBox: { borderRadius: 8, padding: "10px 12px", fontSize: 13, fontWeight: 700, color: "#242019" },

  inlineLink: { background: "none", border: "none", color: "#C1442D", fontSize: 10.5, textDecoration: "underline", cursor: "pointer", padding: 0, marginLeft: 4 },
  inlineAddLink: { display: "flex", alignItems: "center", gap: 5, background: "none", border: "none", color: "#B98A2E", fontSize: 11.5, fontWeight: 600, cursor: "pointer", padding: "6px 0" },
  lockedNote: { display: "flex", alignItems: "center", gap: 5, fontSize: 11, color: "#9A9382", fontStyle: "italic", padding: "6px 0" },

  usuarioRow: { background: "#FFFEFA", border: "1px solid #E4DFCE", borderRadius: 10, padding: "10px 12px", marginBottom: 8 },
  usuarioRowLeft: { display: "flex", alignItems: "center", gap: 7, flexWrap: "wrap" },
  pinRevealBtn: { display: "flex", alignItems: "center", gap: 4, background: "#F0EDE3", border: "none", borderRadius: 7, padding: "3px 8px", fontSize: 10.5, fontFamily: "'IBM Plex Mono', monospace", color: "#5C5644", cursor: "pointer", letterSpacing: 1 },
  usuarioRowActions: { display: "flex", alignItems: "center", gap: 10, marginTop: 6, paddingLeft: 23 },
  resetPinRow: { display: "flex", gap: 8, marginTop: 8, paddingLeft: 23 },
  rolPillsRow: { display: "flex", gap: 6 },
  rolPill: { flex: 1, background: "#EFE9D9", border: "none", borderRadius: 8, padding: "9px", fontSize: 12.5, fontWeight: 600, color: "#8A8272", cursor: "pointer" },
  rolPillActive: { background: "#242019", color: "#F7F4EC" },
  auditText: { fontSize: 11.5, color: "#5C5644", flex: 1 },
  descuentoBox: { background: "#FBEFD9", borderRadius: 10, padding: 10, marginTop: 4 },
  propinaBox: { display: "flex", flexDirection: "column", gap: 6, padding: "6px 0" },
  propinaPillsRow: { display: "flex", gap: 6, alignItems: "center" },
  propinaPill: { background: "#FFFEFA", border: "1px solid #E4DFCE", borderRadius: 8, padding: "7px 11px", fontSize: 12, fontWeight: 600, color: "#5C5644", cursor: "pointer" },
  propinaCustomInput: { width: 60, border: "1px solid #E4DFCE", borderRadius: 8, padding: "7px 8px", fontSize: 12 },
  propinaCustomBtn: { background: "#242019", color: "#F7F4EC", border: "none", borderRadius: 8, padding: "7px 10px", fontSize: 12, fontWeight: 600, cursor: "pointer" },

  subtotalRow: { display: "flex", justifyContent: "space-between", fontSize: 12.5, color: "#5C5644", padding: "4px 0" },
  subtotalRowFinal: { fontSize: 14.5, fontWeight: 700, color: "#242019", borderTop: "1px solid #E4DFCE", marginTop: 4, paddingTop: 8 },

  mesaCuentaCard: { background: "#FFFEFA", border: "1px solid #E4DFCE", borderRadius: 12, marginBottom: 10, overflow: "hidden" },
  mesaCuentaHead: { width: "100%", display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 16px", background: "transparent", border: "none", cursor: "pointer", textAlign: "left" },
  mesaCuentaLeft: { display: "flex", flexDirection: "column", gap: 3 },
  mesaCuentaNum: { fontFamily: "'Oswald', sans-serif", fontSize: 16, fontWeight: 600, color: "#242019" },
  mesaCuentaSince: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: "#9A9382" },
  mesaCuentaRight: { display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 3 },
  mesaCuentaItems: { fontSize: 11, color: "#9A9382" },
  mesaCuentaTotal: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 700, color: "#242019" },
  mesaCuentaBody: { padding: "0 16px 16px", borderTop: "1px solid #EBE5D3" },
  mesaCuentaOrderRow: { display: "flex", gap: 10, padding: "10px 0", borderBottom: "1px solid #EBE5D3", alignItems: "baseline" },
  mesaCuentaOrderTime: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: "#9A9382", flexShrink: 0, width: 44 },
  mesaCuentaOrderItems: { fontSize: 12.5, color: "#5C5644", flex: 1 },
  mesaCuentaOrderTotal: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: "#242019", fontWeight: 600 },

  cajaProductRow: { display: "flex", alignItems: "center", gap: 10, padding: "9px 0", borderBottom: "1px solid #EBE5D3" },
  cajaProductQty: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: "#C1442D", fontWeight: 600, width: 28 },
  cajaProductName: { fontSize: 13.5, color: "#242019", flex: 1 },
  cajaProductSubtotal: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: "#5C5644" },

  closeCajaBtn: { width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, background: "#242019", color: "#F7F4EC", border: "none", borderRadius: 10, padding: "14px", fontSize: 14, fontWeight: 600, cursor: "pointer", marginTop: 10 },
  closeConfirmBox: { background: "#FBEFD9", borderRadius: 12, padding: 14, display: "flex", flexDirection: "column", gap: 10, marginTop: 10 },
  closeConfirmText: { fontSize: 12.5, color: "#6B5218", lineHeight: 1.5 },

  enCursoTag: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: "#B98A2E", fontWeight: 700, marginTop: 2 },
  cobroBox: { background: "#F7F4EC", border: "1px solid #E4DFCE", borderRadius: 12, padding: 14, display: "flex", flexDirection: "column", gap: 10, marginTop: 10 },
  modoChoiceRow: { display: "flex", flexDirection: "column", gap: 8 },
  modoChoiceBtn: { display: "flex", alignItems: "center", gap: 10, background: "#FFFEFA", border: "1px solid #E4DFCE", color: "#242019", borderRadius: 10, padding: "13px 14px", fontSize: 13.5, fontWeight: 600, cursor: "pointer" },
  partesStepperRow: { display: "flex", alignItems: "center", gap: 10 },
  productoSplitRow: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: "1px solid #EBE5D3" },
  pagosRegistradosBox: { background: "#FFFEFA", border: "1px solid #E4DFCE", borderRadius: 10, padding: "10px 12px", marginTop: 10, marginBottom: 4 },
  pagosRegistradosLabel: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: "#8A8272", letterSpacing: 1, marginBottom: 6 },
  pagoRow: { display: "flex", alignItems: "flex-start", gap: 8, padding: "8px 0", borderBottom: "1px solid #EBE5D3" },
  pagoRowMid: { flex: 1, display: "flex", flexDirection: "column", gap: 2 },
  pagoRowEtiqueta: { fontSize: 12.5, color: "#242019", fontWeight: 600 },
  pagoRowItems: { fontSize: 11, color: "#8A8272" },
  pagoRowTime: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: "#9A9382" },
  pagoRowMonto: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: "#242019", fontWeight: 700 },
  dividedTag: { color: "#B98A2E", fontWeight: 700 },

  cierreRow: { display: "flex", alignItems: "center", gap: 10, padding: "9px 0", borderBottom: "1px solid #EBE5D3" },
  cierreDate: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5, color: "#8A8272", flex: 1 },
  cierrePedidos: { fontSize: 11.5, color: "#9A9382" },
  cierreTotal: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: "#242019", fontWeight: 600, width: 70, textAlign: "right" },

  goHistLink: { textAlign: "center", fontSize: 12.5, color: "#B98A2E", fontWeight: 600, padding: "14px 0 4px", cursor: "pointer" },
  exportLink: { display: "flex", alignItems: "center", gap: 5, background: "none", border: "1px solid #D9D2BE", color: "#5C5644", fontSize: 11, fontWeight: 600, cursor: "pointer", padding: "6px 10px", borderRadius: 8 },
  histRow: { width: "100%", display: "flex", alignItems: "center", gap: 10, background: "#FFFEFA", border: "1px solid #E4DFCE", borderRadius: 10, padding: "13px 14px", marginBottom: 8, cursor: "pointer", textAlign: "left", fontFamily: "'Inter', sans-serif" },
  histRowTitle: { fontSize: 14, color: "#242019", fontWeight: 600, flex: 1 },
  histRowSub: { fontSize: 11, color: "#9A9382", fontFamily: "'IBM Plex Mono', monospace" },
  histRowTotal: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 14, color: "#242019", fontWeight: 600 },

  chartBox: { background: "#FFFEFA", border: "1px solid #E4DFCE", borderRadius: 12, padding: "12px 8px 4px", marginBottom: 8 },
  meseroStatRow: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: "1px solid #EBE5D3" },
  meseroStatLeft: { display: "flex", alignItems: "center", gap: 7 },
  meseroStatSub: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: "#9A9382", marginLeft: 4 },
};
