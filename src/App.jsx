import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  Plus, Minus, Send, ChefHat, ClipboardList, Check, Flame, Settings2, Trash2, Clock,
  Receipt, Pencil, Ban, AlertTriangle, Lock, Wallet, Users, Banknote, CreditCard, Smartphone,
  KeyRound, UserCircle2, PauseCircle, ArrowRightLeft, Download,
  Gift, Volume2, MessageSquarePlus, PlayCircle, StopCircle, Eye, EyeOff,
} from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";
import { subscribeShared, saveShared, subscribeCollection, subscribeCollectionSince, getCollectionAll, getCollectionRange, getRecord, saveRecord, saveRecords, saveCollectionArray } from "./firebase.js";

// El día a día (Mesero, Cocina, Caja, Turno, Reportes de hoy/7/30 días) solo
// mantiene en vivo esta cantidad de días recientes — así la app nunca se
// vuelve más lenta con los años, sin importar cuánto historial se acumule.
// Historial y Exportar/Respaldar sí piden todo, pero solo cuando se usan.
const VENTANA_DIAS = 35;
const VENTANA_MS = VENTANA_DIAS * 24 * 60 * 60 * 1000;

const DEFAULT_MESAS = [
  { id: "1", nombre: "Mesa 1" },
  { id: "2", nombre: "Mesa 2" },
  { id: "3", nombre: "Mesa 3" },
  { id: "4", nombre: "Mesa 4" },
];

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
  listo: { label: "Listo para servir", color: "#2F6690" },
  servido: { label: "Servido", color: "#5B7553" },
  cancelado: { label: "Cancelado", color: "#8B80A3" },
};

const METODO_META = {
  efectivo: { label: "Efectivo", icon: Banknote, color: "#5B7553" },
  tarjeta: { label: "Tarjeta", icon: CreditCard, color: "#2F6690" },
  transferencia: { label: "Transferencia", icon: Smartphone, color: "#8A611A" },
};

const PROPINA_OPCIONES = [0, 10, 15, 20];

const ROL_LABELS = { admin: "Admin", cajero: "Cajero", mesero: "Mesero" };
function rolLabel(rol) {
  return ROL_LABELS[rol] || rol;
}
// Un usuario puede tener uno o varios roles a la vez (ej. Mesero + Cajero).
// Esto sigue leyendo el formato viejo (un solo "rol") para no perder nada de
// lo que ya está guardado — pero todo lo nuevo se guarda como "roles": [...].
function rolesDe(usuario) {
  if (!usuario) return [];
  if (Array.isArray(usuario.roles) && usuario.roles.length > 0) return usuario.roles;
  if (usuario.rol) return [usuario.rol];
  return [];
}
function tieneRol(usuario, rol) {
  return rolesDe(usuario).includes(rol);
}
function rolesLabel(usuario) {
  const roles = rolesDe(usuario);
  return roles.length > 0 ? roles.map(rolLabel).join(" + ") : "Sin rol";
}

/* ---------------- mesas (nombres personalizables) ---------------- */

function mesaMatch(a, b) {
  return String(a) === String(b);
}
function mesaNombre(mesas, mesaId) {
  const m = (mesas || []).find((x) => mesaMatch(x.id, mesaId));
  return m ? m.nombre : `Mesa ${mesaId}`;
}
function mesaOrderIndex(mesas, mesaId) {
  const idx = (mesas || []).findIndex((x) => mesaMatch(x.id, mesaId));
  return idx === -1 ? 9999 : idx;
}

// Estado visual de una mesa para la cuadrícula y el plano: distingue entre
// "hay algo pendiente en cocina" (rojo), "ya está listo pero nadie lo ha
// llevado a la mesa" (azul), "cuenta abierta, todo servido" (verde) y libre.
function estadoVisualMesa(cuenta, ordersDeLaCuenta) {
  if (!cuenta) return { tag: "Libre", bg: "#E6DDF5", color: "#8B80A3" };
  const enCocina = ordersDeLaCuenta.some((o) => o.estado === "pendiente" || o.estado === "preparando");
  if (enCocina) return { tag: "En cocina", bg: "#C1442D", color: "#F3EEFB" };
  const listoSinServir = ordersDeLaCuenta.some((o) => o.estado === "listo");
  if (listoSinServir) return { tag: "Listo para servir", bg: "#2F6690", color: "#F3EEFB" };
  return { tag: "Cuenta abierta", bg: "#5B7553", color: "#F3EEFB" };
}

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

// Notificación del sistema — solo funciona mientras el navegador siga abierto
// (aunque sea en segundo plano). Con la pantalla bloqueada NO llega.
function mostrarNotificacionSistema(titulo, cuerpo) {
  try {
    if (!("Notification" in window) || Notification.permission !== "granted") return;
    if (document.visibilityState === "visible") return;
    if (navigator.serviceWorker && navigator.serviceWorker.ready) {
      navigator.serviceWorker.ready.then((reg) => {
        reg.showNotification(titulo, { body: cuerpo, icon: "/icons/icon-192.png", badge: "/icons/icon-192.png" });
      }).catch(() => {});
    }
  } catch (e) {
    // silencioso si el navegador no soporta notificaciones
  }
}


/* La carga/guardado de datos en vivo vive en ./firebase.js (Firebase Realtime Database) */

function descargarRespaldo({ menu, orders, cuentas, config, turnos, mesas }) {
  const data = { version: 2, exportadoTs: Date.now(), menu, orders, cuentas, config, turnos, mesas };
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

function exportarExcel(cuentas, orders, config, mesas) {
  const pagadas = cuentas.filter((c) => c.estado === "pagada").sort((a, b) => a.pagadaTs - b.pagadaTs);
  const filasPagos = [];
  pagadas.forEach((c) => {
    pagosDe(c).forEach((p) => {
      filasPagos.push({
        Fecha: new Date(c.pagadaTs).toLocaleDateString("es-CO"),
        Hora: timeLabel(p.ts),
        Mesa: mesaNombre(mesas, c.mesa),
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

const DEFAULT_CONFIG = { businessName: "La Reserva", usuarios: [], auditLog: [] };

export default function App() {
  const [role, setRole] = useState("mesero");
  const [menu, setMenu] = useState(DEFAULT_MENU);
  const [orders, setOrders] = useState([]);
  const [cuentas, setCuentas] = useState([]);
  const [config, setConfig] = useState(DEFAULT_CONFIG);
  const [turnos, setTurnos] = useState([]);
  const [mesas, setMesas] = useState(DEFAULT_MESAS);
  const [loaded, setLoaded] = useState(false);
  const [, forceTick] = useState(0);
  const [usuarioActual, setUsuarioActual] = useState(null);
  // Si algo falla al guardar en Firebase se muestra un aviso visible.
  const [errorGuardado, setErrorGuardado] = useState(null);
  const avisarErrorGuardado = useCallback((msg) => setErrorGuardado(msg), []);
  const [online, setOnline] = useState(typeof navigator === "undefined" ? true : navigator.onLine);

  // Los navegadores bloquean audio automático hasta que hay una interacción real.
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

  // Estado de conexión real del navegador.
  useEffect(() => {
    const goOnline = () => setOnline(true);
    const goOffline = () => setOnline(false);
    window.addEventListener("online", goOnline);
    window.addEventListener("offline", goOffline);
    return () => { window.removeEventListener("online", goOnline); window.removeEventListener("offline", goOffline); };
  }, []);

  // Suscripción en vivo a Firebase. Pedidos y cuentas solo se suscriben a los
  // últimos VENTANA_DIAS; la ventana se refresca cada hora.
  const [ventanaEpoch, setVentanaEpoch] = useState(0);
  useEffect(() => {
    const refrescar = setInterval(() => setVentanaEpoch((e) => e + 1), 60 * 60 * 1000);
    return () => clearInterval(refrescar);
  }, []);

  useEffect(() => {
    const pending = new Set(["menu", "orders", "cuentas", "config", "turnos", "mesas"]);
    const markLoaded = (key) => { pending.delete(key); if (pending.size === 0) setLoaded(true); };
    const desde = Date.now() - VENTANA_MS;

    const unsubs = [
      subscribeShared("menu", DEFAULT_MENU, (v) => { setMenu(v); markLoaded("menu"); }),
      subscribeCollectionSince("orders", "ts", desde, (v) => { setOrders(v); markLoaded("orders"); }),
      subscribeCollectionSince("cuentas", "ts", desde, (v) => { setCuentas(v); markLoaded("cuentas"); }),
      subscribeShared("config", DEFAULT_CONFIG, (v) => { setConfig({ ...DEFAULT_CONFIG, ...v }); markLoaded("config"); }),
      subscribeCollection("turnos", (v) => { setTurnos(v); markLoaded("turnos"); }),
      subscribeShared("mesas", DEFAULT_MESAS, (v) => { setMesas(v && v.length > 0 ? v : DEFAULT_MESAS); markLoaded("mesas"); }),
    ];
    const clock = setInterval(() => forceTick((t) => t + 1), 1000);
    return () => { unsubs.forEach((unsub) => unsub && unsub()); clearInterval(clock); };
  }, [ventanaEpoch]);

  // Historial completo bajo demanda (una sola vez, no en vivo).
  const [historialCompleto, setHistorialCompleto] = useState(null);
  const [cargandoHistorial, setCargandoHistorial] = useState(false);
  const cargarHistorialCompleto = useCallback(async () => {
    if (historialCompleto || cargandoHistorial) return;
    setCargandoHistorial(true);
    const [todosOrders, todasCuentas] = await Promise.all([getCollectionAll("orders"), getCollectionAll("cuentas")]);
    setHistorialCompleto({ orders: todosOrders, cuentas: todasCuentas });
    setCargandoHistorial(false);
  }, [historialCompleto, cargandoHistorial]);
  const mergeConVentana = (base, ventana) => {
    const byId = Object.fromEntries(base.map((r) => [r.id, r]));
    ventana.forEach((r) => { byId[r.id] = r; });
    return Object.values(byId);
  };
  const ordersParaHistorial = historialCompleto ? mergeConVentana(historialCompleto.orders, orders) : orders;
  const cuentasParaHistorial = historialCompleto ? mergeConVentana(historialCompleto.cuentas, cuentas) : cuentas;

  // Respaldo automático diario: cuando un admin abre la app, crea la copia de
  // AYER si falta, trayendo solo los pedidos y cuentas de ese día.
  useEffect(() => {
    if (!usuarioActual || !tieneRol(usuarioActual, "admin") || !loaded) return;
    let cancelado = false;
    (async () => {
      const ayer = new Date(Date.now() - 24 * 60 * 60 * 1000);
      const dk = dayKey(ayer.getTime());
      const yaExiste = await getRecord(`respaldos/${dk}/hecho`);
      if (cancelado || yaExiste) return;
      const inicioDia = new Date(ayer); inicioDia.setHours(0, 0, 0, 0);
      const finDia = new Date(ayer); finDia.setHours(23, 59, 59, 999);
      const [cuentasDia, ordersDia] = await Promise.all([
        getCollectionRange("cuentas", "ts", inicioDia.getTime(), finDia.getTime()),
        getCollectionRange("orders", "ts", inicioDia.getTime(), finDia.getTime()),
      ]);
      if (cancelado) return;
      await saveShared(`respaldos/${dk}`, {
        hecho: true, generadoTs: Date.now(), dia: dk,
        cuentas: cuentasDia, orders: ordersDia, menu, mesas, config,
      });
    })();
    return () => { cancelado = true; };
  }, [usuarioActual, loaded]);

  const MSG_SIN_CONEXION = "No se pudo guardar — revisa tu conexión a internet e inténtalo de nuevo.";
  const persistMenu = useCallback(async (next) => { setMenu(next); const ok = await saveShared("menu", next); if (!ok) avisarErrorGuardado(MSG_SIN_CONEXION); }, [avisarErrorGuardado]);
  const persistConfig = useCallback(async (next) => { setConfig(next); const ok = await saveShared("config", next); if (!ok) avisarErrorGuardado(MSG_SIN_CONEXION); }, [avisarErrorGuardado]);
  const persistMesas = useCallback(async (next) => { setMesas(next); const ok = await saveShared("mesas", next); if (!ok) avisarErrorGuardado(MSG_SIN_CONEXION); }, [avisarErrorGuardado]);

  // Guardan UN registro a la vez (pedido / cuenta / turno), no la lista completa.
  const saveOrder = useCallback(async (order) => {
    setOrders((prev) => { const i = prev.findIndex((o) => o.id === order.id); if (i === -1) return [...prev, order]; const next = [...prev]; next[i] = order; return next; });
    const ok = await saveRecord("orders", order.id, order);
    if (!ok) avisarErrorGuardado(MSG_SIN_CONEXION);
  }, [avisarErrorGuardado]);
  const saveOrdersMap = useCallback(async (recordsById) => {
    setOrders((prev) => { const byId = Object.fromEntries(prev.map((o) => [o.id, o])); Object.entries(recordsById).forEach(([id, val]) => { byId[id] = val; }); return Object.values(byId); });
    const ok = await saveRecords("orders", recordsById);
    if (!ok) avisarErrorGuardado(MSG_SIN_CONEXION);
  }, [avisarErrorGuardado]);
  const saveCuenta = useCallback(async (cuenta) => {
    setCuentas((prev) => { const i = prev.findIndex((c) => c.id === cuenta.id); if (i === -1) return [...prev, cuenta]; const next = [...prev]; next[i] = cuenta; return next; });
    const ok = await saveRecord("cuentas", cuenta.id, cuenta);
    if (!ok) avisarErrorGuardado(MSG_SIN_CONEXION);
  }, [avisarErrorGuardado]);
  const saveTurno = useCallback(async (turno) => {
    setTurnos((prev) => { const i = prev.findIndex((t) => t.id === turno.id); if (i === -1) return [...prev, turno]; const next = [...prev]; next[i] = turno; return next; });
    const ok = await saveRecord("turnos", turno.id, turno);
    if (!ok) avisarErrorGuardado(MSG_SIN_CONEXION);
  }, [avisarErrorGuardado]);

  const turnoAbierto = turnos.find((t) => t.estado === "abierto") || null;

  if (!loaded) {
    return (
      <div style={styles.loadingScreen}>
        <div style={styles.loadingStamp}>CARGANDO COMANDAS…</div>
      </div>
    );
  }

  // Al iniciar sesión, cada quien cae directo en su sección.
  const onLogin = (usuario) => {
    setUsuarioActual(usuario);
    const esMesero = tieneRol(usuario, "mesero");
    const esCajero = tieneRol(usuario, "cajero") || tieneRol(usuario, "admin");
    if (esMesero && !esCajero) setRole("mesero");
    else if (esCajero && !esMesero) setRole("caja");
  };

  return (
    <div style={styles.app}>
      <style>{fontImports}</style>
      <TopBar role={role} setRole={setRole} config={config} usuarioActual={usuarioActual} onCambiarUsuario={() => setUsuarioActual(null)} online={online} />
      {errorGuardado && (
        <div style={styles.errorGuardadoBanner}>
          <AlertTriangle size={14} />
          <span style={{ flex: 1 }}>{errorGuardado}</span>
          <button style={styles.errorGuardadoCerrar} onClick={() => setErrorGuardado(null)}>✕</button>
        </div>
      )}
      {role === "mesero" && (
        !usuarioActual ? (
          <UsuarioLogin config={config} persistConfig={persistConfig} onEntrar={onLogin} />
        ) : !tieneRol(usuarioActual, "mesero") && !tieneRol(usuarioActual, "admin") ? (
          <AccesoDenegado mensaje="Tu usuario no tiene el rol de Mesero — esto es Mesero." accion="Ir a Caja" onAccion={() => setRole("caja")} onCambiarUsuario={() => setUsuarioActual(null)} />
        ) : (
          <MeseroView
            menu={menu} orders={orders} cuentas={cuentas} usuarioActual={usuarioActual} config={config}
            mesas={mesas} persistMesas={persistMesas}
            saveOrder={saveOrder} saveOrdersMap={saveOrdersMap} saveCuenta={saveCuenta}
            persistMenu={persistMenu} persistConfig={persistConfig}
          />
        )
      )}
      {role === "cocina" && <CocinaView orders={orders} saveOrder={saveOrder} mesas={mesas} />}
      {role === "caja" && (
        !usuarioActual ? (
          <UsuarioLogin config={config} persistConfig={persistConfig} onEntrar={onLogin} />
        ) : !tieneRol(usuarioActual, "cajero") && !tieneRol(usuarioActual, "admin") ? (
          <AccesoDenegado mensaje="Tu usuario no tiene el rol de Cajero — esto es Caja." accion="Ir a Mesero" onAccion={() => setRole("mesero")} onCambiarUsuario={() => setUsuarioActual(null)} />
        ) : (
          <CajaView
            menu={menu} orders={orders} cuentas={cuentas} saveCuenta={saveCuenta} saveOrdersMap={saveOrdersMap}
            turnos={turnos} turnoAbierto={turnoAbierto} saveTurno={saveTurno}
            config={config} persistConfig={persistConfig} usuarioActual={usuarioActual}
            mesas={mesas} persistMesas={persistMesas}
            ordersParaHistorial={ordersParaHistorial} cuentasParaHistorial={cuentasParaHistorial}
            historialCompleto={historialCompleto} cargandoHistorial={cargandoHistorial} cargarHistorialCompleto={cargarHistorialCompleto}
          />
        )
      )}
    </div>
  );
}

function AccesoDenegado({ mensaje, accion, onAccion, onCambiarUsuario }) {
  return (
    <div style={styles.screen}>
      <div style={styles.loginWrap}>
        <div style={styles.pageEyebrow}>SIN ACCESO</div>
        <div style={styles.cajaEmptyText}>{mensaje}</div>
        <button style={styles.addItemBtn} onClick={onAccion}><ArrowRightLeft size={15} /> {accion}</button>
        <button style={styles.inlineAddLink} onClick={onCambiarUsuario}>No soy yo — cambiar de usuario</button>
      </div>
    </div>
  );
}

function TopBar({ role, setRole, config, usuarioActual, onCambiarUsuario, online }) {
  return (
    <div className="no-imprimir" style={styles.topBar}>
      <div style={styles.brand}>
        <span style={styles.brandMark}>◆</span>
        <span style={styles.brandText}>{(config.businessName || "LA COMANDA").toUpperCase()}</span>
        <span style={{ ...styles.syncDot, background: online ? "#5B7553" : "#C1442D" }} title={online ? "En línea" : "Sin conexión"} />
      </div>
      {usuarioActual && (
        <button style={styles.meseroChip} onClick={onCambiarUsuario}>
          <UserCircle2 size={13} /> {usuarioActual.nombre} · {rolesLabel(usuarioActual)}
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

/* ---------------- Login único de la app (aplica a Mesero y Caja) ---------------- */

function UsuarioLogin({ config, persistConfig, onEntrar }) {
  const usuarios = config.usuarios || [];
  const hayUsuarios = usuarios.length > 0;

  const [seleccionado, setSeleccionado] = useState(null);
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [intentosFallidos, setIntentosFallidos] = useState(0);
  const [bloqueadoHasta, setBloqueadoHasta] = useState(0);
  const [, forceTickLogin] = useState(0);

  // Cuenta regresiva visible mientras está bloqueado.
  useEffect(() => {
    if (!bloqueadoHasta) return;
    const t = setInterval(() => forceTickLogin((x) => x + 1), 500);
    return () => clearInterval(t);
  }, [bloqueadoHasta]);

  // solo se usan cuando NO hay usuarios todavía (creación del administrador)
  const [nombreAdmin, setNombreAdmin] = useState("");
  const [pinAdmin, setPinAdmin] = useState("");
  const [pinAdmin2, setPinAdmin2] = useState("");

  if (!hayUsuarios) {
    const crear = async () => {
      if (!nombreAdmin.trim()) { setError("Escribe tu nombre."); return; }
      if (pinAdmin.length < 4) { setError("El PIN debe tener al menos 4 dígitos."); return; }
      if (pinAdmin !== pinAdmin2) { setError("Los PIN no coinciden."); return; }
      const admin = { id: uid(), nombre: nombreAdmin.trim(), pin: pinAdmin, roles: ["admin"] };
      const registro = { ts: Date.now(), accion: "Cuenta administradora creada", actor: nombreAdmin.trim() };
      await persistConfig({ ...config, usuarios: [admin], auditLog: [registro] });
      onEntrar({ nombre: admin.nombre, roles: ["admin"] });
    };

    return (
      <div style={styles.screen}>
        <div style={styles.loginWrap}>
          <div style={styles.pageEyebrow}>CONFIGURAR LA APP</div>
          <div style={styles.pageTitle}>Crea el administrador</div>
          <div style={styles.cajaEmptyText}>Esta primera cuenta queda como Administrador — tiene acceso a Mesero y a Caja con un solo login, puede crear todos los demás usuarios (cajeros y meseros), ver sus PIN, aplicar descuentos y todo lo demás.</div>
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
          <div style={styles.pageEyebrow}>IDENTIFÍCATE</div>
          <div style={styles.pageTitle}>¿Quién eres?</div>
          <div style={styles.loginList}>
            {usuarios.map((u) => (
              <button key={u.id} style={styles.loginBtn} onClick={() => { setSeleccionado(u); setError(""); }}>
                <UserCircle2 size={18} /> {u.nombre}
                <span style={styles.rolBadge}>{rolesLabel(u)}</span>
              </button>
            ))}
          </div>
        </div>
      </div>
    );
  }

  const segundosRestantes = Math.max(0, Math.ceil((bloqueadoHasta - Date.now()) / 1000));
  const bloqueado = segundosRestantes > 0;

  const intentar = () => {
    if (bloqueado) return;
    if (pin === seleccionado.pin) {
      setIntentosFallidos(0);
      onEntrar({ nombre: seleccionado.nombre, roles: rolesDe(seleccionado) });
    } else {
      const siguiente = intentosFallidos + 1;
      setIntentosFallidos(siguiente);
      setPin("");
      if (siguiente >= 5) {
        setBloqueadoHasta(Date.now() + 30000);
        setError("Demasiados intentos — espera 30 segundos.");
        setIntentosFallidos(0);
      } else {
        setError(`PIN incorrecto (intento ${siguiente} de 5).`);
      }
    }
  };

  return (
    <div style={styles.screen}>
      <div style={styles.loginWrap}>
        <button style={styles.backBtn} onClick={() => { setSeleccionado(null); setPin(""); setError(""); setIntentosFallidos(0); setBloqueadoHasta(0); }}>← Elegir otro usuario</button>
        <div style={{ ...styles.pageEyebrow, marginTop: 10 }}>HOLA, {seleccionado.nombre.toUpperCase()}</div>
        <div style={styles.pageTitle}>Tu PIN</div>
        <input
          style={styles.editInput} type="password" inputMode="numeric" placeholder="PIN" disabled={bloqueado}
          value={pin} onChange={(e) => { setPin(e.target.value.replace(/\D/g, "")); setError(""); }}
          onKeyDown={(e) => e.key === "Enter" && intentar()}
        />
        {error && <div style={styles.pinError}>{error}</div>}
        <button style={{ ...styles.addItemBtn, marginTop: 10, opacity: bloqueado ? 0.5 : 1 }} disabled={bloqueado} onClick={intentar}>
          <KeyRound size={15} /> {bloqueado ? `Espera ${segundosRestantes}s` : "Entrar"}
        </button>
      </div>
    </div>
  );
}

/* ---------------- MESERO ---------------- */

function MeseroView({ menu, orders, cuentas, usuarioActual, config, mesas, persistMesas, saveOrder, saveOrdersMap, saveCuenta, persistMenu, persistConfig }) {
  // Editar el menú, precios y mesas queda reservado al administrador.
  const meseroActual = usuarioActual.nombre;
  const esAdminMesero = tieneRol(usuarioActual, "admin");

  // Aviso sonoro cuando cocina marca "listo" un pedido de alguna de SUS mesas.
  const idsListosVistosRef = useRef(null);
  const [avisoListo, setAvisoListo] = useState(null);
  useEffect(() => {
    const misCuentaIds = new Set(cuentas.filter((c) => c.mesero === meseroActual && c.estado === "abierta").map((c) => c.id));
    const listosAhora = orders.filter((o) => o.estado === "listo" && misCuentaIds.has(o.cuentaId));
    const idsActuales = new Set(listosAhora.map((o) => o.id));
    if (idsListosVistosRef.current === null) {
      idsListosVistosRef.current = idsActuales;
      return;
    }
    const nuevos = listosAhora.filter((o) => !idsListosVistosRef.current.has(o.id));
    idsListosVistosRef.current = idsActuales;
    if (nuevos.length > 0) {
      playBeep();
      const nombreMesa = mesaNombre(mesas, nuevos[0].mesa);
      const mensaje = nuevos.length === 1 ? `${nombreMesa}: pedido listo para servir` : `${nuevos.length} pedidos listos para servir`;
      setAvisoListo(mensaje);
      mostrarNotificacionSistema("Comandas", mensaje);
      setTimeout(() => setAvisoListo(null), 4500);
    }
  }, [orders, cuentas, meseroActual, mesas]);

  const [selectedMesa, setSelectedMesa] = useState(null);
  const [draft, setDraft] = useState({}); // { itemId: { qty, nota } }
  const [note, setNote] = useState("");
  const [showMenuEditor, setShowMenuEditor] = useState(false);
  const [showMesasEditor, setShowMesasEditor] = useState(false);
  const [editingOrderId, setEditingOrderId] = useState(null);
  const [confirmCancelId, setConfirmCancelId] = useState(null);
  const [justSent, setJustSent] = useState(false);
  const [activeCat, setActiveCat] = useState(null);
  const [search, setSearch] = useState("");
  const [notingItemId, setNotingItemId] = useState(null);
  const [mesaAction, setMesaAction] = useState(null); // 'trasladar' | 'unir'
  const [vistaMesas, setVistaMesas] = useState("grid"); // 'grid' | 'plano'

  const cuentaAbiertaDe = (mesaId) => cuentas.find((c) => mesaMatch(c.mesa, mesaId) && c.estado === "abierta");

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
    const order = orders.find((o) => o.id === orderId);
    if (order) await saveOrder({ ...order, estado: "cancelado" });
    setConfirmCancelId(null);
    if (editingOrderId === orderId) resetDraft();
  };

  const marcarServido = async (orderId) => {
    const order = orders.find((o) => o.id === orderId);
    if (order) await saveOrder({ ...order, estado: "servido" });
  };

  const enviandoRef = useRef(false);
  const [enviando, setEnviando] = useState(false);
  const enviarPedido = async (mesaNum) => {
    if (totalItems === 0) return;
    if (enviandoRef.current) return; // evita duplicar el pedido con doble toque
    enviandoRef.current = true;
    setEnviando(true);
    try {
      // Si alguien borró un producto del menú desde otro dispositivo, se ignora.
      const items = Object.entries(draft)
        .map(([id, { qty, nota }]) => {
          const m = menu.find((x) => x.id === id);
          if (!m) return null;
          // Se guarda el costo de ESTE momento junto al pedido.
          return { id, name: m.name, price: m.price, cost: m.cost || 0, qty, nota: (nota || "").trim() };
        })
        .filter(Boolean);
      if (items.length === 0) { resetDraft(); return; }

      if (editingOrderId) {
        const original = orders.find((o) => o.id === editingOrderId);
        if (original) await saveOrder({ ...original, items, nota: note.trim(), editado: true });
      } else {
        let cuenta = cuentaAbiertaDe(mesaNum);
        if (!cuenta) {
          cuenta = { id: uid(), mesa: mesaNum, ts: Date.now(), estado: "abierta", mesero: meseroActual };
          await saveCuenta(cuenta);
        }
        const order = { id: uid(), mesa: mesaNum, cuentaId: cuenta.id, items, nota: note.trim(), estado: "pendiente", ts: Date.now() };
        await saveOrder(order);
        setJustSent(true);
        setTimeout(() => setJustSent(false), 2200);
      }
      resetDraft();
    } finally {
      enviandoRef.current = false;
      setEnviando(false);
    }
  };

  const mesasLibres = (excluir) => mesas.filter((m) => !mesaMatch(m.id, excluir) && !cuentaAbiertaDe(m.id));
  const mesasOcupadas = (excluir) => mesas.filter((m) => !mesaMatch(m.id, excluir) && cuentaAbiertaDe(m.id));

  const trasladarA = async (cuenta, nuevaMesa) => {
    const afectadas = cuentaOrders(orders, cuenta.id);
    const patch = {};
    afectadas.forEach((o) => { patch[o.id] = { ...o, mesa: nuevaMesa }; });
    if (Object.keys(patch).length > 0) await saveOrdersMap(patch);
    await saveCuenta({ ...cuenta, mesa: nuevaMesa });
    setMesaAction(null);
    setSelectedMesa(nuevaMesa);
  };

  const unirCon = async (cuentaOrigen, cuentaDestino) => {
    const afectadas = cuentaOrders(orders, cuentaOrigen.id);
    const patch = {};
    afectadas.forEach((o) => { patch[o.id] = { ...o, cuentaId: cuentaDestino.id, mesa: cuentaDestino.mesa }; });
    if (Object.keys(patch).length > 0) await saveOrdersMap(patch);
    await saveCuenta({ ...cuentaOrigen, estado: "fusionada", fusionadaEnId: cuentaDestino.id });
    setMesaAction(null);
    setSelectedMesa(cuentaDestino.mesa);
  };

  if (showMenuEditor && esAdminMesero) {
    return <MenuEditor menu={menu} persistMenu={persistMenu} config={config} persistConfig={persistConfig} actor={meseroActual} onClose={() => setShowMenuEditor(false)} />;
  }

  if (showMesasEditor && esAdminMesero) {
    return <MesasEditor mesas={mesas} persistMesas={persistMesas} cuentas={cuentas} orders={orders} config={config} persistConfig={persistConfig} actor={meseroActual} onClose={() => setShowMesasEditor(false)} />;
  }

  if (selectedMesa) {
    const categories = [...new Set(menu.map((m) => m.cat))];
    const cuenta = cuentaAbiertaDe(selectedMesa);
    const pending = cuenta ? cuentaOrders(orders, cuenta.id).filter((o) => o.estado !== "cancelado") : [];

    return (
      <div style={styles.screen}>
        {avisoListo && <div style={styles.avisoListoBanner}><Volume2 size={14} /> {avisoListo}</div>}
        <div style={styles.subHeader}>
          <button style={styles.backBtn} onClick={exitMesa}>← Mesas</button>
          <div style={styles.mesaTitleWrap}>
            <span style={styles.mesaEyebrow}>MESA</span>
            <span style={styles.mesaTitle}>{mesaNombre(mesas, selectedMesa)}</span>
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
              {mesasLibres(selectedMesa).map((m) => (
                <button key={m.id} style={styles.mesaPickerBtn} onClick={() => trasladarA(cuenta, m.id)}>{m.nombre}</button>
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
              {mesasOcupadas(selectedMesa).map((m) => (
                <button key={m.id} style={styles.mesaPickerBtn} onClick={() => unirCon(cuenta, cuentaAbiertaDe(m.id))}>{m.nombre}</button>
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
                    <div style={styles.miniTicketActions}>
                      <button style={styles.marcarServidoBtn} onClick={() => marcarServido(o.id)}><Check size={12} /> Marcar servido</button>
                    </div>
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
            <button style={{ ...styles.sendBtn, flex: 1, opacity: totalItems === 0 || enviando ? 0.4 : 1 }} disabled={totalItems === 0 || enviando} onClick={() => enviarPedido(selectedMesa)}>
              {editingOrderId ? <Check size={16} /> : <Send size={16} />}
              {enviando ? "Enviando…" : editingOrderId ? `Guardar cambios (${totalItems})` : `Enviar pedido ${totalItems > 0 ? `(${totalItems})` : ""}`}
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={styles.screen}>
      {avisoListo && <div style={styles.avisoListoBanner}><Volume2 size={14} /> {avisoListo}</div>}
      <AvisoActivarNotificaciones />
      <div style={styles.mesaGridHeader}>
        <div>
          <div style={styles.pageEyebrow}>SALÓN</div>
          <div style={styles.pageTitle}>Mesas</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          {esAdminMesero ? (
            <>
              <button style={styles.editMenuBtn} onClick={() => setShowMesasEditor(true)}><Settings2 size={15} /> Mesas</button>
              <button style={styles.editMenuBtn} onClick={() => setShowMenuEditor(true)}><Settings2 size={15} /> Menú</button>
            </>
          ) : (
            <span style={styles.lockedNote}><Lock size={11} /> Menú y mesas los administra el admin</span>
          )}
        </div>
      </div>
      <div style={styles.vistaMesasSwitch}>
        <button style={{ ...styles.vistaMesasBtn, ...(vistaMesas === "grid" ? styles.vistaMesasBtnActive : {}) }} onClick={() => setVistaMesas("grid")}>Cuadrícula</button>
        <button style={{ ...styles.vistaMesasBtn, ...(vistaMesas === "plano" ? styles.vistaMesasBtnActive : {}) }} onClick={() => setVistaMesas("plano")}>Plano</button>
      </div>
      {vistaMesas === "plano" ? (
        <div style={{ padding: "0 16px 16px" }}>
          <PlanoSalon mesas={mesas} cuentas={cuentas} orders={orders} editable={false} onTapMesa={enterMesa} />
        </div>
      ) : (
        <div style={styles.mesaGrid}>
          {mesas.length === 0 && (
            <div style={styles.cajaEmptyText}>No hay mesas creadas todavía. Toca "Mesas" para agregar la primera.</div>
          )}
          {mesas.map((m) => {
            const cuenta = cuentaAbiertaDe(m.id);
            const visual = estadoVisualMesa(cuenta, cuenta ? cuentaOrders(orders, cuenta.id) : []);
            return (
              <button key={m.id} style={styles.mesaCard} onClick={() => enterMesa(m.id)}>
                <span style={styles.mesaCardNum}>{m.nombre}</span>
                <span style={{ ...styles.mesaCardTag, background: visual.bg, color: visual.color }}>{visual.tag}</span>
                {cuenta && <span style={styles.mesaCardSince}>desde {timeLabel(cuenta.ts)}</span>}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

function MenuEditor({ menu, persistMenu, config, persistConfig, actor, onClose }) {
  const [name, setName] = useState("");
  const [price, setPrice] = useState("");
  const [cost, setCost] = useState("");
  const [cat, setCat] = useState("");
  const [businessName, setBusinessName] = useState(config.businessName || "");
  const [savedFlash, setSavedFlash] = useState(false);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [editingItemId, setEditingItemId] = useState(null);
  const [editPriceValue, setEditPriceValue] = useState("");
  const [editCostValue, setEditCostValue] = useState("");

  const registrar = (accion) => {
    const auditLog = config.auditLog || [];
    persistConfig({ ...config, auditLog: [{ ts: Date.now(), accion, actor: actor || "desconocido" }, ...auditLog].slice(0, 50) });
  };

  const addItem = () => {
    if (!name.trim() || !price) return;
    const item = { id: uid(), name: name.trim(), price: parseFloat(price) || 0, cost: parseFloat(cost) || 0, cat: cat.trim() || "Otros" };
    persistMenu([...menu, item]);
    registrar(`Agregó "${item.name}" al menú (${money(item.price)})`);
    setName(""); setPrice(""); setCost(""); setCat("");
  };
  const removeItem = (id) => {
    const item = menu.find((m) => m.id === id);
    persistMenu(menu.filter((m) => m.id !== id));
    if (item) registrar(`Eliminó "${item.name}" del menú`);
    setConfirmDeleteId(null);
  };
  const toggleAgotado = (id) => {
    const item = menu.find((m) => m.id === id);
    persistMenu(menu.map((m) => (m.id === id ? { ...m, agotado: !m.agotado } : m)));
    if (item) registrar(`Marcó "${item.name}" como ${item.agotado ? "disponible" : "agotado"}`);
  };
  const empezarEdicionItem = (item) => {
    setEditingItemId(item.id);
    setEditPriceValue(item.price ? String(item.price) : "");
    setEditCostValue(item.cost ? String(item.cost) : "");
  };
  const guardarEdicionItem = (item) => {
    const nuevoPrecio = parseFloat(editPriceValue) || 0;
    const nuevoCosto = parseFloat(editCostValue) || 0;
    persistMenu(menu.map((m) => (m.id === item.id ? { ...m, price: nuevoPrecio, cost: nuevoCosto } : m)));
    if (nuevoPrecio !== item.price) registrar(`Cambió el precio de "${item.name}" de ${money(item.price)} a ${money(nuevoPrecio)}`);
    if (nuevoCosto !== (item.cost || 0)) registrar(`Actualizó el costo de "${item.name}" a ${money(nuevoCosto)}`);
    setEditingItemId(null); setEditPriceValue(""); setEditCostValue("");
  };

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
            <CampoMonto style={{ ...styles.editInput, flex: 1 }} placeholder="Precio de venta" value={price} onChange={setPrice} />
            <CampoMonto style={{ ...styles.editInput, flex: 1 }} placeholder="Costo (opcional)" value={cost} onChange={setCost} />
          </div>
          <input style={{ ...styles.editInput, flex: 1 }} placeholder="Categoría" value={cat} onChange={(e) => setCat(e.target.value)} />
          <button style={styles.addItemBtn} onClick={addItem}><Plus size={15} /> Agregar al menú</button>
        </div>

        {[...new Set(menu.map((m) => m.cat))].map((c) => (
          <div key={c} style={{ marginBottom: 18 }}>
            <div style={styles.catLabel}>{c}</div>
            {menu.filter((m) => m.cat === c).map((item) => {
              const margen = item.price - (item.cost || 0);
              const margenPct = item.price > 0 ? Math.round((margen / item.price) * 100) : 0;
              return (
                <div key={item.id} style={styles.menuEditItemBlock}>
                  <div style={{ ...styles.editRow, borderBottom: "none" }}>
                    <button
                      style={{ ...styles.agotadoToggle, ...(item.agotado ? styles.agotadoToggleActive : {}) }}
                      onClick={() => toggleAgotado(item.id)}
                      title={item.agotado ? "Marcar disponible" : "Marcar agotado"}
                    >
                      <PauseCircle size={14} />
                    </button>
                    <span style={{ ...styles.menuItemName, ...(item.agotado ? { textDecoration: "line-through", color: "#8B80A3" } : {}) }}>{item.name}</span>
                    {editingItemId !== item.id && (
                      <button style={styles.precioEditableBtn} onClick={() => empezarEdicionItem(item)}>{money(item.price)} <Pencil size={10} /></button>
                    )}
                    {confirmDeleteId === item.id ? (
                      <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                        <button style={styles.confirmYes} onClick={() => removeItem(item.id)}>Sí</button>
                        <button style={styles.confirmNo} onClick={() => setConfirmDeleteId(null)}>No</button>
                      </div>
                    ) : (
                      <button style={styles.deleteBtn} onClick={() => setConfirmDeleteId(item.id)}><Trash2 size={14} /></button>
                    )}
                  </div>
                  {editingItemId === item.id ? (
                    <div style={styles.costoEditRow}>
                      <CampoMonto style={{ ...styles.editInput, flex: 1 }} placeholder="Precio de venta" value={editPriceValue} onChange={setEditPriceValue} autoFocus />
                      <CampoMonto style={{ ...styles.editInput, flex: 1 }} placeholder="Costo" value={editCostValue} onChange={setEditCostValue} />
                      <button style={styles.confirmYes} onClick={() => guardarEdicionItem(item)}>Guardar</button>
                      <button style={styles.confirmNo} onClick={() => setEditingItemId(null)}>Cancelar</button>
                    </div>
                  ) : (
                    <button style={styles.costoPillBtn} onClick={() => empezarEdicionItem(item)}>
                      {item.cost ? (
                        <>Costo {money(item.cost)} · Margen {money(margen)} ({margenPct}%)</>
                      ) : (
                        <>+ Agregar costo de producción (para ver el margen)</>
                      )}
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        ))}
      </div>
    </div>
  );
}

/** Plano visual del salón: cada mesa se dibuja en su posición real (x%, y%).
 * En modo "editable" se puede arrastrar cada mesa; si no, solo se toca para entrar. */
function PlanoSalon({ mesas, cuentas, orders, editable, onTapMesa, onMoverMesa }) {
  const containerRef = useRef(null);
  const [arrastrando, setArrastrando] = useState(null);
  const [posLive, setPosLive] = useState(null);

  const posDefault = (idx) => {
    const cols = 4;
    const col = idx % cols;
    const row = Math.floor(idx / cols);
    return { x: 15 + col * 24, y: 20 + row * 26 };
  };
  const posDe = (m, idx) => (typeof m.x === "number" && typeof m.y === "number") ? { x: m.x, y: m.y } : posDefault(idx);

  const onPointerDown = (e, m) => {
    if (!editable) return;
    e.currentTarget.setPointerCapture(e.pointerId);
    setArrastrando(m.id);
  };
  const onPointerMove = (e) => {
    if (!arrastrando || !containerRef.current) return;
    const rect = containerRef.current.getBoundingClientRect();
    let x = ((e.clientX - rect.left) / rect.width) * 100;
    let y = ((e.clientY - rect.top) / rect.height) * 100;
    x = Math.max(4, Math.min(96, x));
    y = Math.max(6, Math.min(94, y));
    setPosLive({ x, y });
  };
  const onPointerUp = () => {
    if (arrastrando && posLive) onMoverMesa(arrastrando, posLive.x, posLive.y);
    setArrastrando(null);
    setPosLive(null);
  };

  return (
    <div ref={containerRef} style={styles.planoContainer} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerLeave={onPointerUp}>
      {mesas.length === 0 && <div style={styles.planoVacio}>Agrega mesas para verlas aquí.</div>}
      {mesas.map((m, idx) => {
        const cuenta = cuentas.find((c) => mesaMatch(c.mesa, m.id) && c.estado === "abierta");
        const visual = estadoVisualMesa(cuenta, cuenta ? orders.filter((o) => o.cuentaId === cuenta.id) : []);
        const bg = cuenta ? visual.bg : "#FFFFFF";
        const color = cuenta ? visual.color : "#1B1033";
        const pos = arrastrando === m.id && posLive ? posLive : posDe(m, idx);
        return (
          <button
            key={m.id}
            onPointerDown={(e) => onPointerDown(e, m)}
            onClick={() => { if (!editable) onTapMesa(m.id); }}
            style={{
              ...styles.planoMesaTile,
              left: `${pos.x}%`, top: `${pos.y}%`,
              background: bg, color,
              borderColor: cuenta ? bg : "#D9CCEE",
              cursor: editable ? "grab" : "pointer",
              touchAction: editable ? "none" : "auto",
              zIndex: arrastrando === m.id ? 2 : 1,
            }}
          >
            {m.nombre}
          </button>
        );
      })}
    </div>
  );
}

function MesasEditor({ mesas, persistMesas, cuentas, orders, config, persistConfig, actor, onClose }) {
  const [nombre, setNombre] = useState("");
  const [editId, setEditId] = useState(null);
  const [editNombre, setEditNombre] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [error, setError] = useState("");

  const tieneCuentaAbierta = (mesaId) => cuentas.some((c) => mesaMatch(c.mesa, mesaId) && c.estado === "abierta");

  const registrar = (accion) => {
    if (!config || !persistConfig) return;
    const auditLog = config.auditLog || [];
    persistConfig({ ...config, auditLog: [{ ts: Date.now(), accion, actor: actor || "desconocido" }, ...auditLog].slice(0, 50) });
  };

  const agregarMesa = () => {
    if (!nombre.trim()) return;
    const nueva = { id: uid(), nombre: nombre.trim() };
    persistMesas([...mesas, nueva]);
    registrar(`Agregó la mesa/espacio "${nueva.nombre}"`);
    setNombre("");
  };

  const empezarEdicion = (m) => { setEditId(m.id); setEditNombre(m.nombre); setError(""); };
  const guardarEdicion = (m) => {
    if (!editNombre.trim()) return;
    persistMesas(mesas.map((x) => (x.id === m.id ? { ...x, nombre: editNombre.trim() } : x)));
    registrar(`Renombró "${m.nombre}" a "${editNombre.trim()}"`);
    setEditId(null); setEditNombre("");
  };

  const eliminarMesa = (m) => {
    if (tieneCuentaAbierta(m.id)) { setError("Esa mesa tiene una cuenta abierta — ciérrala o cóbrala antes de eliminarla."); setConfirmDeleteId(null); return; }
    persistMesas(mesas.filter((x) => x.id !== m.id));
    registrar(`Eliminó la mesa/espacio "${m.nombre}"`);
    setConfirmDeleteId(null);
  };

  const moverEnPlano = (id, x, y) => {
    persistMesas(mesas.map((m) => (m.id === id ? { ...m, x, y } : m)));
  };

  const mover = (idx, dir) => {
    const next = [...mesas];
    const j = idx + dir;
    if (j < 0 || j >= next.length) return;
    [next[idx], next[j]] = [next[j], next[idx]];
    persistMesas(next);
  };

  return (
    <div style={styles.screen}>
      <div style={styles.subHeader}>
        <button style={styles.backBtn} onClick={onClose}>← Mesas</button>
        <div style={styles.mesaTitleWrap}><span style={styles.mesaEyebrow}>AJUSTES</span><span style={styles.mesaTitle}>Mesas y espacios</span></div>
        <div style={{ width: 70 }} />
      </div>
      <div style={styles.menuScroll}>
        <div style={styles.addItemCard}>
          <div style={styles.catLabel}>AGREGAR MESA O ESPACIO</div>
          <div style={styles.closeConfirmText}>Puedes ponerle cualquier nombre: "Mesa 5", "Barra", "Mesa gris", "Terraza"…</div>
          <input style={styles.editInput} placeholder="Nombre del espacio" value={nombre} onChange={(e) => setNombre(e.target.value)} onKeyDown={(e) => e.key === "Enter" && agregarMesa()} />
          <button style={styles.addItemBtn} onClick={agregarMesa}><Plus size={15} /> Agregar</button>
        </div>

        {error && <div style={styles.pinError}>{error}</div>}

        <div style={styles.catLabel}>UBICACIÓN EN EL PLANO</div>
        <div style={styles.closeConfirmText}>Arrastra cada mesa a su lugar real del salón — así se ve en Mesero cuando cambias a vista "Plano".</div>
        <PlanoSalon mesas={mesas} cuentas={cuentas} orders={orders} editable onMoverMesa={moverEnPlano} />

        <div style={{ ...styles.catLabel, marginTop: 18 }}>TUS MESAS Y ESPACIOS</div>
        {mesas.length === 0 && <div style={styles.cajaEmptyText}>Todavía no has agregado ninguna.</div>}
        {mesas.map((m, idx) => (
          <div key={m.id} style={styles.editRow}>
            <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
              <button style={styles.mesaReorderBtn} onClick={() => mover(idx, -1)} disabled={idx === 0}>▲</button>
              <button style={styles.mesaReorderBtn} onClick={() => mover(idx, 1)} disabled={idx === mesas.length - 1}>▼</button>
            </div>
            {editId === m.id ? (
              <>
                <input autoFocus style={{ ...styles.editInput, flex: 1 }} value={editNombre} onChange={(e) => setEditNombre(e.target.value)} onKeyDown={(e) => e.key === "Enter" && guardarEdicion(m)} />
                <button style={styles.confirmYes} onClick={() => guardarEdicion(m)}>Guardar</button>
                <button style={styles.confirmNo} onClick={() => setEditId(null)}>Cancelar</button>
              </>
            ) : (
              <>
                <span style={{ ...styles.menuItemName, flex: 1 }}>{m.nombre}</span>
                {tieneCuentaAbierta(m.id) && <span style={styles.enCursoTag}>cuenta abierta</span>}
                <button style={styles.deleteBtn} onClick={() => empezarEdicion(m)}><Pencil size={14} /></button>
                {confirmDeleteId === m.id ? (
                  <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    <button style={styles.confirmYes} onClick={() => eliminarMesa(m)}>Sí</button>
                    <button style={styles.confirmNo} onClick={() => setConfirmDeleteId(null)}>No</button>
                  </div>
                ) : (
                  <button style={styles.deleteBtn} onClick={() => { setConfirmDeleteId(m.id); setError(""); }}><Trash2 size={14} /></button>
                )}
              </>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---------------- COCINA ---------------- */

function CocinaView({ orders, saveOrder, mesas }) {
  const active = orders.filter((o) => o.estado !== "listo" && o.estado !== "servido" && o.estado !== "cancelado").sort((a, b) => a.ts - b.ts);
  const recienListos = orders
    .filter((o) => o.estado === "listo" && o.listoTs && Date.now() - o.listoTs < 10 * 60 * 1000)
    .sort((a, b) => b.listoTs - a.listoTs);
  const prevCountRef = useRef(active.length);
  const [flash, setFlash] = useState(false);

  useEffect(() => {
    if (active.length > prevCountRef.current) {
      playBeep();
      setFlash(true);
      mostrarNotificacionSistema("Comandas — Cocina", "Llegó un pedido nuevo");
      setTimeout(() => setFlash(false), 900);
    }
    prevCountRef.current = active.length;
  }, [active.length]);

  const advance = async (id) => {
    const o = orders.find((x) => x.id === id);
    if (!o) return;
    if (o.estado === "pendiente") { await saveOrder({ ...o, estado: "preparando" }); return; }
    await saveOrder({ ...o, estado: "listo", listoTs: Date.now() });
  };
  const retroceder = async (id) => {
    const o = orders.find((x) => x.id === id);
    if (!o) return;
    await saveOrder({ ...o, estado: "pendiente" });
  };
  const deshacerListo = async (id) => {
    const o = orders.find((x) => x.id === id);
    if (!o) return;
    await saveOrder({ ...o, estado: "preparando" });
  };

  return (
    <div style={{ ...styles.screenDark, ...(flash ? styles.screenFlash : {}) }}>
      <AvisoActivarNotificaciones />
      <div style={styles.kdsHeader}>
        <div>
          <div style={styles.kdsEyebrow}>COCINA — PANTALLA EN VIVO <Volume2 size={11} style={{ verticalAlign: -2, marginLeft: 4 }} /></div>
          <div style={styles.kdsTitle}>Comandas activas</div>
        </div>
        <div style={styles.kdsCount}>{active.length}</div>
      </div>

      {recienListos.length > 0 && (
        <div style={styles.recienListosRow}>
          {recienListos.map((o) => (
            <button key={o.id} style={styles.recienListoChip} onClick={() => deshacerListo(o.id)}>
              ↺ {mesaNombre(mesas, o.mesa)} — deshacer "listo"
            </button>
          ))}
        </div>
      )}

      {active.length === 0 ? (
        <div style={styles.emptyKitchen}><Flame size={26} strokeWidth={1.5} color="#6D5BA3" /><div style={styles.emptyKitchenText}>Sin pedidos pendientes</div></div>
      ) : (
        <div style={styles.rail}>
          {active.map((o) => (
            <div key={o.id} style={{ ...styles.ticket, borderLeftColor: urgencyColor(o.ts) }}>
              <div style={styles.ticketHead}>
                <span style={styles.ticketMesa}>{mesaNombre(mesas, o.mesa).toUpperCase()}{o.editado && <span style={styles.editedBadge}>MODIFICADO</span>}</span>
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
              <div style={{ display: "flex", gap: 6 }}>
                {o.estado === "preparando" && (
                  <button style={styles.ticketBackBtn} onClick={() => retroceder(o.id)} title="Volver a pendiente">◀</button>
                )}
                <button style={{ ...styles.ticketBtn, flex: 1, background: o.estado === "pendiente" ? "#B98A2E" : "#5B7553" }} onClick={() => advance(o.id)}>
                  {o.estado === "pendiente" ? <><Flame size={14} /> Empezar preparación</> : <><Check size={14} /> Marcar listo</>}
                </button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ---------------- CAJA ---------------- */

function CajaView({ menu, orders, cuentas, saveCuenta, saveOrdersMap, turnos, turnoAbierto, saveTurno, config, persistConfig, usuarioActual, mesas, persistMesas, ordersParaHistorial, cuentasParaHistorial, historialCompleto, cargandoHistorial, cargarHistorialCompleto }) {
  const [tab, setTab] = useState("mesas");
  const [expandedId, setExpandedId] = useState(null);
  const esAdmin = tieneRol(usuarioActual, "admin");

  const abiertas = cuentas.filter((c) => c.estado === "abierta").sort((a, b) => mesaOrderIndex(mesas, a.mesa) - mesaOrderIndex(mesas, b.mesa));
  const hoyKey = dayKey(Date.now());
  const pagadasHoy = cuentas.filter((c) => c.estado === "pagada" && dayKey(c.pagadaTs) === hoyKey);
  const vendidoHoy = pagadasHoy.reduce((s, c) => s + cuentaGranTotal(orders, c), 0);
  const porMetodoHoy = { efectivo: 0, tarjeta: 0, transferencia: 0 };
  pagadasHoy.forEach((c) => pagosDe(c).forEach((p) => { porMetodoHoy[p.metodoPago] = (porMetodoHoy[p.metodoPago] || 0) + p.monto; }));

  return (
    <div style={styles.screen}>
      <div className="no-imprimir" style={styles.mesaGridHeader}>
        <div><div style={styles.pageEyebrow}>CAJA</div><div style={styles.pageTitle}>Control de consumo</div></div>
      </div>

      <div className="no-imprimir" style={styles.cajaTabSwitch}>
        <button style={{ ...styles.cajaTabBtn, ...(tab === "turno" ? styles.cajaTabBtnActive : {}) }} onClick={() => setTab("turno")}>Turno</button>
        <button style={{ ...styles.cajaTabBtn, ...(tab === "mesas" ? styles.cajaTabBtnActive : {}) }} onClick={() => setTab("mesas")}>Mesas {abiertas.length > 0 ? `(${abiertas.length})` : ""}</button>
        {esAdmin && <button style={{ ...styles.cajaTabBtn, ...(tab === "historial" ? styles.cajaTabBtnActive : {}) }} onClick={() => setTab("historial")}>Historial</button>}
        <button style={{ ...styles.cajaTabBtn, ...(tab === "reportes" ? styles.cajaTabBtnActive : {}) }} onClick={() => setTab("reportes")}>Reportes</button>
      </div>

      {tab === "turno" && (
        <TurnoView turnos={turnos} turnoAbierto={turnoAbierto} saveTurno={saveTurno} orders={orders} cuentas={cuentas} config={config} persistConfig={persistConfig} usuarioActual={usuarioActual} menu={menu} mesas={mesas} ordersParaHistorial={ordersParaHistorial} cuentasParaHistorial={cuentasParaHistorial} historialCompleto={historialCompleto} cargandoHistorial={cargandoHistorial} cargarHistorialCompleto={cargarHistorialCompleto} />
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
                    <Icon size={13} color="#A598C8" /><span style={styles.metodoBreakdownLabel}>{meta.label}</span><span style={styles.metodoBreakdownValue}>{money(porMetodoHoy[key] || 0)}</span>
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
                key={c.id} cuenta={c} orders={orders} saveCuenta={saveCuenta} saveOrdersMap={saveOrdersMap}
                turnoAbierto={!!turnoAbierto} usuarioActual={usuarioActual} mesas={mesas}
                expanded={expandedId === c.id} onToggle={() => setExpandedId(expandedId === c.id ? null : c.id)}
              />
            ))
          )}
          {esAdmin && cuentas.some((c) => c.estado === "pagada") && (
            <div style={styles.goHistLink} onClick={() => setTab("historial")}>Ver historial completo por día, mes y año →</div>
          )}
        </div>
      )}

      {tab === "historial" && esAdmin && <HistorialView cuentasParaHistorial={cuentasParaHistorial} ordersParaHistorial={ordersParaHistorial} historialCompleto={historialCompleto} cargandoHistorial={cargandoHistorial} cargarHistorialCompleto={cargarHistorialCompleto} config={config} mesas={mesas} usuarioActual={usuarioActual} />}
      {tab === "reportes" && <ReportesView cuentas={cuentas} orders={orders} menu={menu} usuarioActual={usuarioActual} />}
    </div>
  );
}

/* ---------------- Turno / arqueo de caja ---------------- */

function TurnoView({ turnos, turnoAbierto, saveTurno, orders, cuentas, config, persistConfig, usuarioActual, menu, mesas, ordersParaHistorial, cuentasParaHistorial, historialCompleto, cargandoHistorial, cargarHistorialCompleto }) {
  const [base, setBase] = useState("");
  const [contado, setContado] = useState("");
  const [observaciones, setObservaciones] = useState("");
  const [confirmandoCierre, setConfirmandoCierre] = useState(false);

  const mesasAbiertas = cuentas.filter((c) => c.estado === "abierta");
  const esAdmin = tieneRol(usuarioActual, "admin");

  const abrirTurno = async () => {
    const baseNum = parseFloat(base) || 0;
    const turno = { id: uid(), aperturaTs: Date.now(), baseEfectivo: baseNum, estado: "abierto", abiertoPor: usuarioActual.nombre };
    await saveTurno(turno);
    setBase("");
  };

  // Efectivo, tarjeta y transferencia recibidos en este turno.
  const totalesPorMetodo = () => {
    const totales = { efectivo: turnoAbierto ? turnoAbierto.baseEfectivo : 0, tarjeta: 0, transferencia: 0 };
    if (!turnoAbierto) return totales;
    cuentas.filter((c) => c.estado === "pagada" && c.pagadaTs >= turnoAbierto.aperturaTs).forEach((c) => {
      pagosDe(c).forEach((p) => { if (totales[p.metodoPago] !== undefined) totales[p.metodoPago] += p.monto; });
    });
    return totales;
  };

  const { efectivo: esperado, tarjeta: tarjetaEsperada, transferencia: transferenciaEsperada } = totalesPorMetodo();
  const contadoNum = parseFloat(contado) || 0;
  const diferencia = contadoNum - esperado;

  const cerrarTurno = async () => {
    await saveTurno({
      ...turnoAbierto, estado: "cerrado", cierreTs: Date.now(),
      esperado, contado: contadoNum, diferencia,
      tarjetaEsperada, transferenciaEsperada,
      observaciones: observaciones.trim() || null,
      cerradoPor: usuarioActual.nombre,
    });
    setContado(""); setObservaciones(""); setConfirmandoCierre(false);
  };

  const cerrados = turnos.filter((t) => t.estado === "cerrado").sort((a, b) => b.cierreTs - a.cierreTs);

  return (
    <div style={styles.menuScroll}>
      {!turnoAbierto ? (
        <div style={styles.turnoBox}>
          <div style={styles.pageEyebrow}>TURNO CERRADO</div>
          <div style={styles.closeConfirmText}>Registra cuánto efectivo hay en caja para empezar el turno.</div>
          <CampoMonto placeholder="Base de efectivo inicial" value={base} onChange={setBase} />
          <button style={{ ...styles.addItemBtn, marginTop: 8 }} onClick={abrirTurno}><PlayCircle size={15} /> Abrir turno</button>
        </div>
      ) : (
        <div style={styles.turnoBox}>
          <div style={styles.pageEyebrow}>TURNO ABIERTO</div>
          <div style={styles.closeConfirmText}>Abierto {dateTimeLabel(turnoAbierto.aperturaTs)} · base {money(turnoAbierto.baseEfectivo)}</div>
          <div style={styles.cajaTotalValue2}>{money(esperado)}<span style={styles.turnoEsperadoLabel}>efectivo esperado en caja ahora</span></div>

          <div style={styles.metodoResumenRow}>
            <div style={styles.metodoResumenItem}>
              <CreditCard size={14} color="#2F6690" />
              <span style={styles.metodoResumenLabel}>Tarjeta</span>
              <span style={styles.metodoResumenValue}>{money(tarjetaEsperada)}</span>
            </div>
            <div style={styles.metodoResumenItem}>
              <Smartphone size={14} color="#8A611A" />
              <span style={styles.metodoResumenLabel}>Transferencia</span>
              <span style={styles.metodoResumenValue}>{money(transferenciaEsperada)}</span>
            </div>
          </div>
          <div style={styles.closeConfirmText}>Compara estos dos contra el cierre de lote del datáfono y tu banco — no hay nada físico que contar, pero sí vale la pena revisar que cuadren.</div>

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
                  <b>Ojo:</b> {mesasAbiertas.length} {mesasAbiertas.length === 1 ? "mesa sigue" : "mesas siguen"} sin cobrar ({mesasAbiertas.map((c) => mesaNombre(mesas, c.mesa)).join(", ")}). Si cierras ahora, ese dinero no quedará contado en este turno.
                </div>
              )}
              <div style={styles.closeConfirmText}>Cuenta el efectivo físico en caja y escríbelo aquí:</div>
              <CampoMonto placeholder="Efectivo contado" value={contado} onChange={setContado} />
              {contado !== "" && (
                <div style={{ ...styles.diferenciaBox, background: diferencia === 0 ? "#E4EEE1" : diferencia > 0 ? "#FBEFD9" : "#FBE2DC" }}>
                  {diferencia === 0 ? "Cuadra exacto ✓" : diferencia > 0 ? `Sobran ${money(diferencia)}` : `Faltan ${money(Math.abs(diferencia))}`}
                </div>
              )}
              <div style={styles.closeConfirmText}>Tarjeta esperada: <b>{money(tarjetaEsperada)}</b> · Transferencia esperada: <b>{money(transferenciaEsperada)}</b></div>
              {contado !== "" && diferencia !== 0 ? (
                <>
                  <div style={styles.closeConfirmText}><b>El efectivo no cuadró — cuéntanos por qué (obligatorio):</b></div>
                  <input autoFocus style={styles.editInput} placeholder="Ej: se dio mal un vuelto, faltó registrar un pago…" value={observaciones} onChange={(e) => setObservaciones(e.target.value)} />
                </>
              ) : (
                <input style={styles.editInput} placeholder="Observaciones si algo de tarjeta/transferencia no cuadra (opcional)" value={observaciones} onChange={(e) => setObservaciones(e.target.value)} />
              )}
              <div style={{ display: "flex", gap: 8 }}>
                <button style={styles.cancelEditBtn} onClick={() => setConfirmandoCierre(false)}>Cancelar</button>
                <button
                  style={{ ...styles.sendBtn, flex: 1, opacity: contado !== "" && (diferencia === 0 || observaciones.trim()) ? 1 : 0.4 }}
                  disabled={contado === "" || (diferencia !== 0 && !observaciones.trim())}
                  onClick={cerrarTurno}
                >
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
            <div key={t.id} style={styles.turnoRowCol}>
              <div style={{ ...styles.turnoRow, border: "none", padding: 0 }}>
                <span style={styles.cierreDate}>{dateTimeLabel(t.aperturaTs)} → {timeLabel(t.cierreTs)}</span>
                <span style={{ ...styles.turnoDiff, color: t.diferencia === 0 ? "#5B7553" : "#C1442D" }}>
                  {t.diferencia === 0 ? "Cuadró" : t.diferencia > 0 ? `+${money(t.diferencia)}` : money(t.diferencia)}
                </span>
              </div>
              {(t.tarjetaEsperada > 0 || t.transferenciaEsperada > 0) && (
                <div style={styles.turnoSubRow}>Tarjeta {money(t.tarjetaEsperada || 0)} · Transferencia {money(t.transferenciaEsperada || 0)}</div>
              )}
              {t.observaciones && <div style={styles.turnoObsRow}>"{t.observaciones}"</div>}
            </div>
          ))}
        </>
      )}

      {esAdmin && (
        <RespaldoBox menu={menu} orders={orders} cuentas={cuentas} config={config} turnos={turnos} mesas={mesas} persistConfig={persistConfig} usuarioActual={usuarioActual} ordersParaHistorial={ordersParaHistorial} cuentasParaHistorial={cuentasParaHistorial} historialCompleto={historialCompleto} cargandoHistorial={cargandoHistorial} cargarHistorialCompleto={cargarHistorialCompleto} />
      )}

      {esAdmin && <PanelSenalesAlerta cuentas={cuentasParaHistorial} turnos={turnos} />}

      <div style={{ ...styles.catLabel, marginTop: 24 }}>USUARIOS Y SEGURIDAD</div>
      {esAdmin ? (
        <GestionUsuarios config={config} persistConfig={persistConfig} usuarioActual={usuarioActual} />
      ) : (
        <div style={styles.lockedNote}><Lock size={11} /> Solo un administrador puede gestionar usuarios y ver el registro de seguridad.</div>
      )}
    </div>
  );
}

function RespaldoBox({ menu, orders, cuentas, config, turnos, mesas, persistConfig, usuarioActual, ordersParaHistorial, cuentasParaHistorial, historialCompleto, cargandoHistorial, cargarHistorialCompleto }) {
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
    await saveCollectionArray("orders", archivo.orders);
    await saveCollectionArray("cuentas", archivo.cuentas);
    await saveCollectionArray("turnos", archivo.turnos || []);
    if (archivo.mesas && archivo.mesas.length > 0) await saveShared("mesas", archivo.mesas);
    const nuevoConfig = {
      ...archivo.config,
      auditLog: [{ ts: Date.now(), accion: "Restauró un respaldo desde archivo", actor: usuarioActual.nombre }, ...(archivo.config.auditLog || [])].slice(0, 50),
    };
    await saveShared("config", nuevoConfig);
    window.location.reload();
  };

  const descargarPendiente = useRef(false);
  // Una vez que el historial completo termina de cargar, dispara la descarga.
  useEffect(() => {
    if (descargarPendiente.current && historialCompleto) {
      descargarPendiente.current = false;
      descargarRespaldo({ menu, orders: ordersParaHistorial, cuentas: cuentasParaHistorial, config, turnos, mesas });
    }
  }, [historialCompleto]);
  const onClickDescargar = () => {
    if (historialCompleto) {
      descargarRespaldo({ menu, orders: ordersParaHistorial, cuentas: cuentasParaHistorial, config, turnos, mesas });
    } else {
      descargarPendiente.current = true;
      cargarHistorialCompleto();
    }
  };

  return (
    <div style={{ ...styles.catLabel, marginTop: 24 }}>
      RESPALDO
      <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 8 }}>
        <button style={styles.modoChoiceBtn} onClick={onClickDescargar} disabled={cargandoHistorial}>
          <Download size={15} /> {cargandoHistorial ? "Preparando respaldo completo…" : "Descargar copia de seguridad completa"}
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
      <RespaldosAutomaticosBox />
    </div>
  );
}

function RespaldosAutomaticosBox() {
  const [dias, setDias] = useState(null); // null = sin cargar todavía
  const [cargando, setCargando] = useState(false);

  const cargarLista = async () => {
    setCargando(true);
    const hoy = new Date();
    const candidatos = [];
    for (let i = 1; i <= 14; i++) {
      const d = new Date(hoy.getTime() - i * 24 * 60 * 60 * 1000);
      candidatos.push(dayKey(d.getTime()));
    }
    const resultados = await Promise.all(candidatos.map((dk) => getRecord(`respaldos/${dk}`)));
    setDias(resultados.filter(Boolean).sort((a, b) => b.dia.localeCompare(a.dia)));
    setCargando(false);
  };

  const descargarDia = (r) => {
    const data = { version: 2, exportadoTs: Date.now(), menu: r.menu, orders: r.orders, cuentas: r.cuentas, config: r.config, turnos: [], mesas: r.mesas };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `respaldo_automatico_${r.dia}.json`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div style={{ marginTop: 14 }}>
      {dias === null ? (
        <button style={styles.inlineAddLink} onClick={cargarLista} disabled={cargando}>
          {cargando ? "Buscando…" : "Ver copias automáticas de los últimos 14 días"}
        </button>
      ) : dias.length === 0 ? (
        <div style={styles.cajaEmptyText}>Todavía no hay copias automáticas (se generan solas cada día que un admin abre la app — la de hoy se hace mañana).</div>
      ) : (
        <>
          <div style={styles.catLabel}>COPIAS AUTOMÁTICAS RECIENTES</div>
          {dias.map((r) => (
            <div key={r.dia} style={styles.turnoRow}>
              <span style={styles.cierreDate}>{r.dia} · {r.cuentas.length} {r.cuentas.length === 1 ? "cuenta" : "cuentas"}</span>
              <button style={styles.inlineLink} onClick={() => descargarDia(r)}><Download size={11} style={{ verticalAlign: -1 }} /> descargar</button>
            </div>
          ))}
        </>
      )}
    </div>
  );
}

/** Junta en un solo lugar las señales de alerta (cancelaciones, descuentos,
 * descuadres de turno) para que el admin decida si algo amerita revisión. */
function PanelSenalesAlerta({ cuentas, turnos }) {
  const [expandido, setExpandido] = useState(false);

  const porPersona = {};
  const tocar = (nombre) => {
    if (!porPersona[nombre]) porPersona[nombre] = { cancelaciones: 0, montoCancelado: 0, descuentos: 0, montoDescontado: 0, turnosConFaltante: 0, montoFaltante: 0 };
    return porPersona[nombre];
  };

  cuentas.filter((c) => c.estado === "cancelada" && (c.montoCancelado || 0) > 0 && c.canceladaPor).forEach((c) => {
    const p = tocar(c.canceladaPor);
    p.cancelaciones += 1;
    p.montoCancelado += c.montoCancelado || 0;
  });
  cuentas.filter((c) => c.descuento && c.descuento.aplicadoPor).forEach((c) => {
    const p = tocar(c.descuento.aplicadoPor);
    p.descuentos += 1;
    p.montoDescontado += c.descuento.monto || 0;
  });
  turnos.filter((t) => t.estado === "cerrado" && t.diferencia < 0 && t.cerradoPor).forEach((t) => {
    const p = tocar(t.cerradoPor);
    p.turnosConFaltante += 1;
    p.montoFaltante += Math.abs(t.diferencia);
  });

  const filas = Object.entries(porPersona)
    .map(([nombre, d]) => ({ nombre, ...d, alerta: d.cancelaciones >= 3 || d.turnosConFaltante >= 2 }))
    .sort((a, b) => (b.alerta ? 1 : 0) - (a.alerta ? 1 : 0) || (b.montoCancelado + b.montoFaltante) - (a.montoCancelado + a.montoFaltante));

  const hayAlertas = filas.some((f) => f.alerta);

  if (filas.length === 0) return null;

  return (
    <div style={{ marginTop: 24 }}>
      <button style={styles.canceladasToggle} onClick={() => setExpandido((v) => !v)}>
        <AlertTriangle size={13} color={hayAlertas ? "#C1442D" : "#B98A2E"} />
        <span style={{ color: hayAlertas ? "#8A3A2B" : "#4A3F66" }}>Señales de alerta (últimos 35 días){hayAlertas ? " — hay algo que revisar" : ""}</span>
        <span>{expandido ? "▲" : "▼"}</span>
      </button>
      {expandido && (
        <div style={{ marginTop: 8 }}>
          {filas.map((f) => (
            <div key={f.nombre} style={{ ...styles.turnoRowCol, ...(f.alerta ? { background: "#FBE2DC", borderRadius: 8, padding: "8px 10px" } : {}) }}>
              <div style={{ ...styles.turnoRow, border: "none", padding: 0 }}>
                <span style={styles.pagoRowEtiqueta}>{f.nombre}</span>
                {f.alerta && <AlertTriangle size={13} color="#C1442D" />}
              </div>
              <div style={styles.turnoSubRow}>
                {f.cancelaciones > 0 && <>· {f.cancelaciones} {f.cancelaciones === 1 ? "cancelación" : "cancelaciones"} ({money(f.montoCancelado)}) </>}
                {f.descuentos > 0 && <>· {f.descuentos} {f.descuentos === 1 ? "descuento" : "descuentos"} ({money(f.montoDescontado)}) </>}
                {f.turnosConFaltante > 0 && <>· {f.turnosConFaltante} {f.turnosConFaltante === 1 ? "turno con faltante" : "turnos con faltante"} ({money(f.montoFaltante)})</>}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function GestionUsuarios({ config, persistConfig, usuarioActual }) {
  const usuarios = config.usuarios || [];
  const auditLog = config.auditLog || [];
  const [nombre, setNombre] = useState("");
  const [pin, setPin] = useState("");
  const [roles, setRoles] = useState(["mesero"]);
  const [error, setError] = useState("");
  const [verPin, setVerPin] = useState({});
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [verLog, setVerLog] = useState(false);

  const registrar = (accion) => ({ ts: Date.now(), accion, actor: usuarioActual.nombre });

  const toggleRol = (r) => setRoles((prev) => (prev.includes(r) ? prev.filter((x) => x !== r) : [...prev, r]));

  const agregar = async () => {
    if (!nombre.trim()) { setError("Escribe el nombre."); return; }
    if (pin.length < 4) { setError("El PIN debe tener al menos 4 dígitos."); return; }
    if (roles.length === 0) { setError("Elige al menos un rol."); return; }
    if (usuarios.some((u) => u.nombre.toLowerCase() === nombre.trim().toLowerCase())) { setError("Ya existe un usuario con ese nombre."); return; }
    const nuevo = { id: uid(), nombre: nombre.trim(), pin, roles };
    await persistConfig({
      ...config,
      usuarios: [...usuarios, nuevo],
      auditLog: [registrar(`Creó el usuario "${nuevo.nombre}" (${rolesLabel(nuevo)})`), ...auditLog].slice(0, 50),
    });
    setNombre(""); setPin(""); setRoles(["mesero"]); setError("");
  };

  const eliminar = async (u) => {
    const admins = usuarios.filter((x) => tieneRol(x, "admin"));
    if (tieneRol(u, "admin") && admins.length <= 1) { setError("No puedes eliminar al único administrador."); setConfirmDeleteId(null); return; }
    if (u.nombre === usuarioActual.nombre) { setError("No puedes eliminar tu propio usuario mientras estás dentro."); setConfirmDeleteId(null); return; }
    await persistConfig({
      ...config,
      usuarios: usuarios.filter((x) => x.id !== u.id),
      auditLog: [registrar(`Eliminó el usuario "${u.nombre}"`), ...auditLog].slice(0, 50),
    });
    setConfirmDeleteId(null); setError("");
  };

  const cambiarRol = async (u, rol) => {
    const actuales = rolesDe(u);
    const nuevos = actuales.includes(rol) ? actuales.filter((r) => r !== rol) : [...actuales, rol];
    if (nuevos.length === 0) { setError("Cada usuario necesita al menos un rol."); return; }
    if (actuales.includes("admin") && !nuevos.includes("admin") && usuarios.filter((x) => tieneRol(x, "admin")).length <= 1) {
      setError("Tiene que quedar al menos un administrador."); return;
    }
    const lista = usuarios.map((x) => (x.id === u.id ? { ...x, roles: nuevos, rol: undefined } : x));
    await persistConfig({ ...config, usuarios: lista, auditLog: [registrar(`Cambió los roles de "${u.nombre}" a ${nuevos.map(rolLabel).join(" + ")}`), ...auditLog].slice(0, 50) });
    setError("");
  };

  return (
    <div>
      <div style={styles.addItemCard}>
        <input style={styles.editInput} placeholder="Nombre del nuevo usuario" value={nombre} onChange={(e) => setNombre(e.target.value)} />
        <input style={styles.editInput} type="password" inputMode="numeric" placeholder="PIN (mínimo 4 dígitos)" value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} />
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {Object.keys(ROL_LABELS).map((r) => (
            <button key={r} style={{ ...styles.catPill, ...(roles.includes(r) ? styles.catPillActive : {}) }} onClick={() => toggleRol(r)}>{rolLabel(r)}</button>
          ))}
        </div>
        {error && <div style={styles.pinError}>{error}</div>}
        <button style={styles.addItemBtn} onClick={agregar}><Plus size={15} /> Crear usuario</button>
      </div>

      {usuarios.map((u) => (
        <div key={u.id} style={styles.turnoRowCol}>
          <div style={{ ...styles.turnoRow, border: "none", padding: 0 }}>
            <span style={styles.pagoRowEtiqueta}><UserCircle2 size={14} style={{ verticalAlign: -2 }} /> {u.nombre}</span>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <button style={styles.iconBtn} onClick={() => setVerPin((v) => ({ ...v, [u.id]: !v[u.id] }))} title="Ver PIN">
                {verPin[u.id] ? <EyeOff size={14} /> : <Eye size={14} />}
              </button>
              <span style={styles.cierreDate}>{verPin[u.id] ? u.pin : "••••"}</span>
              {confirmDeleteId === u.id ? (
                <>
                  <button style={styles.confirmYes} onClick={() => eliminar(u)}>Sí</button>
                  <button style={styles.confirmNo} onClick={() => setConfirmDeleteId(null)}>No</button>
                </>
              ) : (
                <button style={styles.deleteBtn} onClick={() => setConfirmDeleteId(u.id)}><Trash2 size={14} /></button>
              )}
            </div>
          </div>
          <div style={{ display: "flex", gap: 6, marginTop: 6, flexWrap: "wrap" }}>
            {Object.keys(ROL_LABELS).map((r) => (
              <button key={r} style={{ ...styles.catPill, padding: "4px 10px", fontSize: 11, ...(tieneRol(u, r) ? styles.catPillActive : {}) }} onClick={() => cambiarRol(u, r)}>{rolLabel(r)}</button>
            ))}
          </div>
        </div>
      ))}

      <button style={{ ...styles.canceladasToggle, marginTop: 14 }} onClick={() => setVerLog((v) => !v)}>
        <Lock size={13} /> <span>Registro de seguridad</span> <span>{verLog ? "▲" : "▼"}</span>
      </button>
      {verLog && (
        <div style={{ marginTop: 8 }}>
          {auditLog.length === 0 && <div style={styles.cajaEmptyText}>Sin movimientos todavía.</div>}
          {auditLog.map((a, i) => (
            <div key={i} style={styles.turnoRowCol}>
              <div style={styles.turnoSubRow}>{dateTimeLabel(a.ts)} · {a.actor}</div>
              <div style={{ fontSize: 13 }}>{a.accion}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ---------------- Campo de monto con separador de miles ---------------- */

function CampoMonto({ value, onChange, placeholder, style, autoFocus }) {
  const mostrado = value ? Number(value).toLocaleString("es-CO") : "";
  return (
    <input
      style={style || styles.editInput}
      inputMode="numeric"
      placeholder={placeholder}
      autoFocus={autoFocus}
      value={mostrado}
      onChange={(e) => onChange(e.target.value.replace(/\D/g, ""))}
    />
  );
}

function AvisoActivarNotificaciones() {
  const soportado = typeof window !== "undefined" && "Notification" in window;
  const [permiso, setPermiso] = useState(soportado ? Notification.permission : "denied");
  if (!soportado || permiso !== "default") return null;
  const activar = async () => {
    try { setPermiso(await Notification.requestPermission()); } catch (e) { setPermiso("denied"); }
  };
  return (
    <button style={styles.avisoNotifBtn} onClick={activar}>
      <Volume2 size={14} /> Toca aquí para activar los avisos del navegador
    </button>
  );
}

/* ---------------- Tarjeta de cuenta (cobro) ---------------- */

function MesaCuentaCard({ cuenta, orders, saveCuenta, saveOrdersMap, turnoAbierto, usuarioActual, mesas, expanded, onToggle }) {
  const esAdmin = tieneRol(usuarioActual, "admin");
  const [metodo, setMetodo] = useState("efectivo");
  const [modo, setModo] = useState("completa");
  const [seleccion, setSeleccion] = useState({});
  const [descInput, setDescInput] = useState("");
  const [confirmCancel, setConfirmCancel] = useState(false);
  const [procesando, setProcesando] = useState(false);

  const subtotal = cuentaSubtotal(orders, cuenta.id);
  const pct = (cuenta.propina && cuenta.propina.pct) || 0;
  const propinaMonto = Math.round((subtotal * pct) / 100);
  const descuento = (cuenta.descuento && cuenta.descuento.monto) || 0;
  const granTotal = Math.max(0, subtotal - descuento + propinaMonto);
  const pagado = pagosTotal(cuenta);
  const restante = Math.max(0, granTotal - pagado);
  const items = itemsAgregados(orders, cuenta.id);
  const asignada = qtyAsignadaPorProducto(cuenta);
  const hayPagos = pagosDe(cuenta).length > 0;
  const sinServir = cuentaOrders(orders, cuenta.id).filter((o) => o.estado !== "cancelado" && o.estado !== "servido").length;
  const visual = estadoVisualMesa(cuenta, cuentaOrders(orders, cuenta.id));

  const setPropina = async (p) => {
    await saveCuenta({ ...cuenta, propina: { pct: p, monto: Math.round((subtotal * p) / 100) } });
  };
  const aplicarDescuento = async () => {
    const monto = Math.min(parseFloat(descInput) || 0, subtotal);
    await saveCuenta({ ...cuenta, descuento: monto > 0 ? { monto, aplicadoPor: usuarioActual.nombre } : null });
    setDescInput("");
  };

  const registrarPago = async (monto, extra) => {
    if (procesando) return;
    setProcesando(true);
    try {
      const pago = { id: uid(), ts: Date.now(), monto, metodoPago: metodo, procesadoPor: usuarioActual.nombre, ...extra };
      const cerrada = pagado + monto >= granTotal;
      await saveCuenta({
        ...cuenta,
        propina: { pct, monto: propinaMonto },
        pagos: [...pagosDe(cuenta), pago],
        ...(cerrada ? { estado: "pagada", pagadaTs: Date.now() } : {}),
      });
      setSeleccion({});
    } finally {
      setProcesando(false);
    }
  };

  const cobrarCompleta = () => registrarPago(restante, { tipo: "completa", etiqueta: hayPagos ? "Saldo restante" : "Cuenta completa" });

  const montoSeleccion = items.reduce((s, it) => s + (seleccion[it.key] || 0) * it.price, 0);
  const cobrarProductos = () => {
    const elegidos = items.filter((it) => (seleccion[it.key] || 0) > 0).map((it) => ({ key: it.key, name: it.name, price: it.price, qty: seleccion[it.key] }));
    if (elegidos.length === 0) return;
    const todoAsignado = items.every((it) => (asignada[it.key] || 0) + (seleccion[it.key] || 0) >= it.qty);
    const monto = todoAsignado ? restante : Math.min(montoSeleccion, restante);
    registrarPago(monto, { tipo: "por_producto", etiqueta: "Por producto", items: elegidos });
  };
  const cambiarSel = (it, d) => {
    const libre = it.qty - (asignada[it.key] || 0);
    setSeleccion((prev) => ({ ...prev, [it.key]: Math.max(0, Math.min(libre, (prev[it.key] || 0) + d)) }));
  };

  const cancelarCuenta = async () => {
    const patch = {};
    cuentaOrders(orders, cuenta.id).forEach((o) => { patch[o.id] = { ...o, estado: "cancelado" }; });
    if (Object.keys(patch).length > 0) await saveOrdersMap(patch);
    await saveCuenta({ ...cuenta, estado: "cancelada", canceladaTs: Date.now(), canceladaPor: usuarioActual.nombre, montoCancelado: subtotal });
    setConfirmCancel(false);
  };

  const puedeCobrar = turnoAbierto && restante > 0 && !procesando;

  return (
    <div style={styles.cuentaCard}>
      <button style={styles.cuentaHead} onClick={onToggle}>
        <div style={{ textAlign: "left" }}>
          <div style={styles.cuentaMesa}>{mesaNombre(mesas, cuenta.mesa)}</div>
          <div style={styles.cuentaSub}>{cuenta.mesero || "—"} · desde {timeLabel(cuenta.ts)} · {cuentaItemCount(orders, cuenta.id)} ítems</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={styles.cuentaTotal}>{money(granTotal)}</div>
          <span style={{ ...styles.mesaCardTag, background: visual.bg, color: visual.color }}>{visual.tag}</span>
        </div>
      </button>

      {expanded && (
        <div style={styles.cuentaBody}>
          {items.length === 0 && <div style={styles.cajaEmptyText}>Esta cuenta no tiene productos.</div>}
          {items.map((it) => (
            <div key={it.key} style={styles.cuentaLinea}>
              <span>{it.qty}× {it.name}</span>
              <span>{money(it.qty * it.price)}</span>
            </div>
          ))}
          {sinServir > 0 && <div style={styles.warnBanner}><AlertTriangle size={14} /> {sinServir} {sinServir === 1 ? "pedido" : "pedidos"} aún sin servir.</div>}

          <div style={styles.cuentaLinea}><span>Subtotal</span><span>{money(subtotal)}</span></div>
          {descuento > 0 && <div style={styles.cuentaLinea}><span>Descuento ({cuenta.descuento.aplicadoPor})</span><span>-{money(descuento)}</span></div>}

          <div style={styles.catLabel}>PROPINA</div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {PROPINA_OPCIONES.map((p) => (
              <button key={p} disabled={hayPagos} style={{ ...styles.catPill, ...(pct === p ? styles.catPillActive : {}), opacity: hayPagos ? 0.5 : 1 }} onClick={() => setPropina(p)}>
                {p === 0 ? "Sin propina" : `${p}%`}
              </button>
            ))}
          </div>
          {propinaMonto > 0 && <div style={styles.cuentaLinea}><span>Propina</span><span>{money(propinaMonto)}</span></div>}

          {esAdmin && !hayPagos && (
            <>
              <div style={styles.catLabel}>DESCUENTO (SOLO ADMIN)</div>
              <div style={{ display: "flex", gap: 8 }}>
                <CampoMonto style={{ ...styles.editInput, flex: 1 }} placeholder="Monto a descontar" value={descInput} onChange={setDescInput} />
                <button style={styles.confirmYes} onClick={aplicarDescuento}><Gift size={13} /> Aplicar</button>
              </div>
            </>
          )}

          <div style={{ ...styles.cuentaLinea, fontWeight: 700, fontSize: 16 }}><span>Total a cobrar</span><span>{money(granTotal)}</span></div>

          {hayPagos && (
            <>
              <div style={styles.catLabel}>PAGOS REGISTRADOS</div>
              {pagosDe(cuenta).map((p) => (
                <div key={p.id} style={styles.cuentaLinea}>
                  <span style={styles.pagoRowEtiqueta}>{p.etiqueta} · {METODO_META[p.metodoPago]?.label} · {timeLabel(p.ts)}</span>
                  <span>{money(p.monto)}</span>
                </div>
              ))}
              <div style={styles.cuentaLinea}><span>Falta por cobrar</span><b>{money(restante)}</b></div>
            </>
          )}

          {!turnoAbierto && <div style={styles.warnBanner}><AlertTriangle size={14} /> Abre el turno para poder cobrar.</div>}

          {restante > 0 && (
            <>
              <div style={styles.catLabel}>MÉTODO DE PAGO</div>
              <div style={{ display: "flex", gap: 6 }}>
                {Object.entries(METODO_META).map(([key, meta]) => {
                  const Icon = meta.icon;
                  return (
                    <button key={key} style={{ ...styles.catPill, flex: 1, justifyContent: "center", ...(metodo === key ? styles.catPillActive : {}) }} onClick={() => setMetodo(key)}>
                      <Icon size={13} /> {meta.label}
                    </button>
                  );
                })}
              </div>

              <div style={{ display: "flex", gap: 6, marginTop: 4 }}>
                <button style={{ ...styles.vistaMesasBtn, ...(modo === "completa" ? styles.vistaMesasBtnActive : {}) }} onClick={() => setModo("completa")}>Cobrar todo</button>
                <button style={{ ...styles.vistaMesasBtn, ...(modo === "productos" ? styles.vistaMesasBtnActive : {}) }} onClick={() => setModo("productos")}>Dividir por producto</button>
              </div>

              {modo === "completa" ? (
                <button style={{ ...styles.sendBtn, opacity: puedeCobrar ? 1 : 0.4 }} disabled={!puedeCobrar} onClick={cobrarCompleta}>
                  <Check size={16} /> Cobrar {money(restante)}
                </button>
              ) : (
                <>
                  {items.map((it) => {
                    const libre = it.qty - (asignada[it.key] || 0);
                    if (libre <= 0) return null;
                    return (
                      <div key={it.key} style={styles.menuRow}>
                        <div>
                          <div style={styles.menuItemName}>{it.name}</div>
                          <div style={styles.menuItemPrice}>{money(it.price)} · quedan {libre}</div>
                        </div>
                        <div style={styles.stepper}>
                          <button style={styles.stepBtn} onClick={() => cambiarSel(it, -1)}><Minus size={14} /></button>
                          <span style={styles.stepVal}>{seleccion[it.key] || 0}</span>
                          <button style={{ ...styles.stepBtn, ...styles.stepBtnPlus }} onClick={() => cambiarSel(it, 1)}><Plus size={14} /></button>
                        </div>
                      </div>
                    );
                  })}
                  <button style={{ ...styles.sendBtn, opacity: puedeCobrar && montoSeleccion > 0 ? 1 : 0.4 }} disabled={!puedeCobrar || montoSeleccion === 0} onClick={cobrarProductos}>
                    <Check size={16} /> Cobrar selección ({money(montoSeleccion)})
                  </button>
                </>
              )}
            </>
          )}

          <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
            <button style={styles.cancelEditBtn} onClick={() => window.print()}>Imprimir cuenta</button>
            {!hayPagos && (
              confirmCancel ? (
                <>
                  <button style={styles.confirmYes} onClick={cancelarCuenta}>Sí, cancelar cuenta</button>
                  <button style={styles.confirmNo} onClick={() => setConfirmCancel(false)}>No</button>
                </>
              ) : (
                <button style={{ ...styles.cancelEditBtn, color: "#C1442D" }} onClick={() => setConfirmCancel(true)}><Ban size={13} /> Cancelar cuenta</button>
              )
            )}
          </div>

          <div className="solo-imprimir" style={{ fontFamily: "monospace", fontSize: 12 }}>
            <div style={{ textAlign: "center", fontWeight: 700 }}>{mesaNombre(mesas, cuenta.mesa)}</div>
            {items.map((it) => <div key={it.key}>{it.qty}× {it.name} — {money(it.qty * it.price)}</div>)}
            <div>Subtotal: {money(subtotal)}</div>
            {descuento > 0 && <div>Descuento: -{money(descuento)}</div>}
            {propinaMonto > 0 && <div>Propina: {money(propinaMonto)}</div>}
            <div style={{ fontWeight: 700 }}>TOTAL: {money(granTotal)}</div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------- Historial (admin) ---------------- */

function HistorialView({ cuentasParaHistorial, ordersParaHistorial, historialCompleto, cargandoHistorial, cargarHistorialCompleto, config, mesas, usuarioActual }) {
  const [abiertos, setAbiertos] = useState({});
  const toggle = (k) => setAbiertos((prev) => ({ ...prev, [k]: !prev[k] }));

  useEffect(() => { cargarHistorialCompleto(); }, []);

  const { days, months, years } = buildHistory(cuentasParaHistorial, ordersParaHistorial);
  const yearList = Object.values(years).sort((a, b) => b.key.localeCompare(a.key));

  return (
    <div style={styles.menuScroll}>
      {cargandoHistorial && <div style={styles.cajaEmptyText}>Cargando historial completo…</div>}
      {!historialCompleto && !cargandoHistorial && <div style={styles.cajaEmptyText}>Mostrando solo los últimos {VENTANA_DIAS} días.</div>}
      {yearList.length === 0 && <div style={styles.cajaEmptyText}>Todavía no hay cuentas pagadas.</div>}

      {yearList.length > 0 && (
        <button style={styles.modoChoiceBtn} onClick={() => exportarExcel(cuentasParaHistorial, ordersParaHistorial, config, mesas)}>
          <Download size={15} /> Exportar a Excel
        </button>
      )}

      {yearList.map((y) => (
        <div key={y.key} style={{ marginTop: 12 }}>
          <button style={styles.histHead} onClick={() => toggle(y.key)}>
            <b>{y.key}</b><span>{money(y.total)} {abiertos[y.key] ? "▲" : "▼"}</span>
          </button>
          {abiertos[y.key] && y.monthKeys.slice().sort().reverse().map((mk) => {
            const m = months[mk];
            return (
              <div key={mk} style={{ marginLeft: 10 }}>
                <button style={styles.histHead} onClick={() => toggle(mk)}>
                  <span>{monthLabel(m.ts)}</span><span>{money(m.total)} {abiertos[mk] ? "▲" : "▼"}</span>
                </button>
                {abiertos[mk] && m.dayKeys.slice().sort().reverse().map((dk) => {
                  const d = days[dk];
                  return (
                    <div key={dk} style={{ marginLeft: 10 }}>
                      <button style={styles.histHead} onClick={() => toggle(dk)}>
                        <span>{dayLabel(d.ts)}</span><span>{money(d.total)} {abiertos[dk] ? "▲" : "▼"}</span>
                      </button>
                      {abiertos[dk] && (
                        <div style={styles.histDetalle}>
                          <div style={styles.turnoSubRow}>
                            {d.cuentasCount} {d.cuentasCount === 1 ? "cuenta" : "cuentas"} ·{" "}
                            {Object.entries(METODO_META).map(([k, meta]) => `${meta.label} ${money(d.porMetodo[k] || 0)}`).join(" · ")}
                          </div>
                          {d.cuentas.sort((a, b) => a.pagadaTs - b.pagadaTs).map((c) => (
                            <div key={c.id} style={styles.cuentaLinea}>
                              <span>{timeLabel(c.pagadaTs)} · {mesaNombre(mesas, c.mesa)} · {c.mesero || "—"}</span>
                              <span>{money(c.total)}</span>
                            </div>
                          ))}
                          <div style={styles.catLabel}>PRODUCTOS DEL DÍA</div>
                          {Object.entries(d.productos).sort((a, b) => b[1].subtotal - a[1].subtotal).map(([name, p]) => (
                            <div key={name} style={styles.cuentaLinea}><span>{p.qty}× {name}</span><span>{money(p.subtotal)}</span></div>
                          ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      ))}
    </div>
  );
}

/* ---------------- Reportes ---------------- */

function ReportesView({ cuentas, orders, menu, usuarioActual }) {
  const [rango, setRango] = useState(1);
  const esAdmin = tieneRol(usuarioActual, "admin");

  const inicio = new Date();
  inicio.setHours(0, 0, 0, 0);
  inicio.setDate(inicio.getDate() - (rango - 1));
  const pagadas = cuentas.filter((c) => c.estado === "pagada" && c.pagadaTs >= inicio.getTime());

  const total = pagadas.reduce((s, c) => s + cuentaGranTotal(orders, c), 0);
  const ticketPromedio = pagadas.length > 0 ? total / pagadas.length : 0;
  const propinas = pagadas.reduce((s, c) => s + ((c.propina && c.propina.monto) || 0), 0);
  const descuentos = pagadas.reduce((s, c) => s + ((c.descuento && c.descuento.monto) || 0), 0);

  const productos = {};
  let costoTotal = 0;
  let ventaProductos = 0;
  pagadas.forEach((c) => {
    cuentaOrders(orders, c.id).filter((o) => o.estado !== "cancelado").forEach((o) =>
      o.items.forEach((it) => {
        if (!productos[it.name]) productos[it.name] = { qty: 0, subtotal: 0 };
        productos[it.name].qty += it.qty;
        productos[it.name].subtotal += it.qty * it.price;
        costoTotal += (it.cost || 0) * it.qty;
        ventaProductos += it.price * it.qty;
      })
    );
  });
  const top = Object.entries(productos).sort((a, b) => b[1].qty - a[1].qty).slice(0, 10);

  const grupos = {};
  pagadas.forEach((c) => {
    const k = rango === 1 ? hourLabel(new Date(c.pagadaTs).getHours()) : dayKey(c.pagadaTs).slice(5);
    grupos[k] = (grupos[k] || 0) + cuentaGranTotal(orders, c);
  });
  const dataGrafica = Object.entries(grupos).sort((a, b) => a[0].localeCompare(b[0])).map(([k, v]) => ({ k, ventas: v }));

  return (
    <div style={styles.menuScroll}>
      <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
        {[[1, "Hoy"], [7, "7 días"], [30, "30 días"]].map(([n, label]) => (
          <button key={n} style={{ ...styles.vistaMesasBtn, ...(rango === n ? styles.vistaMesasBtnActive : {}) }} onClick={() => setRango(n)}>{label}</button>
        ))}
      </div>

      <div style={styles.cajaTotalCard}>
        <span style={styles.cajaTotalLabel}>Ventas ({pagadas.length} {pagadas.length === 1 ? "cuenta" : "cuentas"})</span>
        <span style={styles.cajaTotalValue}>{money(total)}</span>
        <div style={styles.metodoBreakdownRow}>
          <div style={styles.metodoBreakdownItem}><span style={styles.metodoBreakdownLabel}>Ticket promedio</span><span style={styles.metodoBreakdownValue}>{money(ticketPromedio)}</span></div>
          <div style={styles.metodoBreakdownItem}><span style={styles.metodoBreakdownLabel}>Propinas</span><span style={styles.metodoBreakdownValue}>{money(propinas)}</span></div>
          <div style={styles.metodoBreakdownItem}><span style={styles.metodoBreakdownLabel}>Descuentos</span><span style={styles.metodoBreakdownValue}>{money(descuentos)}</span></div>
        </div>
      </div>

      {esAdmin && costoTotal > 0 && (
        <div style={styles.turnoBox}>
          <div style={styles.catLabel}>MARGEN ESTIMADO</div>
          <div style={styles.cuentaLinea}><span>Ventas de productos</span><span>{money(ventaProductos)}</span></div>
          <div style={styles.cuentaLinea}><span>Costo de producción</span><span>{money(costoTotal)}</span></div>
          <div style={{ ...styles.cuentaLinea, fontWeight: 700 }}>
            <span>Margen</span>
            <span>{money(ventaProductos - costoTotal)} ({ventaProductos > 0 ? Math.round(((ventaProductos - costoTotal) / ventaProductos) * 100) : 0}%)</span>
          </div>
        </div>
      )}

      {dataGrafica.length > 0 ? (
        <div style={{ width: "100%", height: 200, marginTop: 14 }}>
          <div style={styles.catLabel}>{rango === 1 ? "VENTAS POR HORA" : "VENTAS POR DÍA"}</div>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={dataGrafica}>
              <CartesianGrid strokeDasharray="3 3" stroke="#D9CCEE" />
              <XAxis dataKey="k" tick={{ fontSize: 10 }} />
              <YAxis tick={{ fontSize: 10 }} width={48} tickFormatter={(v) => `${Math.round(v / 1000)}k`} />
              <Tooltip formatter={(v) => money(v)} />
              <Bar dataKey="ventas" fill="#4A2C8F" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      ) : (
        <div style={styles.cajaEmptyText}>No hay ventas en este período.</div>
      )}

      {top.length > 0 && (
        <>
          <div style={{ ...styles.catLabel, marginTop: 40 }}>MÁS VENDIDOS</div>
          {top.map(([name, p]) => (
            <div key={name} style={styles.cuentaLinea}><span>{p.qty}× {name}</span><span>{money(p.subtotal)}</span></div>
          ))}
        </>
      )}
    </div>
  );
}

/* ---------------- Fuentes y estilos ---------------- */

const fontImports = `
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
* { box-sizing: border-box; }
body { margin: 0; background: #F3EEFB; }
button { font-family: inherit; }
.solo-imprimir { display: none; }
@media print {
  .no-imprimir { display: none !important; }
  .solo-imprimir { display: block !important; }
  body { background: #fff; }
}
`;

const C = { bg: "#F3EEFB", ink: "#1B1033", line: "#D9CCEE", mute: "#8B80A3", accent: "#4A2C8F", card: "#FFFFFF", soft: "#E6DDF5", dark: "#1B1033" };
const btnBase = { border: "none", borderRadius: 10, cursor: "pointer", fontFamily: "inherit", display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 6 };
const inputBase = { width: "100%", padding: "10px 12px", border: `1px solid ${C.line}`, borderRadius: 10, fontSize: 14, fontFamily: "inherit", background: "#fff", color: C.ink };
const smallBtn = { ...btnBase, padding: "6px 10px", fontSize: 12, fontWeight: 600 };
const eyebrow = { fontSize: 11, letterSpacing: 1.2, color: C.mute, fontWeight: 700 };

const styles = {
  app: { fontFamily: "'DM Sans', system-ui, sans-serif", background: C.bg, color: C.ink, minHeight: "100vh", display: "flex", flexDirection: "column" },
  loadingScreen: { minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: C.bg },
  loadingStamp: { fontFamily: "'DM Sans', sans-serif", fontWeight: 700, letterSpacing: 2, color: C.accent },
  errorGuardadoBanner: { display: "flex", alignItems: "center", gap: 8, background: "#FBE2DC", color: "#8A3A2B", padding: "8px 14px", fontSize: 13 },
  errorGuardadoCerrar: { ...btnBase, background: "transparent", color: "#8A3A2B", padding: 4 },
  avisoListoBanner: { display: "flex", alignItems: "center", gap: 8, background: "#2F6690", color: "#fff", padding: "10px 14px", fontSize: 14, fontWeight: 600 },
  avisoNotifBtn: { ...btnBase, width: "100%", background: "#FBEFD9", color: "#8A611A", padding: "8px 14px", fontSize: 13, borderRadius: 0 },

  topBar: { display: "flex", alignItems: "center", flexWrap: "wrap", gap: 8, padding: "10px 14px", background: C.card, borderBottom: `1px solid ${C.line}` },
  brand: { display: "flex", alignItems: "center", gap: 6, marginRight: "auto" },
  brandMark: { color: C.accent },
  brandText: { fontWeight: 700, letterSpacing: 1.5, fontSize: 14 },
  syncDot: { width: 8, height: 8, borderRadius: 4, display: "inline-block" },
  meseroChip: { ...smallBtn, background: C.soft, color: C.ink },
  roleSwitch: { display: "flex", gap: 4, width: "100%" },
  roleBtn: { ...btnBase, flex: 1, padding: "8px 6px", background: C.soft, color: C.ink, fontSize: 13, fontWeight: 600 },
  roleBtnActive: { background: C.accent, color: "#F3EEFB" },
  roleBtnActiveDark: { background: C.dark, color: "#F3EEFB" },

  screen: { flex: 1, display: "flex", flexDirection: "column", background: C.bg, minHeight: 0 },
  screenDark: { flex: 1, display: "flex", flexDirection: "column", background: C.dark, color: "#F3EEFB", transition: "background .3s" },
  screenFlash: { background: "#3A2468" },
  menuScroll: { flex: 1, overflowY: "auto", padding: "12px 16px 24px" },

  loginWrap: { padding: 20, display: "flex", flexDirection: "column", gap: 10, maxWidth: 420, margin: "0 auto", width: "100%" },
  pageEyebrow: eyebrow,
  pageTitle: { fontSize: 24, fontWeight: 700 },
  cajaEmptyText: { fontSize: 13, color: C.mute, lineHeight: 1.5, padding: "6px 0" },
  loginList: { display: "flex", flexDirection: "column", gap: 8 },
  loginBtn: { ...btnBase, justifyContent: "flex-start", padding: "12px 14px", background: C.card, border: `1px solid ${C.line}`, fontSize: 15, fontWeight: 600, color: C.ink },
  rolBadge: { marginLeft: "auto", fontSize: 11, color: C.mute, fontWeight: 600 },
  pinError: { color: "#C1442D", fontSize: 13, padding: "4px 0" },
  editInput: inputBase,
  addItemBtn: { ...btnBase, background: C.accent, color: "#F3EEFB", padding: "11px 14px", fontSize: 14, fontWeight: 600 },
  inlineAddLink: { ...btnBase, background: "transparent", color: C.accent, fontSize: 13, textDecoration: "underline", padding: 6 },
  inlineLink: { ...btnBase, background: "transparent", color: C.accent, fontSize: 12, textDecoration: "underline" },
  backBtn: { ...btnBase, background: "transparent", color: C.accent, fontSize: 14, fontWeight: 600, padding: 6 },

  subHeader: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 14px", background: C.card, borderBottom: `1px solid ${C.line}` },
  mesaTitleWrap: { display: "flex", flexDirection: "column", alignItems: "center" },
  mesaEyebrow: eyebrow,
  mesaTitle: { fontSize: 18, fontWeight: 700 },
  cuentaOpenSub: { fontSize: 11, color: C.mute },
  mesaActionBtn: { ...btnBase, width: 34, height: 34, background: C.soft, color: C.ink },
  mesaActionPanel: { padding: 12, background: C.card, borderBottom: `1px solid ${C.line}`, display: "flex", flexDirection: "column", gap: 8 },
  modoChoiceBtn: { ...btnBase, justifyContent: "flex-start", padding: "11px 14px", background: C.card, border: `1px solid ${C.line}`, color: C.ink, fontSize: 14, fontWeight: 600, width: "100%" },
  closeConfirmText: { fontSize: 13, color: "#4A3F66", lineHeight: 1.5 },
  mesaPickerRow: { display: "flex", flexWrap: "wrap", gap: 6 },
  mesaPickerBtn: { ...smallBtn, background: C.soft, color: C.ink },
  cancelEditBtn: { ...btnBase, padding: "10px 14px", background: C.soft, color: C.ink, fontSize: 13, fontWeight: 600 },

  ticketStripCol: { padding: "8px 14px", display: "flex", flexDirection: "column", gap: 6, background: C.card, borderBottom: `1px solid ${C.line}` },
  miniTicketRow: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, flexWrap: "wrap" },
  miniTicket: { display: "flex", alignItems: "center", gap: 6, fontSize: 12 },
  stateDot: { width: 8, height: 8, borderRadius: 4, display: "inline-block" },
  miniTicketTime: { color: C.mute, marginLeft: 4 },
  editedTag: { background: "#FBEFD9", color: "#8A611A", fontSize: 10, padding: "1px 6px", borderRadius: 6, marginLeft: 4 },
  miniTicketActions: { display: "flex", alignItems: "center", gap: 6 },
  confirmText: { fontSize: 12, color: "#C1442D" },
  confirmYes: { ...smallBtn, background: "#C1442D", color: "#fff" },
  confirmNo: { ...smallBtn, background: C.soft, color: C.ink },
  iconBtn: { ...btnBase, width: 28, height: 28, background: C.soft, color: C.ink },
  marcarServidoBtn: { ...smallBtn, background: "#5B7553", color: "#fff" },
  miniTicketTotal: { fontSize: 12, color: "#4A3F66", paddingTop: 4 },
  editingBanner: { display: "flex", alignItems: "center", gap: 8, background: "#FBEFD9", color: "#8A611A", padding: "8px 14px", fontSize: 13 },
  sentBanner: { display: "flex", alignItems: "center", gap: 8, background: "#E4EEE1", color: "#3F5A38", padding: "8px 14px", fontSize: 13 },
  repeatBtn: { ...btnBase, margin: "8px 14px 0", padding: "8px 12px", background: C.soft, color: C.ink, fontSize: 13, fontWeight: 600 },

  catPillsWrap: { padding: "10px 14px 4px", display: "flex", flexDirection: "column", gap: 8 },
  searchInput: inputBase,
  catPills: { display: "flex", gap: 6, overflowX: "auto", paddingBottom: 4 },
  catPill: { ...btnBase, padding: "6px 12px", background: C.soft, color: C.ink, fontSize: 13, fontWeight: 600, whiteSpace: "nowrap" },
  catPillActive: { background: C.accent, color: "#F3EEFB" },
  catPillBadge: { background: "#C1442D", color: "#fff", borderRadius: 8, fontSize: 10, padding: "1px 6px" },
  menuRow: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "10px 0", borderBottom: `1px solid ${C.line}` },
  menuItemName: { fontSize: 14, fontWeight: 600 },
  agotadoTag: { background: "#8B80A3", color: "#fff", fontSize: 10, padding: "1px 6px", borderRadius: 6, marginLeft: 6 },
  menuItemCat: { fontSize: 11, color: C.mute },
  menuItemPrice: { fontSize: 13, color: "#4A3F66" },
  itemNoteInput: { ...inputBase, padding: "6px 8px", fontSize: 12, marginTop: 4 },
  itemNoteBtn: { ...btnBase, background: "transparent", color: C.accent, fontSize: 11, padding: "4px 0", justifyContent: "flex-start" },
  stepper: { display: "flex", alignItems: "center", gap: 8 },
  stepBtn: { ...btnBase, width: 32, height: 32, background: C.soft, color: C.ink },
  stepBtnPlus: { background: C.accent, color: "#F3EEFB" },
  stepVal: { minWidth: 20, textAlign: "center", fontWeight: 700 },
  noResults: { textAlign: "center", color: C.mute, padding: 20, fontSize: 13 },
  orderBar: { padding: 12, background: C.card, borderTop: `1px solid ${C.line}`, display: "flex", flexDirection: "column", gap: 8 },
  noteInput: inputBase,
  sendBtn: { ...btnBase, padding: "12px 14px", background: C.accent, color: "#F3EEFB", fontSize: 15, fontWeight: 700 },

  mesaGridHeader: { display: "flex", alignItems: "flex-end", justifyContent: "space-between", padding: "14px 16px 6px", gap: 8 },
  editMenuBtn: { ...smallBtn, background: C.soft, color: C.ink },
  lockedNote: { display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: C.mute },
  vistaMesasSwitch: { display: "flex", gap: 6, padding: "6px 16px 10px" },
  vistaMesasBtn: { ...btnBase, flex: 1, padding: "8px 10px", background: C.soft, color: C.ink, fontSize: 13, fontWeight: 600 },
  vistaMesasBtnActive: { background: C.accent, color: "#F3EEFB" },
  mesaGrid: { display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 10, padding: "0 16px 16px" },
  mesaCard: { ...btnBase, flexDirection: "column", alignItems: "flex-start", padding: 14, background: C.card, border: `1px solid ${C.line}`, color: C.ink, gap: 6, minHeight: 90 },
  mesaCardNum: { fontSize: 17, fontWeight: 700 },
  mesaCardTag: { fontSize: 11, fontWeight: 700, padding: "3px 8px", borderRadius: 8, display: "inline-block" },
  mesaCardSince: { fontSize: 11, color: C.mute },

  addItemCard: { background: C.card, border: `1px solid ${C.line}`, borderRadius: 12, padding: 12, display: "flex", flexDirection: "column", gap: 8, marginBottom: 14 },
  catLabel: { ...eyebrow, margin: "12px 0 6px" },
  menuEditItemBlock: { borderBottom: `1px solid ${C.line}`, padding: "4px 0" },
  editRow: { display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderBottom: `1px solid ${C.line}` },
  agotadoToggle: { ...btnBase, width: 28, height: 28, background: C.soft, color: C.mute },
  agotadoToggleActive: { background: "#8B80A3", color: "#fff" },
  precioEditableBtn: { ...btnBase, marginLeft: "auto", background: "transparent", color: C.ink, fontSize: 13, fontWeight: 600, gap: 4 },
  deleteBtn: { ...btnBase, width: 28, height: 28, background: "transparent", color: "#C1442D" },
  costoEditRow: { display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap", paddingBottom: 6 },
  costoPillBtn: { ...btnBase, background: "transparent", color: C.mute, fontSize: 11, padding: "2px 0", justifyContent: "flex-start" },
  enCursoTag: { background: "#FBEFD9", color: "#8A611A", fontSize: 10, padding: "2px 6px", borderRadius: 6 },
  mesaReorderBtn: { ...btnBase, width: 22, height: 16, fontSize: 9, background: C.soft, color: C.ink, padding: 0 },

  planoContainer: { position: "relative", width: "100%", height: 320, background: C.card, border: `1px dashed ${C.line}`, borderRadius: 12, overflow: "hidden", marginBottom: 12 },
  planoVacio: { position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", color: C.mute, fontSize: 13 },
  planoMesaTile: { position: "absolute", transform: "translate(-50%, -50%)", padding: "10px 12px", borderRadius: 10, border: "2px solid", fontSize: 12, fontWeight: 700, userSelect: "none", fontFamily: "inherit" },

  kdsHeader: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 16px" },
  kdsEyebrow: { ...eyebrow, color: "#A598C8" },
  kdsTitle: { fontSize: 22, fontWeight: 700 },
  kdsCount: { fontSize: 32, fontWeight: 700, color: "#F3EEFB" },
  recienListosRow: { display: "flex", gap: 6, flexWrap: "wrap", padding: "0 16px 8px" },
  recienListoChip: { ...smallBtn, background: "#3A2468", color: "#E6DDF5" },
  emptyKitchen: { flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 8, padding: 40 },
  emptyKitchenText: { color: "#A598C8", fontSize: 14 },
  rail: { display: "flex", flexDirection: "column", gap: 12, padding: "0 16px 20px" },
  ticket: { background: "#2A1B4D", borderLeft: "6px solid", borderRadius: 10, padding: 12, display: "flex", flexDirection: "column", gap: 6 },
  ticketHead: { display: "flex", justifyContent: "space-between", alignItems: "center" },
  ticketMesa: { fontSize: 18, fontWeight: 700 },
  editedBadge: { background: "#B98A2E", color: "#fff", fontSize: 10, padding: "2px 6px", borderRadius: 6, marginLeft: 8 },
  ticketTime: { fontSize: 14 },
  ticketExactTime: { fontSize: 11, color: "#A598C8" },
  ticketDivider: { height: 1, background: "#4A3F66", margin: "4px 0" },
  ticketLine: { display: "flex", gap: 8, fontSize: 16 },
  ticketQty: { fontWeight: 700, color: "#E6B35A" },
  ticketItemName: { fontWeight: 600 },
  ticketItemNote: { fontSize: 13, color: "#E6B35A", marginLeft: 28 },
  ticketNote: { fontSize: 13, fontStyle: "italic", color: "#E6DDF5" },
  ticketBackBtn: { ...btnBase, padding: "0 14px", background: "#4A3F66", color: "#fff" },
  ticketBtn: { ...btnBase, padding: "12px 14px", color: "#fff", fontSize: 14, fontWeight: 700 },

  cajaTabSwitch: { display: "flex", gap: 4, padding: "6px 16px 10px", overflowX: "auto" },
  cajaTabBtn: { ...btnBase, flex: 1, padding: "8px 10px", background: C.soft, color: C.ink, fontSize: 13, fontWeight: 600 },
  cajaTabBtnActive: { background: C.accent, color: "#F3EEFB" },
  cajaTotalCard: { background: C.dark, color: "#F3EEFB", borderRadius: 14, padding: 16, display: "flex", flexDirection: "column", gap: 6, marginBottom: 14 },
  cajaTotalLabel: { fontSize: 12, color: "#A598C8" },
  cajaTotalValue: { fontSize: 30, fontWeight: 700 },
  cajaTotalValue2: { fontSize: 28, fontWeight: 700, display: "flex", flexDirection: "column", margin: "6px 0" },
  turnoEsperadoLabel: { fontSize: 12, fontWeight: 400, color: C.mute },
  metodoBreakdownRow: { display: "flex", flexDirection: "column", gap: 4, marginTop: 6 },
  metodoBreakdownItem: { display: "flex", alignItems: "center", gap: 6, fontSize: 13 },
  metodoBreakdownLabel: { color: "#A598C8", flex: 1 },
  metodoBreakdownValue: { fontWeight: 600 },
  warnBanner: { display: "flex", alignItems: "center", gap: 8, background: "#FBEFD9", color: "#8A611A", padding: "10px 12px", borderRadius: 10, fontSize: 13, margin: "8px 0" },
  goHistLink: { color: C.accent, fontSize: 13, textDecoration: "underline", cursor: "pointer", padding: "12px 0", textAlign: "center" },

  turnoBox: { background: C.card, border: `1px solid ${C.line}`, borderRadius: 12, padding: 14, display: "flex", flexDirection: "column", gap: 8, marginBottom: 14 },
  metodoResumenRow: { display: "flex", gap: 8 },
  metodoResumenItem: { flex: 1, display: "flex", flexDirection: "column", gap: 2, background: C.bg, borderRadius: 10, padding: 10 },
  metodoResumenLabel: { fontSize: 11, color: C.mute },
  metodoResumenValue: { fontSize: 15, fontWeight: 700 },
  closeCajaBtn: { ...btnBase, padding: "11px 14px", background: C.dark, color: "#F3EEFB", fontSize: 14, fontWeight: 600 },
  closeConfirmBox: { display: "flex", flexDirection: "column", gap: 8 },
  diferenciaBox: { padding: "10px 12px", borderRadius: 10, fontWeight: 700, fontSize: 14 },
  turnoRowCol: { padding: "10px 0", borderBottom: `1px solid ${C.line}` },
  turnoRow: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, padding: "8px 0", borderBottom: `1px solid ${C.line}` },
  cierreDate: { fontSize: 12, color: "#4A3F66" },
  turnoDiff: { fontSize: 13, fontWeight: 700 },
  turnoSubRow: { fontSize: 12, color: C.mute, marginTop: 2 },
  turnoObsRow: { fontSize: 12, fontStyle: "italic", color: "#4A3F66", marginTop: 2 },
  canceladasToggle: { ...btnBase, width: "100%", justifyContent: "space-between", padding: "10px 12px", background: C.soft, color: C.ink, fontSize: 13, fontWeight: 600 },
  pagoRowEtiqueta: { fontSize: 13, fontWeight: 600 },

  cuentaCard: { background: C.card, border: `1px solid ${C.line}`, borderRadius: 12, marginBottom: 10, overflow: "hidden" },
  cuentaHead: { ...btnBase, width: "100%", justifyContent: "space-between", padding: 14, background: "transparent", color: C.ink, borderRadius: 0 },
  cuentaMesa: { fontSize: 17, fontWeight: 700 },
  cuentaSub: { fontSize: 12, color: C.mute },
  cuentaTotal: { fontSize: 18, fontWeight: 700, marginBottom: 4 },
  cuentaBody: { padding: "0 14px 14px", display: "flex", flexDirection: "column", gap: 6 },
  cuentaLinea: { display: "flex", justifyContent: "space-between", gap: 8, fontSize: 14, padding: "3px 0" },

  histHead: { ...btnBase, width: "100%", justifyContent: "space-between", padding: "10px 12px", background: C.card, border: `1px solid ${C.line}`, color: C.ink, fontSize: 14, marginTop: 6 },
  histDetalle: { padding: "6px 10px 10px", background: C.card, borderRadius: 8, marginTop: 4 },
};
