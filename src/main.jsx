import React from "react";
import ReactDOM from "react-dom/client";
import "./index.css";

const root = ReactDOM.createRoot(document.getElementById("root"));

// Si falta la configuración de Firebase (o algo falla al arrancar), en vez de
// dejar la pantalla en blanco se muestra aquí qué pasó.
const REQUERIDAS = [
  "VITE_FIREBASE_API_KEY",
  "VITE_FIREBASE_DATABASE_URL",
  "VITE_FIREBASE_PROJECT_ID",
  "VITE_FIREBASE_APP_ID",
];

function Aviso({ titulo, lineas, detalle }) {
  return (
    <div style={{ minHeight: "100dvh", background: "#171614", color: "#CFC9BB", fontFamily: "Inter, sans-serif", padding: 24, boxSizing: "border-box" }}>
      <div style={{ maxWidth: 460, margin: "0 auto" }}>
        <div style={{ color: "#D2A64E", fontSize: 12, letterSpacing: 1.5, fontWeight: 700 }}>NO SE PUDO INICIAR</div>
        <div style={{ fontSize: 22, fontWeight: 700, margin: "6px 0 12px" }}>{titulo}</div>
        {lineas && lineas.map((l) => (
          <div key={l} style={{ fontFamily: "monospace", fontSize: 13, color: "#CF7B6F", padding: "3px 0" }}>{l}</div>
        ))}
        {detalle && <div style={{ fontFamily: "monospace", fontSize: 12, color: "#918B7D", marginTop: 8 }}>{detalle}</div>}
        <div style={{ fontSize: 13, color: "#ADA798", marginTop: 16, lineHeight: 1.5 }}>
          Revisa que estas variables existan en Netlify (Site configuration → Environment variables), con el nombre exacto, y vuelve a desplegar el sitio.
        </div>
      </div>
    </div>
  );
}

const faltantes = REQUERIDAS.filter((n) => !import.meta.env[n]);

if (faltantes.length > 0) {
  root.render(<Aviso titulo="Falta la configuración de Firebase" lineas={faltantes} />);
} else {
  import("./App.jsx")
    .then((m) => {
      const App = m.default;
      root.render(<App />);
    })
    .catch((e) => {
      root.render(<Aviso titulo="La app tuvo un error al arrancar" detalle={String((e && e.message) || e)} />);
    });
}

// Registra el service worker para que la app se pueda "instalar" en el
// celular (ícono propio, se abre en su ventana). No cachea datos.
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("/sw.js").catch((e) => console.error("Error registrando el service worker:", e));
  });
}

