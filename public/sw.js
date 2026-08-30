// Service worker mínimo — solo existe para que el navegador considere la app
// "instalable" (ícono en la pantalla de inicio, se abre en su propia
// ventana). A propósito NO cachea nada de Firebase ni datos de la app: el
// modo offline de verdad es una decisión aparte, todavía pendiente.

self.addEventListener("install", (event) => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(self.clients.claim());
});

// Sin "fetch" handler a propósito: todo pasa directo a la red, como si el
// service worker no existiera. Esto evita cualquier riesgo de mostrar datos
// viejos o desactualizados (pedidos, cuentas) mientras no haya modo offline.
