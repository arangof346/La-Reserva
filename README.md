# Comandas — listo para Netlify

Este proyecto ya está armado para desplegarse en Netlify sin que tengas que
modificar nada de estructura. Solo falta un paso tuyo: conectar una base de
datos gratuita (Firebase) para que los datos se vean en vivo entre
dispositivos y no se pierdan al cerrar la página.

## Por qué hace falta Firebase

Netlify solo sirve archivos — no guarda información por sí solo. Para que el
Mesero, la Cocina y la Caja (en celulares/tablets distintos) vean los mismos
pedidos y cuentas **en tiempo real**, la app necesita una base de datos en la
nube. Usamos **Firebase Realtime Database**, que tiene un plan gratuito más
que suficiente para un negocio de este tamaño.

## Paso 1 — Crear el proyecto de Firebase (5 minutos, gratis)

1. Ve a https://console.firebase.google.com y crea un proyecto nuevo.
2. En el menú lateral: **Compilación → Realtime Database → Crear base de
   datos**. Elige la región más cercana a ti. Cuando pregunte por las reglas
   de seguridad, elige **"Modo de prueba"** para arrancar rápido.
3. Ve a **⚙️ Configuración del proyecto → General → Tus apps** y haz clic en
   el ícono **`</>`** (Web) para registrar una app web. No hace falta
   Firebase Hosting, solo necesitas el objeto de configuración que te
   muestra, algo así:

   ```js
   const firebaseConfig = {
     apiKey: "AIza...",
     authDomain: "tu-proyecto.firebaseapp.com",
     databaseURL: "https://tu-proyecto-default-rtdb.firebaseio.com",
     projectId: "tu-proyecto",
     storageBucket: "tu-proyecto.appspot.com",
     messagingSenderId: "123456789",
     appId: "1:123456789:web:abc123",
   };
   ```

4. Abre `src/firebase.js` en este proyecto y reemplaza el objeto
   `firebaseConfig` de ejemplo por el que copiaste.

Sobre seguridad: el archivo `database.rules.json` que incluí deja la base
abierta (cualquiera con la URL puede leer/escribir), lo cual es aceptable
para una herramienta interna pequeña — la URL no es pública en ningún lado.
Si más adelante quieres cerrarla más, en la consola de Firebase puedes pegar
reglas más estrictas en Realtime Database → Reglas.

## Paso 2 — Subir el proyecto a Netlify

**Opción A (recomendada): conectar un repositorio Git**

1. Sube esta carpeta completa a un repositorio de GitHub (o GitLab/Bitbucket).
2. En Netlify: **Add new site → Import an existing project**, elige el
   repositorio.
3. Netlify va a detectar automáticamente el archivo `netlify.toml` — el
   comando de build (`npm run build`) y la carpeta a publicar (`dist`) ya
   quedan configurados, no cambies nada ahí.
4. Deploy. Netlify instala dependencias, compila y publica solo.

**Opción B: sin Git, deploy manual**

1. En tu computador, con [Node.js](https://nodejs.org) instalado, entra a
   esta carpeta en la terminal y corre:
   ```
   npm install
   npm run build
   ```
2. Esto crea una carpeta `dist/`. Ve a https://app.netlify.com/drop y
   arrastra esa carpeta `dist` — Netlify la publica al instante.
3. Si luego cambias algo del código, repite `npm run build` y vuelve a
   arrastrar `dist`.

## Verificar que todo funciona

1. Abre el link de Netlify en dos pestañas o dos dispositivos distintos.
2. En una, entra como Mesero y envía un pedido.
3. En la otra, entra como Cocina — el pedido debe aparecer solo, sin
   recargar la página.
4. Cierra ambas pestañas y vuelve a abrir el link: todo lo que hiciste debe
   seguir ahí (nada se borra).

Si algo no aparece en vivo, lo más probable es que falte pegar el
`firebaseConfig` real en `src/firebase.js` — revisa la consola del navegador
(F12) por errores de Firebase.

## Estructura del proyecto

```
├── index.html          punto de entrada HTML
├── netlify.toml         configuración de build para Netlify
├── database.rules.json  reglas de seguridad sugeridas para Firebase
├── package.json          dependencias (React, Firebase, recharts, xlsx, lucide-react)
├── vite.config.js        configuración del bundler
└── src/
    ├── main.jsx          monta la app de React
    ├── App.jsx           toda la app (Mesero, Cocina, Caja)
    └── firebase.js        conexión a la base de datos en vivo — AQUÍ pegas tu config
```

No necesitas tocar nada más que `src/firebase.js`. El resto del código está
listo tal cual para producción.
