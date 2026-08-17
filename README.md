# Comandas — Dulce & Café

Este proyecto ya trae todo el sistema (mesero, cocina, caja) listo para subir
a internet con dominio propio, usando **Firebase** como base de datos y
**Netlify** como hospedaje.

## Parte 1 — Crear el proyecto de Firebase (una sola vez)

1. Entra a https://console.firebase.google.com y crea una cuenta gratis si no
   tienes (con tu cuenta de Google normal).
2. Clic en **"Crear un proyecto"**. Ponle el nombre que quieras (ej. "dulce-cafe").
   No hace falta activar Google Analytics — puedes dejarlo desactivado.
3. Una vez creado, en el menú de la izquierda entra a **"Compilación" → "Firestore Database"**.
4. Clic en **"Crear base de datos"**. Elige la ubicación más cercana (ej. `southamerica-east1`
   o la que te sugiera). Empieza en **modo de producción**.
5. Ve a la pestaña **"Reglas"** dentro de Firestore, borra lo que haya y pega esto:
   ```
   rules_version = '2';
   service cloud.firestore {
     match /databases/{database}/documents {
       match /data/{document=**} {
         allow read, write: if true;
       }
     }
   }
   ```
   Clic en **"Publicar"**.

   > Nota importante: esto deja la base de datos abierta a cualquiera que tenga
   > el link de tu sitio — es el mismo nivel de acceso que ya tenías con el
   > artifact de Claude. Es razonable para un negocio pequeño, pero si más
   > adelante quieres restringirlo con login, se puede hacer después.

6. Ahora ve a **"Configuración del proyecto"** (el ícono de engranaje arriba a
   la izquierda) → pestaña **"General"** → baja hasta "Tus apps" → clic en el
   ícono `</>` (Web) para registrar una app web.
7. Ponle un apodo (ej. "comandas-web") y clic en **"Registrar app"**. Firebase
   te muestra un bloque de código con 6 valores como estos:
   ```
   apiKey: "AIzaSy...",
   authDomain: "dulce-cafe.firebaseapp.com",
   projectId: "dulce-cafe",
   storageBucket: "dulce-cafe.appspot.com",
   messagingSenderId: "123456789",
   appId: "1:123456789:web:abc123"
   ```
   Guarda esos 6 valores — los necesitas ahora.

## Parte 2 — Configurar el proyecto en tu computador

1. Instala Node.js si no lo tienes: https://nodejs.org (versión LTS).
2. Abre una terminal dentro de esta carpeta y corre:
   ```
   npm install
   ```
3. Copia el archivo `.env.example` y renómbralo a `.env`.
4. Abre `.env` y pega los 6 valores de Firebase que guardaste, así:
   ```
   VITE_FIREBASE_API_KEY=AIzaSy...
   VITE_FIREBASE_AUTH_DOMAIN=dulce-cafe.firebaseapp.com
   VITE_FIREBASE_PROJECT_ID=dulce-cafe
   VITE_FIREBASE_STORAGE_BUCKET=dulce-cafe.appspot.com
   VITE_FIREBASE_MESSAGING_SENDER_ID=123456789
   VITE_FIREBASE_APP_ID=1:123456789:web:abc123
   ```
5. Para probarlo en tu computador antes de publicarlo:
   ```
   npm run dev
   ```
   Te da un link tipo `http://localhost:5173` — ábrelo en el navegador y
   prueba que todo funcione (crear un pedido, etc.) antes de publicar.

## Parte 3 — Publicar en Netlify

**Opción fácil (arrastrar y soltar, sin necesidad de GitHub):**

1. En la terminal, dentro de esta carpeta, corre:
   ```
   npm run build
   ```
   Esto crea una carpeta `dist/` con la versión final del sitio.
2. Entra a https://app.netlify.com y crea una cuenta gratis.
3. En el panel principal, busca la zona que dice **"Drag and drop your site
   output folder here"** (arrastra aquí la carpeta de tu sitio).
4. Arrastra la carpeta `dist` (la que se generó en el paso 1) a esa zona.
5. Netlify te da un link público al instante (algo como
   `https://nombre-random.netlify.app`). Ya tu sitio está en línea.
6. **Muy importante:** ese primer despliegue NO va a funcionar todavía porque
   le faltan las variables de Firebase (el archivo `.env` no se sube con
   `dist`, es solo para tu computador). Ve a:
   **Site settings → Environment variables → Add a variable**, y agrega ahí
   las mismas 6 variables `VITE_FIREBASE_...` con sus valores.
7. Luego ve a **Deploys → Trigger deploy → Clear cache and deploy site** para
   que tome las variables nuevas. Ahora sí debería funcionar completo.

**Opción recomendada a futuro (con GitHub):** si luego quieres que cada
cambio se publique solo, sube esta carpeta a un repositorio de GitHub y
conecta ese repositorio desde Netlify ("Import from Git") — así Netlify
reconstruye el sitio automáticamente cada vez que subas un cambio, sin tener
que repetir `npm run build` y arrastrar la carpeta a mano.

## Poner tu propio dominio (opcional)

Si tienes o compras un dominio (ej. `dulceycafe.com`), en Netlify:
**Site settings → Domain management → Add a custom domain**, y sigue las
instrucciones para apuntar tu dominio ahí. Netlify también te regala un
certificado de seguridad (https) automático.

## Notas

- Los datos (menú, pedidos, cuentas, usuarios) ahora viven en Firebase, no en
  Claude — puedes cerrar esta conversación y el sitio sigue funcionando para
  siempre.
- El PIN de Caja y de cada mesero, el menú, el historial: todo se mantiene
  igual que en la versión que ya conoces, solo cambió dónde se guarda.
- Si algún día quieres mover esto a otro hospedaje (Vercel, etc.), el mismo
  proyecto sirve — Netlify no es la única opción, solo la que armamos ahora.
