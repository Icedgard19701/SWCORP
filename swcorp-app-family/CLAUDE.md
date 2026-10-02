# CLAUDE.md — Familia de apps SWCorp

Guía para crear y estandarizar las apps web internas de SWCorp (IIS en `THE-MANNYS-SERV`,
embebidas en Acumatica). Esta carpeta es la **fuente de verdad visual** de la familia.

| Archivo | Qué es |
|---|---|
| `CLAUDE.md` | Reglas, razones, preferencias del usuario, registro de decisiones |
| `css/swcorp-tokens.css` | Fuentes + tokens (colores, forma, tipo, movimiento, sombras). Toda app lo carga primero |
| `css/swcorp.css` | Componentes de la familia (ventana, header, botones, pills, tablas…). Apps nuevas lo cargan segundo |
| `guide.html` | Referencia viva: cada componente renderizado solo con los dos CSS |
| `fonts/` | Geist + Geist Mono (OFL) |
| `assets/` | Logo blanco/navy, favicon |

La carpeta imita `static/` de una app (`css/` y `fonts/` hermanas), así los CSS se copian sin tocar rutas.

Versión: **v1 (2026-10-02)** — v0 se extrajo de Lowe's Invoice Reconciler; v1 unifica los nombres de clase
con los que usan las apps (`.app-frame`, `.header`, `.side-menu`, `.pop-out`, `.toast-stack`) y **las 5 apps cargan
`swcorp.css`**. Se actualiza con cada indicación del usuario.

## App de referencia

**Lowe's Invoice Reconciler** (`C:\Publish\SWCORP\Lowe's Invoice Reconciler`) es la más
trabajada: ante cualquier duda visual, su `WebApp/static/css/app.css` manda. Las guías viejas por app
(`SWCorp-Web-App-Style-Guide.md`, `style-reference*.html`, `STYLE_GUIDE.md` de ScanShip…) se borraron
el 2026-10-02: esta carpeta es la única referencia.

## Cómo usar en una app nueva o existente

1. Copiar `css/*` a `static/css/` y `fonts/*` a `static/fonts/` tal cual. Logo y favicon a `static/`.
2. Orden de carga: `swcorp-tokens.css` → `swcorp.css` → `app.css` (solo reglas propias de la app; nunca copias
   de reglas de `swcorp.css`: si una regla de familia no sirve, se cambia aquí y se propaga). React/Vite (Rebate):
   los tres se importan en `main.jsx` en ese orden y `#root { display: contents }`. **Prohibido** `:root` con colores propios en una app:
   si falta un color, se agrega a `swcorp-tokens.css` y se propaga. `:root` de la app solo para valores locales (safe-area).
3. **JS siempre en archivos externos** (`static/js/app.js`). Valores del servidor por un `<script>` mínimo
   antes del archivo: `window.<APP>_CONFIG = { base: {{ base_url|tojson }}, … }`. Nada de `{{ }}` dentro del `.js`.
4. **CSS siempre en archivos externos** (`static/css/`), nunca inline en `<style>` ni en el template
   (decisión 2026-10-01, buenas prácticas). Assets con `{{ base_url }}static/...?v={{ asset_v }}`;
   `asset_v` = mtime del CSS desde un `context_processor`. En `app.py`, fijar
   `mimetypes.add_type('text/css', '.css')` y `('text/javascript', '.js')`: bajo la identidad del
   app pool de IIS el registro puede devolver `text/plain` y el navegador descarta la hoja.
5. Esqueleto de página: el de `guide.html` (ver "Estructura").
6. Verificar contra `guide.html` lado a lado.
7. **Deploy en IIS:** templates/static se ven al instante; cambios de Python requieren reciclar el pool.
   Forma acordada: editar en `web.config` la línea `Last deploy: AAAA-MM-DD HH:MM:SS` (cambiar la hora).
   IIS reinicia la app al detectar el cambio. Toda app debe tener esa línea en su comentario de cabecera.
   Nunca `iisreset`, nunca reiniciar el servidor.
8. Al escribir archivos con Python en Windows usar `write_bytes` o `newline=""`: `write_text` con `\r\n`
   produce `\r\r\n` y duplica líneas en blanco.

## Estructura (obligatoria)

```html
<html>                       <!-- script: si window.self !== window.top -> class="embedded" -->
<body>
  <a class="skip" href="#content">Skip to content</a>
  <svg class="sprite">…symbols…</svg>
  <div class="app-frame">
    <header class="header">a.header-home(img.header-logo-img) · .header-sep · .header-title · (.run-chip) · .header-grow · button.menu-toggle</header>
    <div class="menu-scrim"></div>
    <nav class="side-menu" inert> .menu-body > .menu-glide + ul.menu-list > li.menu-anim > a.menu-item </nav>
    <div class="app-body">               <!-- o .content directo (LIR) -->
      <div id="content" class="content" tabindex="-1">
        <main class="main">  .titlebar(h2 + p) · contenido · .spacer </main>
        <footer class="foot"> .msg · .grow · .total · .btn.primary </footer>   <!-- opcional -->
      </div>
    </div>
    <div class="proc-screen">…</div>       <!-- overlays: absolute, dentro de .app-frame -->
  </div>
  <a class="pop-out" target="_blank" rel="noopener" aria-label="Open … in a new tab">↗</a>   <!-- solo visible embebida -->
  <div id="flash" class="toast-stack" aria-live="assertive"></div>
</body>
```

## Reglas y por qué

### Ventana flotante
- `.window` = tarjeta `min(1480px,100%)` × `min(920px, 100dvh − 2·gap)`, gap `clamp(20px,4vmin,48px)`,
  radius 20, sombra `0 8px 40px rgba(0,0,0,.18)`, fondo de página blanco.
  **Por qué:** la app antes vivía dentro de Acumatica; debe sentirse ventana sobre fondo neutro.
- `≤1180px` ancho o `≤720px` alto, o dentro de iframe (`html.embedded`): full-bleed, sin radius ni sombra.
  **Por qué:** en Acumatica el iframe mide ~1650×940 (1080p) o ~1100×630 (laptop); doble marco desperdicia espacio.
- La página nunca hace scroll. Solo `.main` (o la tabla interna). Overlays `position:absolute` dentro de `.window`.

### Responsive — todas las apps, 4 tiers (decisión 2026-10-01)
Toda app de la familia funciona en teléfono, compacto, escritorio y ventanas bajas. Reglas en
`swcorp.css` sección "responsive tiers". Dentro de un tier se escala con `clamp()`, no con más queries.

| Tier | Condición | Ventana | Gutter `.main` | Cambios |
|---|---|---|---|---|
| Teléfono | `< 600px` ancho | full-bleed + safe-area | 20 / 16 | header: logo 22, h1 14 con elipsis, chip oculto, acciones solo icono 36×36; título 18; inputs 16px (evita zoom iOS); botones 40, primario del footer 44 a todo el ancho; pills 2 por fila (impar final ocupa la fila), alto 56; toasts a todo el ancho |
| Compacto | `600–1180px` ancho | full-bleed | 24 / 24 | — |
| Escritorio | `> 1180 × 720` | tarjeta flotante | 28 / 32 | — |
| Bajo | `≤ 820px` alto | (según ancho) | 18 / 24 | footer 12/24 |
| Muy bajo | `≤ 720px` alto | full-bleed | 18 / 24 | zona de carga min 120, icono 36, gap 10, pills padding 10/12 |

Iframe de Acumatica (`html.embedded`) = full-bleed siempre. Usar `100dvh`, nunca `100vh`.
Meta viewport obligatorio: `width=device-width, initial-scale=1, viewport-fit=cover` (sin `viewport-fit` la safe-area vale 0).
Hover solo en `(hover:hover) and (pointer:fine)`.

### Zona de trabajo (dentro de la ventana, pantalla de carga)
Proporciones de LIR: título → zona de carga → pills separados 20px; padding 28/32 con el mismo espacio
bajo las pills que sobre el título. Zona de carga: borde 1.5px `--line`, radio 10, min-alto 150, padding
32/24/28, crece (`flex:1`) para llenar el alto libre; icono 48px, 14px hasta el prompt (14px `--ink`, link
navy 600), hint 12px `--ink-3` 6px debajo.
**Un solo icono para toda zona de carga** (bandeja con flecha hacia arriba, el de LIR/Rebate), en todas las
pantallas y apps: lo que cambia de una pantalla a otra es el texto, no el icono. Subtítulo 13px `--ink-3`, line-height 1.65, 4px bajo el título.
**Subtítulo = una sola línea, en inglés, minimalista**: qué hace la pantalla y con qué datos, sin relleno
("Validates carrier invoices against Acumatica and PaceJet, ready to import.").

### Header y menú lateral (decisión 2026-10-01)
- Navy 56px, padding 24px. **Logo a la izquierda** (blanco, 26px, `swcorp-logo-white.svg`) + divisor 1×24px
  + nombre de la app (15/600). Chip de contexto opcional (`.run-chip`).
- **El logo lleva a la pantalla principal de la app** (en Freight Bill: la vista Freight Bills). Es un `<a>` al
  inicio como respaldo, pero el JS navega sin recargar (no se pierden archivos cargados): desde otra vista vuelve con
  la transición normal; desde la pantalla final reinicia; ya en inicio o a mitad de una corrida, no hace nada (nunca
  interrumpe un proceso). Primera parada de Tab; `aria-label` con destino; hover opacidad .85, press .97.
- **Dentro de Acumatica el header no repite el nombre** (decisión 2026-10-01): la pantalla de Acumatica ya lo
  muestra y la página también; con el nombre del header serían tres. Un `<script>` en `<head>` (antes de pintar)
  marca `html.embedded` si `window.self !== window.top`, y el header oculta divisor y título: queda logo + menú.
  Abierta directo en el navegador, el header completo.
- **Burbuja "abrir en pestaña nueva" (solo dentro de Acumatica):** círculo navy de 32px abajo-derecha (14px del
  borde), icono ↗ 15px, casi invisible en reposo (opacidad .22, escala .92) y completo con hover o foco de teclado;
  en táctil reposa en .6. `<a target="_blank" rel="noopener">` a la raíz de la app con `aria-label` (sin tooltip).
  Abre una sesión nueva (no lleva los archivos ya cargados). Dentro de Acumatica los toasts suben para no taparla.
- **La esquina derecha es solo para el botón de menú** (`.menu-toggle`, 36×36, tres líneas que se doblan en X).
  Nada de iconos ni acciones sueltas en el header.
- **Vistas y enlaces van en el menú lateral** (`.side-menu`), dentro de la ventana, bajo el header, a la derecha:
  grupos con label 10px uppercase ("Views", "Links"); cada fila con icono en círculo 32px, título 13/600 y
  descripción 12px; la vista actual con `aria-current` (fondo navy-50, icono navy relleno); enlaces externos
  con flecha ↗. **Animación = motion.dev "variants" en CSS puro** (comparada cuadro por cuadro con el original
  en Playwright): en reposo, **solo tres líneas blancas, sin disco** (el círculo no se ve hasta presionar). El
  panel blanco ocupa todo el alto de la ventana, **también sobre el header**, recortado a un círculo de 0px
  **centrado en el botón, medido con JS** (`--cx/--cy`, nunca asumido: el padding cambia por tier); al abrir
  crece hasta llenarlo (640ms, curva drawer `cubic-bezier(.32,.72,0,1)`), las filas suben 24px escalonadas
  60ms tras 180ms, y las líneas de arriba y abajo **se doblan en una X** moviendo sus extremos (propiedad CSS
  `d`, 320ms) mientras la del medio se desvanece. Al cerrar: filas fuera de inmediato, X de vuelta a líneas, y
  el círculo regresa al disco (420ms, tras 140ms). El resto de la ventana se oscurece (scrim navy 18%).
  Esc, clic fuera o elegir cierra; cerrado es `inert`; salir con Tab cierra; el foco vuelve al botón.
  Las líneas pasan a navy al abrir y vuelven a blanco cuando el círculo ya se cerró. **La ventana debe verse
  siempre igual (todo adentro):** nada translúcido bajo el panel (sin sombra, el scrim termina donde empieza el
  panel), y como cada capa recortada por la curva deja pasar la de abajo en el borde suavizado, la esquina del
  header se redondea más **solo mientras el panel ya la cubre** (100 ms tras abrir, hasta 370 ms tras empezar a
  cerrar). Verificado midiendo píxeles en Playwright: abierto, la esquina queda blanca pura. Teléfono: ancho completo. Reduced motion: sin círculo, filas con fade.
- **Cierre rápido al elegir una opción:** el menú se cierra sin espera en 220ms ease-out (filas fuera en 80ms), a la
  par de la entrada de la pantalla nueva (ambos terminan ~190ms). El cierre con X/Esc conserva la animación completa.
  Motivo medido: con el cierre completo la pantalla llegaba a 235ms pero el menú tapaba un tercio de la ventana hasta
  564ms y todo se sentía lento. La esquina del header vuelve a los 90ms (el círculo aún la cubre): 0 cuadros expuestos.
- **Contenido del menú (decisión 2026-10-01):** cada opción lleva el **logo propio de su destino** (icono de la
  app/vista, SharePoint, Acumatica...) en una baldosa de 34px radio 9, pintado en navy con `mask` (un solo color;
  la vista actual lo invierte a blanco sobre navy). Las opciones van **centradas verticalmente** en el panel.
  **Reacción al mouse:** un único resaltado suave que se desliza a la opción bajo el cursor (240ms ease-out) y se
  desvanece al salir; la baldosa solo colorea su borde (sin moverse). **Una sola lista**, sin
  títulos de grupo ("Views", "Links"): vistas primero, enlaces externos después (con ↗). Solo con mouse (`hover:hover`), nada de inclinaciones 3D,
  imanes ni cursores custom (herramienta interna, criterio taste-skill: movimiento con propósito).
- `.hbtn` (acciones con texto en el header) eliminado de la familia (2026-10-02); LIR ya usa el menú lateral.
- **Ningún elemento del header lleva `view-transition-name`** ni z-index: un nombre lo vuelve contexto de
  apilamiento y el botón de menú queda bajo el panel abierto (pasó en LIR).

### Color
- Navy `#1C3D5A` único color de marca. Texto base `--ink` `#212529`; títulos en navy.
- Escalera de significado: info navy · pendiente ámbar `#C8A44D` · aviso `--warn` · ok verde · bloqueo rojo.
  **Rojo solo para fallo duro.** Un archivo equivocado o un dato faltante no es rojo.
- `--line-strong` en inputs/búsqueda (contraste 3:1, WCAG 1.4.11).

### Tipografía
- **Geist + Geist Mono en todas las apps** (decisión 2026-10-01), vendorizadas en `static/fonts/` (sin CDN,
  sin Inter ni Google Fonts). Body 14px/1.45.
- **Títulos en Title Case** (cada palabra con mayúscula inicial, el resto minúscula): "Freight Bill Processor",
  "Small Parcels", "Stop Processing?". Nunca títulos en MAYÚSCULAS sostenidas. Aplica a títulos de página,
  diálogos, toasts, pills y botones. Mayúsculas completas solo en labels de 10–11px.
- Bloque título + subtítulo (todas las pantallas, todas las apps):

  | Pieza | Tamaño / peso | Interlineado | Tracking | Color | Espacio |
  |---|---|---|---|---|---|
  | Título de página | 20px / 800 (teléfono 18px) | 1.45 | −.01em | navy | — |
  | Subtítulo | 13px / 400, una línea | 1.65 | — | `--ink-3` | 4px bajo el título |
  | Bloque → contenido | — | — | — | — | 20px (≤720px alto: 10–12px) |
  | Margen de pantalla | — | — | — | — | igual arriba y abajo: 20/16 teléfono · 24/24 compacto · 28/32 escritorio · 18/24 ventana baja |

  Tamaños fijos por tier, sin `clamp(…vw…)` en títulos (hacía 16px el título en teléfono).
  `text-wrap: balance` en el título, `pretty` en el subtítulo.
- Diálogo 16/800 · sección 14/700 · cuerpo 13–14 · secundario 12 · labels 10–11/700 uppercase. **Mínimo 10px.**
- 800 solo para títulos y números grandes. Mayúsculas solo 10–11px/700 con tracking.
- Cifras y IDs con `tabular-nums` (`.mono`, `.num`, `.acct`).

### Componentes (ver `guide.html`)
- Botón 36px, radius 8, 13/600; `primary` navy, `sec` borde 1.5px, `danger`, `small` 30px. Un primario por pantalla.
- Bordes de contenedores 1.5px `--line`; radius 10 caja / 8 control / 14 diálogo / 20 ventana.
- Tablas: header navy, títulos 10px uppercase centrados que envuelven (nunca "PAYMENT AMO…"), filas 38px,
  tinte por estado. Miles de filas: Tabulator con el mismo look, paginación/filtros en servidor.
- **Buscador = "tuck" de LIR** (`label.search.tuck`): en reposo solo el ícono al final de la fila; se abre
  (260px, crece hacia la izquierda) con hover o foco y queda abierto mientras tiene texto; × lo vacía y lo
  vuelve a cerrar aunque el mouse siga encima (`.shut` hasta `pointerleave`); Esc limpia. Filtra en vivo
  (debounce), sin botón "Search". En teléfono un tap en el ícono enfoca y abre. JS: `tuckSearch`/`syncSearch`.
- **Input pills** (`.slots` > `.slot`): una por archivo/fuente bajo la zona de carga. Alto 70px,
  padding 13/16, gap 12, aro 24px (borde 2px) con check trazado que se dibuja, origen 10px/700 uppercase
  sobre detalle 11px/500. Estados: reposo = pendiente (**sin color extra**: vacía basta para saber que falta) ·
  `.loading` (aro gira) · `.done` (navy, texto blanco/.72) · `.failed` (rojo suave + x).
  **× para quitar** (`.slot-x`) en todo archivo que el usuario cargó y que aún no se usó: visible en hover/focus
  (siempre en táctil); al quitar, el check se retrae por su trazo y luego se desvanece el navy. Sin × en
  archivos ya consumidos (p. ej. facturas WWEX/FedEx agregadas al workbook) ni en fuentes automáticas (Acumatica).
  El servidor borra la copia y su hash de duplicado, o el reemplazo se rechazaría. Cantidad fija: grid de 3; variable: `.slots.flex`.
  **El detalle debe ser link** a donde se obtiene el archivo (portal, SharePoint, pantalla Acumatica),
  sin subrayado en reposo, subrayado en hover, `target=_blank rel=noopener`. Texto plano solo si no hay a dónde enlazar.
- **Pills que entran o salen de la fila** (p. ej. 2 ↔ 4): nunca de golpe. Todos los cambios del mismo
  instante son un solo movimiento: se mide la fila antes y después y cada pill anima su ancho (420 ms,
  ease-in-out `cubic-bezier(.45,0,.2,1)`; una curva que arranca a toda velocidad se lee como salto). La que
  entra abre desde 0 (ancho, padding y opacidad, gap compensado con margen negativo); la que sale se cierra
  y se desvanece sin salir del flujo hasta terminar; las demás cambian de ancho. En teléfono (la fila se
  parte) entran con pop; con reduced motion, instantáneo. Verificado midiendo anchos por cuadro.
- Etiqueta pequeña de ID: `.tag` (antes `.pill`, renombrada para no chocar con las input pills).
- Avisos (`.notice`) y toasts: tarjeta blanca + borde izquierdo 3px con el significado.
- **Todo aviso flotante es toast abajo-derecha** (pila única fuera de la tarjeta, el más nuevo abajo):
  errores, resúmenes y avisos de acción pendiente. Nada de avisos arriba-centro ni barras metidas en el
  flujo de la pantalla (decisión 2026-10-01). **Todos duran 6s**, sin toasts fijos; una línea de 2px al pie
  (`.toast-timer`, color del borde) se vacía en esos 6s y su `animationend` es el temporizador, así hover/focus
  pausa ambos. Volver a mostrarlo reinicia. Acción pendiente ("Additional Files Required"): ámbar, además se
  va en cuanto se cumple; lo que falta lo siguen mostrando las pills. **Teléfono (< 600px): arriba**, bajo el
  header, a todo el ancho, entran desde arriba y salen hacia arriba (decisión 2026-10-01); escritorio y compacto
  siguen abajo-derecha.
  **Un solo diseño de toast**: tarjeta blanca, borde izquierdo 3px, icono 16px, título 13/700, texto 12px,
  × gris arriba-derecha, línea de tiempo al pie. Nada de toasts navy ni números gigantes: un resumen va como
  título con la cifra ("351 Pending Invoices") y una lista nombre ↔ cantidad (12px, cifra 600 tabular).
- Diálogo: `<dialog>` nativo, 440px, backdrop navy blur 4px, entrada con rebote, salida rápida.
- Procesando: `.proc-screen` navy dentro de la ventana, loader blanco que muta, texto que cruza.
- **Resumen con odómetro antes de la espera larga** (origen LIR Process): en cuanto están todos los archivos,
  pantalla blanca con título/subtítulo y tarjetas centradas (radio 16, borde 1.5px, label 12/600 `--ink-3`,
  cifra 24/600 tabular; la principal con borde y label navy). Cada dígito gira como odómetro (tira 0–9,
  1–3 vueltas de izquierda a derecha, 1.5 s, escalonado 110 ms por tarjeta y 45 ms por dígito). Abajo el
  nombre del siguiente paso sobre una barra de 140×3px que se llena durante el giro. Al terminar, **2 s de
  lectura solo si el trabajo de fondo ya acabó**; si no, pasa directo a la pantalla de carga (sin pausa delante
  de una espera). Avanza sola; clic o tecla detiene y aparece **Continue**. **El trabajo largo corre detrás** desde que aparece:
  al terminar el odómetro se muestra el resultado si ya está, si no la pantalla de carga hasta que esté.
  Con reduced motion: sin giro, 1 s + lectura.
- Estados obligatorios en cada pantalla: vacío, cargando (skeleton), error.

### Movimiento (revisado con criterio Emil Kowalski, 2026-10-01)
- **Solo con propósito**: retroalimentación, estado, evitar cambios bruscos. Si se ve decenas de veces al día,
  mínimo o nada; las celebraciones raras (pantalla final) pueden lucirse.
- **Curvas** (tokens): entrar/salir y presionar → `--ease-out` `cubic-bezier(.23,1,.32,1)`; mover/morphear en
  pantalla → `--ease-in-out` `cubic-bezier(.45,0,.2,1)`; color/hover → `ease`; progreso → `linear`.
  **Nunca ease-in** (ni `cubic-bezier(.4,0,1,1)`): arranca lento justo cuando el usuario mira.
  Sin rebote (`--ease-overshoot`) en UI frecuente; solo celebraciones.
- **Duraciones** (tokens): presión `--dur-press` 160ms · salida `--dur-exit` 180ms · entrada de pantalla
  `--dur-enter` 260ms · toast `--dur-toast` 320ms · morph `--dur-morph` 350ms. Salida siempre más rápida que entrada.
- **Presionar**: todo lo pulsable lleva `:active { transform: scale(.97) }` (× pequeñas .95), 160ms ease-out.
- **Nunca desde la nada**: entradas desde `scale(.95)` + opacidad (máximo .7 en celebraciones), nunca `scale(0)`.
- **Hover** solo bajo `@media (hover:hover) and (pointer:fine)` (en táctil se queda pegado). Propiedades
  explícitas en `transition`, nunca `all`; preferir `transform`/`opacity`.
- **Cambio de pantalla (decisión final 2026-10-01): la pantalla nueva entra de izquierda a derecha**, siempre en
  esa única dirección, deslizándose 28px y apareciendo (opacidad), 280ms `cubic-bezier(.23,1,.32,1)`, **solo con
  `transform` y `opacity`**. La vieja se va al instante; header, menú y toasts no se tocan. Hacia la pantalla navy de
  procesando: fundido 200ms. Reduced motion: cambio directo.
  **Descartados, medidos en Playwright (screencast cuadro a cuadro):** cortina de pantalla completa (tapaba ~490ms,
  cubría el header) y mask wipe con View Transitions (congelaba ~180ms al hacer clic y corría a 15–20fps por repintar
  la máscara en cada cuadro; en RDP sin GPU, peor). Regla: **una transición de pantalla nunca repinta áreas grandes
  por cuadro**; verificar con screencast que no haya cuadros > 34ms.
- **Cruce de textos** (mensajes que rotan): opacidad + `filter: blur(2px)` durante el cambio.
- **Listas que cambian** (pila de toasts, pills): los demás elementos se deslizan a su nuevo lugar (FLIP con
  `transform`, 240ms ease-out); entradas escalonadas 30–80ms entre elementos.
- **Temporizadores** (toasts) se pausan en hover/focus y con la pestaña oculta (`visibilitychange`).
- **Sin pulsos redundantes**: si hay spinner, no además un borde que late.
- **Reduced motion**: menos, no cero: lo que se desplaza pasa a fade de 150–200ms; opacidad y color se quedan;
  loaders mantienen su velocidad. La sesión RDP del usuario probablemente reporta `reduce`.

### Teclado (obligatorio en todas las apps, decisión 2026-10-01)
- **Todo se puede usar sin mouse.** Cada control llega con Tab en orden visual y se activa con Enter/Espacio.
  Elementos clicables que no son `<button>`/`<a>` (zonas de carga, tarjetas) llevan `role="button"`,
  `tabindex="0"`, `aria-label` y manejan Enter/Espacio (respetando el mismo estado deshabilitado/ocupado).
- Un control dentro de otro control no es una segunda parada: el interior lleva `tabindex="-1"`.
- `:focus-visible` siempre visible: navy 2px con offset 2px; blanco sobre navy (header, pantallas navy).
  Clic con mouse no muestra anillo.
- Controles que solo aparecen en hover (× de las pills) aparecen también con foco de teclado.
- Al cambiar de pantalla el foco pasa a la pantalla nueva (`tabindex="-1"`, sin anillo propio), nunca se
  queda en un control oculto.
- Menús y diálogos: abren con foco en el primer elemento útil (o el actual), Esc cierra, el foco vuelve al
  disparador; salir con Tab cierra el menú. Cerrados, `inert`.
- Verificar cada pantalla recorriéndola con Tab (Playwright: listar paradas y comprobar anillo).

### Sin tooltips (decisión 2026-10-01)
- **Prohibido `title=""` y tooltips de hover** (los "cuadritos negros" del navegador, `tooltip: true` de
  Tabulator, `<title>` en SVG de UI). La información que importa se escribe en pantalla; para lectores de
  pantalla, `aria-label` o texto `.vh`.
- **Única excepción, solo si el usuario lo pide:** encabezados de columna que expliquen de dónde sale o cómo
  se calcula un dato. Aun así, accesible por teclado (no solo hover).
- **Excepción 2 (decisión 2026-10-01, AmazonReturns):** celda recortada (texto largo con `…`) que no tiene
  otra forma de mostrar su valor completo. Se usa la **tarjeta propia de la app** (`td.tip[data-tip]`,
  una sola tarjeta para toda la página, por delegación), nunca `title`. Requisitos: la celda lleva
  `tabindex="0"`; la tarjeta aparece con hover **y con foco**, se va con blur, Esc, scroll o resize; el
  hint dice "Click to copy this row" (puntero) o "Press Enter to copy this row" (teclado), y Enter/Space
  copia la fila. El valor completo de las demás celdas, si hace falta para copiar, va en `data-full`.
- Encabezados ordenables y filas que abren detalle: `tabindex="0"`, Enter/Space = click; tras re-render
  el foco vuelve al encabezado nuevo (no cae al `body`).

### Accesibilidad
- Skip link cuando hay más de un bloque navegable antes del contenido. `aria-live` en toasts y estados.
- Áreas de clic ≥ 30px; iconos solos llevan `aria-label`.

### Embebido e IIS
- Subpath siempre (`{{ base_url }}`), nunca rutas absolutas.
- No `X-Frame-Options: DENY`; CSP con `frame-ancestors 'self' https://swcorp.acumatica.com`.
- Tras cambios de templates: reciclar solo el app pool propio. **Nunca** reiniciar el servidor ni `iisreset`.

### Pill de Acumatica y consultas OData (patrón obligatorio, origen LIR `recon/odata.py`)
1. **Ping = check.** Al abrir la página, un hilo pide **1 fila** del Generic Inquiry
   (`$top=1&$select=<un campo>`, ~1–2 s). Si responde, la pill pasa a check; si no, x roja + error.
   La pill indica *conexión*, no "datos descargados". Nunca bajar el GI completo para poner el check.
2. **Búsqueda filtrada, nunca el GI completo.** Las filas se piden por los valores de los archivos:
   `$filter=Campo eq 'v1' or Campo eq 'v2' …`, un campo por request, en lotes que dejan la URL
   **< 1900 caracteres** (el IIS de Acumatica da 404 sin explicación pasado ~2048), `$select` solo con
   las columnas que usa el cruce, comillas escapadas (`'` → `''`), varios lotes en paralelo (2–4 hilos).
   Acumatica compara sin distinguir mayúsculas.
3. **Se pide al cargar el archivo**, en segundo plano (prefetch), y el paso final (Process/Import) solo
   completa lo que falte. Cada sesión guarda los valores ya pedidos (en mayúsculas) y las filas: nunca se
   pide dos veces lo mismo; si una consulta falla, sus valores quedan sin pedir y el reintento los incluye.
4. **Mismo resultado que el GI completo.** El conjunto de valores replica la cascada de cruce (superconjunto:
   todo valor que cualquier rama podría buscar), y las filas tienen las mismas columnas, así la lógica de
   cruce no cambia. Al migrar una app, **probar lado a lado** completo vs filtrado sobre datos reales y
   exigir salida idéntica antes de publicar.
5. Antes de escribir algo irreversible (p. ej. agregar filas a un workbook compartido) se exige el ping OK.

Medido en Freight Bill (2026-10-01): GI completo 158 017 filas en 69.7 s → filtrado 337 filas en 5.7 s;
check de la pill en 1.7 s; factura de 351 registros idéntica celda por celda.

## Organización de proyectos (recomendación 2026-10-02, aún no aplicada)

Hoy cada app se organiza distinto (código en la raíz en Freight Bill y ScanShip, en `WebApp/` en LIR y Amazon
Returns, en `webapp/backend` + `webapp/frontend` en Rebate; tres formas de leer configuración). Objetivo común:

```
<app>/
  WebApp/            app.py (o main.py), templates/, static/{css,js,img,fonts}
  tests/             regresión con datos sintéticos (nunca cifras reales en repos públicos)
  deploy/            setup_iis.ps1, sync-family.ps1
  docs/              notas de la app (sin datos reales)
  data/              uploads/, output/, logs/   (ignorado por git)
  .venv/             (ignorado) — web.config apunta a este python, no a uno global
  requirements.txt · .env.example · web.config.example · README.md
```

- **Familia compartida:** esta carpeta como repo propio; cada app recibe copias de `css/`, `fonts/` y `assets/` con un
  script (`deploy/sync-family.ps1`) que también compara hashes, para que ninguna copia diverja (hoy dos favicons
  difieren). Sin submódulos ni carpeta compartida servida por IIS: cada repo queda autocontenido.
- **Logos y SVG:** la fuente única es `assets/` (blanco, navy, favicon). Los originales de diseño (logo principal y
  cuadrado) deben vivir solo en `assets/source/`; los iconos de cada menú, en el `static/img/` de su app.
- **Secretos:** fuera de los repos, en una carpeta del usuario del servidor con permisos restringidos, un `.env` por
  app; la app lo localiza por una variable de entorno puesta en `web.config`. Se versionan solo `.env.example` y
  `web.config.example`; `.env`, `*.secret`, `credentials.json` y `web.config` en el `.gitignore` de todas.
  Nunca secretos escritos en `config.py`.

## Preferencias del usuario (se va llenando)

- Responder en español, conciso.
- La guía se actualiza a la par de cada cambio que pida en las apps.
- Playwright solo si lo pide o es estrictamente necesario; él revisa visualmente.

## Estado de la familia

| App | Ruta | Estado vs. guía |
|---|---|---|
| Lowe's Invoice Reconciler | `C:\Publish\SWCORP\Lowe's Invoice Reconciler` | Referencia (v0 sale de aquí) |
| Freight Bill Processor | `C:\Publish\SWCORP\freight-bill` | En revisión (ver abajo) |
| Amazon Returns | `D:\Proyects\AmazonReturns` (dev.swcorp.com/AmazonReturns/) | En revisión (ver abajo) |
| ScanShip | `C:\Publish\ScanShip` (dev.swcorp.com/scanship/) | En revisión (ver abajo) |
| Lowe's Rebate Validation | `C:\Publish\Lowes Rebate Validation` (dev.swcorp.com/Lowes-Rebate-Validation/) | Estandarizada 2026-10-02 (ver abajo) |

### ScanShip — hecho (2026-10-01)
- Templates sueltos (index 1049 / log 875 líneas, logo SVG inline ×3) → `base.html` + `index.html` (Scanner)
  + `log.html` (Scan Log); `static/css/app.css`, `static/js/{app,scanner,log}.js`. Shell, header, menú, toasts,
  tabla, buscador tuck y pager copiados de Amazon Returns (mismo código). `asset_v` + mimetypes en `app.py`;
  `Last deploy` agregado a `web.config`. Aquí `base_url` NO termina en `/` (`{{ base_url }}/static/…`).
- Borrados: `/logo.svg` y `templates/SWCorp_logo_Main.svg`, `header-left/right-icon.svg`, logs viejos de IIS en la raíz.
- **Excepción de alineación (decisión 2026-10-01):** el Scanner es herramienta de celular, su columna va
  **centrada** (título, subtítulo, campo, orden; máx. 720px). Las demás vistas de la familia alinean a la izquierda.
- Scanner: campo de orden siempre visible (no tuck: el lector de códigos escribe donde está el foco);
  ancho máx. 720px en escritorio (los toasts nunca tapan un check); check = botón (teclado) con relleno
  navy + palomita que se dibuja; tras escanear la fila se actualiza en su lugar (no se redibuja la lista);
  pantalla "Order Completed" con pop + palomita; diálogo de confirmación de familia; × para cerrar la orden.
- Toasts: icono por tipo (✓ ok, ⓘ neutro, ⚠ error), títulos en Title Case.
- Pruebas: `/api/search` INSERTA filas Pending (no es solo lectura) y scan/unscan/complete escriben;
  en Playwright se simulan todas con `route`, solo `/api/log` va al servidor.

### Amazon Returns — hecho (2026-10-01)
- CSS/JS fuera de `base.html` (1950 → ~100 líneas): `static/css/app.css`, `static/js/{app,returns,monitoring}.js`;
  `asset_v` + mimetypes en `app.py`; `Last deploy` en `web.config`. Ojo: aquí `base_url` termina en `/`
  (`{{ base_url }}static/…`).
- Tokens de familia + Geist; títulos/subtítulos/espacios estándar; "Returns" en Title Case.
- Header de familia + menú lateral circular (Returns / Monitoring con logos), modo embebido + pop-out.
- Notice bars y toast de resumen eliminados: errores = toast rojo 6s de la pila abajo-derecha.
- Transición de pantalla = entrada izquierda→derecha, también al llegar desde el menú.
- Sin `title`; tarjeta de celda recortada accesible con teclado (Excepción 2); orden y filas por teclado,
  anillo `:focus-visible` global.
- CSS heredado de Freight sin uso borrado (upload zone, pills, popups, pantallas de proceso; 1402 → ~1090 líneas).
- Hover solo bajo `(hover: hover) and (pointer: fine)`; press `scale(.97)` en botones, `.98` en tarjetas.
- Teléfono: toasts arriba; tarjetas del pipeline en una fila cada una (nombre + punto + último import);
  buscador tuck como LIR (ícono que se abre con hover/foco).
- Entrada de tablas rota (se cortaba a ~40%): la 1ª página se pide con límite de sondeo (30), se mide y se
  recargaba en silencio a mitad de la animación. Ahora se recortan las filas sobrantes sin redibujar
  (`trimRows`). **Regla: nunca reemplazar una tabla mientras su entrada anima.**
- Buscador "tuck" de LIR en Returns y Monitoring (antes input + botón navy). Arreglado de paso: tras una
  búsqueda con pocos resultados la tabla se quedaba con ese número de filas al limpiar (`fitLimit`).

### Freight Bill — hecho (2026-10-01)
- Input pills con medidas/aspecto de LIR, links en el detalle conservados.
- CSS movido a `static/css/app.css` (672 líneas); `asset_v` y mimetypes en `app.py` (Python: vivo en IIS tras reciclar el pool).
- Zona de trabajo con proporciones de LIR + tiers bajo (≤820 / ≤720) + pills 2 por fila en teléfono.

- Fuente Geist vendorizada (`static/fonts/`), Inter eliminada.
- Títulos en Title Case: "Freight Bill Processor", "Small Parcels".

- JS (1073 líneas) movido a `static/js/app.js`; `window.FB_CONFIG` con `base` y `sid`. Template: 2172 → 381 líneas.
- Colores: carga `swcorp-tokens.css`; `--gray-*`/`--green`/`--red` y hex sueltos reemplazados por tokens
  de familia (`--line`, `--surface-2`, `--ink-3`, `--ok`, `--amber`…). Texto base `--ink`, títulos navy.
  Efecto visible: verde de éxito más oscuro (`#1E7A4A`), avisos azules `--navy-50`.
- `index.html` de la raíz (copia vieja, no servida, fuera de git) eliminado; respaldo en scratchpad de la sesión.
- `web.config` con línea `Last deploy`. Verificado en dev.swcorp.com: CSS/JS `text/*`, fuente `font/woff2`.

- Toasts: "Additional Files Required" (ambas vistas), error y resumen de carrier en una pila `.toast-stack`
  abajo-derecha. Antes: barra ámbar sobre las pills y toasts arriba-centro.
- (2026-10-02) Aviso suave del servidor ("Missing columns in … — those tabs will be skipped.") = toast ámbar
  "Some Tabs Skipped" (`#warn-toast`, `.notice-bar`) en la pila, 6s. Eliminado `#error-toast` del flujo y su CSS.

### Rebate Validation — hecho (2026-10-02, 8 fases, un commit por fase)
- Stack distinto: **React + Vite** (`webapp/frontend/src`, estilos en `src/styles/`), FastAPI sirve `frontend/dist`.
  Deploy = `npm run build` en `webapp/frontend` (sin reciclar IIS; solo cambios de Python lo requieren).
  CSS de familia importados en `main.jsx` (`swcorp-tokens.css` primero); fuentes en `src/fonts/` (Vite las empaqueta).
- Fase 1 hecha: tokens de familia + Geist (Inter/Google Fonts fuera), `:root` propio eliminado; colores `--match-*`
  (MatchType del inventory map) movidos a `swcorp-tokens.css` y propagados; texto `--ink`; títulos 20/800
  (18 teléfono), "Rebate Validation" en Title Case, subtítulos de una línea; mínimo 10px.
- 2 Header de familia + menú lateral en React (`components/AppFrame.jsx`, CSS portado de Amazon Returns):
  Rebate Validation / Inventory Map / SharePoint ↗; los iconos viejos del header pasaron a logos del menú
  (`assets/menu-*.svg`). `html.embedded` en `index.html` + burbuja pop-out. Arreglado: entre 768 y 1180px y en
  el iframe la ventana no llenaba el alto (`align-items:center` del body fuera del modo tarjeta).
- 3 **Una sola pila de toasts** (`components/Toast.jsx`): había tres (App, UploadScreen, `SkuCopiedToast`). API
  `toast({kind,title,detail})` / `notify(msg, kind)` por evento global. `kind`: error rojo (fallo duro) · warn ámbar
  (archivo equivocado, algo pendiente) · ok verde · info navy. Timer = animación CSS de `.toast-timer`
  (`onAnimationEnd`), pausa en hover/foco y `html.tab-hidden`. Click copia el texto (sin `title`).
- 4 Pills = `.slots.flex` (4 fijas), estado `.failed` para Acumatica (antes parecía "esperando"), × por teclado.
- 5 Movimiento: entrada izquierda→derecha 280ms, salida instantánea (antes salida ease-in .28s y luego entrada);
  sin rebote salvo la pantalla Complete (desde .7); escalonado de filas con tope de 8.
- 6 Sin `title=`: tarjeta propia `src/hoverTip.js` (solo si el texto está recortado; `data-tip-always` para
  información que la celda no muestra). El foco pasa a la pantalla nueva al cambiar de vista/pantalla.
- 7 Buscador tuck en Inventory Map; márgenes por tier; títulos de tabla envuelven; el diálogo devuelve el foco.
- 8 Limpieza: iconos y CSS sin uso.
- Pendiente: `web.config` usa `last-touched:` en vez de la línea `Last deploy:` de familia; cambiarla en el próximo
  deploy de Python (editarla recicla la app). Detalles de las pills sin link (no se sabe de dónde sale cada archivo).

### Unificación de swcorp.css (hecho 2026-10-02)
- `swcorp.css` v1 usa los nombres de las 4 apps hermanas; las 5 apps lo cargan y borraron sus copias de reset,
  ventana, header, menú y pop-out. Efecto visible: Amazon Returns, ScanShip y Freight Bill centraban el `body` desde
  768px, así que entre 768 y 1180px y **dentro de Acumatica** la ventana no llenaba el alto (o salía como tarjeta);
  ahora la tarjeta flotante solo existe fuera de iframe y > 1180×720 (`html:not(.embedded)`).
- `swcorp.css` suma: tarjeta de celda recortada (`.hover-tip`), toast único con tipos (`.toast.warn/.error/.ok`,
  `.toast-icon/-body/-title/-detail/-x/-timer`), pantalla de celebración (`.completed`, de ScanShip).
- Arreglado en las 3 implementaciones de la tarjeta (AR, Rebate, guía): al llegar con Tab el navegador hace scroll
  y el scroll la cerraba; ahora una celda con foco conserva su tarjeta.
- Deploy: **Flask sin `TEMPLATES_AUTO_RELOAD` (ScanShip) cachea templates**: un cambio de template exige reciclar
  (`Last deploy` en web.config). ScanShip estuvo ~5 min sin estilos de shell por eso.
- Pendiente: las tarjetas de toast de Freight Bill, Amazon Returns y ScanShip siguen con nombres propios
  (`.err-fixed-toast`, `.err-toast-*`, `.notice-bar`); migrarlas a `.toast` de familia (JS + CSS por app).

### Freight Bill — pendiente (revisado 2026-10-02)
- Toasts con nombres propios (ver arriba). Ya resueltos: header de familia, texto `--ink`, toasts abajo-derecha,
  tiers 600/1180, `html.embedded`, aviso en el flujo, carga de `swcorp.css`, favicon solo SWCorp.
- Ya resueltos (quitados de esta lista): header de familia, texto `--ink`, toasts abajo-derecha, tiers 600/1180, `html.embedded`, aviso en el flujo.

### LIR — hecho (2026-10-02)
- Tier teléfono (< 600px) = bloque "responsive tiers" de `swcorp.css` al final de `app.css`, más lo propio:
  inputs de `table.edit` y edición de Tabulator a 16px, `.form-grid.four` y diálogos a todo el ancho.
  `viewport-fit=cover` en `base.html` (sin él `env(safe-area-inset-*)` vale 0). Toasts arriba en teléfono:
  `toast-in/out` usan `--toast-dy`. Quitada la regla de 1 columna de pills a ≤900px (3 por fila hasta 600px).
  Verificado a 390px en las 6 vistas: sin scroll horizontal, sin inputs < 16px.

- (2026-10-02) Movimiento: toasts con línea de tiempo 6s (`.toast-timer`, su `animationend` = temporizador; pausa
  en hover/foco y `html.tab-hidden`), FLIP en `#flash` (MutationObserver de hijos; la pila está anclada abajo, así
  que solo se mueven los de arriba cuando sale uno de abajo), sin rebote en toasts ni tarjetas DM, sin ease-in.
  Los diálogos conservan su entrada con rebote (regla de Componentes: diálogo = entrada con rebote).

### LIR — hecho (2026-10-02, header)
- Header de familia + menú lateral (Remit / Setup; el logo va a la remit, nunca a la raíz, que crea un workspace
  nuevo) + pop-out. Carga `swcorp-tokens.css` y `swcorp.css`; su `:root` propio, fuentes, reset, ventana y header
  se borraron de `app.css` (`--muted-bg` → `--surface-3`). Toasts en `.toast-stack` (tarjeta propia aún).

## Registro de decisiones

| Fecha | Decisión | Origen |
|---|---|---|
| 2026-10-01 | v0 de la guía = sistema de LIR (Geist, ventana 1480×920, header logo-izq., toasts abajo-der.) | Usuario: LIR es la app más trabajada |
| 2026-10-01 | Input pills con medidas y aspecto de LIR `.slot`; detalle como link de preferencia. Aplicado a Freight Bill (clases `.pill` propias conservadas por el JS) | Usuario |
| 2026-10-01 | CSS separado del template en todas las apps | Usuario: buenas prácticas |
| 2026-10-01 | Soporte móvil/iPhone obligatorio para toda la familia (4 tiers) | Usuario |
| 2026-10-01 | Proporciones de la zona de trabajo = LIR | Usuario |
| 2026-10-01 | Geist para todas las apps | Usuario |
| 2026-10-01 | Títulos en Title Case, nunca MAYÚSCULAS | Usuario |
| 2026-10-01 | JS separado del template; colores solo desde tokens de familia (`swcorp-tokens.css`) | Usuario |
| 2026-10-01 | Reciclar IIS = cambiar hora de `Last deploy` en `web.config` | Usuario |
| 2026-10-01 | Borrar archivos que la app no usa | Usuario |
| 2026-10-01 | Avisos en el flujo y toasts arriba-centro pasan a la pila de toasts abajo-derecha | Usuario |
| 2026-10-01 | Todos los toasts 6s con línea de tiempo restante al pie | Usuario |
| 2026-10-01 | × para quitar archivos cargados + animación de des-check; pendientes sin navy | Usuario |
| 2026-10-01 | Resumen de pendientes por carrier = pantalla odómetro entre carga de archivos y loading (proceso ya corriendo detrás); toast del resumen eliminado | Usuario |
| 2026-10-01 | Apps navegables con teclado; sin tooltips `title` salvo pedido explícito (encabezados de columna). Aplicado a Freight Bill y LIR | Usuario |
| 2026-10-01 | Header: logo + divisor + nombre a la izquierda, solo botón de menú a la derecha; acciones al menú lateral circular | Usuario |
| 2026-10-01 | Revisión de movimiento (Emil Kowalski) aplicada a Freight Bill; reglas en Movimiento | Usuario |
| 2026-10-01 | Transición entre pantallas = entrada izquierda→derecha (transform+opacity), una sola dirección; cortina y mask wipe descartados por rendimiento | Usuario |
| 2026-10-01 | Pill de Acumatica = ping (check) + búsqueda filtrada por valores, nunca el GI completo (patrón LIR). Freight Bill migrado y verificado idéntico | Usuario |
| 2026-10-01 | Título/subtítulo/espacios estandarizados (tabla en Tipografía); mínimo 10px | Usuario |
| 2026-10-01 | Vistas tipo scanner de celular (ScanShip) van centradas; el resto alineado a la izquierda | Usuario |
| 2026-10-01 | Buscador = tuck de LIR (ícono, se abre con hover/foco) en todas las apps | Usuario |
| 2026-10-01 | Tarjeta propia para el valor completo de una celda recortada = excepción permitida, accesible con teclado (AmazonReturns, opción B) | Usuario |
| 2026-10-01 | Subtítulos: una línea, inglés, minimalista ("Validates carrier invoices against Acumatica and PaceJet, ready to import.") | Usuario |
| 2026-10-02 | Colores MatchType de Rebate (`--match-*`) pasan a `swcorp-tokens.css`: ninguna app define colores propios | Estandarización Rebate |
| 2026-10-02 | swcorp.css v1 con los nombres de las apps hermanas (`.app-frame/.header/.side-menu/.pop-out/.toast-stack`); las 5 apps lo cargan y borran sus copias | Usuario: "unifícalo" |
| 2026-10-02 | LIR: header de familia + menú lateral (Remit / Setup); `.hbtn` eliminado | Usuario |
| 2026-10-02 | Guías viejas por app borradas; esta carpeta es la única referencia | Usuario: "borra todo" |
