# SWCorp — Guía de estilo para aplicativos web

Spec compartido entre **ScanShip** (`dev.swcorp.com/scanship/`) y **Freight Bill Processor**
(`dev.swcorp.com/freight-bill-processor/`). Usar como base de cualquier app nueva.

Origen: `C:\Publish\ScanShip\templates\index.html`, `log.html` y
`C:\Publish\SWCORP\freight-bill\templates\index.html`.
Última estandarización: 2026-08-14.

---

## 1. Stack y estructura

- Flask + Jinja, un solo `templates/index.html` por app.
- CSS **inline en `<style>`** dentro del `<head>` — sin archivos externos, sin frameworks.
- Rutas de assets siempre vía `{{ base_url }}` (las apps corren en subdirectorio de IIS,
  nunca en la raíz del dominio).
- SVGs inline o en `templates/` / `Icons/`.

### Head obligatorio

```html
<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
```

`viewport-fit=cover` + safe-area insets = necesario para iPhone con notch.

---

## 2. Tokens

```css
*, *::before, *::after { margin: 0; padding: 0; box-sizing: border-box; }

:root {
  --navy:      #1C3D5A;   /* marca — header, botones primarios, acentos */
  --navy-dark: #152E45;   /* estado :active de botones navy */
  --navy-light:#2A5278;   /* opcional */
  --white:     #FFFFFF;
  --gray-50:   #F8F9FA;   /* fondo de fila seleccionada */
  --gray-100:  #F1F3F5;
  --gray-200:  #DEE2E6;   /* bordes de tarjeta/fila */
  --gray-300:  #CED4DA;   /* bordes de input, iconos vacíos */
  --gray-400:  #ADB5BD;   /* placeholder, texto secundario tenue */
  --gray-500:  #868E96;   /* subtítulos */
  --gray-700:  #495057;   /* labels */
  --gray-900:  #212529;   /* texto principal */
  --green:     #2E9E5E;   /* éxito */
  --red:       #DC3545;   /* error */

  --sat: env(safe-area-inset-top,    0px);
  --sab: env(safe-area-inset-bottom, 0px);
  --sal: env(safe-area-inset-left,   0px);
  --sar: env(safe-area-inset-right,  0px);
}
```

**Tipografía:** Inter (Google Fonts, pesos `400;500;600;700;800`), fallback
`-apple-system, BlinkMacSystemFont, sans-serif`.

```css
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
```

---

## 3. Shell — viewport bloqueado, scroll solo interno

Regla clave: **la página nunca hace scroll**. Solo hace scroll la lista interna
(`overflow-y: auto` en un contenedor con `flex:1; min-height:0`).

```css
html { height: 100dvh; overflow: hidden; }

body {
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
  color: var(--gray-900);
  background: var(--white);
  height: 100dvh;
  overflow: hidden;
  display: flex;
  flex-direction: column;
  padding-top:    var(--sat);
  padding-bottom: var(--sab);
  padding-left:   var(--sal);
  padding-right:  var(--sar);
  -webkit-tap-highlight-color: transparent;
  -webkit-text-size-adjust: 100%;
}

.app-frame {          /* móvil: edge-to-edge */
  flex: 1;
  min-height: 0;
  width: 100%;
  background: var(--white);
  position: relative;
  display: flex;
  flex-direction: column;
}

.app-body {           /* zona bajo el header */
  flex: 1;
  min-height: 0;
  overflow: hidden;
  display: flex;
  flex-direction: column;
  background: var(--white);
}
```

Jerarquía: `body > .app-frame > .header + .app-body > [pantallas]`.

---

## 4. Header

```css
.header {
  background: var(--navy);
  padding: 12px 16px;
  display: flex;
  align-items: center;
  min-height: 56px;
  height: 56px;
  flex-shrink: 0;
}
.header-logo-img { height: 30px; width: auto; filter: brightness(0) invert(1); }
```

- Altura fija 56px, siempre navy.
- Logo `SWCorp_logo_Main.svg` con `filter: brightness(0) invert(1)` para volverlo blanco.

### 4.1 Iconos del header — caja única obligatoria

Todos los iconos del header (SharePoint, tabla/log, cubo, back…) usan **la misma caja**:

```css
.header-icon {                 /* en freight-bill se llama .sp-icon-link */
  display: flex; align-items: center; justify-content: center;
  width: 36px; height: 36px;   /* caja tocable */
  border-radius: 8px;
  flex-shrink: 0;
  text-decoration: none;
  cursor: pointer;
  -webkit-tap-highlight-color: transparent;
}
.header-icon:hover,
.header-icon:active { background: rgba(255,255,255,0.12); }

/* glifo outline (viewBox 0 0 24 24) */
.header-icon svg {
  width: 26px; height: 26px;
  stroke: var(--white); fill: none;
  stroke-width: 1.5; stroke-linecap: round; stroke-linejoin: round;
}
/* glifo relleno (logos tipo SharePoint) — mismo 26px, se blanquea con filter */
.sp-icon-link svg { width: 26px; height: 26px; filter: brightness(0) invert(1); }
```

**Posición: siempre por flex, nunca `position: absolute`.** El icono absoluto con
`right: 14px` fue justo el bug — no seguía el `padding` del header, así que el inset
del icono no coincidía entre apps al cambiar de breakpoint. Con flex, el inset lo
define solo el padding del header (16 → 20 → 24px) y queda idéntico en las dos apps.

Layout del header:

```css
.header { display: flex; align-items: center; justify-content: space-between; }
```

```html
<div class="header">
  <div class="header-spacer" aria-hidden="true"></div>  <!-- si no hay icono izq -->
  <div class="header-logo">…logo…</div>
  <a class="header-icon" href="…">…svg…</a>
</div>
```

Si la app solo tiene un icono (derecha), poner a la izquierda un
**spacer de 36×36** para que el logo quede centrado óptimamente:

```css
.header-spacer { width: 36px; height: 36px; flex-shrink: 0; }
```

Si el orden en el DOM no puede cambiar, ordenar con flex `order`:
`.header-spacer{order:1}` · `.header-logo{order:2}` · `.header-icon{order:3}`.

Nota visual: un glifo **relleno** a 26px pesa más que uno **outline** a 26px.
Si se ve más grande al comparar lado a lado, bajar solo el relleno a 24px —
la caja de 36px nunca cambia.

---

## 5. VENTANA FLOTANTE — spec exacto (≥768px)

Esto es lo que hay que copiar textual para que todas las apps midan igual lado a lado:

```css
@media (min-width: 768px) {
  body {
    background: var(--white);   /* fondo blanco — decisión 2026-08-14 */
    padding: 0;                 /* fuera safe-area: la tarjeta va centrada */
    align-items: center;
    justify-content: center;
  }
  .app-frame {
    width: clamp(700px, 85vw, 1150px);
    height: 85dvh;
    max-height: 800px;
    flex: none;                 /* anula el flex:1 base */
    border-radius: 20px;
    overflow: hidden;
    box-shadow: 0 8px 40px rgba(0,0,0,0.18);
  }
  .app-body {
    border-bottom-left-radius: 20px;
    border-bottom-right-radius: 20px;
  }
  /* Overlays a pantalla completa deben quedar dentro de la tarjeta */
  .completed-screen { position: absolute; border-radius: 20px; }

  .header          { padding: 14px 24px; }
  .header-logo-img { height: 36px; }
  /* gutter interno estándar: 32px horizontal */
  .search-section  { padding: 28px 32px 20px; }
}
```

Notas:
- Fondo blanco + tarjeta blanca = solo la sombra delimita. Si queda muy plano,
  agregar `border: 1px solid var(--gray-200);` a `.app-frame` en este media query.
- Overlays fullscreen en móvil usan `position: fixed; inset: 0;` — en desktop hay que
  volverlos `position: absolute` o se salen de la tarjeta.

### Breakpoint intermedio (≥430px) — teléfonos grandes

```css
@media (min-width: 430px) {
  .header          { padding: 14px 20px; }
  .header-logo-img { height: 34px; }
  /* gutters 14px -> 18px, inputs 48px -> 52px */
}
```

Escala de gutter horizontal: **14px móvil · 18px ≥430px · 32px ≥768px**.
Escala de logo: **30px · 34px · 36px**.
Escala de input/botón alto: **48px · 52px · 52px**.

---

## 6. Componentes

### Input + botón fusionados (barra de búsqueda)

```css
.search-row { display: flex; border: 2px solid var(--gray-300); border-radius: 10px;
              overflow: hidden; background: var(--white); }
.search-row:focus-within { border-color: var(--navy); }
.search-input { flex: 1; min-width: 0; padding: 0 14px; height: 48px; border: none;
                font-size: 16px;   /* 16px obligatorio: evita zoom en iOS */
                font-family: inherit; font-weight: 500; outline: none; background: transparent; }
.search-input::placeholder { color: var(--gray-400); font-weight: 400; }
.btn-search { width: 52px; height: 48px; background: var(--navy); border: none;
              display: flex; align-items: center; justify-content: center; cursor: pointer; }
.btn-search:active { background: var(--navy-dark); }
```

En desktop la barra se limita: `.search-row { max-width: 580px; margin: 0 auto; }`.

### Tarjeta de contexto (navy)

`background: var(--navy); border-radius: 12px; padding: 12px 14px; margin-bottom: 12px;`
Texto blanco 15px/700. Input interno blanco, 44px alto, `border-radius: 8px`.

### Fila de lista

```css
.item-row { background: var(--white); border: 1.5px solid var(--gray-200);
            border-radius: 10px; padding: 11px 14px; margin-bottom: 6px;
            min-height: 52px; display: flex; align-items: center;
            justify-content: space-between;
            transition: border-color .25s, background .25s; }
.item-row.scanned { border-color: var(--navy); background: var(--gray-50); }
```

Título 14px/600 `--gray-900`, subtítulo 11px `--gray-400`, ambos con
`white-space: nowrap; overflow: hidden; text-overflow: ellipsis;`.
Check: círculo 28px, borde 2px `--gray-300`; activo se rellena navy y el SVG pasa de
`opacity: 0` a `1`.

### Label de sección

11px, `font-weight: 700`, `text-transform: uppercase`, `letter-spacing: 0.8px`,
color `--gray-700`.

### Botones

| Tipo | Estilo |
|---|---|
| Primario | `background: var(--navy)`, texto blanco, sin borde, `:active` navy-dark |
| Secundario | fondo blanco, `border: 1.5px solid var(--gray-200)`, texto `--gray-700`; `:active` borde+texto navy |
| Chip / acción compacta | alto 28px, `padding: 0 12px`, 11px/700 uppercase, radius 8px |
| Deshabilitado | `opacity: 0.5` |

Radius de botón: **8px**. Altura de acción en diálogo: **44px**.

### Popup de confirmación

Backdrop `rgba(0,0,0,0.45)`, z-index 600, padding 24px.
Card blanca, radius 12px, padding 20px, `width: min(88%, 340px)`,
`box-shadow: 0 8px 32px rgba(0,0,0,0.24)`. Título 16px/700 navy, texto 13px/500
`--gray-700` `line-height: 1.45`. Acciones en flex, gap 10px, `flex: 1` cada una.

### Toast

```css
.toast { position: fixed; top: calc(var(--sat) + 20px); left: 50%;
         transform: translateX(-50%) translateY(-20px);
         padding: 12px 22px; border-radius: 10px; color: var(--white);
         font-weight: 600; font-size: 14px; z-index: 999; opacity: 0;
         transition: all .25s ease; max-width: min(85%, 340px); text-align: center; }
.toast.show    { transform: translateX(-50%) translateY(0); opacity: 1; }
.toast-success { background: var(--green); }
.toast-error   { background: var(--red); }
```

### Pantalla de éxito fullscreen

Fondo navy sólido, logo blanco arriba (40px, `opacity: .9`), círculo 110px
`rgba(255,255,255,0.12)` con círculo interno blanco 80px y check navy 40px
(`stroke-width: 2.5`), título 26px/700 blanco, subtítulo 16px/600
`rgba(255,255,255,0.6)`. Padding `32px 24px max(32px, var(--sab))`.

### Estado vacío

Icono 56px `stroke: var(--gray-300)` `stroke-width: 1`, texto 13px `--gray-400`
`line-height: 1.6`, centrado. Contenedor `flex:1` con centrado para ocupar el resto.

### Skeleton

```css
.skel-bar { height: 12px; border-radius: 6px;
  background: linear-gradient(90deg, var(--gray-100) 25%, var(--gray-200) 50%, var(--gray-100) 75%);
  background-size: 200% 100%; animation: shimmer 1.2s infinite; }
@keyframes shimmer { to { background-position: -200% 0; } }
```

---

## 7. Escala general

| Cosa | Valor |
|---|---|
| Radius | 20px tarjeta flotante · 12px tarjeta/popup · 10px input/fila/toast · 8px botón/input interno · 50% círculo |
| Altura de control | 56px header · 52px input desktop · 48px input móvil · 44px acción de diálogo · 28px chip |
| Sombras | tarjeta flotante `0 8px 40px rgba(0,0,0,.18)` · popup `0 8px 32px rgba(0,0,0,.24)` |
| Transiciones | `.25s` estados · `.2s` hover |
| Tamaños de texto | 26 título éxito · 20/18 título de pantalla · 16 input · 15/14 principal · 13 cuerpo · 12/11 secundario y labels |

---

## 8. Reglas móviles no negociables

1. `font-size: 16px` en todo `input` — menos que eso hace zoom en iOS.
2. `-webkit-tap-highlight-color: transparent` en `body` y en cada elemento tocable.
3. Usar `:active`, no `:hover`, para feedback táctil (hover se queda pegado en touch).
4. Área tocable ≥ 44px (usar márgenes negativos si visualmente debe verse menor,
   ej. `.order-arrow { width:44px; height:44px; margin:-8px -8px -8px 0; }`).
5. `100dvh`, nunca `100vh` (la barra de URL de móvil rompe `vh`).
6. `-webkit-overflow-scrolling: touch` en la región con scroll.
7. Respetar safe-area con `var(--sat/--sab/--sal/--sar)` en móvil; quitarla en desktop.

---

## 9. Checklist para app nueva

- [ ] Copiar bloque tokens `:root` + `@import` de Inter tal cual.
- [ ] Copiar shell (`html`, `body`, `.app-frame`, `.app-body`) tal cual.
- [ ] Copiar el media query 768px de la sección 5 **sin cambiar números**.
- [ ] Header navy 56px con logo SWCorp invertido.
- [ ] Iconos del header en caja 36×36 / glifo 26px / radius 8px, posicionados por
      flex (spacer 36px si falta el icono izquierdo) — nunca `position: absolute`.
- [ ] Un solo contenedor con scroll; todo lo demás `flex-shrink: 0`.
- [ ] Overlays: `fixed` en móvil, `absolute` + radius 20px en desktop.
- [ ] Assets y links vía `{{ base_url }}`.
- [ ] Verificar lado a lado con ScanShip a 1920px y en iPhone.
- [ ] Tras deploy: reiniciar app pool de IIS (Jinja cachea templates).

---

## 10. Historial de estandarización (2026-08-14)

Valores que estaban desalineados entre las dos apps y a qué quedaron. Sirve de
referencia de qué revisar primero cuando dos apps "no se ven iguales".

### Ventana flotante

| Propiedad | ScanShip (antes) | Freight (antes) | Estándar |
|---|---|---|---|
| `.app-frame width` | `clamp(700px, 85vw, 1150px)` | `clamp(680px, 82vw, 1100px)` | `clamp(700px, 85vw, 1150px)` |
| `.app-frame max-height` | `800px` | `780px` | `800px` |
| `body background` (≥768px) | `var(--gray-100)` | `var(--white)` | `var(--white)` |
| gutter interno ≥768px | `28px 32px 20px` | `32px 36px 28px` | `28px 32px 20px` |
| logo base / 430 / 768 | `30 / 34 / 36px` | `28 / 32 / 34px` | `30 / 34 / 36px` |

### Iconos del header

| Propiedad | ScanShip (antes) | Freight (antes) | Estándar |
|---|---|---|---|
| Posición | flex `space-between` | `position: absolute; right: 14px` | flex `space-between` |
| Caja | `36×36`, radius `8px` | `32×32`, radius `6px` | `36×36`, radius `8px` |
| Glifo | `26px` | `22px` | `26px` |
| Opacidad base | `1` | `.7` | `1` |
| Hueco izquierdo | icono decorativo | ninguno | icono o `.header-spacer` 36×36 |

Causa raíz del desfase de iconos: el `position: absolute` con inset fijo `right: 14px`
no seguía el `padding` del header, así que al cambiar de breakpoint (16 → 20 → 24px)
el icono de una app quedaba en una `x` distinta que el de la otra. Con flex el inset
lo hereda del padding y coincide solo.

Aviso pendiente: el glifo de SharePoint es **relleno** (viewBox `0 0 50 50`) y los de
ScanShip son **outline** (viewBox `0 0 24 24`). A 26px el relleno pesa más a la vista.
Si molesta al comparar lado a lado, bajar solo ese SVG a 24px — la caja de 36px no se toca.
