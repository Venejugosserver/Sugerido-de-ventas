# SUGERIDO DE VENTAS · Venejugos 2015

Analítico de ventas y sugerido de compra por marca, a partir del histórico consolidado, las ventas diarias de
MobilVendor y del Módulo B, y el inventario general. Misma familia visual que CORTE DE VENTAS.

| Archivo | Para qué sirve |
|---|---|
| `index.html` | **Analítico**: venta por día, semana, mes o trimestre, por módulo, marca, producto, presentación, cliente y asesor. |
| `sugerido.html` | **Sugerido por marca**: necesidades de compra, ajuste por producto, alertas, borrador y aprobación del pedido. |
| `admin.html` | **Panel administrativo** (llave `DIGIMARKET2026`): carga de archivos, agrupador, equivalencias, parámetros y publicación. |
| `nucleo.js` | Lectura de Excel (el histórico grande se lee por partes), agrupador y cálculos. |
| `ui.js`, `estilos.css`, `archivo.woff2`, `logo.jpg` | Interfaz compartida. |
| `xlsx.full.min.js`, `fflate.min.js`, `chart.umd.min.js` | Librerías incluidas: no requieren instalación. |
| `datos/ventas.json`, `datos/inventario.json`, `datos/config.json` | Datos publicados. El panel los actualiza. |
| `datos/pedidos.json` | Lo crea el sugerido al aprobar pedidos, si el navegador tiene el token de GitHub. |

## Publicar en GitHub Pages (una sola vez)

1. Crea un repositorio (por ejemplo `Venejugosserver/SUGERIDO-DE-VENTAS`) y sube todos los archivos y la carpeta `datos`
   (*Add file → Upload files*, arrastra todo junto).
2. *Settings → Pages → Build and deployment → Deploy from a branch → main / (root) → Save*.
3. Crea un token de alcance fino: foto de perfil → *Settings → Developer settings → Personal access tokens →
   Fine-grained tokens → Generate new token*. *Repository access*: **Only select repositories** (este repositorio).
   *Permissions → Repository permissions → Contents*: **Read and write**. Copia el token (`github_pat_…`); GitHub lo muestra una sola vez.
4. En `admin.html` → *Publicar*: escribe `usuario/repositorio`, rama `main`, deja vacía la carpeta del sitio si está en la raíz,
   pega el token y pulsa **Probar conexión** y luego **Publicar cambios**.

Con una cuenta gratuita, GitHub Pages solo publica repositorios públicos: los archivos de `datos/` quedan visibles para quien tenga el enlace.
La llave del panel evita el acceso casual, pero quien sepa leer el código puede verla.

## Uso diario

1. Panel → *Carga de archivos*: suelta las ventas de MobilVendor, las del Módulo B y el inventario (y el histórico cuando lo actualicen).
2. **Incorporar archivos**. El histórico reemplaza todos los días que cubre; cada archivo diario reemplaza solo sus días y su módulo.
3. Revisa *Agrupador → Sugerencias* y *Nuevos*.
4. *Publicar → Publicar cambios*. GitHub Pages tarda alrededor de un minuto.
5. En el sugerido, cada evaluador revisa su marca, ajusta cantidades y pulsa **Aprobar y enviar pedido**.

## Reglas de cálculo

- **Venta neta**: `total_venta_dolar` del sistema y `PrecioTotal` de MobilVendor (sin IVA). Notas de crédito y devoluciones restan.
- **Unidades**: con unidad suelta (UNI, UND., KG) se usa `venta_neta`; con empaque (CAJ, DIS, BUL…) `total_emp_aux`.
- **Producto agrupado**: mismo código en ambos módulos (en el Módulo B, la columna `alterno`), misma descripción, o unión manual.
- **Marca**: unidad de negocio del histórico; para productos nuevos, equivalencia de `DescrMarca` (MobilVendor) o `unidad_negocio` (Módulo B).
- **Sugerido** = venta por día hábil × (cobertura + seguridad) − stock, en unidad de compra. Venta por día = 50 % últimos 30 días + 50 % últimos 90.
- **Costo**: costo del inventario por unidad de compra; si no es coherente con el precio o falta, se estima con el margen típico.
- **Clientes excluidos**: INLUVENCA (configurable).

Las llaves del navegador usan el prefijo `vjs_`, así que conviven con CORTE DE VENTAS (`vjc_`) y GLOBAL KPI en el mismo dominio.
