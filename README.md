# KNO · Avance de Inventarios

App estática (sin build) para el control de auditoría de almacén por sucursal.

## Páginas

- `index.html` — Home con el "deck" de cargas guardadas. Cada tarjeta muestra el avance de esa carga; el botón **+ Nueva carga** crea una carga vacía y navega a su detalle.
- `carga.html?id=<uuid>` — Vista de una carga: subir Excel (columna requerida `areas del almacen`, opcional `cotizacion`) o empezar en blanco, editar el checklist por área, y ver el **Informe por Zona** (tarjetas + gráfico) en una pestaña separada del detalle.

## Backend (Supabase)

Proyecto dedicado `KNO-Avance-Inventarios` (no compartido con otros proyectos). Tablas:

- `cargas` — una fila por sucursal/auditoría guardada.
- `areas_almacen` — una fila por área del almacén dentro de una carga (`carga_id` FK, cascada al borrar la carga).
- `vista_resumen_cargas` — vista usada por el Home para mostrar el avance de cada carga sin hacer una consulta por tarjeta.

El acceso es público (sin login), igual que la versión anterior basada en `localStorage`: cualquiera con la URL puede leer y editar. Si se necesita restringir el acceso más adelante, el camino es agregar Supabase Auth y ajustar las políticas RLS de ambas tablas.

Las credenciales del cliente (`js/supabase-client.js`) usan la *publishable key* pública del proyecto — es segura de exponer en el navegador, el control de acceso real vive en las políticas RLS.

## Desarrollo local

No requiere instalación. Basta con servir la carpeta con cualquier servidor estático, por ejemplo:

```bash
npx serve .
```

## Despliegue

Cada push a `main` publica automáticamente el sitio en GitHub Pages vía `.github/workflows/pages.yml`.
