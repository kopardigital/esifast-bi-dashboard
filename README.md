# ESIFAST · Sistema BI Operativo

Sistema de Business Intelligence para análisis operativo y económico de logística y almacenaje. 100% estático, sin backend, publicable directamente en **GitHub Pages**.

![ESIFAST BI](https://img.shields.io/badge/ESIFAST-BI%20Sistema-2563eb?style=for-the-badge)
![GitHub Pages](https://img.shields.io/badge/GitHub_Pages-Ready-10b981?style=for-the-badge)
![No backend](https://img.shields.io/badge/Sin_backend-Client--side_only-f59e0b?style=for-the-badge)

---

## 🚀 Demo rápido

1. Abrí `index.html` en cualquier navegador moderno.
2. Ingresá con usuario `admin` y contraseña `Cambiar123!`.
3. Subí el archivo `ESIFAST.xlsx`.
4. Explorá el dashboard automático.

---

## ⚠ Advertencia de seguridad importante

> **Este sistema usa login del lado cliente (JavaScript puro).**
>
> Las credenciales están visibles en el código fuente (`app.js`). Cualquier persona con acceso al repositorio puede verlas. Esta barrera es únicamente **visual** — no representa seguridad real.
>
> **Antes de publicar:**
> - Cambiar las credenciales en `app.js` → variable `CONFIG`.
> - Nunca subir credenciales de producción a GitHub.
> - Si se requiere seguridad real: implementar autenticación con backend (Node.js, Python Flask, etc.) y no publicar como sitio estático.

---

## 📁 Estructura del proyecto

```
esifast-bi/
├── index.html     ← Estructura HTML completa
├── style.css      ← Estilos del dashboard
├── app.js         ← Lógica, parseo, gráficos, exportación
└── README.md      ← Esta documentación
```

---

## 📊 Funcionalidades

### Login
- Pantalla de acceso con usuario y contraseña configurable en `app.js`.
- Sesión guardada en `sessionStorage` (se limpia al cerrar el navegador).
- Botón de cerrar sesión en el sidebar.

### Carga del Excel
- Botón de carga + drag & drop.
- Lee todas las hojas: `RESUMEN`, `INGRESOS`, `EGRESOS ENERGY`, `EGRESOS ESIFAR`.
- Valida que existan las hojas esperadas.
- Procesamiento 100% local — el archivo nunca sale del navegador.

### Dashboard Principal
- **12 KPIs**: Total IN, OUT, Almacenaje, Facturado, Operaciones, Pallets IN/OUT, Saldo, Clientes, OP abiertas, Días promedio, Promedio por OP.
- **5 gráficos Chart.js**: Ingresos por cliente, Almacenaje por cliente, Pallets comparativo, Distribución de conceptos (doughnut), Evolución mensual.
- Tabla de top 10 OPs por almacenaje.

### Operaciones
- Tabla completa con todas las OPs consolidadas.
- **Filtros**: cliente, fecha desde/hasta, estado, Nro OP, mes.
- Búsqueda en tiempo real.
- Ordenamiento por columna (click en encabezado).
- Paginación de 20 registros por página.
- Exportación a CSV y Excel (.xlsx).

### Análisis Ejecutivo
- Comparativa CLADAN ENERGY vs ESIFAR.
- Análisis de facturación, pallets, permanencia y datos incompletos.
- Generado automáticamente al procesar el archivo.

### Recomendaciones
- 8 recomendaciones automáticas con prioridad (alta/media/baja).
- Detección de OPs sin salida, saldos pendientes, facturación potencial, remitos faltantes.

### Resumen
- Tabla de la hoja RESUMEN del Excel por empresa y mes.
- Totales por concepto (IN, OUT, Almacenaje, Desconsolidado).

### Exportar PDF
- Informe ejecutivo completo con: KPIs, gráficos, tabla resumen, análisis y recomendaciones.
- Generado con jsPDF directamente en el navegador.

---

## 📦 Tecnologías

| Librería | Versión | CDN |
|----------|---------|-----|
| SheetJS (xlsx) | 0.18.5 | cdnjs.cloudflare.com |
| Chart.js | 4.4.1 | cdnjs.cloudflare.com |
| html2canvas | 1.4.1 | cdnjs.cloudflare.com |
| jsPDF | 2.5.1 | cdnjs.cloudflare.com |

---

## 🌐 Publicar en GitHub Pages

1. Crear un repositorio en GitHub.
2. Subir los archivos: `index.html`, `style.css`, `app.js`, `README.md`.
3. Ir a **Settings → Pages → Source: main branch / root**.
4. El sitio estará disponible en `https://TU-USUARIO.github.io/NOMBRE-REPO/`.

### ⚠ Antes de publicar

- [ ] Cambiar credenciales en `app.js` (variable `CONFIG`).
- [ ] Agregar el repositorio al `.gitignore` los archivos Excel reales (`*.xlsx`).
- [ ] Revisar que no haya datos sensibles en el código.

---

## 📋 Hojas Excel requeridas

| Hoja | Descripción |
|------|-------------|
| `RESUMEN` | Tarifarios y totales por empresa y mes |
| `INGRESOS` | Registro de operaciones de ingreso |
| `EGRESOS ENERGY` | Movimientos de salida CLADAN ENERGY |
| `EGRESOS ESIFAR` | Movimientos de salida ESIFAR |
| `DESCONSOLIDADOS` | Operaciones de desconsolidado (puede estar vacía) |

---

## 🔧 Configuración

En `app.js`, modificar el objeto `CONFIG`:

```javascript
const CONFIG = {
  user: 'admin',          // ← Cambiar usuario
  pass: 'Cambiar123!',    // ← Cambiar contraseña (NO usar en producción)
  rowsPerPage: 20,        // Filas por página en tabla
  expectedSheets: [...],  // Hojas requeridas en el Excel
};
```

---

## 💡 Limitaciones

- El login frontend **no es seguro**. Ver advertencia arriba.
- El archivo Excel no se guarda — hay que volver a cargarlo si se recarga la página.
- Los gráficos en el PDF pueden verse diferentes según el navegador (renderizado canvas).
- Para datasets muy grandes (+500 OPs) la generación del PDF puede ser lenta.

---

## 📄 Licencia

MIT — libre uso con atribución.
