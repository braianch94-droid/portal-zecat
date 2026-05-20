# Dashboard de Productividad — Zecat

Dashboard interactivo HTML generado automáticamente desde datos Excel de picking, reclamos y producción de Zecat Artículos Promocionales SA.

---

## Archivos del proyecto

| Archivo | Descripción |
|---------|-------------|
| `Update-Dashboard.ps1` | Script principal — genera el HTML y el Excel de salida |
| `Actualizar-Dashboard.bat` | Doble clic para actualizar el dashboard |
| `Setup-Tareas.ps1` | Configuración de tareas programadas (opcional) |
| `chartjs.min.js` | Chart.js 4.4.0 embebido (funciona sin internet) |
| `Dashboard_Productividad.html` | Dashboard generado — abrir en navegador |

> `productividad.xlsx` (datos fuente) **no se incluye en el repositorio** — contiene datos privados.

---

## Requisitos

- Windows con **PowerShell 5+**
- **Microsoft Excel** instalado (el script lo usa vía COM para leer el `.xlsx`)
- Archivo `productividad.xlsx` en la misma carpeta que el script

---

## Uso

1. Asegurarse de que `productividad.xlsx` esté actualizado y **cerrado**
2. Ejecutar `Actualizar-Dashboard.bat` (o correr el `.ps1` directamente)
3. Abrir `Dashboard_Productividad.html` en el navegador

---

## Pestañas del dashboard

| Pestaña | Contenido |
|---------|-----------|
| **Productividad Picking** | KPIs del equipo, ranking de operarios, evolución mensual, staffing |
| **Pie de Máquina** | Lin/Día por turno (T1 vs T2) |
| **Muestra Simple** | Indicadores de Lezcano Agustín |
| **Reclamos** | Tasa, ranking por operario y por categoría, Sin Identificar |
| **Control** | En desarrollo |

---

## Configuración rápida

Al inicio de `Update-Dashboard.ps1` se pueden ajustar estos parámetros:

```powershell
# Target de lineas/día para picking regular
$TARGET = 84

# Operarios excluidos del cálculo de productividad del equipo
# (extras, no-pickers regulares) — usar fragmento del nombre
$ExcludeFromProd = @("AIRALA", "Falta definir")
```

---

## Tema claro / oscuro

El dashboard incluye un botón ☾ en la esquina superior derecha del header para cambiar entre modo claro y oscuro. La preferencia se guarda en el navegador.
