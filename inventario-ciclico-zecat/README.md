# Inventario Ciclico - Zecat ARG 2025-26

Sistema automatizado de inventario ciclico para **Zecat - Articulos Promocionales SA**.
Selecciona articulos a contar diariamente, envia email con la lista y genera un dashboard HTML con el progreso.

---

## Que hace el sistema

1. **Cada dia laboral a las 7:00 AM** analiza el stock disponible y selecciona 20 articulos para contar
2. Envia un email con la lista del dia + resumen del conteo anterior
3. Adjunta un archivo `.csv` con los SKUs listos para importar
4. **A las 8:00 AM** regenera el dashboard HTML con los datos del archivo Ciclico

### Logica de seleccion de articulos

- **Nunca repite** un articulo hasta haber recorrido todos los elegibles (sistema de ciclos)
- Prioriza articulos con **bajo stock web** (Tramo 1 primero, luego T2, T3, T4)
- Excluye articulos ya contados en el Ciclico y los ya recomendados en el ciclo actual
- Al completar el ciclo completo, reinicia automaticamente con Ciclo 2, 3, etc.

### Dias en que NO corre

- Sabados y domingos
- Feriados nacionales de Argentina (lista hardcodeada, actualizable anualmente)
- Periodos de pausa configurados (vacaciones, etc.)

---

## Estructura del proyecto

```
Ciclico/
├── Generar_Conteo_Diario.ps1           # Script principal - seleccion y email
├── Actualizar_Dashboard.ps1            # Genera el Dashboard HTML
├── ENVIAR CONTEO DIARIO.bat            # Ejecucion manual del conteo
├── INSTALAR TAREA AUTOMATICA.bat       # Instala las tareas en el Programador de Windows
├── .gitignore
└── README.md

# Archivos de datos (NO incluidos en el repo, van en la misma carpeta):
├── Ciclico ARG 2025-26.xlsx            # Registro de conteos del periodo
├── Stock a HOY.xlsx                    # Stock actual exportado del sistema
├── Dashboard_Ciclico.html              # Generado automaticamente
├── ciclo_actual.json                   # Estado del ciclo (generado al primer run)
├── historial_recomendaciones.json      # Historial de recomendaciones diarias
└── cobertura_familias.json             # Cobertura por familia (generado por el conteo)
```

---

## Requisitos

- **Windows 10/11**
- **PowerShell 5.1** (incluido en Windows)
- **Microsoft Excel** instalado (para leer los .xlsx via COM)
- **Microsoft Outlook** instalado y configurado (para enviar emails)
- Los archivos de datos en la misma carpeta que los scripts

---

## Instalacion desde cero

### 1. Clonar el repositorio

```powershell
git clone https://github.com/TU_USUARIO/inventario-ciclico.git
cd inventario-ciclico
```

### 2. Colocar los archivos de datos

Copiar en la misma carpeta:
- `Ciclico ARG 2025-26.xlsx`
- `Stock a HOY.xlsx`

> Los archivos JSON de estado (`ciclo_actual.json`, etc.) se crean solos en el primer run.

### 3. Instalar las tareas automaticas

Hacer clic derecho sobre `INSTALAR TAREA AUTOMATICA.bat` y elegir **Ejecutar como administrador**.

Esto crea dos tareas en el Programador de Windows:
- `Zecat - Conteo Ciclico Diario y Email` → 7:00 AM, lunes a viernes
- `Zecat - Actualizar Dashboard Ciclico` → 8:00 AM, lunes a viernes

### 4. Configurar destinatarios (opcional)

Editar `Generar_Conteo_Diario.ps1` y modificar la linea:

```powershell
$Destinatarios = "email1@zecat.com; email2@zecat.com; email3@zecat.com"
```

---

## Ejecucion manual

Para enviar el conteo del dia sin esperar la tarea programada:

```
Doble click en: ENVIAR CONTEO DIARIO.bat
```

---

## Configurar feriados y pausas

### Agregar feriados o puentes turisticos

En `Generar_Conteo_Diario.ps1` y `Actualizar_Dashboard.ps1`, buscar el bloque:

```powershell
$feriadosARG = @(
    # 2026
    "01/01/2026","16/02/2026", ...
    # Agregar puentes turisticos cuando el gobierno los anuncie
)
```

Agregar la fecha en formato `"dd/MM/yyyy"`.
**Actualizar esta lista al inicio de cada año.**

### Configurar una pausa de vacaciones

Buscar el bloque `PAUSA VACACIONES` y modificar las fechas:

```powershell
$pausaDesde = [datetime]"2026-05-18"
$pausaHasta = [datetime]"2026-05-26"
```

---

## Como funciona el sistema de ciclos

El archivo `ciclo_actual.json` guarda el estado del ciclo actual:

```json
{
  "cicloNumero": 1,
  "fechaInicio": "dd/MM/yyyy",
  "totalElegibles": 1558,
  "contadosEnCiclico": 486,
  "recomendadosEsteCiclo": 60,
  "pendientes": 1012,
  "yaRecomendados": ["articulo1", "articulo2", ...]
}
```

- **totalElegibles**: universo de articulos activos en web
- **contadosEnCiclico**: ya tienen conteo registrado en el archivo Ciclico
- **recomendadosEsteCiclo**: enviados por email en el ciclo actual (aun no contados)
- **pendientes**: quedan por recomendar

Cuando `pendientes` llega a 0, el ciclo se reinicia automaticamente.

---

## Dashboard

El archivo `Dashboard_Ciclico.html` se genera automaticamente cada dia habil a las 8 AM.
Incluye:
- Progreso total del inventario ciclico
- Exactitud por mes y por familia
- Impacto economico (diferencias en valor DDP)
- Cobertura por familia
- Toggle de tema claro/oscuro (preferencia guardada en el browser)

---

## Feriados nacionales Argentina - lista actual

| Fecha | Feriado |
|---|---|
| 01/01 | Año Nuevo |
| Lun/Mar previos al Mier de Ceniza | Carnaval |
| 24/03 | Dia de la Memoria |
| 02/04 | Malvinas |
| Viernes previo a Pascua | Viernes Santo |
| 01/05 | Dia del Trabajador |
| 25/05 | Revolucion de Mayo |
| 3er lunes de junio (o lunes mas cercano al 17/06) | Guemes |
| 20/06 | Belgrano |
| 09/07 | Independencia |
| 3er lunes de agosto | San Martin |
| 2do lunes de octubre | Diversidad Cultural |
| 4to lunes de noviembre | Soberania Nacional |
| 08/12 | Inmaculada Concepcion |
| 25/12 | Navidad |

---

## Flujo de datos

```
Stock a HOY.xlsx  ──┐
                    ├──► Generar_Conteo_Diario.ps1 ──► Email con CSV adjunto
Ciclico ARG.xlsx  ──┘                                │
                                                     └──► ciclo_actual.json
                                                          historial.json
                                                          cobertura.json

Ciclico ARG.xlsx ──► Actualizar_Dashboard.ps1 ──► Dashboard_Ciclico.html
```
