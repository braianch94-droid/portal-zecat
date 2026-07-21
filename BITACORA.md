# Bitácora — Portal Zecat / Dashboard Productividad

Registro simple de temas, cambios y pendientes. Se va sumando a medida que aparecen (no es changelog automático de git, es para tener memoria de qué pasó y por qué).

Formato: fecha · tema · qué pasó · resuelto / pendiente.

---

## 2026-07-08

**Bonos — nuevas reglas por sector.** Se agregaron pestañas de Bono: Presentismo-only (Flavio Malagrino, Maria Ledesma, Adrian Romero), Muestra Simple (Agustin Lezcano), Despacho, Pie de Máquina (Jose Penida, Brian Ocampo) y Producción (42 personas de Evaluación de Desempeño, fórmula propia: Presentismo 5% + Disciplina 5% [MAG>80% y PBI+20%] + Reclamos/Mermas 5% [reclamos<1% y mermas≤0,05%], escalonado, máx 15%, sin bono grupal). Resuelto.

**Bug: Reclamos Control en 0.** El nombre en la columna "Control" del Excel de Reclamos venía en Title Case y no matcheaba con los nombres ALL CAPS de la hoja Control → los reclamos de Control siempre daban 0. Fix: uppercase al leer `$ctrlName` en `Update-Dashboard.ps1`. Resuelto.

**N1/N2 Pickeadores — fórmula corregida.** No es por % de cumplimiento sino por valores absolutos: N1 = ≥84 líneas/día y ≤1 reclamo, N2 = ≥100 líneas/día y 0 reclamos. Resuelto.

**Toggle Activo/Inactivo por persona en Bonos.** Para poder ocultar gente del cálculo sin borrar datos. Resuelto.

**Bug: badge de filtro stale.** Al cambiar de sección (Bonos/Control/Maquinistas/Resumen) el cartel "Filtro: ..." quedaba con el valor viejo por un `return` temprano en `applyFilter()`. Resuelto.

**Login — email obligatorio y válido.** Se detectó una cuenta real (id=8) con email inválido (`b0da+ga$34`), lo que rompía la recuperación de contraseña. Se agregó validación de formato en registro/alta, endpoint `PUT /me/email` para autoactualizar, y un flujo obligatorio de "actualizá tu email" al loguearse si el email no es válido. Resuelto (la cuenta rota se autocorrige la próxima vez que esa persona entre).

**PENDIENTE — Tarea programada de descarga se corta seguido.** `Zecat_DescargaProductividad` tiene activado "Iniciar la tarea solo si el equipo funciona con corriente alterna" y la notebook se queda sin correr cuando no está enchufada. Pasó 3+ veces esta sesión, cada vez hubo que correr el pipeline a mano (`run_descarga.bat` + `Actualizar-Dashboard.bat`). Requiere que el usuario entre a Task Scheduler → la tarea → Properties → Conditions → destildar esa opción (pide contraseña de Windows, no lo puedo hacer yo). **Sigue sin resolverse.**

**Transporte/Logística — botón "Analizar y generar dashboard".** Reportado como que no respondía para Milagros Sanchez. Investigado: el botón funciona bien, fue error de carga del archivo Rutas (no error de código). No se aplicó fix en el repo fuente.

---

## 2026-07-16

**Descarga automática cortada + contraseña de Infor cambiada.** La tarea no corría desde el 29/06 (sigue el tema batería). Además, al correrla a mano, el login a Infor WMS fallaba: la contraseña había cambiado y el script tenía la vieja en `descarga-querys.py` línea 24. El usuario la actualizó y la descarga corrió OK (98 Picking, 12 Control, 2 Maquinista). Resuelto (data al día).

**Bug: `Actualizar-Dashboard.bat` no parseaba.** Los comentarios del .bat tenían guiones unicode (── ) y un paréntesis suelto que rompían el parseo de cmd.exe en este entorno. Se limpiaron los comentarios (sin tocar la lógica). Resuelto.

**Bonos ahora se guardan POR MES (historial).** Antes los datos manuales (presentismo, errores de Muestra Simple/Despacho/Pie de Máquina, checks de Producción, bono grupal) se guardaban en localStorage con clave fija, así que al cambiar de mes seguías viendo lo último cargado. Ahora cada mes guarda su propia carga: al cambiar a otro mes se lee la de ese mes, y si el mes no tiene nada (ej. uno recién habilitado) aparece en blanco. Migración una-sola-vez: lo ya cargado quedó asignado a julio 2026 para no perderlo. Activo/Inactivo se dejó global (nivel roster, no por mes). En `Update-Dashboard.ps1`: `_bonPeriodKey()`/`_bonKey()`/`_bonLoadPeriod()` + saves con clave namespaced por período. Validado en navegador (aislamiento por mes + migración). Resuelto y publicado.

**QA del cambio de bonos + arreglos.** Corrí QA sobre el bono-por-mes (validado sobre el artefacto deployado, sin ensuciar datos). 0 bugs, 0 errores de consola. Se arreglaron 3 hallazgos: (1) colisión de período cuando Año=Todos + mes puntual → ahora bucket propio `todos-MM`; (2) limpieza de claves viejas huérfanas de localStorage (`zecat-bon-*` sin `::mes`); (3) el `.bat` ahora versiona también `Update-Dashboard.ps1`, no solo el HTML generado. Pendiente de decisión del usuario: default de presentismo (hoy viene "presente" en mes nuevo).

**Bug: "difBadge is not defined" en Dashboard Cíclico CHILE.** El dashboard de Inventario Cíclico de Chile no cargaba (error `difBadge is not defined`). Causa real (NO era cache, aunque al principio lo parecía): el archivo `inventario-ciclico-chile/index.html` **usaba** `difBadge(...)` en 3 lugares pero **nunca definía la función** — port incompleto desde la versión ARG (a Chile se le copiaron `swColor`/`tierBadge` pero se olvidó `difBadge`). El cíclico de ARG siempre estuvo bien. Fix: agregar la definición de `difBadge` en el archivo de Chile (idéntica a ARG). Validado en producción. Resuelto.
De paso, se agregó cache-busting a los iframes lazy del portal (`_ifrSrc` con token por carga en `index.html`) para que futuras actualizaciones de dashboards embebidos no queden pegadas con versiones viejas. Nota: los errores `onLoad is not defined` en consola son ruido pre-existente benigno (carrera de timing de iframes), no rompen nada.

**Descarga Infor rota por editor Monaco (Run query deshabilitado).** La descarga fallaba consistentemente en la 1ra query: el editor Monaco de Compass no terminaba de inicializar cuando el script pegaba el SQL (`setValue` no entraba), y el botón "Run query" quedaba deshabilitado → Timeout. Fix en `descarga-querys.py` `enter_sql()`: reintentar `setValue` hasta que Monaco esté listo y verificar que el SQL entró, + "nudge" con teclado real (espacio+backspace) para disparar el evento que habilita Run, + esperar a que el botón esté habilitado antes de clickear. Validado: bajó 119 Picking / 15 Control / 3 Maquinista (días 14-16/07). Resuelto.
