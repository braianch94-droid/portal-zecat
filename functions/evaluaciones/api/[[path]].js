// Pages Function: API de Evaluaciones de Desempeño (Cloudflare Pages + D1)
// Ruta: /evaluaciones/api/*   |   Binding D1: env.DB

const JSON_HEADERS = { "content-type": "application/json; charset=utf-8" };
const json = (data, status = 200) => new Response(JSON.stringify(data), { status, headers: JSON_HEADERS });
const bad = (msg, status = 400) => json({ error: msg }, status);

export async function onRequest(context) {
  const { request, env } = context;
  const url = new URL(request.url);
  // Normalizo: /evaluaciones/api/meta -> /api/meta (reuso el router tal cual)
  url.pathname = url.pathname.replace(/^\/evaluaciones/, "") || "/";
  try {
    return await api(request, env, url);
  } catch (err) {
    return json({ error: "Error del servidor", detail: String(err) }, 500);
  }
}

// ---------- Autenticación (contraseña compartida) ----------
const AUTH_SALT = "evaluaciones-auth-v1";
const COOKIE = "ev_sess";

function resolvePassword(env) {
  return (env.APP_PASSWORD && String(env.APP_PASSWORD)) ||
         (env.APP_PASSWORD_DEFAULT && String(env.APP_PASSWORD_DEFAULT)) ||
         "zecat2025";
}
async function token(pw) {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey("raw", enc.encode(pw), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const sig = await crypto.subtle.sign("HMAC", key, enc.encode(AUTH_SALT));
  return [...new Uint8Array(sig)].map(b => b.toString(16).padStart(2, "0")).join("");
}
function timingSafeEqual(a, b) {
  if (typeof a !== "string" || typeof b !== "string" || a.length !== b.length) return false;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}
function getCookie(request, name) {
  const c = request.headers.get("cookie") || "";
  const m = c.match(new RegExp("(?:^|;\\s*)" + name + "=([^;]+)"));
  return m ? m[1] : null;
}
function sessionCookie(value, url, maxAge) {
  const secure = url.protocol === "https:" ? " Secure;" : "";
  return `${COOKIE}=${value}; HttpOnly; Path=/evaluaciones; SameSite=Lax;${secure} Max-Age=${maxAge}`;
}

async function api(request, env, url) {
  const db = env.DB;
  const path = url.pathname;
  const method = request.method;
  const seg = path.split("/").filter(Boolean); // ["api", ...]

  const expected = await token(resolvePassword(env));

  // POST /api/login  { password }
  if (path === "/api/login" && method === "POST") {
    const b = await request.json().catch(() => ({}));
    const ok = timingSafeEqual(await token(String(b.password || "")), expected);
    if (!ok) return bad("Contraseña incorrecta", 401);
    return new Response(JSON.stringify({ ok: true }), {
      status: 200,
      headers: { ...JSON_HEADERS, "set-cookie": sessionCookie(expected, url, 2592000) },
    });
  }
  // POST /api/logout
  if (path === "/api/logout" && method === "POST") {
    return new Response(JSON.stringify({ ok: true }), {
      status: 200,
      headers: { ...JSON_HEADERS, "set-cookie": sessionCookie("", url, 0) },
    });
  }

  // A partir de acá, todo requiere sesión válida
  if (!timingSafeEqual(getCookie(request, COOKIE) || "", expected)) {
    return bad("No autenticado", 401);
  }

  // GET /api/meta
  if (path === "/api/meta" && method === "GET") {
    const objetivos = (await db.prepare("SELECT id, nombre, orden FROM objetivos WHERE activo=1 ORDER BY orden").all()).results;
    const dims = (await db.prepare("SELECT DISTINCT sector, area, tecnica, turno FROM empleados WHERE activo=1").all()).results;
    const sectores = [...new Set(dims.map(d => d.sector).filter(Boolean))].sort();
    const areas = [...new Set(dims.map(d => d.area).filter(Boolean))].sort();
    const tecnicas = [...new Set(dims.map(d => d.tecnica).filter(Boolean))].sort();
    const turnos = [...new Set(dims.map(d => d.turno).filter(Boolean))].sort();
    const periodos = (await db.prepare("SELECT DISTINCT periodo FROM evaluaciones ORDER BY periodo DESC").all()).results.map(r => r.periodo);
    return json({
      objetivos, sectores, areas, tecnicas, turnos, periodos,
      escala: [
        { v: 1, label: "No Cumple" },
        { v: 2, label: "Cumple Parcialmente" },
        { v: 3, label: "Cumple" },
        { v: 4, label: "Supera" },
        { v: 5, label: "Excede" },
      ],
    });
  }

  // GET /api/empleados
  if (path === "/api/empleados" && method === "GET") {
    const sector = url.searchParams.get("sector");
    const area = url.searchParams.get("area");
    const turno = url.searchParams.get("turno");
    const tecnica = url.searchParams.get("tecnica");
    const periodo = url.searchParams.get("periodo");
    const q = (url.searchParams.get("q") || "").trim().toLowerCase();

    let sql = `
      SELECT e.*,
        (SELECT COUNT(*) FROM evaluaciones ev WHERE ev.empleado_id = e.id ${periodo ? "AND ev.periodo = ?p" : ""}) AS n_eval,
        (SELECT ev.promedio FROM evaluaciones ev WHERE ev.empleado_id = e.id ${periodo ? "AND ev.periodo = ?p" : ""} ORDER BY ev.fecha DESC, ev.id DESC LIMIT 1) AS ult_promedio,
        (SELECT ev.periodo FROM evaluaciones ev WHERE ev.empleado_id = e.id ORDER BY ev.fecha DESC, ev.id DESC LIMIT 1) AS ult_periodo,
        (SELECT ev.fecha FROM evaluaciones ev WHERE ev.empleado_id = e.id ORDER BY ev.fecha DESC, ev.id DESC LIMIT 1) AS ult_fecha
      FROM empleados e WHERE e.activo=1`;
    const binds = [];
    sql = sql.replaceAll("?p", "?");
    if (periodo) { binds.push(periodo, periodo); }
    if (sector) { sql += " AND e.sector = ?"; binds.push(sector); }
    if (area) { sql += " AND e.area = ?"; binds.push(area); }
    if (turno) { sql += " AND e.turno = ?"; binds.push(turno); }
    if (tecnica) { sql += " AND e.tecnica = ?"; binds.push(tecnica); }
    sql += " ORDER BY e.apellido, e.nombre";

    let rows = (await db.prepare(sql).bind(...binds).all()).results;
    if (q) {
      rows = rows.filter(r =>
        (`${r.apellido} ${r.nombre}`).toLowerCase().includes(q) ||
        (r.tecnica || "").toLowerCase().includes(q) ||
        (r.area || "").toLowerCase().includes(q));
    }
    return json(rows);
  }

  // GET /api/empleados/:id
  if (seg[1] === "empleados" && seg[2] && method === "GET") {
    const id = Number(seg[2]);
    const emp = await db.prepare("SELECT * FROM empleados WHERE id=?").bind(id).first();
    if (!emp) return bad("Empleado no encontrado", 404);
    const evals = (await db.prepare(
      "SELECT id, fecha, periodo, evaluador, promedio FROM evaluaciones WHERE empleado_id=? ORDER BY fecha DESC, id DESC"
    ).bind(id).all()).results;
    return json({ ...emp, evaluaciones: evals });
  }

  // GET /api/evaluaciones/:id
  if (seg[1] === "evaluaciones" && seg[2] && method === "GET") {
    const id = Number(seg[2]);
    const ev = await db.prepare(`
      SELECT ev.*, e.nombre, e.apellido, e.sector, e.area, e.tecnica, e.turno
      FROM evaluaciones ev JOIN empleados e ON e.id = ev.empleado_id WHERE ev.id=?`).bind(id).first();
    if (!ev) return bad("Evaluación no encontrada", 404);
    const puntajes = (await db.prepare(
      "SELECT p.objetivo_id, p.puntaje, o.nombre, o.orden FROM evaluacion_puntajes p JOIN objetivos o ON o.id=p.objetivo_id WHERE p.evaluacion_id=? ORDER BY o.orden"
    ).bind(id).all()).results;
    return json({ ...ev, puntajes });
  }

  // POST /api/evaluaciones
  if (path === "/api/evaluaciones" && method === "POST") {
    const b = await request.json();
    const err = validarEval(b);
    if (err) return bad(err);
    const prom = avg(b.puntajes.map(p => p.puntaje));
    const now = new Date().toISOString();
    const res = await db.prepare(`
      INSERT INTO evaluaciones (empleado_id, fecha, periodo, evaluador, posicion_evaluador, posicion_desempenada, aspectos_mejorar, comentarios_evaluador, comentarios_evaluado, promedio, created_at)
      VALUES (?,?,?,?,?,?,?,?,?,?,?)`).bind(
        b.empleado_id, b.fecha, b.periodo, b.evaluador,
        b.posicion_evaluador || null, b.posicion_desempenada || null,
        b.aspectos_mejorar || null, b.comentarios_evaluador || null, b.comentarios_evaluado || null,
        prom, now
      ).run();
    const evalId = res.meta.last_row_id;
    await guardarPuntajes(db, evalId, b.puntajes);
    return json({ id: evalId, promedio: prom }, 201);
  }

  // PUT /api/evaluaciones/:id
  if (seg[1] === "evaluaciones" && seg[2] && method === "PUT") {
    const id = Number(seg[2]);
    const b = await request.json();
    const err = validarEval(b);
    if (err) return bad(err);
    const prom = avg(b.puntajes.map(p => p.puntaje));
    const now = new Date().toISOString();
    const upd = await db.prepare(`
      UPDATE evaluaciones SET empleado_id=?, fecha=?, periodo=?, evaluador=?, posicion_evaluador=?, posicion_desempenada=?, aspectos_mejorar=?, comentarios_evaluador=?, comentarios_evaluado=?, promedio=?, updated_at=?
      WHERE id=?`).bind(
        b.empleado_id, b.fecha, b.periodo, b.evaluador,
        b.posicion_evaluador || null, b.posicion_desempenada || null,
        b.aspectos_mejorar || null, b.comentarios_evaluador || null, b.comentarios_evaluado || null,
        prom, now, id
      ).run();
    if (upd.meta.changes === 0) return bad("Evaluación no encontrada", 404);
    await db.prepare("DELETE FROM evaluacion_puntajes WHERE evaluacion_id=?").bind(id).run();
    await guardarPuntajes(db, id, b.puntajes);
    return json({ id, promedio: prom });
  }

  // DELETE /api/evaluaciones/:id
  if (seg[1] === "evaluaciones" && seg[2] && method === "DELETE") {
    const id = Number(seg[2]);
    await db.prepare("DELETE FROM evaluacion_puntajes WHERE evaluacion_id=?").bind(id).run();
    const r = await db.prepare("DELETE FROM evaluaciones WHERE id=?").bind(id).run();
    if (r.meta.changes === 0) return bad("Evaluación no encontrada", 404);
    return json({ ok: true });
  }

  // GET /api/dashboard
  if (path === "/api/dashboard" && method === "GET") {
    const periodo = url.searchParams.get("periodo");
    const sector = url.searchParams.get("sector");
    const where = [];
    const binds = [];
    if (periodo) { where.push("ev.periodo = ?"); binds.push(periodo); }
    if (sector) { where.push("e.sector = ?"); binds.push(sector); }
    const wsql = where.length ? "WHERE " + where.join(" AND ") : "";

    const totalEmpleados = (await db.prepare("SELECT COUNT(*) c FROM empleados WHERE activo=1" + (sector ? " AND sector=?" : "")).bind(...(sector ? [sector] : [])).first()).c;

    const base = `FROM evaluaciones ev JOIN empleados e ON e.id=ev.empleado_id ${wsql}`;
    const resumen = await db.prepare(`SELECT COUNT(*) n_evaluaciones, COUNT(DISTINCT ev.empleado_id) n_evaluados, AVG(ev.promedio) prom_global, MIN(ev.promedio) min_prom, MAX(ev.promedio) max_prom ${base}`).bind(...binds).first();

    const porSector = (await db.prepare(`SELECT e.sector AS k, AVG(ev.promedio) prom, COUNT(*) n ${base} GROUP BY e.sector ORDER BY prom DESC`).bind(...binds).all()).results;
    const porArea = (await db.prepare(`SELECT COALESCE(NULLIF(e.area,''), e.tecnica) AS k, AVG(ev.promedio) prom, COUNT(*) n ${base} GROUP BY k ORDER BY prom DESC`).bind(...binds).all()).results;
    const porTurno = (await db.prepare(`SELECT e.turno AS k, AVG(ev.promedio) prom, COUNT(*) n ${base} GROUP BY e.turno ORDER BY prom DESC`).bind(...binds).all()).results;

    const baseP = `FROM evaluacion_puntajes p JOIN evaluaciones ev ON ev.id=p.evaluacion_id JOIN empleados e ON e.id=ev.empleado_id JOIN objetivos o ON o.id=p.objetivo_id ${wsql}`;
    const porObjetivo = (await db.prepare(`SELECT o.nombre AS k, o.orden, AVG(p.puntaje) prom, COUNT(*) n ${baseP} GROUP BY o.id ORDER BY o.orden`).bind(...binds).all()).results;

    const ranking = (await db.prepare(`
      SELECT e.id, e.nombre, e.apellido, e.sector, COALESCE(NULLIF(e.area,''), e.tecnica) AS area, AVG(ev.promedio) prom, COUNT(*) n
      ${base} GROUP BY e.id ORDER BY prom DESC`).bind(...binds).all()).results;

    const dist = (await db.prepare(`
      SELECT CASE
        WHEN ev.promedio < 2 THEN '1-2'
        WHEN ev.promedio < 3 THEN '2-3'
        WHEN ev.promedio < 4 THEN '3-4'
        ELSE '4-5' END AS bucket, COUNT(*) n
      ${base} GROUP BY bucket`).bind(...binds).all()).results;

    return json({ totalEmpleados, resumen, porSector, porArea, porTurno, porObjetivo, ranking, dist });
  }

  // GET /api/export
  if (path === "/api/export" && method === "GET") {
    const rows = (await db.prepare(`
      SELECT ev.id, e.apellido, e.nombre, e.sector, e.area, e.tecnica, e.turno,
        ev.periodo, ev.fecha, ev.evaluador, ev.promedio,
        ev.aspectos_mejorar, ev.comentarios_evaluador, ev.comentarios_evaluado
      FROM evaluaciones ev JOIN empleados e ON e.id=ev.empleado_id
      ORDER BY ev.fecha DESC, ev.id DESC`).all()).results;
    return json(rows);
  }

  return bad("Ruta no encontrada", 404);
}

function validarEval(b) {
  if (!b || typeof b !== "object") return "Cuerpo inválido";
  if (!b.empleado_id) return "Falta el empleado";
  if (!b.fecha) return "Falta la fecha";
  if (!b.periodo) return "Falta el período";
  if (!b.evaluador || !String(b.evaluador).trim()) return "Falta el evaluador";
  if (!Array.isArray(b.puntajes) || b.puntajes.length === 0) return "Faltan los puntajes";
  for (const p of b.puntajes) {
    if (!p.objetivo_id) return "Puntaje sin objetivo";
    const n = Number(p.puntaje);
    if (!Number.isInteger(n) || n < 1 || n > 5) return "Cada puntaje debe ser un entero de 1 a 5";
  }
  return null;
}
async function guardarPuntajes(db, evalId, puntajes) {
  const stmt = db.prepare("INSERT INTO evaluacion_puntajes (evaluacion_id, objetivo_id, puntaje) VALUES (?,?,?)");
  const batch = puntajes.map(p => stmt.bind(evalId, p.objetivo_id, Number(p.puntaje)));
  await db.batch(batch);
}
function avg(arr) {
  if (!arr.length) return 0;
  return Math.round((arr.reduce((a, b) => a + Number(b), 0) / arr.length) * 100) / 100;
}
