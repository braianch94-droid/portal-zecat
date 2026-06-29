const DB_BINDING = 'CICLICO_DB';

async function hashPassword(password) {
  const enc = new TextEncoder();
  const km = await crypto.subtle.importKey('raw', enc.encode(password), 'PBKDF2', false, ['deriveBits']);
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const bits = await crypto.subtle.deriveBits(
    { name: 'PBKDF2', salt, iterations: 100000, hash: 'SHA-256' }, km, 256
  );
  return btoa(String.fromCharCode(...salt)) + ':' + btoa(String.fromCharCode(...new Uint8Array(bits)));
}

async function verifyPassword(password, stored) {
  if (!stored || !stored.includes(':')) return false;
  try {
    const [sb, hb] = stored.split(':');
    const salt = Uint8Array.from(atob(sb), c => c.charCodeAt(0));
    const enc = new TextEncoder();
    const km = await crypto.subtle.importKey('raw', enc.encode(password), 'PBKDF2', false, ['deriveBits']);
    const bits = await crypto.subtle.deriveBits(
      { name: 'PBKDF2', salt, iterations: 100000, hash: 'SHA-256' }, km, 256
    );
    return btoa(String.fromCharCode(...new Uint8Array(bits))) === hb;
  } catch { return false; }
}

function makeResp(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      'Content-Type': 'application/json',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS'
    }
  });
}

async function getSession(DB, request) {
  const auth = request.headers.get('Authorization') || '';
  const token = auth.replace('Bearer ', '').trim();
  if (!token) return null;
  const s = await DB.prepare(
    'SELECT s.token, u.id as uid, u.email, u.nombre, u.role, u.modulos, u.activo FROM sessions s JOIN users u ON s.user_id=u.id WHERE s.token=? AND s.expires_at > datetime("now") AND u.activo=1'
  ).bind(token).first();
  if (!s) return null;
  return { token, userId: s.uid, email: s.email, nombre: s.nombre, role: s.role, modulos: safeJSON(s.modulos, []) };
}

function safeJSON(v, fallback) {
  if (!v) return fallback;
  if (typeof v !== 'string') return v;
  try { return JSON.parse(v); } catch { return fallback; }
}

async function safeJson(request) {
  try { return await request.json(); } catch { return {}; }
}

export async function onRequest(context) {
  const { request, env, params } = context;
  if (request.method === 'OPTIONS') return makeResp(null, 204);
  const DB = env[DB_BINDING];
  if (!DB) return makeResp({ error: 'DB not configured' }, 503);
  const seg = Array.isArray(params.path) ? params.path : (params.path || '').split('/').filter(Boolean);
  const method = request.method;

  // INIT — solo si no hay usuarios
  if (seg[0] === 'init' && method === 'POST') {
    const cnt = await DB.prepare('SELECT COUNT(*) as n FROM users').first();
    if (cnt && cnt.n > 0) return makeResp({ error: 'Already initialized' }, 400);
    const { email, password, nombre = '' } = await safeJson(request);
    if (!email || !password) return makeResp({ error: 'Email y contraseña requeridos' }, 400);
    const hash = await hashPassword(password);
    const now = new Date().toISOString();
    await DB.prepare("INSERT INTO users (email,nombre,password_hash,role,modulos,activo,created_at) VALUES (?,?,?,'superadmin','[]',1,?)").bind(email.toLowerCase(), nombre, hash, now).run();
    return makeResp({ ok: true });
  }

  // LOGIN
  if (seg[0] === 'login' && method === 'POST') {
    const { email, password } = await safeJson(request);
    if (!email || !password) return makeResp({ error: 'Credenciales requeridas' }, 400);
    const user = await DB.prepare('SELECT * FROM users WHERE email=? COLLATE NOCASE').bind(email.toLowerCase()).first();
    if (!user || !user.activo) return makeResp({ error: 'Credenciales inválidas' }, 401);
    if (!await verifyPassword(password, user.password_hash)) return makeResp({ error: 'Credenciales inválidas' }, 401);
    const token = crypto.randomUUID() + '-' + crypto.randomUUID();
    const expires = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString();
    await DB.prepare('INSERT INTO sessions (token,user_id,expires_at) VALUES (?,?,?)').bind(token, user.id, expires).run();
    await DB.prepare('UPDATE users SET last_login=? WHERE id=?').bind(new Date().toISOString(), user.id).run();
    return makeResp({ token, user: { id: user.id, email: user.email, nombre: user.nombre, role: user.role, modulos: safeJSON(user.modulos, []) } });
  }

  // ME
  if (seg[0] === 'me' && method === 'GET') {
    const sess = await getSession(DB, request);
    if (!sess) return makeResp({ error: 'No autenticado' }, 401);
    return makeResp({ user: { id: sess.userId, email: sess.email, nombre: sess.nombre, role: sess.role, modulos: sess.modulos } });
  }

  // LOGOUT
  if (seg[0] === 'logout' && method === 'POST') {
    const sess = await getSession(DB, request);
    if (sess) await DB.prepare('DELETE FROM sessions WHERE token=?').bind(sess.token).run();
    return makeResp({ ok: true });
  }

  // REGISTER — auto-registro queda inactivo hasta que superadmin active
  if (seg[0] === 'register' && method === 'POST') {
    const { email, password, nombre = '' } = await safeJson(request);
    if (!email || !password) return makeResp({ error: 'Email y contraseña requeridos' }, 400);
    if (password.length < 8) return makeResp({ error: 'La contraseña debe tener al menos 8 caracteres' }, 400);
    const hash = await hashPassword(password);
    const now = new Date().toISOString();
    try {
      await DB.prepare("INSERT INTO users (email,nombre,password_hash,role,modulos,activo,created_at) VALUES (?,?,?,'viewer','[]',0,?)").bind(email.toLowerCase(), nombre, hash, now).run();
      return makeResp({ ok: true, message: 'Solicitud enviada. Esperá que el administrador active tu cuenta.' });
    } catch(e) {
      return makeResp({ error: 'El email ya está registrado' }, 409);
    }
  }

  // USERS (solo superadmin)
  if (seg[0] === 'users') {
    const sess = await getSession(DB, request);
    if (!sess) return makeResp({ error: 'No autenticado' }, 401);
    if (sess.role !== 'superadmin') return makeResp({ error: 'Sin permiso' }, 403);

    if (method === 'GET') {
      const { results } = await DB.prepare('SELECT id,email,nombre,role,modulos,activo,created_at,last_login FROM users ORDER BY id').all();
      return makeResp(results.map(u => ({ ...u, modulos: safeJSON(u.modulos, []) })));
    }

    if (method === 'POST') {
      const { email, password, nombre = '', role = 'viewer', modulos = [] } = await safeJson(request);
      if (!email || !password) return makeResp({ error: 'Email y contraseña requeridos' }, 400);
      try {
        const hash = await hashPassword(password);
        const now = new Date().toISOString();
        const r = await DB.prepare('INSERT INTO users (email,nombre,password_hash,role,modulos,activo,created_at) VALUES (?,?,?,?,?,1,?)').bind(email.toLowerCase(), nombre, hash, role, JSON.stringify(modulos), now).run();
        return makeResp({ id: r.meta.last_row_id, ok: true });
      } catch(e) { return makeResp({ error: 'El email ya existe' }, 409); }
    }

    if (method === 'PUT' && seg[1]) {
      const id = parseInt(seg[1]);
      const target = await DB.prepare('SELECT role FROM users WHERE id=?').bind(id).first();
      if (target?.role === 'superadmin' && sess.userId !== id) return makeResp({ error: 'No podés modificar otro superadmin' }, 403);
      const body = await safeJson(request);
      const upd = [], binds = [];
      if (body.nombre !== undefined) { upd.push('nombre=?'); binds.push(body.nombre); }
      if (body.role !== undefined) { upd.push('role=?'); binds.push(body.role); }
      if (body.modulos !== undefined) { upd.push('modulos=?'); binds.push(JSON.stringify(body.modulos)); }
      if (body.activo !== undefined) { upd.push('activo=?'); binds.push(body.activo ? 1 : 0); }
      if (body.password) { upd.push('password_hash=?'); binds.push(await hashPassword(body.password)); }
      if (!upd.length) return makeResp({ ok: true });
      binds.push(id);
      await DB.prepare(`UPDATE users SET ${upd.join(',')} WHERE id=?`).bind(...binds).run();
      if (body.activo === false) await DB.prepare('DELETE FROM sessions WHERE user_id=?').bind(id).run();
      return makeResp({ ok: true });
    }
  }

  // CHANGE PASSWORD
  if (seg[0] === 'password' && method === 'PUT') {
    const sess = await getSession(DB, request);
    if (!sess) return makeResp({ error: 'No autenticado' }, 401);
    const { current_password, new_password } = await safeJson(request);
    if (!new_password || new_password.length < 8) return makeResp({ error: 'La nueva contraseña debe tener al menos 8 caracteres' }, 400);
    const user = await DB.prepare('SELECT password_hash FROM users WHERE id=?').bind(sess.userId).first();
    if (!await verifyPassword(current_password, user.password_hash)) return makeResp({ error: 'Contraseña actual incorrecta' }, 401);
    await DB.prepare('UPDATE users SET password_hash=? WHERE id=?').bind(await hashPassword(new_password), sess.userId).run();
    return makeResp({ ok: true });
  }

  // FORGOT PASSWORD
  if (seg[0] === 'forgot' && method === 'POST') {
    const { email } = await safeJson(request);
    if (!email) return makeResp({ error: 'Email requerido' }, 400);
    const user = await DB.prepare("SELECT id, nombre FROM users WHERE email=? COLLATE NOCASE AND activo=1").bind(email.toLowerCase()).first();
    if (!user) return makeResp({ ok: true }); // no revelar si existe o no
    await DB.prepare("DELETE FROM password_resets WHERE user_id=? OR expires_at < datetime('now')").bind(user.id).run();
    const token = crypto.randomUUID() + '-' + crypto.randomUUID();
    const expires = new Date(Date.now() + 60 * 60 * 1000).toISOString();
    await DB.prepare('INSERT INTO password_resets (token,user_id,expires_at,used) VALUES (?,?,?,0)').bind(token, user.id, expires).run();
    const origin = new URL(request.url).origin;
    const resetUrl = origin + '/?reset=' + token;
    const resendKey = env.RESEND_API_KEY;
    if (resendKey) {
      try {
        await fetch('https://api.resend.com/emails', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + resendKey },
          body: JSON.stringify({
            from: 'Zecat Portal <noreply@zecat.com>',
            to: [email],
            subject: 'Recuperar contraseña — Zecat Portal',
            html: `<p>Hola ${user.nombre||''},</p><p>Hacé click acá para recuperar tu contraseña (válido 1 hora):</p><p><a href="${resetUrl}">${resetUrl}</a></p><p>Si no solicitaste esto, ignorá este email.</p>`
          })
        });
        return makeResp({ ok: true, sent: true });
      } catch(e) {}
    }
    return makeResp({ ok: true, sent: false, reset_url: resetUrl });
  }

  // RESET PASSWORD
  if (seg[0] === 'reset' && method === 'POST') {
    const { token, password } = await safeJson(request);
    if (!token || !password) return makeResp({ error: 'Datos incompletos' }, 400);
    if (password.length < 8) return makeResp({ error: 'La contraseña debe tener al menos 8 caracteres' }, 400);
    const reset = await DB.prepare("SELECT * FROM password_resets WHERE token=? AND expires_at > datetime('now') AND used=0").bind(token).first();
    if (!reset) return makeResp({ error: 'Link inválido o expirado' }, 400);
    await DB.prepare('UPDATE users SET password_hash=? WHERE id=?').bind(await hashPassword(password), reset.user_id).run();
    await DB.prepare('UPDATE password_resets SET used=1 WHERE token=?').bind(token).run();
    await DB.prepare('DELETE FROM sessions WHERE user_id=?').bind(reset.user_id).run();
    return makeResp({ ok: true });
  }

  return makeResp({ error: 'Not found' }, 404);
}
