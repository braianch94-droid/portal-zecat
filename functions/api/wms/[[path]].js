const API_KEY = 'zecat-wms-2026';

function cors(resp) {
  resp.headers.set('Access-Control-Allow-Origin', '*');
  resp.headers.set('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  resp.headers.set('Access-Control-Allow-Headers', 'Content-Type, X-Wms-Key');
  return resp;
}
function json(data, status = 200) {
  return cors(new Response(JSON.stringify(data), {
    status, headers: {'Content-Type': 'application/json'}
  }));
}
// Fecha en hora de Argentina (UTC-3, sin DST) → "YYYY-MM-DD HH:MM:SS"
function ahora() {
  return new Date(Date.now() - 3 * 3600 * 1000).toISOString().slice(0, 19).replace('T', ' ');
}

async function getState(DB) {
  const [ubis, arts, stk, movs] = await Promise.all([
    DB.prepare('SELECT * FROM ubicaciones ORDER BY codigo').all(),
    DB.prepare('SELECT id, sku, codigo_bejerman, descripcion, unidad, stock_minimo FROM articulos ORDER BY sku').all(),
    DB.prepare('SELECT articulo_id, ubicacion_id, cantidad FROM stock WHERE cantidad > 0').all(),
    DB.prepare('SELECT id, fecha, tipo, articulo_id, ubicacion_origen_id AS origen_id, ubicacion_destino_id AS destino_id, cantidad, usuario, nota FROM movimientos ORDER BY id DESC LIMIT 1000').all(),
  ]);
  return {
    ubicaciones: ubis.results,
    articulos: arts.results,
    stock: stk.results,
    movimientos: movs.results,
  };
}

export async function onRequest(context) {
  try {
    const {request, env, params} = context;
    const DB = env.WMS_DB;

    if (request.method === 'OPTIONS') return cors(new Response(null, {status: 204}));
    if (request.headers.get('X-Wms-Key') !== API_KEY) return json({error: 'Unauthorized'}, 401);
    if (!DB) return json({error: 'DB not configured'}, 503);

    const seg = Array.isArray(params.path) ? params.path : (params.path || '').split('/').filter(Boolean);
    const method = request.method;
    const body = (method === 'POST' || method === 'PUT') ? await request.json().catch(() => ({})) : {};

    // === ESTADO COMPLETO ===
    if (seg[0] === 'state' && method === 'GET') {
      return json(await getState(DB));
    }

    // === UBICACIONES ===
    if (seg[0] === 'ubicaciones') {
      if (method === 'POST' && seg[1] === 'toggle') {
        await DB.prepare('UPDATE ubicaciones SET activa = 1 - activa WHERE id = ?').bind(+body.id).run();
        return json(await getState(DB));
      }
      if (method === 'POST') {
        const cod = (body.codigo || '').trim().toUpperCase();
        if (!cod) return json({error: 'El código de ubicación es obligatorio.'}, 400);
        try {
          await DB.prepare('INSERT INTO ubicaciones (codigo, descripcion, zona, activa) VALUES (?, ?, ?, 1)')
            .bind(cod, (body.descripcion || '').trim(), (body.zona || '').trim()).run();
        } catch (e) {
          if (String(e).includes('UNIQUE')) return json({error: `Ya existe una ubicación con el código ${cod}.`}, 409);
          throw e;
        }
        return json(await getState(DB));
      }
    }

    // === ARTÍCULOS ===
    if (seg[0] === 'articulos') {
      // Carga masiva (maestro). Idempotente: INSERT OR IGNORE por SKU único.
      if (method === 'POST' && seg[1] === 'bulk') {
        const items = Array.isArray(body.items) ? body.items : [];
        let inserted = 0;
        // Upsert por SKU: crea o actualiza descripción/unidad/código Bejerman (no toca stock_minimo)
        const stmt = DB.prepare(
          'INSERT INTO articulos (sku, codigo_bejerman, descripcion, unidad, stock_minimo) VALUES (?, ?, ?, ?, ?) ' +
          'ON CONFLICT(sku) DO UPDATE SET codigo_bejerman = excluded.codigo_bejerman, descripcion = excluded.descripcion, unidad = excluded.unidad'
        );
        for (let i = 0; i < items.length; i += 50) {
          const chunk = items.slice(i, i + 50);
          const res = await DB.batch(chunk.map(it => {
            let min = parseFloat(it.stock_minimo);
            if (!(min >= 0) || isNaN(min)) min = 0;
            return stmt.bind((it.sku || '').trim().toUpperCase(), (it.codigo_bejerman || '').trim(), (it.descripcion || '').trim(), (it.unidad || 'UN').trim() || 'UN', min);
          }));
          for (const r of res) inserted += (r.meta && r.meta.changes) ? r.meta.changes : 0;
        }
        return json({ok: true, received: items.length, inserted});
      }
      if (method === 'POST') {
        const sku = (body.sku || '').trim().toUpperCase();
        if (!sku) return json({error: 'El SKU es obligatorio.'}, 400);
        let min = parseFloat(body.stock_minimo);
        if (!(min >= 0) || isNaN(min)) min = 0;
        try {
          await DB.prepare('INSERT INTO articulos (sku, codigo_bejerman, descripcion, unidad, stock_minimo) VALUES (?, ?, ?, ?, ?)')
            .bind(sku, (body.codigo_bejerman || '').trim(), (body.descripcion || '').trim(), (body.unidad || 'UN').trim() || 'UN', min).run();
        } catch (e) {
          if (String(e).includes('UNIQUE')) return json({error: `Ya existe un artículo con el SKU ${sku}.`}, 409);
          throw e;
        }
        return json(await getState(DB));
      }
      if (method === 'PUT' && seg[1]) {
        let min = parseFloat(body.stock_minimo);
        if (!(min >= 0) || isNaN(min)) min = 0;
        await DB.prepare('UPDATE articulos SET codigo_bejerman = ?, descripcion = ?, unidad = ?, stock_minimo = ? WHERE id = ?')
          .bind((body.codigo_bejerman || '').trim(), (body.descripcion || '').trim(), (body.unidad || 'UN').trim() || 'UN', min, +seg[1]).run();
        return json(await getState(DB));
      }
    }

    // === MOVIMIENTO (ingreso / egreso / transferencia) ===
    if (seg[0] === 'mover' && method === 'POST') {
      const tipo = body.tipo;
      const art = +body.articulo_id;
      const cant = parseFloat(body.cantidad);
      const ori = body.origen_id ? +body.origen_id : null;
      const dest = body.destino_id ? +body.destino_id : null;
      const usuario = (body.usuario || '').trim();
      const nota = (body.nota || '').trim();

      if (!art || !(cant > 0)) return json({error: 'Completá artículo y cantidad válida.'}, 400);

      // Baja de stock en origen (atómica: solo si hay suficiente)
      async function bajarOrigen(uid) {
        const r = await DB.prepare(
          'UPDATE stock SET cantidad = cantidad - ? WHERE articulo_id = ? AND ubicacion_id = ? AND cantidad >= ?'
        ).bind(cant, art, uid, cant).run();
        if (r.meta.changes !== 1) throw new Error('Stock insuficiente en la ubicación.');
      }
      // Alta de stock en destino (upsert incremental)
      async function subirDestino(uid) {
        await DB.prepare(
          'INSERT INTO stock (articulo_id, ubicacion_id, cantidad) VALUES (?, ?, ?) ' +
          'ON CONFLICT(articulo_id, ubicacion_id) DO UPDATE SET cantidad = cantidad + ?'
        ).bind(art, uid, cant, cant).run();
      }

      try {
        if (tipo === 'INGRESO') {
          if (!dest) throw new Error('Elegí la ubicación de destino.');
          await subirDestino(dest);
        } else if (tipo === 'EGRESO') {
          if (!ori) throw new Error('Elegí la ubicación de origen.');
          await bajarOrigen(ori);
        } else if (tipo === 'TRANSFERENCIA') {
          if (!ori || !dest) throw new Error('Elegí origen y destino.');
          if (ori === dest) throw new Error('Origen y destino no pueden ser iguales.');
          await bajarOrigen(ori);
          await subirDestino(dest);
        } else {
          throw new Error('Tipo de movimiento inválido.');
        }
      } catch (e) {
        return json({error: e.message}, 400);
      }

      await DB.prepare(
        'INSERT INTO movimientos (fecha, tipo, articulo_id, ubicacion_origen_id, ubicacion_destino_id, cantidad, usuario, nota) VALUES (?, ?, ?, ?, ?, ?, ?, ?)'
      ).bind(ahora(), tipo, art, tipo === 'INGRESO' ? null : ori, tipo === 'EGRESO' ? null : dest, cant, usuario, nota).run();

      return json(await getState(DB));
    }

    return json({error: 'Not found'}, 404);
  } catch (e) {
    return json({error: 'Internal error', detail: String(e)}, 500);
  }
}
