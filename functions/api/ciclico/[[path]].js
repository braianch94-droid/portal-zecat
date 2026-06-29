const API_KEY = 'zecat-ciclico-2026';

function cors(resp) {
  resp.headers.set('Access-Control-Allow-Origin', '*');
  resp.headers.set('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  resp.headers.set('Access-Control-Allow-Headers', 'Content-Type, X-Ciclico-Key');
  return resp;
}

function json(data, status = 200) {
  return cors(new Response(JSON.stringify(data), {
    status,
    headers: {'Content-Type': 'application/json'}
  }));
}

export async function onRequest(context) {
  try {
  const {request, env, params} = context;
  const DB = env.CICLICO_DB;

  if (request.method === 'OPTIONS') return cors(new Response(null, {status: 204}));

  const key = request.headers.get('X-Ciclico-Key');
  if (key !== API_KEY) return json({error: 'Unauthorized'}, 401);

  if (!DB) return json({error: 'DB not configured'}, 503);

  const url = new URL(request.url);
  const seg = Array.isArray(params.path) ? params.path : (params.path || '').split('/').filter(Boolean);
  const t = url.searchParams.get('tabla') === 'cl' ? 'cl' : 'arg';
  const method = request.method;

  // === CONTEOS ===
  if (seg[0] === 'conteos') {
    const tbl = `conteos_${t}`;

    if (method === 'GET') {
      const fecha = url.searchParams.get('fecha');
      let results;
      if (fecha) {
        ({results} = await DB.prepare(`SELECT * FROM ${tbl} WHERE dia = ? ORDER BY id DESC`).bind(fecha).all());
      } else {
        ({results} = await DB.prepare(`SELECT * FROM ${tbl} ORDER BY id DESC`).all());
      }
      return json(results);
    }

    if (method === 'POST') {
      const d = await request.json().catch(() => ({}));
      if (!d.articulo) return json({error: 'Campo articulo requerido'}, 400);
      const r = await DB.prepare(
        `INSERT INTO ${tbl} (articulo,sku_wms,familia,sistema,fisico,ms,total,dif_cont,diferencia,ajustado,motivo,observacion,mes,dia,ddp,stock_ddp,dif_ddp,fecha_ingreso) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`
      ).bind(d.articulo,d.sku_wms,d.familia,d.sistema,d.fisico,d.ms??0,d.total,d.dif_cont,d.diferencia,d.ajustado??'NO',d.motivo,d.observacion,d.mes,d.dia,d.ddp??0,d.stock_ddp,d.dif_ddp,d.fecha_ingreso).run();
      return json({id: r.meta.last_row_id, success: true});
    }

    if (method === 'PUT' && seg[1]) {
      const id = parseInt(seg[1]);
      const d = await request.json();
      await DB.prepare(
        `UPDATE ${tbl} SET articulo=?,sku_wms=?,familia=?,sistema=?,fisico=?,ms=?,total=?,dif_cont=?,diferencia=?,ajustado=?,motivo=?,observacion=?,mes=?,dia=?,ddp=?,stock_ddp=?,dif_ddp=? WHERE id=?`
      ).bind(d.articulo,d.sku_wms,d.familia,d.sistema,d.fisico,d.ms??0,d.total,d.dif_cont,d.diferencia,d.ajustado,d.motivo,d.observacion,d.mes,d.dia,d.ddp??0,d.stock_ddp,d.dif_ddp,id).run();
      return json({success: true});
    }

    if (method === 'DELETE' && seg[1]) {
      const id = parseInt(seg[1]);
      await DB.prepare(`DELETE FROM ${tbl} WHERE id=?`).bind(id).run();
      return json({success: true});
    }
  }

  // === STOCK ===
  if (seg[0] === 'stock') {
    const tbl = `stock_${t}`;

    if (method === 'GET') {
      const {results} = await DB.prepare(`SELECT * FROM ${tbl} ORDER BY articulo`).all();
      return json(results);
    }

    if (method === 'POST') {
      const items = await request.json();
      if (!Array.isArray(items) || !items.length) return json({success: true, total: 0});
      await DB.prepare(`DELETE FROM ${tbl}`).run();
      for (let i = 0; i < items.length; i += 100) {
        const batch = items.slice(i, i + 100);
        await DB.batch(batch.map(s =>
          DB.prepare(`INSERT INTO ${tbl} (articulo,familia,stock_web,ddp,stock_valorizado,sku,situacion,fecha_actualizacion) VALUES (?,?,?,?,?,?,?,?)`)
            .bind(s.articulo, s.familia||'', s.stock_web||0, s.ddp||0, s.stock_valorizado||0, s.sku||s.sku_wms||'', s.situacion||'', s.fecha_actualizacion||'')
        ));
      }
      return json({success: true, total: items.length});
    }
  }

  // === CICLO ===
  if (seg[0] === 'ciclo') {
    const tbl = `ciclo_${t}`;

    if (method === 'GET') {
      let r = await DB.prepare(`SELECT * FROM ${tbl} WHERE id=1`).first();
      if (!r) {
        const today = new Date().toISOString().split('T')[0];
        await DB.prepare(`INSERT INTO ${tbl} (id,ciclo_numero,fecha_inicio,total_elegibles,ya_recomendados) VALUES (1,1,?,0,'[]')`).bind(today).run();
        r = {id:1, ciclo_numero:1, fecha_inicio:today, total_elegibles:0, ya_recomendados:'[]'};
      }
      return json(r);
    }

    if (method === 'PUT') {
      const d = await request.json();
      const yr = typeof d.ya_recomendados === 'string' ? d.ya_recomendados : JSON.stringify(d.ya_recomendados || []);
      await DB.prepare(`INSERT OR REPLACE INTO ${tbl} (id,ciclo_numero,fecha_inicio,total_elegibles,ya_recomendados) VALUES (1,?,?,?,?)`)
        .bind(d.ciclo_numero||1, d.fecha_inicio||'', d.total_elegibles||0, yr).run();
      return json({success: true});
    }
  }

  // === RECOMENDACIONES ===
  if (seg[0] === 'recomendaciones') {
    const tbl = `recomendaciones_${t}`;
    const fecha = url.searchParams.get('fecha');

    if (method === 'GET') {
      if (!fecha) return json(null);
      const r = await DB.prepare(`SELECT * FROM ${tbl} WHERE fecha=?`).bind(fecha).first();
      return json(r || null);
    }

    if (method === 'PUT') {
      const d = await request.json();
      const items = typeof d.items === 'string' ? d.items : JSON.stringify(d.items || []);
      await DB.prepare(`INSERT OR REPLACE INTO ${tbl} (fecha,items) VALUES (?,?)`).bind(d.fecha, items).run();
      return json({success: true});
    }
  }

  return json({error: 'Not found'}, 404);
  } catch(e) {
    return json({error: 'Internal error'}, 500);
  }
}
