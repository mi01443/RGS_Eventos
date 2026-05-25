// ═══════════════════════════════════════════════════════════════════
//  JANTAR DOS CASAIS — Google Apps Script Backend v4
//
//  COMO PUBLICAR:
//  1. Planilha → Extensões → Apps Script
//  2. Cole este código (apague o anterior)
//  3. Salvar → Executar → "setupSheets" (autorize quando pedir)
//  4. Implantar → Nova implantação → Aplicativo da Web
//     • Executar como: Eu mesmo
//     • Quem tem acesso: Qualquer pessoa (inclusive anônimos)
//  5. Copie a URL e cole no index.html (variável SHEET_URL)
//  ⚠️  Sempre que editar: crie uma NOVA implantação
// ═══════════════════════════════════════════════════════════════════

var TAB_CONFIG = 'Config';
var TAB_FOODS  = 'Foods';
var TAB_RES    = 'Reservations';

// ── Resposta JSON ────────────────────────────────────────────────────
function resp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok(extra) { return resp(Object.assign({ ok: true }, extra || {})); }
function fail(msg) { return resp({ ok: false, error: String(msg) }); }

// ── Roteador principal (tudo via GET) ────────────────────────────────
function doGet(e) {
  try {
    var action  = (e.parameter.action  || '').trim();
    var payload = (e.parameter.payload || '{}');
    var d = {};
    try { d = JSON.parse(decodeURIComponent(payload)); } catch(_) {}

    switch (action) {
      case 'ping':              return ok({ msg: 'pong' });
      case 'setupSheets':       return ok(setupSheets());
      case 'getConfig':         return ok(getConfig());
      case 'saveConfig':        return ok(saveConfig(d));
      case 'savePixConfig':     return ok(savePixConfig(d));
      case 'getFoods':          return ok({ foods: getFoods() });
      case 'addFood':           return ok(addFood(d));
      case 'addFoodsBatch':     return ok(addFoodsBatch(d));
      case 'removeFood':        return ok(removeFood(d));
      case 'getReservations':   return ok({ reservations: getReservations() });
      case 'reserveFoods':      return ok(reserveFoods(d));      // múltiplos itens, atômico
      case 'confirmPayment':    return ok(confirmPayment(d));
      case 'releaseReservation':return ok(releaseReservation(d)); // libera todos os itens da reserva
      case 'getSummary':        return ok(getSummary());
      default:                  return ok({ msg: 'API ok — ação desconhecida: ' + action });
    }
  } catch(ex) {
    return fail(ex.message);
  }
}

// ════════════════════════════════════════════════════════════════════
//  SETUP — cria abas se não existirem
// ════════════════════════════════════════════════════════════════════
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function ensure(name, headers) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    return sh;
  }

  var cfg = ensure(TAB_CONFIG, ['key', 'value']);
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, 11, 2).setValues([
      ['title','Jantar dos Casais'],
      ['date',''], ['time',''], ['local',''],
      ['valor','100'], ['desc',''],
      ['whatsapp',''], ['senha','admin123'],
      ['pix_key',''], ['pix_name',''], ['pix_bank',''],
    ]);
  }

  // Foods: id | name | type | status | person1 | person2 | phone | timestamp
  ensure(TAB_FOODS, ['id','name','type','status','person1','person2','phone','timestamp']);

  // Reservations: id | foodIds | foodNames | person1 | person2 | phone | status | valor | timestamp
  // foodIds e foodNames são listas separadas por vírgula (suporte a múltiplos itens)
  ensure(TAB_RES, ['id','foodIds','foodNames','person1','person2','phone','status','valor','timestamp']);

  return { msg: 'Planilha configurada com sucesso!' };
}

// ════════════════════════════════════════════════════════════════════
//  CONFIG
// ════════════════════════════════════════════════════════════════════
function getConfig() {
  var sh   = sheet(TAB_CONFIG);
  var rows = sh.getDataRange().getValues();
  var m    = {};
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var key = String(rows[i][0]);
    var val = rows[i][1];
    // Sheets converte datas/horas em objetos Date — serializa corretamente
    if (val instanceof Date && !isNaN(val.getTime())) {
      if (key === 'date') {
        var dd = val.getDate(), mm = val.getMonth()+1, yy = val.getFullYear();
        val = yy + '-' + (mm<10?'0':'') + mm + '-' + (dd<10?'0':'') + dd;
      } else if (key === 'time') {
        var hh = val.getHours(), mi = val.getMinutes();
        val = (hh<10?'0':'') + hh + ':' + (mi<10?'0':'') + mi;
      } else {
        val = String(val);
      }
    } else {
      val = String(val !== undefined && val !== null ? val : '');
    }
    m[key] = val;
  }
  return {
    event: {
      title:    m['title']    || '',
      date:     m['date']     || '',
      time:     m['time']     || '',
      local:    m['local']    || '',
      valor:    m['valor']    || '0',
      desc:     m['desc']     || '',
      whatsapp: m['whatsapp'] || '',
      senha:    m['senha']    || 'admin123',
    },
    pix: {
      key:  m['pix_key']  || '',
      name: m['pix_name'] || '',
      bank: m['pix_bank'] || '',
    }
  };
}

function saveConfig(d) {
  var map = {};
  if (d.title    !== undefined) map['title']    = d.title;
  if (d.date     !== undefined) map['date']     = d.date;
  if (d.time     !== undefined) map['time']     = d.time;
  if (d.local    !== undefined) map['local']    = d.local;
  if (d.valor    !== undefined) map['valor']    = d.valor;
  if (d.desc     !== undefined) map['desc']     = d.desc;
  if (d.whatsapp !== undefined) map['whatsapp'] = d.whatsapp;
  if (d.senha    !== undefined && d.senha !== '') map['senha'] = d.senha;
  setConfigKeys(map);
  return { saved: true };
}

function savePixConfig(d) {
  var map = {};
  if (d.key  !== undefined) map['pix_key']  = d.key;
  if (d.name !== undefined) map['pix_name'] = d.name;
  if (d.bank !== undefined) map['pix_bank'] = d.bank;
  setConfigKeys(map);
  return { saved: true };
}

function setConfigKeys(map) {
  var sh   = sheet(TAB_CONFIG);
  var rows = sh.getDataRange().getValues();
  Object.keys(map).forEach(function(k) {
    if (map[k] === undefined || map[k] === null) return;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === k) {
        sh.getRange(i+1, 2).setValue(map[k]);
        break;
      }
    }
  });
}

// ════════════════════════════════════════════════════════════════════
//  FOODS
// ════════════════════════════════════════════════════════════════════
function getFoods() {
  var rows = sheet(TAB_FOODS).getDataRange().getValues().slice(1);
  return rows.filter(function(r){ return !!r[0]; }).map(function(r){
    return {
      id:        String(r[0]),
      name:      String(r[1] || ''),
      type:      String(r[2] || ''),
      status:    String(r[3] || 'free'),
      person1:   String(r[4] || ''),
      person2:   String(r[5] || ''),
      phone:     String(r[6] || ''),
      timestamp: String(r[7] || ''),
    };
  });
}

function addFood(d) {
  var id = d.id || String(Date.now());
  sheet(TAB_FOODS).appendRow([id, d.name||'', d.type||'Outro', 'free', '', '', '', '']);
  return { id: id };
}

function addFoodsBatch(d) {
  var items = d.items || [];
  if (!items.length) return { added: 0 };
  var sh   = sheet(TAB_FOODS);
  var rows = items.map(function(item, i) {
    return [String(Date.now()+i), item.name||'', item.type||'Outro', 'free', '', '', '', ''];
  });
  sh.getRange(sh.getLastRow()+1, 1, rows.length, 8).setValues(rows);
  return { added: rows.length };
}

function removeFood(d) {
  var sh = sheet(TAB_FOODS), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(d.id)) {
      if (String(rows[i][3]) !== 'free') return { removed: false, error: 'Item já reservado — libere primeiro' };
      sh.deleteRow(i+1);
      return { removed: true };
    }
  }
  return { removed: false, error: 'Não encontrado' };
}

function setFoodRow(sh, foodId, status, person1, person2, phone) {
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(foodId)) {
      sh.getRange(i+1, 4, 1, 5).setValues([[
        status, person1||'', person2||'', phone||'', new Date().toISOString()
      ]]);
      return true;
    }
  }
  return false;
}

// ════════════════════════════════════════════════════════════════════
//  RESERVA ATÔMICA — múltiplos itens com LockService
//  Garante que dois requests simultâneos não reservem o mesmo item
// ════════════════════════════════════════════════════════════════════
function reserveFoods(d) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Servidor ocupado, tente em instantes.' };

  try {
    var foodIds   = d.foodIds   || [];
    var foodNames = d.foodNames || [];
    var fsh  = sheet(TAB_FOODS);
    var rows = fsh.getDataRange().getValues();

    // Verifica todos os itens antes de reservar qualquer um
    for (var k = 0; k < foodIds.length; k++) {
      var fid = String(foodIds[k]);
      var found = false;
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === fid) {
          found = true;
          if (String(rows[i][3]) !== 'free') {
            return {
              ok: false, conflict: true,
              conflictItem: String(rows[i][1]),
              error: 'Item "' + rows[i][1] + '" já foi reservado.'
            };
          }
          break;
        }
      }
      if (!found) return { ok: false, error: 'Item ' + fid + ' não encontrado.' };
    }

    // Todos livres — reserva todos dentro do lock
    for (var k = 0; k < foodIds.length; k++) {
      setFoodRow(fsh, foodIds[k], 'reserved', d.person1||'', d.person2||'', d.phone||'');
    }

    // Obtém valor do evento
    var cfg   = getConfig();
    var valor = cfg.event.valor || '0';

    // Cria UMA reserva com todos os itens
    var resId     = 'R' + Date.now();
    var idsStr    = foodIds.join(',');
    var namesStr  = (Array.isArray(foodNames) ? foodNames : []).join(',');

    sheet(TAB_RES).appendRow([
      resId, idsStr, namesStr,
      d.person1||'', d.person2||'', d.phone||'',
      'pending', valor, new Date().toISOString(),
    ]);

    return { ok: true, reservationId: resId };
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════════════
//  RESERVATIONS
// ════════════════════════════════════════════════════════════════════
function getReservations() {
  var rows = sheet(TAB_RES).getDataRange().getValues().slice(1);
  return rows.filter(function(r){ return !!r[0]; }).map(function(r){
    var idsStr   = String(r[1] || '');
    var namesStr = String(r[2] || '');
    return {
      id:        String(r[0]),
      foodIds:   idsStr   ? idsStr.split(',')   : [],
      foodNames: namesStr,          // string separada por vírgula para exibição
      person1:   String(r[3] || ''),
      person2:   String(r[4] || ''),
      phone:     String(r[5] || ''),
      status:    String(r[6] || 'pending'),
      valor:     String(r[7] || '0'),
      ts:        String(r[8] || ''),
    };
  });
}

function confirmPayment(d) {
  var lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    var rsh   = sheet(TAB_RES);
    var rrows = rsh.getDataRange().getValues();
    var foodIds = [], p1 = '', p2 = '', phone = '', valor = '0';

    for (var i = 1; i < rrows.length; i++) {
      if (String(rrows[i][0]) === String(d.reservationId)) {
        rsh.getRange(i+1, 7).setValue('confirmed');
        var idsStr = String(rrows[i][1] || '');
        foodIds = idsStr ? idsStr.split(',') : (d.foodIds || []);
        p1 = String(rrows[i][3]); p2 = String(rrows[i][4]);
        phone = String(rrows[i][5]); valor = String(rrows[i][7]);
        break;
      }
    }

    // Atualiza todos os itens para 'confirmed'
    var fsh = sheet(TAB_FOODS);
    foodIds.forEach(function(fid) {
      setFoodRow(fsh, fid.trim(), 'confirmed', p1, p2, phone);
    });

    return { confirmed: true, person1: p1, person2: p2, phone: phone, valor: valor };
  } finally { lock.releaseLock(); }
}

function releaseReservation(d) {
  var lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    var rsh   = sheet(TAB_RES);
    var rrows = rsh.getDataRange().getValues();

    // Normaliza foodIds do payload: aceita array ou string separada por vírgula
    var foodIds = [];
    if (d.foodIds) {
      if (Array.isArray(d.foodIds)) {
        foodIds = d.foodIds.map(function(id){ return String(id).trim(); }).filter(Boolean);
      } else {
        foodIds = String(d.foodIds).split(',').map(function(id){ return id.trim(); }).filter(Boolean);
      }
    }

    // Busca a reserva na planilha
    for (var i = 1; i < rrows.length; i++) {
      if (String(rrows[i][0]) === String(d.reservationId)) {
        rsh.getRange(i+1, 7).setValue('cancelled');
        // SEMPRE lê os ids da planilha — fonte de verdade garantida
        var idsStr = String(rrows[i][1] || '');
        var idsFromSheet = idsStr ? idsStr.split(',').map(function(s){ return s.trim(); }).filter(Boolean) : [];
        // Merge: usa planilha como base, adiciona payload (garante cobertura total)
        idsFromSheet.forEach(function(id){
          if (id && foodIds.indexOf(id) === -1) foodIds.push(id);
        });
        break;
      }
    }

    if (!foodIds.length) {
      return { released: false, error: 'Nenhum item encontrado para liberar' };
    }

    // Libera TODOS os itens encontrados
    var fsh = sheet(TAB_FOODS);
    foodIds.forEach(function(fid) {
      if (fid) setFoodRow(fsh, fid, 'free', '', '', '');
    });

    return { released: true, count: foodIds.length };
  } finally { lock.releaseLock(); }
}

// ════════════════════════════════════════════════════════════════════
//  RESUMO / DASHBOARD
// ════════════════════════════════════════════════════════════════════
function getSummary() {
  var res   = getReservations().filter(function(r){ return r.status !== 'cancelled'; });
  var foods = getFoods();

  var confirmados = res.filter(function(r){ return r.status === 'confirmed'; });
  var pendentes   = res.filter(function(r){ return r.status === 'pending'; });

  return {
    totalCasais:  res.length,
    totalPessoas: res.length * 2,
    confirmados:  confirmados.length,
    pendentes:    pendentes.length,
    arrecadado:   confirmados.reduce(function(s,r){ return s + Number(r.valor||0); }, 0),
    aConfirmar:   pendentes.reduce(function(s,r){ return s + Number(r.valor||0); }, 0),
    totalItems:   foods.length,
    livres:       foods.filter(function(f){ return f.status === 'free'; }).length,
    reservados:   foods.filter(function(f){ return f.status !== 'free'; }).length,
  };
}

// ════════════════════════════════════════════════════════════════════
//  HELPER
// ════════════════════════════════════════════════════════════════════
function sheet(name) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('Aba "' + name + '" não encontrada. Execute setupSheets primeiro.');
  return s;
}
