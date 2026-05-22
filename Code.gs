// ═══════════════════════════════════════════════════════════════════
//  JANTAR DOS CASAIS — Google Apps Script Backend v3
//
//  COMO PUBLICAR:
//  1. Planilha → Extensões → Apps Script
//  2. Cole este código (apague o anterior)
//  3. Salve → Executar → "setupSheets" (autorize quando pedir)
//  4. Implantar → Nova implantação → Aplicativo da Web
//     • Executar como: Eu mesmo
//     • Quem tem acesso: Qualquer pessoa (inclusive anônimos)
//  5. Copie a URL e cole no index.html (variável SHEET_URL)
//  ⚠️  Sempre que editar: crie uma NOVA implantação (não edite a antiga)
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
function ok(extra)  { return resp(Object.assign({ ok: true }, extra || {})); }
function fail(msg)  { return resp({ ok: false, error: String(msg) }); }

// ── Roteador (tudo via GET) ──────────────────────────────────────────
function doGet(e) {
  try {
    var action  = (e.parameter.action  || '').trim();
    var payload = (e.parameter.payload || '{}');
    var d = {};
    try { d = JSON.parse(decodeURIComponent(payload)); } catch(_) {}

    switch (action) {
      case 'ping':             return ok({ msg: 'pong' });
      case 'setupSheets':      return ok(setupSheets());
      case 'getConfig':        return ok(getConfig());
      case 'saveConfig':       return ok(saveConfig(d));
      case 'savePixConfig':    return ok(savePixConfig(d));
      case 'getFoods':         return ok({ foods: getFoods() });
      case 'addFood':          return ok(addFood(d));
      case 'addFoodsBatch':    return ok(addFoodsBatch(d));   // importação txt
      case 'removeFood':       return ok(removeFood(d));
      case 'getReservations':  return ok({ reservations: getReservations() });
      case 'reserveFood':      return ok(reserveFood(d));     // atômico
      case 'confirmPayment':   return ok(confirmPayment(d));
      case 'releaseFood':      return ok(releaseFood(d));
      case 'getSummary':       return ok(getSummary());
      default:                 return ok({ msg: 'API ok' });
    }
  } catch(ex) {
    return fail(ex.message);
  }
}

// ════════════════════════════════════════════════════════════════════
//  SETUP
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

  // Reservations: id | foodId | foodName | person1 | person2 | phone | status | valor | timestamp
  ensure(TAB_RES, ['id','foodId','foodName','person1','person2','phone','status','valor','timestamp']);

  return { msg: 'Planilha configurada!' };
}

// ════════════════════════════════════════════════════════════════════
//  CONFIG  — leitura e escrita campo a campo para evitar perda de dados
// ════════════════════════════════════════════════════════════════════
function getConfig() {
  var sh   = sheet(TAB_CONFIG);
  var rows = sh.getDataRange().getValues();
  var m    = {};
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var key = String(rows[i][0]);
    var val = rows[i][1];
    // Sheets pode converter campos de data/hora em objetos Date
    if (val instanceof Date && !isNaN(val.getTime())) {
      if (key === 'date') {
        // Campo de data: formata como YYYY-MM-DD
        var dd = val.getDate(), mm = val.getMonth()+1, yy = val.getFullYear();
        val = yy + '-' + (mm<10?'0':'') + mm + '-' + (dd<10?'0':'') + dd;
      } else if (key === 'time') {
        // Campo de hora: formata como HH:MM
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
  // Salva somente os campos enviados — nunca apaga o que não veio
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
    var val = map[k];
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === k) {
        sh.getRange(i + 1, 2).setValue(val !== undefined && val !== null ? val : '');
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
  sheet(TAB_FOODS).appendRow([id, d.name || '', d.type || 'Outro', 'free', '', '', '', '']);
  return { id: id };
}

// Importação em lote via txt (array de {name, type})
function addFoodsBatch(d) {
  var items = d.items || [];
  if (!items.length) return { added: 0 };
  var sh   = sheet(TAB_FOODS);
  var rows = items.map(function(item, i) {
    return [String(Date.now() + i), item.name || '', item.type || 'Outro', 'free', '', '', '', ''];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
  return { added: rows.length };
}

function removeFood(d) {
  var sh = sheet(TAB_FOODS), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(d.id)) {
      if (String(rows[i][3]) !== 'free') return { removed: false, error: 'Item já reservado — libere primeiro' };
      sh.deleteRow(i + 1);
      return { removed: true };
    }
  }
  return { removed: false, error: 'Não encontrado' };
}

function setFoodRow(sh, foodId, status, person1, person2, phone) {
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(foodId)) {
      sh.getRange(i+1, 4, 1, 5).setValues([[status, person1||'', person2||'', phone||'', new Date().toISOString()]]);
      return true;
    }
  }
  return false;
}

// ════════════════════════════════════════════════════════════════════
//  RESERVA ATÔMICA com LockService
// ════════════════════════════════════════════════════════════════════
function reserveFood(d) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Servidor ocupado, tente em instantes.' };

  try {
    var fsh  = sheet(TAB_FOODS);
    var rows = fsh.getDataRange().getValues();
    var idx  = -1;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(d.foodId)) { idx = i; break; }
    }
    if (idx < 0) return { ok: false, error: 'Item não encontrado.' };

    if (String(rows[idx][3]) !== 'free') {
      return { ok: false, conflict: true, takenBy: String(rows[idx][4] || ''), error: 'Item já reservado.' };
    }

    // Marca como reservado dentro do lock
    fsh.getRange(idx+1, 4, 1, 5).setValues([['reserved', d.person1||'', d.person2||'', d.phone||'', new Date().toISOString()]]);

    // Obtém valor do evento para gravar na reserva
    var cfg   = getConfig();
    var valor = cfg.event.valor || '0';

    var resId = 'R' + Date.now();
    sheet(TAB_RES).appendRow([
      resId, d.foodId, d.foodName||'',
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
    return {
      id:       String(r[0]),
      foodId:   String(r[1]),
      foodName: String(r[2] || ''),
      person1:  String(r[3] || ''),
      person2:  String(r[4] || ''),
      phone:    String(r[5] || ''),
      status:   String(r[6] || 'pending'),
      valor:    String(r[7] || '0'),
      ts:       String(r[8] || ''),
    };
  });
}

function confirmPayment(d) {
  var lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    var rsh = sheet(TAB_RES), rrows = rsh.getDataRange().getValues();
    var foodId='', p1='', p2='', phone='', valor='0';
    for (var i = 1; i < rrows.length; i++) {
      if (String(rrows[i][0]) === String(d.reservationId)) {
        rsh.getRange(i+1, 7).setValue('confirmed');
        foodId = String(rrows[i][1]); p1 = String(rrows[i][3]);
        p2 = String(rrows[i][4]); phone = String(rrows[i][5]); valor = String(rrows[i][7]);
        break;
      }
    }
    if (foodId) setFoodRow(sheet(TAB_FOODS), foodId, 'confirmed', p1, p2, phone);
    return { confirmed: true, person1: p1, person2: p2, phone: phone, valor: valor };
  } finally { lock.releaseLock(); }
}

function releaseFood(d) {
  var lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    var rsh = sheet(TAB_RES), rrows = rsh.getDataRange().getValues();
    for (var i = 1; i < rrows.length; i++) {
      if (String(rrows[i][0]) === String(d.reservationId)) {
        rsh.getRange(i+1, 7).setValue('cancelled'); break;
      }
    }
    setFoodRow(sheet(TAB_FOODS), d.foodId, 'free', '', '', '');
    return { released: true };
  } finally { lock.releaseLock(); }
}

// ════════════════════════════════════════════════════════════════════
//  RESUMO / DASHBOARD
// ════════════════════════════════════════════════════════════════════
function getSummary() {
  var res   = getReservations().filter(function(r){ return r.status !== 'cancelled'; });
  var foods = getFoods();

  var totalCasais    = res.length;
  var totalPessoas   = totalCasais * 2;
  var confirmados    = res.filter(function(r){ return r.status === 'confirmed'; }).length;
  var pendentes      = res.filter(function(r){ return r.status === 'pending'; }).length;
  var arrecadado     = res.reduce(function(s, r){ return s + Number(r.valor || 0); }, 0);
  var aConfirmar     = res.filter(function(r){ return r.status === 'pending'; })
                         .reduce(function(s, r){ return s + Number(r.valor || 0); }, 0);

  var livres    = foods.filter(function(f){ return f.status === 'free'; }).length;
  var reservados= foods.filter(function(f){ return f.status !== 'free'; }).length;

  return {
    totalCasais:    totalCasais,
    totalPessoas:   totalPessoas,
    confirmados:    confirmados,
    pendentes:      pendentes,
    arrecadado:     arrecadado,
    aConfirmar:     aConfirmar,
    totalItems:     foods.length,
    livres:         livres,
    reservados:     reservados,
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
