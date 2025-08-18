/***********************
 * MODELO Y ENCABEZADOS
 ***********************/
const SH = {
  ASOCIACIONES        : 'ASOCIACIONES',
  INSPECCIONES        : 'INSPECCIONES',
  CAUCES              : 'CAUCES',
  AUTORIDAD_VIEW      : 'AUTORIDAD_VIEW',   // derivada
  PERSONAS            : 'PERSONAS',
  INSPECCION_AUT      : 'INSPECCION_AUT',   // asignación de Inspector a Inspección
  ASOCIACION_AUT      : 'ASOCIACION_AUT',   // asignación de Gerente a Asociación
  CAUCE_TOMERO        : 'CAUCE_TOMERO',     // asignación de Tomero a Cauce
  PERSONA_TELEFONOS : 'PERSONA_TELEFONOS',// detalle de teléfonos por persona
  ROLES               : 'ROLES',            // catálogo de roles
  META                : 'META'
};

const HDR = {
  ASOCIACIONES: [
    'ID_ASOCIACION','ASOCIACION_NOMBRE','CUENCA',
    'DIRECCION','LOCALIDAD','DEPARTAMENTO','PROVINCIA','CP','UBICACION',
    'TELEFONO','MAIL',
    'LOGO' // NUEVO
  ],
  INSPECCIONES: [
    'ID_INSPECCION','INSPECCION','FK_ID_ASOCIACION','CUENCA',
    'DIRECCION','LOCALIDAD','DEPARTAMENTO','PROVINCIA','CP','UBICACION',
    'TELEFONO','MAIL',
    'FK_ID_AUTORIDAD',
    'LOGO' // NUEVO
  ],
  CAUCES: [
    'ID_CAUCE','FK_ID_INSPECCION','FK_ID_ASOCIACION','CUENCA',
    'DENOM_CAUCE','CLASE','ORIGEN','DESAGUE','COD_NOMBRE',
    'FK_ID_AUTORIDAD','FK_TOMERO_PRINCIPAL',
    'LOGO' // OPCIONAL
  ],
  AUTORIDAD_VIEW: [
    'ID_AUTORIDAD','ID_INSPECCION','FK_ID_ASOCIACION',
    'AUTORIDAD_NOMBRE','AUTORIDAD_MAIL','AUTORIDAD_TELEFONO'
  ],
 PERSONAS: [
  'ID_PERSONA','NOMBRE','MAIL','TELEFONO',
  'ROL',
  'DIRECCION','LOCALIDAD','DEPARTAMENTO','PROVINCIA','CP','UBICACION',
  'FOTO', // NUEVO

  // NUEVOS CAMPOS PARA SELECCIÓN JERÁRQUICA
  'CUENCA_SELECCIONADA',
  'ASOCIACION_SELECCIONADA',
  'INSPECCION_SELECCIONADA',
  'CAUCE_SELECCIONADO'
],

  INSPECCION_AUT: [
    'ID_INSPECCION_AUT','ID_INSPECCION','FK_ID_ASOCIACION','ID_PERSONA',
    'ROL','PRINCIPAL','VIG_DESDE','VIG_HASTA'
  ],
  ASOCIACION_AUT: [
    'ID_ASOCIACION_AUT','FK_ID_ASOCIACION','ID_PERSONA',
    'ROL','PRINCIPAL','VIG_DESDE','VIG_HASTA'
  ],
  CAUCE_TOMERO: [
    'ID_CAUCE_TOMERO','FK_ID_CAUCE','ID_PERSONA',
    'ROL','PRINCIPAL','VIG_DESDE','VIG_HASTA'
  ],
  PERSONA_TELEFONOS: [
    'ID_TELEFONO','ID_PERSONA','NUMERO','TIPO','PRINCIPAL','NOTAS'
  ],
  ROLES: [
    'ID_ROL','ROL_KEY','ROL_LABEL','ORDEN',
    'ES_INSPECTOR','ES_GERENTE_TECNICO','ES_TOMERO','ACTIVO','NOTAS'
  ],
  META: ['KEY','VALUE']
};

const ID_RANGES = {
  PERSONA          : 700000,
  INSPECCION_AUT   : 910000,
  ASOCIACION_AUT   : 920000,
  CAUCE_TOMERO     : 930000,
  TELEFONO         : 740000,
  ROL              : 980000
};

/***********************
 * INICIALIZACIÓN + MENÚ
 ***********************/
function tryGetUi_(){ try{ return SpreadsheetApp.getUi(); }catch(e){ return null; } }

function buildMenu_(){
  const ui = tryGetUi_(); if (!ui) return;
  const m = ui.createMenu('Modelo Riego');
  m.addItem('Reordenar INSPECCIONES (ubicación)','reorderInspecciones');
  m.addItem('Secuencia completa (auto)','runFullSetup');
  m.addItem('Init paso 1 (rápido)','initModel_step1');
  m.addItem('Init paso 2 (rápido)','initModel_step2');
  m.addItem('Inicializar/Actualizar modelo (todo)','initModel');
  m.addSeparator();

  const subTel = ui.createMenu('Teléfonos')
    .addItem('Reconstruir detalle desde PERSONAS','rebuildTelefonosFromPersonas')
    .addItem('Actualizar PERSONAS.TELEFONO desde detalle','writePhonesToPersonas_');
  m.addSubMenu(subTel);

  const subRol = ui.createMenu('Roles')
    .addItem('Inicializar/precargar lista','rolesInit_')
    .addItem('Forzar semillas por defecto','rolesSeedForce_');
  m.addSubMenu(subRol);

  m.addSeparator();
  m.addItem('Sincronizar FKs','syncFKsAll');
  m.addSeparator();
  m.addItem('Auditar encabezados','auditHeaders');
  m.addItem('Reparar encabezados (seguro)','fixHeadersSafe');
  m.addItem('Reordenar INSPECCIONES (ubicación)','reorderInspecciones');
  m.addSeparator();
  m.addItem('Formatear encabezados','styleAllHeaders');
  m.addItem('Colorear pestañas','colorizeTabs');
  m.addToUi();
}
function onOpen(){ buildMenu_(); }

/***********************
 * UTILIDADES BASE
 ***********************/
function getSh_(name){
  const sh=SpreadsheetApp.getActive().getSheetByName(name);
  if(!sh) throw new Error('Falta hoja: '+name);
  return sh;
}
function asText_(v){ return (v==null)?'':String(v).trim().replace(/\.0$/,''); }
function idx_(headers, name){
  const i=headers.indexOf(name);
  if(i===-1) throw new Error('Falta columna: '+name);
  return i;
}
function readTable_(sh){
  const v=sh.getDataRange().getValues();
  const h=v.shift().map(asText_);
  return {headers:h, rows:v};
}
function writeTable_(sh, headers, rows){
  sh.clear();
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if(rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  sh.getDataRange().setNumberFormat('@');
}
function styleHeader_(sh){
  const lc=Math.max(sh.getLastColumn(),1);
  if(lc<1) return;
  sh.getRange(1,1,1,lc)
    .setBackground('#E6F4FF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true,true,true,true,true,true,'#000',SpreadsheetApp.BorderStyle.SOLID);
  sh.setFrozenRows(1);
}
function styleAllHeaders(){
  Object.values(SH).forEach(name=>{
    const sh = SpreadsheetApp.getActive().getSheetByName(name);
    if (sh) try{ styleHeader_(sh); }catch(e){}
  });
}
function colorizeTabs(){
  const ss = SpreadsheetApp.getActive();
  const EDITABLE = new Set([SH.ASOCIACIONES, SH.INSPECCIONES, SH.PERSONAS, SH.CAUCES, SH.PERSONA_TELEFONOS, SH.ROLES]);
  const blue  = '#90CAF9';
  const gray  = '#BDBDBD';
  ss.getSheets().forEach(sh=>{
    const name = sh.getName();
    sh.setTabColor( EDITABLE.has(name) ? blue : gray );
  });
}

/********** CABECERAS (helpers) **********/
function ensureSheetWithHeaders_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureHeaders_(sh, headers);
  sh.getDataRange().setNumberFormat('@');
}
function ensureHeaders_(sh, headers) {
  const lc  = Math.max(sh.getLastColumn(), headers.length);
  const row = sh.getRange(1,1,1,lc).getValues()[0].map(v=>asText_(v));
  let changed=false;
  headers.forEach((h,i)=>{ if (row[i]!==h){ row[i]=h; changed=true; }});
  if (changed) sh.getRange(1,1,1,lc).setValues([row]);
}
function ensureHeadersExist_(){
  ensureSheetWithHeaders_(SH.ASOCIACIONES,      HDR.ASOCIACIONES);
  ensureSheetWithHeaders_(SH.INSPECCIONES,      HDR.INSPECCIONES);
  ensureSheetWithHeaders_(SH.CAUCES,            HDR.CAUCES);
  ensureSheetWithHeaders_(SH.AUTORIDAD_VIEW,    HDR.AUTORIDAD_VIEW);
  ensureSheetWithHeaders_(SH.PERSONAS,          HDR.PERSONAS);
  ensureSheetWithHeaders_(SH.INSPECCION_AUT,    HDR.INSPECCION_AUT);
  ensureSheetWithHeaders_(SH.ASOCIACION_AUT,    HDR.ASOCIACION_AUT);
  ensureSheetWithHeaders_(SH.CAUCE_TOMERO,      HDR.CAUCE_TOMERO);
  ensureSheetWithHeaders_(SH.PERSONA_TELEFONOS, HDR.PERSONA_TELEFONOS);
  ensureSheetWithHeaders_(SH.ROLES,             HDR.ROLES);
  ensureSheetWithHeaders_(SH.META,              HDR.META);
}

/***********************
 * CONTADORES Y IDs
 ***********************/
function calcNextId_(sheetName, _col0, fallback) {
  try {
    const { rows } = readTable_(getSh_(sheetName));
    if (!rows.length) return fallback;
    const ids = rows.map(r => parseInt(asText_(r[0])||'0',10)).filter(n=>!isNaN(n));
    return Math.max(...ids, fallback-1) + 1;
  } catch(e) { return fallback; }
}
function normalizeMetaCounters(){
  ensureHeadersExist_();
  const meta = getSh_(SH.META);
  const t = readTable_(meta);
  const iKey = idx_(t.headers,'KEY');
  const iVal = idx_(t.headers,'VALUE');
  const targets = [
    { key:'NEXT_ID_PERSONA',      sheet:SH.PERSONAS,       base:ID_RANGES.PERSONA },
    { key:'NEXT_ID_INSPECCION_AUT', sheet:SH.INSPECCION_AUT, base:ID_RANGES.INSPECCION_AUT },
    { key:'NEXT_ID_ASOCIACION_AUT', sheet:SH.ASOCIACION_AUT, base:ID_RANGES.ASOCIACION_AUT },
    { key:'NEXT_ID_CAUCE_TOMERO',   sheet:SH.CAUCE_TOMERO,   base:ID_RANGES.CAUCE_TOMERO },
    { key:'NEXT_ID_TELEFONO',       sheet:SH.PERSONA_TELEFONOS, base:ID_RANGES.TELEFONO }
  ];
  targets.forEach(tg=>{
    const desired = Math.max(calcNextId_(tg.sheet,0,tg.base), tg.base);
    let row = t.rows.find(r => asText_(r[iKey])===tg.key);
    if (!row) t.rows.push([tg.key, String(desired)]);
    else {
      const cur = parseInt(asText_(row[iVal])||'0',10);
      if (isNaN(cur) || cur < desired) row[iVal] = String(desired);
    }
  });
  writeTable_(meta, t.headers, t.rows);
}
function nextId_(key, base) {
  ensureHeadersExist_();
  const SOURCE = {
    'NEXT_ID_PERSONA'      : SH.PERSONAS,
    'NEXT_ID_INSPECCION_AUT': SH.INSPECCION_AUT,
    'NEXT_ID_ASOCIACION_AUT': SH.ASOCIACION_AUT,
    'NEXT_ID_CAUCE_TOMERO'   : SH.CAUCE_TOMERO,
    'NEXT_ID_TELEFONO'       : SH.PERSONA_TELEFONOS
  };
  const meta = getSh_(SH.META);
  const t = readTable_(meta);
  const iKey = idx_(t.headers,'KEY');
  const iVal = idx_(t.headers,'VALUE');

  let row = t.rows.find(r => asText_(r[iKey]) === key);
  if (!row) {
    const start = calcNextId_(SOURCE[key]||'', 0, base);
    row = [key, String(start)];
    t.rows.push(row);
    writeTable_(meta, t.headers, t.rows);
  }
  let n = parseInt(asText_(row[iVal])||String(base),10);
  if (isNaN(n) || n < base) n = base;
  const curr = n;
  row[iVal] = String(n+1);
  writeTable_(meta, t.headers, t.rows);
  return String(curr);
}

/***********************
 * PERSONAS (búsquedas)
 ***********************/
function indexBy_(sheetName, keyCol) {
  const t = readTable_(getSh_(sheetName)); const iKey = idx_(t.headers,keyCol);
  const map = {};
  t.rows.forEach(r=>{
    const k = asText_(r[iKey]); if(!k) return;
    const o = {}; t.headers.forEach((h,j)=>o[h]=asText_(r[j]));
    map[k]=o;
  });
  return map;
}
function resolvePersonaId_(idOrName) {
  if (!idOrName) return '';
  const persons = indexBy_(SH.PERSONAS,'ID_PERSONA');
  if (persons[idOrName]) return idOrName;
  const t = readTable_(getSh_(SH.PERSONAS));
  const iId = idx_(t.headers,'ID_PERSONA'); const iNm = idx_(t.headers,'NOMBRE');
  for (const r of t.rows)
    if (asText_(r[iNm]).toUpperCase()===idOrName.toUpperCase())
      return asText_(r[iId]);
  return '';
}
function quickCreatePersonaByName_(nombre) {
  const sh = getSh_(SH.PERSONAS);
  const t  = readTable_(sh);
  const row = new Array(t.headers.length).fill('');
  row[idx_(t.headers,'ID_PERSONA')] = nextId_('NEXT_ID_PERSONA', ID_RANGES.PERSONA);
  row[idx_(t.headers,'NOMBRE')]      = nombre;
  t.rows.push(row);
  writeTable_(sh, t.headers, t.rows);
  return row[idx_(t.headers,'ID_PERSONA')];
}

/***********************
 * AUTORIDAD/RELACIONES
 ***********************/
function keyInspAsoc_(idIn, idAs){ return `${asText_(idIn)}__${asText_(idAs)}`; }
function computeAutoridadId_(idIns, idAs){
  const a = String(idIns).replace(/\D/g,'').padStart(6,'0');
  const b = String(idAs).replace(/\D/g,'').padStart(6,'0');
  return `9${a}${b}`;
}
function principalInspectorMap_() {
  const t = readTable_(getSh_(SH.INSPECCION_AUT));
  if (!t.rows.length) return {};
  const iIn = idx_(t.headers,'ID_INSPECCION');
  const iAs = idx_(t.headers,'FK_ID_ASOCIACION');
  const iPe = idx_(t.headers,'ID_PERSONA');
  const iRo = idx_(t.headers,'ROL');
  const iPr = idx_(t.headers,'PRINCIPAL');
  const map = {};
  t.rows.forEach(r=>{
    if (asText_(r[iRo]).toUpperCase()!=='INSPECTOR') return;
    if (!/^SI$/i.test(asText_(r[iPr]))) return;
    const idIns = asText_(r[iIn]); const idAs = asText_(r[iAs]);
    if (!idIns||!idAs) return;
    const k = keyInspAsoc_(idIns,idAs);
    map[k] = { idInspeccion:idIns, idAsociacion:idAs, idPersona:asText_(r[iPe]), idAutoridad:computeAutoridadId_(idIns,idAs) };
  });
  return map;
}
function writeFkAutoridadToInspecciones_(inspMap) {
  const sh = getSh_(SH.INSPECCIONES);
  const { headers, rows } = readTable_(sh);
  const iIn  = idx_(headers,'ID_INSPECCION');
  const iAs  = idx_(headers,'FK_ID_ASOCIACION');
  const iAut = idx_(headers,'FK_ID_AUTORIDAD');
  rows.forEach(r=>{
    const k = keyInspAsoc_(r[iIn], r[iAs]);
    r[iAut] = inspMap[k]?.idAutoridad || '';
  });
  sh.getRange(2,1,rows.length,headers.length).setValues(rows).setNumberFormat('@');
}
function writeFkAutoridadToCauces_(inspMap) {
  const sh = getSh_(SH.CAUCES);
  const { headers, rows } = readTable_(sh);
  const iIn  = idx_(headers,'FK_ID_INSPECCION');
  const iAs  = idx_(headers,'FK_ID_ASOCIACION');
  const iAut = idx_(headers,'FK_ID_AUTORIDAD');
  rows.forEach(r=>{
    const k = keyInspAsoc_(r[iIn], r[iAs]);
    r[iAut] = inspMap[k]?.idAutoridad || '';
  });
  sh.getRange(2,1,rows.length,headers.length).setValues(rows).setNumberFormat('@');
}
function rebuildAutoridadView_(inspMap) {
  const sh = getSh_(SH.AUTORIDAD_VIEW);
  const persons = indexBy_(SH.PERSONAS,'ID_PERSONA');
  const out = [];
  Object.values(inspMap).forEach(v=>{
    const p = persons[v.idPersona]||{};
    out.push([v.idAutoridad, v.idInspeccion, v.idAsociacion, p.NOMBRE||'', p.MAIL||'', p.TELEFONO||'']);
  });
  writeTable_(sh, HDR.AUTORIDAD_VIEW, out);
}
function writeTomeroPrincipalToCauces_() {
  const shCT = getSh_(SH.CAUCE_TOMERO);
  const shC  = getSh_(SH.CAUCES);
  const { headers: hCT, rows: rCT } = readTable_(shCT);
  const iCauce = idx_(hCT,'FK_ID_CAUCE');
  const iPers  = idx_(hCT,'ID_PERSONA');
  const iRol   = idx_(hCT,'ROL');
  const iPrinc = idx_(hCT,'PRINCIPAL');

  const m = {};
  rCT.forEach(r=>{
    if (asText_(r[iRol]).toUpperCase()!=='TOMERO') return;
    if (/^SI$/i.test(asText_(r[iPrinc]))) m[asText_(r[iCauce])] = asText_(r[iPers]);
  });

  const { headers: hC, rows: rC } = readTable_(shC);
  const iIdC = idx_(hC,'ID_CAUCE');
  const iFkT = idx_(hC,'FK_TOMERO_PRINCIPAL');
  rC.forEach(r=> r[iFkT] = m[asText_(r[iIdC])] || '' );
  shC.getRange(2,1,rC.length,hC.length).setValues(rC).setNumberFormat('@');
}
function syncFKsAll() {
  ensureHeadersExist_();
  const inp = principalInspectorMap_();
  writeFkAutoridadToInspecciones_(inp);
  writeFkAutoridadToCauces_(inp);
  rebuildAutoridadView_(inp);
  writeTomeroPrincipalToCauces_();
}

/***********************
 * ASIGNACIONES (UI-lite opcional)
 ***********************/
function assignInspector_(idInspeccion, idAsociacion, idPersona, principal, force){
  const sh = getSh_(SH.INSPECCION_AUT);
  const t = readTable_(sh); const h=t.headers; const r=t.rows;
  const iId=idx_(h,'ID_INSPECCION_AUT'), iIn=idx_(h,'ID_INSPECCION');
  const iAs=idx_(h,'FK_ID_ASOCIACION'), iPe=idx_(h,'ID_PERSONA');
  const iRo=idx_(h,'ROL'), iPr=idx_(h,'PRINCIPAL');

  const existsPrincipal = r.some(x =>
    asText_(x[iIn])===idInspeccion &&
    asText_(x[iAs])===idAsociacion &&
    asText_(x[iRo]).toUpperCase()==='INSPECTOR' &&
    /^SI$/i.test(asText_(x[iPr]))
  );
  if (principal && existsPrincipal && !force){
    return { needsConfirm:true, message:'Ya existe un INSPECTOR principal. ¿Reemplazarlo?' };
  }
  if (principal && existsPrincipal && force){
    r.forEach(x=>{
      if (asText_(x[iIn])===idInspeccion && asText_(x[iAs])===idAsociacion &&
          asText_(x[iRo]).toUpperCase()==='INSPECTOR' && /^SI$/i.test(asText_(x[iPr]))){
        x[iPr]='NO';
      }
    });
  }
  let updated=false;
  r.forEach(x=>{
    if (asText_(x[iIn])===idInspeccion && asText_(x[iAs])===idAsociacion &&
        asText_(x[iPe])===idPersona && asText_(x[iRo]).toUpperCase()==='INSPECTOR'){
      x[iPr] = principal?'SI':'NO'; updated=true;
    }
  });
  if (!updated){
    const idNew = nextId_('NEXT_ID_INSPECCION_AUT', ID_RANGES.INSPECCION_AUT);
    r.push([idNew, idInspeccion, idAsociacion, idPersona, 'INSPECTOR', principal?'SI':'NO', '', '']);
  }
  writeTable_(sh,h,r);
  syncFKsAll();
  return { ok:true };
}
function assignGerente_(idAsociacion, idPersona, principal, force){
  const sh = getSh_(SH.ASOCIACION_AUT);
  const t = readTable_(sh); const h=t.headers; const r=t.rows;
  const iId=idx_(h,'ID_ASOCIACION_AUT'), iAs=idx_(h,'FK_ID_ASOCIACION');
  const iPe=idx_(h,'ID_PERSONA'), iRo=idx_(h,'ROL'), iPr=idx_(h,'PRINCIPAL');

  const existsPrincipal = r.some(x =>
    asText_(x[iAs])===idAsociacion && asText_(x[iRo]).toUpperCase()==='GERENTE_TECNICO' && /^SI$/i.test(asText_(x[iPr]))
  );
  if (principal && existsPrincipal && !force){
    return { needsConfirm:true, message:'Ya existe un GERENTE TÉCNICO principal. ¿Reemplazarlo?' };
  }
  if (principal && existsPrincipal && force){
    r.forEach(x=>{
      if (asText_(x[iAs])===idAsociacion && asText_(x[iRo]).toUpperCase()==='GERENTE_TECNICO' && /^SI$/i.test(asText_(x[iPr]))){
        x[iPr]='NO';
      }
    });
  }
  let updated=false;
  r.forEach(x=>{
    if (asText_(x[iAs])===idAsociacion && asText_(x[iPe])===idPersona && asText_(x[iRo]).toUpperCase()==='GERENTE_TECNICO'){
      x[iPr] = principal?'SI':'NO'; updated=true;
    }
  });
  if (!updated){
    const idNew = nextId_('NEXT_ID_ASOCIACION_AUT', ID_RANGES.ASOCIACION_AUT);
    r.push([idNew, idAsociacion, idPersona, 'GERENTE_TECNICO', principal?'SI':'NO', '', '']);
  }
  writeTable_(sh,h,r);
  return { ok:true };
}
function assignTomero_(idCauce, idPersona, principal, force){
  const sh = getSh_(SH.CAUCE_TOMERO);
  const t = readTable_(sh); const h=t.headers; const r=t.rows;
  const iId=idx_(h,'ID_CAUCE_TOMERO'), iCa=idx_(h,'FK_ID_CAUCE');
  const iPe=idx_(h,'ID_PERSONA'), iRo=idx_(h,'ROL'), iPr=idx_(h,'PRINCIPAL');

  const existsPrincipal = r.some(x =>
    asText_(x[iCa])===idCauce && asText_(x[iRo]).toUpperCase()==='TOMERO' && /^SI$/i.test(asText_(x[iPr]))
  );
  if (principal && existsPrincipal && !force){
    return { needsConfirm:true, message:'Ya hay un TOMERO principal asignado. ¿Agregar otro principal?' };
  }
  let updated=false;
  r.forEach(x=>{
    if (asText_(x[iCa])===idCauce && asText_(x[iPe])===idPersona && asText_(x[iRo]).toUpperCase()==='TOMERO'){
      x[iPr] = principal?'SI':'NO'; updated=true;
    }
  });
  if (!updated){
    const idNew = nextId_('NEXT_ID_CAUCE_TOMERO', ID_RANGES.CAUCE_TOMERO);
    r.push([idNew, idCauce, idPersona, 'TOMERO', principal?'SI':'NO', '', '']);
  }
  writeTable_(sh,h,r);
  writeTomeroPrincipalToCauces_();
  return { ok:true };
}

/***********************
 * TELÉFONOS (concat → PERSONAS)
 ***********************/
function phonesConcatForPerson_(idPersona){
  const t = readTable_(getSh_(SH.PERSONA_TELEFONOS));
  const iP = idx_(t.headers,'ID_PERSONA');
  const iN = idx_(t.headers,'NUMERO');
  const iPr = idx_(t.headers,'PRINCIPAL');
  const iT = idx_(t.headers,'TIPO');

  const list = t.rows
    .filter(r=>asText_(r[iP])===idPersona && asText_(r[iN]))
    .sort((a,b)=>{
      const ap = /^TRUE$/i.test(asText_(a[iPr])) ? 0 : 1;
      const bp = /^TRUE$/i.test(asText_(b[iPr])) ? 0 : 1;
      if (ap!==bp) return ap-bp;
      return asText_(a[iT]).localeCompare(asText_(b[iT]));
    })
    .map(r=>asText_(r[iN]));
  return list.join(' / ');
}
function updatePersonaPhonesFromDetail_(){
  ensureHeadersExist_();
  const shP = getSh_(SH.PERSONAS);
  const t = readTable_(shP);
  const iId = idx_(t.headers,'ID_PERSONA');
  const iTel = idx_(t.headers,'TELEFONO');

  t.rows.forEach(r=>{
    const id = asText_(r[iId]);
    r[iTel] = phonesConcatForPerson_(id);
  });
  writeTable_(shP, t.headers, t.rows);
  const ui = tryGetUi_(); if (ui) ui.alert('PERSONAS.TELEFONO sincronizado desde PERSONA_TELEFONOS ✓');
}

/***********************
 * ROLES (semillas + helpers)
 ***********************/
function rolesDefaultSeeds_(){
  return [
    ['980000','NO_ASIGNADO','No asignado',0,'FALSE','FALSE','FALSE','TRUE',''],
    ['980001','INSPECTOR','Inspector',10,'TRUE','FALSE','FALSE','TRUE',''],
    ['980002','GERENTE_TECNICO','Gerente Técnico',20,'FALSE','TRUE','FALSE','TRUE',''],
    ['980003','TOMERO','Tomero',30,'FALSE','FALSE','TRUE','TRUE',''],
    ['980004','SUBDELEGADO','Subdelegado',40,'FALSE','FALSE','FALSE','TRUE',''],
    ['980005','CONSEJERO','Consejero',50,'FALSE','FALSE','FALSE','TRUE',''],
    ['980006','TECNICO','Técnico',60,'FALSE','FALSE','FALSE','TRUE',''],
    ['980007','ADMINISTRATIVO','Administrativo',70,'FALSE','FALSE','FALSE','TRUE',''],
    ['980008','INGENIERO','Ingeniero',80,'FALSE','FALSE','FALSE','TRUE',''],
    ['980009','CONTADOR','Contador',90,'FALSE','FALSE','FALSE','TRUE',''],
    ['980010','PRESIDENTE','Presidente',100,'FALSE','FALSE','FALSE','TRUE','']
  ];
}
function rolesInit_(){
  ensureHeadersExist_();
  const sh = getSh_(SH.ROLES);
  const t = readTable_(sh);
  if (!t.rows.length){
    writeTable_(sh, HDR.ROLES, rolesDefaultSeeds_());
  }else{
    writeTable_(sh, HDR.ROLES, t.rows);
  }
  styleHeader_(sh);
  const ui = tryGetUi_(); if (ui) ui.alert('ROLES inicializado/verificado ✓');
}
function rolesSeedForce_(){
  ensureHeadersExist_();
  const sh = getSh_(SH.ROLES);
  const cur = indexBy_(SH.ROLES,'ROL_KEY');
  const seeds = rolesDefaultSeeds_();
  seeds.forEach(row=>{
    const key=row[1];
    if (!cur[key]){
      const t = readTable_(sh);
      t.rows.push(row);
      writeTable_(sh, t.headers, t.rows);
    }
  });
  const ui = tryGetUi_(); if (ui) ui.alert('Semillas de ROLES forzadas/aplicadas ✓');
}

/***********************
 * ENCABEZADOS: auditoría básica
 ***********************/
function auditHeaders(){
  ensureHeadersExist_();
  const ss = SpreadsheetApp.getActive();
  const repName = 'AUDIT_ENCABEZADOS';
  const old = ss.getSheetByName(repName);
  if (old) ss.deleteSheet(old);
  const rep = ss.insertSheet(repName).setTabColor('#EF6C00');
  rep.getRange(1,1,rep.getMaxRows(),rep.getMaxColumns()).clearDataValidations();

  const out = [['SHEET','OK','HEADERS_ACTUALES','HEADERS_ESPERADOS','FALTAN','SOBRAN']];
  Object.keys(HDR).forEach(name=>{
    const sh = ss.getSheetByName(name);
    if(!sh){
      out.push([name,'NO','','',HDR[name].join(' | '),'(no existe)']);
      return;
    }
    const current = sh.getRange(1,1,1,Math.max(1,sh.getLastColumn()))
                       .getValues()[0].map(v=>String(v||'').trim());
    const expected = HDR[name];
    const faltan = expected.filter(h=>!current.includes(h));
    const sobran = current.filter(h=>!expected.includes(h));
    out.push([
      name,
      (faltan.length===0 && sobran.length===0) ? 'SI' : 'NO',
      current.join(' | '),
      expected.join(' | '),
      faltan.join(' | '),
      sobran.join(' | ')
    ]);
  });
  writeTable_(rep, out[0], out.slice(1));
  styleHeader_(rep);
}

function fixHeadersSafe(){
  ensureHeadersExist_();
  const ss=SpreadsheetApp.getActive();
  Object.keys(HDR).forEach(name=>{
    const sh=ss.getSheetByName(name); if(!sh) return;
    const expected=HDR[name];
    const data=sh.getDataRange().getValues();
    const curr=data[0].map(v=>String(v||'').trim());
    const headers=curr.slice();
    const faltan=expected.filter(h=>!headers.includes(h));
    if (faltan.length){
      const extra=Array(faltan.length).fill('');
      data[0]=headers.concat(faltan);
      for(let i=1;i<data.length;i++) data[i]=data[i].concat(extra);
      writeTable_(sh, data[0], data.slice(1));
    }else{
      writeTable_(sh, headers, data.slice(1));
    }
    styleHeader_(sh);
  });
}

/**
 * Reordena INSPECCIONES al orden canónico (HDR.INSPECCIONES) en bloques.
 */
function reorderInspecciones(){
  return reorderInspeccionesFast_(2000);
}
function reorderInspeccionesFast_(CHUNK){
  const ss   = SpreadsheetApp.getActive();
  const src  = getSh_(SH.INSPECCIONES);
  const need = HDR.INSPECCIONES.slice();

  const lastCol = src.getLastColumn();
  const lastRow = src.getLastRow();
  if (lastCol === 0) return;

  const header = src.getRange(1,1,1,lastCol)
                     .getValues()[0]
                     .map(v => String(v||'').trim());

  let already = true;
  for (let i=0; i<need.length; i++){
    if ((header[i]||'') !== need[i]) { already = false; break; }
  }
  if (already) return;

  const map    = need.map(h => header.indexOf(h));
  const extraIdx = header.map((_,i)=>i).filter(i => !map.includes(i));
  const newHeader= need.concat(extraIdx.map(i => header[i]));

  const tmpName = SH.INSPECCIONES + '__NEW';
  const tmpOld  = ss.getSheetByName(tmpName);
  if (tmpOld) ss.deleteSheet(tmpOld);
  const tmp = ss.insertSheet(tmpName);
  tmp.getRange(1,1,1,newHeader.length).setValues([newHeader]);

  if (lastRow > 1){
    let srcRow = 2;
    let dstRow = 2;
    while (srcRow <= lastRow){
      const n = Math.min(CHUNK, lastRow - srcRow + 1);
      const chunk = src.getRange(srcRow, 1, n, lastCol).getValues();
      const newRows = chunk.map(row => {
        const base  = need.map((_,k)=> map[k] >= 0 ? row[ map[k] ] : '' );
        const extra = extraIdx.map(i => row[i]);
        return base.concat(extra);
      });
      tmp.getRange(dstRow, 1, n, newHeader.length).setValues(newRows);
      srcRow += n;
      dstRow += n;
      SpreadsheetApp.flush();
    }
  }

  const color = src.getTabColor();
  ss.deleteSheet(src);
  tmp.setName(SH.INSPECCIONES);
  if (color) tmp.setTabColor(color);
  styleHeader_(tmp);
}

/***********************
 * INICIALIZAR MODELO Y SECUENCIA COMPLETA
 ***********************/
function initModel(){
  ensureHeadersExist_();
  rolesInit_();
  normalizeMetaCounters();
  reorderInspecciones();
  styleAllHeaders();
  colorizeTabs();
}
function initModel_step1(){
  ensureHeadersExist_();
  rolesInit_();
  normalizeMetaCounters();
}
function initModel_step2(){
  reorderInspecciones();
  styleAllHeaders();
  colorizeTabs();
}
function runFullSetup(){
  initModel();
  syncFKsAll();
  writePhonesToPersonas_();
  auditHeaders();
  fixHeadersSafe();
  styleAllHeaders();
  colorizeTabs();
  const ui=tryGetUi_(); if(ui) ui.alert('Secuencia completa lista ✅');
}

/***********************
 * PERSONAS (búsquedas)
 ***********************/
function indexBy_(sheetName, keyCol) {
  const t = readTable_(getSh_(sheetName)); const iKey = idx_(t.headers,keyCol);
  const map = {};
  t.rows.forEach(r=>{
    const k = asText_(r[iKey]); if(!k) return;
    const o = {}; t.headers.forEach((h,j)=>o[h]=asText_(r[j]));
    map[k]=o;
  });
  return map;
}
function resolvePersonaId_(idOrName) {
  if (!idOrName) return '';
  const persons = indexBy_(SH.PERSONAS,'ID_PERSONA');
  if (persons[idOrName]) return idOrName;
  const t = readTable_(getSh_(SH.PERSONAS));
  const iId = idx_(t.headers,'ID_PERSONA'); const iNm = idx_(t.headers,'NOMBRE');
  for (const r of t.rows)
    if (asText_(r[iNm]).toUpperCase()===idOrName.toUpperCase())
      return asText_(r[iId]);
  return '';
}
function quickCreatePersonaByName_(nombre) {
  const sh = getSh_(SH.PERSONAS);
  const t  = readTable_(sh);
  const row = new Array(t.headers.length).fill('');
  row[idx_(t.headers,'ID_PERSONA')] = nextId_('NEXT_ID_PERSONA', ID_RANGES.PERSONA);
  row[idx_(t.headers,'NOMBRE')]      = nombre;
  t.rows.push(row);
  writeTable_(sh, t.headers, t.rows);
  return row[idx_(t.headers,'ID_PERSONA')];
}

/***********************
 * AUTORIDAD/RELACIONES
 ***********************/
function keyInspAsoc_(idIn, idAs){ return `${asText_(idIn)}__${asText_(idAs)}`; }
function computeAutoridadId_(idIns, idAs){
  const a = String(idIns).replace(/\D/g,'').padStart(6,'0');
  const b = String(idAs).replace(/\D/g,'').padStart(6,'0');
  return `9${a}${b}`;
}
function principalInspectorMap_() {
  const t = readTable_(getSh_(SH.INSPECCION_AUT));
  if (!t.rows.length) return {};
  const iIn = idx_(t.headers,'ID_INSPECCION');
  const iAs = idx_(t.headers,'FK_ID_ASOCIACION');
  const iPe = idx_(t.headers,'ID_PERSONA');
  const iRo = idx_(t.headers,'ROL');
  const iPr = idx_(t.headers,'PRINCIPAL');
  const map = {};
  t.rows.forEach(r=>{
    if (asText_(r[iRo]).toUpperCase()!=='INSPECTOR') return;
    if (!/^SI$/i.test(asText_(r[iPr]))) return;
    const idIns = asText_(r[iIn]); const idAs = asText_(r[iAs]);
    if (!idIns||!idAs) return;
    const k = keyInspAsoc_(idIns,idAs);
    map[k] = { idInspeccion:idIns, idAsociacion:idAs, idPersona:asText_(r[iPe]), idAutoridad:computeAutoridadId_(idIns,idAs) };
  });
  return map;
}
function writeFkAutoridadToInspecciones_(inspMap) {
  const sh = getSh_(SH.INSPECCIONES);
  const { headers, rows } = readTable_(sh);
  const iIn  = idx_(headers,'ID_INSPECCION');
  const iAs  = idx_(headers,'FK_ID_ASOCIACION');
  const iAut = idx_(headers,'FK_ID_AUTORIDAD');
  rows.forEach(r=>{
    const k = keyInspAsoc_(r[iIn], r[iAs]);
    r[iAut] = inspMap[k]?.idAutoridad || '';
  });
  sh.getRange(2,1,rows.length,headers.length).setValues(rows).setNumberFormat('@');
}
function writeFkAutoridadToCauces_(inspMap) {
  const sh = getSh_(SH.CAUCES);
  const { headers, rows } = readTable_(sh);
  const iIn  = idx_(headers,'FK_ID_INSPECCION');
  const iAs  = idx_(headers,'FK_ID_ASOCIACION');
  const iAut = idx_(headers,'FK_ID_AUTORIDAD');
  rows.forEach(r=>{
    const k = keyInspAsoc_(r[iIn], r[iAs]);
    r[iAut] = inspMap[k]?.idAutoridad || '';
  });
  sh.getRange(2,1,rows.length,headers.length).setValues(rows).setNumberFormat('@');
}
function rebuildAutoridadView_(inspMap) {
  const sh = getSh_(SH.AUTORIDAD_VIEW);
  const persons = indexBy_(SH.PERSONAS,'ID_PERSONA');
  const out = [];
  Object.values(inspMap).forEach(v=>{
    const p = persons[v.idPersona]||{};
    out.push([v.idAutoridad, v.idInspeccion, v.idAsociacion, p.NOMBRE||'', p.MAIL||'', p.TELEFONO||'']);
  });
  writeTable_(sh, HDR.AUTORIDAD_VIEW, out);
}
function writeTomeroPrincipalToCauces_() {
  const shCT = getSh_(SH.CAUCE_TOMERO);
  const shC  = getSh_(SH.CAUCES);
  const { headers: hCT, rows: rCT } = readTable_(shCT);
  const iCauce = idx_(hCT,'FK_ID_CAUCE');
  const iPers  = idx_(hCT,'ID_PERSONA');
  const iRol   = idx_(hCT,'ROL');
  const iPrinc = idx_(hCT,'PRINCIPAL');

  const m = {};
  rCT.forEach(r=>{
    if (asText_(r[iRol]).toUpperCase()!=='TOMERO') return;
    if (/^SI$/i.test(asText_(r[iPrinc]))) m[asText_(r[iCauce])] = asText_(r[iPers]);
  });

  const { headers: hC, rows: rC } = readTable_(shC);
  const iIdC = idx_(hC,'ID_CAUCE');
  const iFkT = idx_(hC,'FK_TOMERO_PRINCIPAL');
  rC.forEach(r=> r[iFkT] = m[asText_(r[iIdC])] || '' );
  shC.getRange(2,1,rC.length,hC.length).setValues(rC).setNumberFormat('@');
}
function syncFKsAll() {
  ensureHeadersExist_();
  const inp = principalInspectorMap_();
  writeFkAutoridadToInspecciones_(inp);
  writeFkAutoridadToCauces_(inp);
  rebuildAutoridadView_(inp);
  writeTomeroPrincipalToCauces_();
}

/***********************
 * ASIGNACIONES (UI-lite opcional)
 ***********************/
function assignInspector_(idInspeccion, idAsociacion, idPersona, principal, force){
  const sh = getSh_(SH.INSPECCION_AUT);
  const t = readTable_(sh); const h=t.headers; const r=t.rows;
  const iId=idx_(h,'ID_INSPECCION_AUT'), iIn=idx_(h,'ID_INSPECCION');
  const iAs=idx_(h,'FK_ID_ASOCIACION'), iPe=idx_(h,'ID_PERSONA');
  const iRo=idx_(h,'ROL'), iPr=idx_(h,'PRINCIPAL');

  const existsPrincipal = r.some(x =>
    asText_(x[iIn])===idInspeccion &&
    asText_(x[iAs])===idAsociacion &&
    asText_(x[iRo]).toUpperCase()==='INSPECTOR' &&
    /^SI$/i.test(asText_(x[iPr]))
  );
  if (principal && existsPrincipal && !force){
    return { needsConfirm:true, message:'Ya existe un INSPECTOR principal. ¿Reemplazarlo?' };
  }
  if (principal && existsPrincipal && force){
    r.forEach(x=>{
      if (asText_(x[iIn])===idInspeccion && asText_(x[iAs])===idAsociacion &&
          asText_(x[iRo]).toUpperCase()==='INSPECTOR' && /^SI$/i.test(asText_(x[iPr]))){
        x[iPr]='NO';
      }
    });
  }
  let updated=false;
  r.forEach(x=>{
    if (asText_(x[iIn])===idInspeccion && asText_(x[iAs])===idAsociacion &&
        asText_(x[iPe])===idPersona && asText_(x[iRo]).toUpperCase()==='INSPECTOR'){
      x[iPr] = principal?'SI':'NO'; updated=true;
    }
  });
  if (!updated){
    const idNew = nextId_('NEXT_ID_INSPECCION_AUT', ID_RANGES.INSPECCION_AUT);
    r.push([idNew, idInspeccion, idAsociacion, idPersona, 'INSPECTOR', principal?'SI':'NO', '', '']);
  }
  writeTable_(sh,h,r);
  syncFKsAll();
  return { ok:true };
}
function assignGerente_(idAsociacion, idPersona, principal, force){
  const sh = getSh_(SH.ASOCIACION_AUT);
  const t = readTable_(sh); const h=t.headers; const r=t.rows;
  const iId=idx_(h,'ID_ASOCIACION_AUT'), iAs=idx_(h,'FK_ID_ASOCIACION');
  const iPe=idx_(h,'ID_PERSONA'), iRo=idx_(h,'ROL'), iPr=idx_(h,'PRINCIPAL');

  const existsPrincipal = r.some(x =>
    asText_(x[iAs])===idAsociacion && asText_(x[iRo]).toUpperCase()==='GERENTE_TECNICO' && /^SI$/i.test(asText_(x[iPr]))
  );
  if (principal && existsPrincipal && !force){
    return { needsConfirm:true, message:'Ya existe un GERENTE TÉCNICO principal. ¿Reemplazarlo?' };
  }
  if (principal && existsPrincipal && force){
    r.forEach(x=>{
      if (asText_(x[iAs])===idAsociacion && asText_(x[iRo]).toUpperCase()==='GERENTE_TECNICO' && /^SI$/i.test(asText_(x[iPr]))){
        x[iPr]='NO';
      }
    });
  }
  let updated=false;
  r.forEach(x=>{
    if (asText_(x[iAs])===idAsociacion && asText_(x[iPe])===idPersona && asText_(x[iRo]).toUpperCase()==='GERENTE_TECNICO'){
      x[iPr] = principal?'SI':'NO'; updated=true;
    }
  });
  if (!updated){
    const idNew = nextId_('NEXT_ID_ASOCIACION_AUT', ID_RANGES.ASOCIACION_AUT);
    r.push([idNew, idAsociacion, idPersona, 'GERENTE_TECNICO', principal?'SI':'NO', '', '']);
  }
  writeTable_(sh,h,r);
  return { ok:true };
}
function assignTomero_(idCauce, idPersona, principal, force){
  const sh = getSh_(SH.CAUCE_TOMERO);
  const t = readTable_(sh); const h=t.headers; const r=t.rows;
  const iId=idx_(h,'ID_CAUCE_TOMERO'), iCa=idx_(h,'FK_ID_CAUCE');
  const iPe=idx_(h,'ID_PERSONA'), iRo=idx_(h,'ROL'), iPr=idx_(h,'PRINCIPAL');

  const existsPrincipal = r.some(x =>
    asText_(x[iCa])===idCauce && asText_(x[iRo]).toUpperCase()==='TOMERO' && /^SI$/i.test(asText_(x[iPr]))
  );
  if (principal && existsPrincipal && !force){
    return { needsConfirm:true, message:'Ya hay un TOMERO principal asignado. ¿Agregar otro principal?' };
  }
  let updated=false;
  r.forEach(x=>{
    if (asText_(x[iCa])===idCauce && asText_(x[iPe])===idPersona && asText_(x[iRo]).toUpperCase()==='TOMERO'){
      x[iPr] = principal?'SI':'NO'; updated=true;
    }
  });
  if (!updated){
    const idNew = nextId_('NEXT_ID_CAUCE_TOMERO', ID_RANGES.CAUCE_TOMERO);
    r.push([idNew, idCauce, idPersona, 'TOMERO', principal?'SI':'NO', '', '']);
  }
  writeTable_(sh,h,r);
  writeTomeroPrincipalToCauces_();
  return { ok:true };
}

/***********************
 * TELÉFONOS (concat → PERSONAS)
 ***********************/
function phonesConcatForPerson_(idPersona){
  const t = readTable_(getSh_(SH.PERSONA_TELEFONOS));
  const iP = idx_(t.headers,'ID_PERSONA');
  const iN = idx_(t.headers,'NUMERO');
  const iPr = idx_(t.headers,'PRINCIPAL');
  const iT = idx_(t.headers,'TIPO');

  const list = t.rows
    .filter(r=>asText_(r[iP])===idPersona && asText_(r[iN]))
    .sort((a,b)=>{
      const ap = /^TRUE$/i.test(asText_(a[iPr])) ? 0 : 1;
      const bp = /^TRUE$/i.test(asText_(b[iPr])) ? 0 : 1;
      if (ap!==bp) return ap-bp;
      return asText_(a[iT]).localeCompare(asText_(b[iT]));
    })
    .map(r=>asText_(r[iN]));
  return list.join(' / ');
}
function updatePersonaPhonesFromDetail_(){
  ensureHeadersExist_();
  const shP = getSh_(SH.PERSONAS);
  const t = readTable_(shP);
  const iId = idx_(t.headers,'ID_PERSONA');
  const iTel = idx_(t.headers,'TELEFONO');

  t.rows.forEach(r=>{
    const id = asText_(r[iId]);
    r[iTel] = phonesConcatForPerson_(id);
  });
  writeTable_(shP, t.headers, t.rows);
  const ui = tryGetUi_(); if (ui) ui.alert('PERSONAS.TELEFONO sincronizado desde PERSONA_TELEFONOS ✓');
}

/***********************
 * ROLES (semillas + helpers)
 ***********************/
function rolesDefaultSeeds_(){
  return [
    ['980000','NO_ASIGNADO','No asignado',0,'FALSE','FALSE','FALSE','TRUE',''],
    ['980001','INSPECTOR','Inspector',10,'TRUE','FALSE','FALSE','TRUE',''],
    ['980002','GERENTE_TECNICO','Gerente Técnico',20,'FALSE','TRUE','FALSE','TRUE',''],
    ['980003','TOMERO','Tomero',30,'FALSE','FALSE','TRUE','TRUE',''],
    ['980004','SUBDELEGADO','Subdelegado',40,'FALSE','FALSE','FALSE','TRUE',''],
    ['980005','CONSEJERO','Consejero',50,'FALSE','FALSE','FALSE','TRUE',''],
    ['980006','TECNICO','Técnico',60,'FALSE','FALSE','FALSE','TRUE',''],
    ['980007','ADMINISTRATIVO','Administrativo',70,'FALSE','FALSE','FALSE','TRUE',''],
    ['980008','INGENIERO','Ingeniero',80,'FALSE','FALSE','FALSE','TRUE',''],
    ['980009','CONTADOR','Contador',90,'FALSE','FALSE','FALSE','TRUE',''],
    ['980010','PRESIDENTE','Presidente',100,'FALSE','FALSE','FALSE','TRUE','']
  ];
}
function rolesInit_(){
  ensureHeadersExist_();
  const sh = getSh_(SH.ROLES);
  const t = readTable_(sh);
  if (!t.rows.length){
    writeTable_(sh, HDR.ROLES, rolesDefaultSeeds_());
  }else{
    writeTable_(sh, HDR.ROLES, t.rows);
  }
  styleHeader_(sh);
  const ui = tryGetUi_(); if (ui) ui.alert('ROLES inicializado/verificado ✓');
}
function rolesSeedForce_(){
  ensureHeadersExist_();
  const sh = getSh_(SH.ROLES);
  const cur = indexBy_(SH.ROLES,'ROL_KEY');
  const seeds = rolesDefaultSeeds_();
  seeds.forEach(row=>{
    const key=row[1];
    if (!cur[key]){
      const t = readTable_(sh);
      t.rows.push(row);
      writeTable_(sh, t.headers, t.rows);
    }
  });
  const ui = tryGetUi_(); if (ui) ui.alert('Semillas de ROLES forzadas/aplicadas ✓');
}

/***********************
 * ENCABEZADOS: auditoría básica
 ***********************/
function auditHeaders(){
  ensureHeadersExist_();
  const ss = SpreadsheetApp.getActive();
  const repName = 'AUDIT_ENCABEZADOS';
  const old = ss.getSheetByName(repName);
  if (old) ss.deleteSheet(old);
  const rep = ss.insertSheet(repName).setTabColor('#EF6C00');
  rep.getRange(1,1,rep.getMaxRows(),rep.getMaxColumns()).clearDataValidations();

  const out = [['SHEET','OK','HEADERS_ACTUALES','HEADERS_ESPERADOS','FALTAN','SOBRAN']];
  Object.keys(HDR).forEach(name=>{
    const sh = ss.getSheetByName(name);
    if(!sh){
      out.push([name,'NO','','',HDR[name].join(' | '),'(no existe)']);
      return;
    }
    const current = sh.getRange(1,1,1,Math.max(1,sh.getLastColumn()))
                       .getValues()[0].map(v=>String(v||'').trim());
    const expected = HDR[name];
    const faltan = expected.filter(h=>!current.includes(h));
    const sobran = current.filter(h=>!expected.includes(h));
    out.push([
      name,
      (faltan.length===0 && sobran.length===0) ? 'SI' : 'NO',
      current.join(' | '),
      expected.join(' | '),
      faltan.join(' | '),
      sobran.join(' | ')
    ]);
  });
  writeTable_(rep, out[0], out.slice(1));
  styleHeader_(rep);
}

function fixHeadersSafe(){
  ensureHeadersExist_();
  const ss=SpreadsheetApp.getActive();
  Object.keys(HDR).forEach(name=>{
    const sh=ss.getSheetByName(name); if(!sh) return;
    const expected=HDR[name];
    const data=sh.getDataRange().getValues();
    const curr=data[0].map(v=>String(v||'').trim());
    const headers=curr.slice();
    const faltan=expected.filter(h=>!headers.includes(h));
    if (faltan.length){
      const extra=Array(faltan.length).fill('');
      data[0]=headers.concat(faltan);
      for(let i=1;i<data.length;i++) data[i]=data[i].concat(extra);
      writeTable_(sh, data[0], data.slice(1));
    }else{
      writeTable_(sh, headers, data.slice(1));
    }
    styleHeader_(sh);
  });
}

/**
 * Reordena INSPECCIONES al orden canónico (HDR.INSPECCIONES) en bloques.
 */
function reorderInspecciones(){
  return reorderInspeccionesFast_(2000);
}
function reorderInspeccionesFast_(CHUNK){
  const ss   = SpreadsheetApp.getActive();
  const src  = getSh_(SH.INSPECCIONES);
  const need = HDR.INSPECCIONES.slice();

  const lastCol = src.getLastColumn();
  const lastRow = src.getLastRow();
  if (lastCol === 0) return;

  const header = src.getRange(1,1,1,lastCol)
                     .getValues()[0]
                     .map(v => String(v||'').trim());

  let already = true;
  for (let i=0; i<need.length; i++){
    if ((header[i]||'') !== need[i]) { already = false; break; }
  }
  if (already) return;

  const map    = need.map(h => header.indexOf(h));
  const extraIdx = header.map((_,i)=>i).filter(i => !map.includes(i));
  const newHeader= need.concat(extraIdx.map(i => header[i]));

  const tmpName = SH.INSPECCIONES + '__NEW';
  const tmpOld  = ss.getSheetByName(tmpName);
  if (tmpOld) ss.deleteSheet(tmpOld);
  const tmp = ss.insertSheet(tmpName);
  tmp.getRange(1,1,1,newHeader.length).setValues([newHeader]);

  if (lastRow > 1){
    let srcRow = 2;
    let dstRow = 2;
    while (srcRow <= lastRow){
      const n = Math.min(CHUNK, lastRow - srcRow + 1);
      const chunk = src.getRange(srcRow, 1, n, lastCol).getValues();
      const newRows = chunk.map(row => {
        const base  = need.map((_,k)=> map[k] >= 0 ? row[ map[k] ] : '' );
        const extra = extraIdx.map(i => row[i]);
        return base.concat(extra);
      });
      tmp.getRange(dstRow, 1, n, newHeader.length).setValues(newRows);
      srcRow += n;
      dstRow += n;
      SpreadsheetApp.flush();
    }
  }

  const color = src.getTabColor();
  ss.deleteSheet(src);
  tmp.setName(SH.INSPECCIONES);
  if (color) tmp.setTabColor(color);
  styleHeader_(tmp);
}

/***********************
 * INICIALIZAR MODELO Y SECUENCIA COMPLETA
 ***********************/
function initModel(){
  ensureHeadersExist_();
  rolesInit_();
  normalizeMetaCounters();
  reorderInspecciones();
  styleAllHeaders();
  colorizeTabs();
}
function initModel_step1(){
  ensureHeadersExist_();
  rolesInit_();
  normalizeMetaCounters();
}
function initModel_step2(){
  reorderInspecciones();
  styleAllHeaders();
  colorizeTabs();
}
function runFullSetup(){
  initModel();
  syncFKsAll();
  writePhonesToPersonas_();
  auditHeaders();
  fixHeadersSafe();
  styleAllHeaders();
  colorizeTabs();
  const ui=tryGetUi_(); if(ui) ui.alert('Secuencia completa lista ✅');
}

/***********************
 * PERSONAS (búsquedas)
 ***********************/
function indexBy_(sheetName, keyCol) {
  const t = readTable_(getSh_(sheetName)); const iKey = idx_(t.headers,keyCol);
  const map = {};
  t.rows.forEach(r=>{
    const k = asText_(r[iKey]); if(!k) return;
    const o = {}; t.headers.forEach((h,j)=>o[h]=asText_(r[j]));
    map[k]=o;
  });
  return map;
}
function resolvePersonaId_(idOrName) {
  if (!idOrName) return '';
  const persons = indexBy_(SH.PERSONAS,'ID_PERSONA');
  if (persons[idOrName]) return idOrName;
  const t = readTable_(getSh_(SH.PERSONAS));
  const iId = idx_(t.headers,'ID_PERSONA'); const iNm = idx_(t.headers,'NOMBRE');
  for (const r of t.rows)
    if (asText_(r[iNm]).toUpperCase()===idOrName.toUpperCase())
      return asText_(r[iId]);
  return '';
}
function quickCreatePersonaByName_(nombre) {
  const sh = getSh_(SH.PERSONAS);
  const t  = readTable_(sh);
  const row = new Array(t.headers.length).fill('');
  row[idx_(t.headers,'ID_PERSONA')] = nextId_('NEXT_ID_PERSONA', ID_RANGES.PERSONA);
  row[idx_(t.headers,'NOMBRE')]      = nombre;
  t.rows.push(row);
  writeTable_(sh, t.headers, t.rows);
  return row[idx_(t.headers,'ID_PERSONA')];
}

/***********************
 * AUTORIDAD/RELACIONES
 ***********************/
function keyInspAsoc_(idIn, idAs){ return `${asText_(idIn)}__${asText_(idAs)}`; }
function computeAutoridadId_(idIns, idAs){
  const a = String(idIns).replace(/\D/g,'').padStart(6,'0');
  const b = String(idAs).replace(/\D/g,'').padStart(6,'0');
  return `9${a}${b}`;
}
function principalInspectorMap_() {
  const t = readTable_(getSh_(SH.INSPECCION_AUT));
  if (!t.rows.length) return {};
  const iIn = idx_(t.headers,'ID_INSPECCION');
  const iAs = idx_(t.headers,'FK_ID_ASOCIACION');
  const iPe = idx_(t.headers,'ID_PERSONA');
  const iRo = idx_(t.headers,'ROL');
  const iPr = idx_(t.headers,'PRINCIPAL');
  const map = {};
  t.rows.forEach(r=>{
    if (asText_(r[iRo]).toUpperCase()!=='INSPECTOR') return;
    if (!/^SI$/i.test(asText_(r[iPr]))) return;
    const idIns = asText_(r[iIn]); const idAs = asText_(r[iAs]);
    if (!idIns||!idAs) return;
    const k = keyInspAsoc_(idIns,idAs);
    map[k] = { idInspeccion:idIns, idAsociacion:idAs, idPersona:asText_(r[iPe]), idAutoridad:computeAutoridadId_(idIns,idAs) };
  });
  return map;
}
function writeFkAutoridadToInspecciones_(inspMap) {
  const sh = getSh_(SH.INSPECCIONES);
  const { headers, rows } = readTable_(sh);
  const iIn  = idx_(headers,'ID_INSPECCION');
  const iAs  = idx_(headers,'FK_ID_ASOCIACION');
  const iAut = idx_(headers,'FK_ID_AUTORIDAD');
  rows.forEach(r=>{
    const k = keyInspAsoc_(r[iIn], r[iAs]);
    r[iAut] = inspMap[k]?.idAutoridad || '';
  });
  sh.getRange(2,1,rows.length,headers.length).setValues(rows).setNumberFormat('@');
}
function writeFkAutoridadToCauces_(inspMap) {
  const sh = getSh_(SH.CAUCES);
  const { headers, rows } = readTable_(sh);
  const iIn  = idx_(headers,'FK_ID_INSPECCION');
  const iAs  = idx_(headers,'FK_ID_ASOCIACION');
  const iAut = idx_(headers,'FK_ID_AUTORIDAD');
  rows.forEach(r=>{
    const k = keyInspAsoc_(r[iIn], r[iAs]);
    r[iAut] = inspMap[k]?.idAutoridad || '';
  });
  sh.getRange(2,1,rows.length,headers.length).setValues(rows).setNumberFormat('@');
}
function rebuildAutoridadView_(inspMap) {
  const sh = getSh_(SH.AUTORIDAD_VIEW);
  const persons = indexBy_(SH.PERSONAS,'ID_PERSONA');
  const out = [];
  Object.values(inspMap).forEach(v=>{
    const p = persons[v.idPersona]||{};
    out.push([v.idAutoridad, v.idInspeccion, v.idAsociacion, p.NOMBRE||'', p.MAIL||'', p.TELEFONO||'']);
  });
  writeTable_(sh, HDR.AUTORIDAD_VIEW, out);
}
function writeTomeroPrincipalToCauces_() {
  const shCT = getSh_(SH.CAUCE_TOMERO);
  const shC  = getSh_(SH.CAUCES);
  const { headers: hCT, rows: rCT } = readTable_(shCT);
  const iCauce = idx_(hCT,'FK_ID_CAUCE');
  const iPers  = idx_(hCT,'ID_PERSONA');
  const iRol   = idx_(hCT,'ROL');
  const iPrinc = idx_(hCT,'PRINCIPAL');

  const m = {};
  rCT.forEach(r=>{
    if (asText_(r[iRol]).toUpperCase()!=='TOMERO') return;
    if (/^SI$/i.test(asText_(r[iPrinc]))) m[asText_(r[iCauce])] = asText_(r[iPers]);
  });

  const { headers: hC, rows: rC } = readTable_(shC);
  const iIdC = idx_(hC,'ID_CAUCE');
  const iFkT = idx_(hC,'FK_TOMERO_PRINCIPAL');
  rC.forEach(r=> r[iFkT] = m[asText_(r[iIdC])] || '' );
  shC.getRange(2,1,rC.length,hC.length).setValues(rC).setNumberFormat('@');
}
function syncFKsAll() {
  ensureHeadersExist_();
  const inp = principalInspectorMap_();
  writeFkAutoridadToInspecciones_(inp);
  writeFkAutoridadToCauces_(inp);
  rebuildAutoridadView_(inp);
  writeTomeroPrincipalToCauces_();
}

/***********************
 * ASIGNACIONES (UI-lite opcional)
 ***********************/
function assignInspector_(idInspeccion, idAsociacion, idPersona, principal, force){
  const sh = getSh_(SH.INSPECCION_AUT);
  const t = readTable_(sh); const h=t.headers; const r=t.rows;
  const iId=idx_(h,'ID_INSPECCION_AUT'), iIn=idx_(h,'ID_INSPECCION');
  const iAs=idx_(h,'FK_ID_ASOCIACION'), iPe=idx_(h,'ID_PERSONA');
  const iRo=idx_(h,'ROL'), iPr=idx_(h,'PRINCIPAL');

  const existsPrincipal = r.some(x =>
    asText_(x[iIn])===idInspeccion &&
    asText_(x[iAs])===idAsociacion &&
    asText_(x[iRo]).toUpperCase()==='INSPECTOR' &&
    /^SI$/i.test(asText_(x[iPr]))
  );
  if (principal && existsPrincipal && !force){
    return { needsConfirm:true, message:'Ya existe un INSPECTOR principal. ¿Reemplazarlo?' };
  }
  if (principal && existsPrincipal && force){
    r.forEach(x=>{
      if (asText_(x[iIn])===idInspeccion && asText_(x[iAs])===idAsociacion &&
          asText_(x[iRo]).toUpperCase()==='INSPECTOR' && /^SI$/i.test(asText_(x[iPr]))){
        x[iPr]='NO';
      }
    });
  }
  let updated=false;
  r.forEach(x=>{
    if (asText_(x[iIn])===idInspeccion && asText_(x[iAs])===idAsociacion &&
        asText_(x[iPe])===idPersona && asText_(x[iRo]).toUpperCase()==='INSPECTOR'){
      x[iPr] = principal?'SI':'NO'; updated=true;
    }
  });
  if (!updated){
    const idNew = nextId_('NEXT_ID_INSPECCION_AUT', ID_RANGES.INSPECCION_AUT);
    r.push([idNew, idInspeccion, idAsociacion, idPersona, 'INSPECTOR', principal?'SI':'NO', '', '']);
  }
  writeTable_(sh,h,r);
  syncFKsAll();
  return { ok:true };
}
function assignGerente_(idAsociacion, idPersona, principal, force){
  const sh = getSh_(SH.ASOCIACION_AUT);
  const t = readTable_(sh); const h=t.headers; const r=t.rows;
  const iId=idx_(h,'ID_ASOCIACION_AUT'), iAs=idx_(h,'FK_ID_ASOCIACION');
  const iPe=idx_(h,'ID_PERSONA'), iRo=idx_(h,'ROL'), iPr=idx_(h,'PRINCIPAL');

  const existsPrincipal = r.some(x =>
    asText_(x[iAs])===idAsociacion && asText_(x[iRo]).toUpperCase()==='GERENTE_TECNICO' && /^SI$/i.test(asText_(x[iPr]))
  );
  if (principal && existsPrincipal && !force){
    return { needsConfirm:true, message:'Ya existe un GERENTE TÉCNICO principal. ¿Reemplazarlo?' };
  }
  if (principal && existsPrincipal && force){
    r.forEach(x=>{
      if (asText_(x[iAs])===idAsociacion && asText_(x[iRo]).toUpperCase()==='GERENTE_TECNICO' && /^SI$/i.test(asText_(x[iPr]))){
        x[iPr]='NO';
      }
    });
  }
  let updated=false;
  r.forEach(x=>{
    if (asText_(x[iAs])===idAsociacion && asText_(x[iPe])===idPersona && asText_(x[iRo]).toUpperCase()==='GERENTE_TECNICO'){
      x[iPr] = principal?'SI':'NO'; updated=true;
    }
  });
  if (!updated){
    const idNew = nextId_('NEXT_ID_ASOCIACION_AUT', ID_RANGES.ASOCIACION_AUT);
    r.push([idNew, idAsociacion, idPersona, 'GERENTE_TECNICO', principal?'SI':'NO', '', '']);
  }
  writeTable_(sh,h,r);
  return { ok:true };
}
function assignTomero_(idCauce, idPersona, principal, force){
  const sh = getSh_(SH.CAUCE_TOMERO);
  const t = readTable_(sh); const h=t.headers; const r=t.rows;
  const iId=idx_(h,'ID_CAUCE_TOMERO'), iCa=idx_(h,'FK_ID_CAUCE');
  const iPe=idx_(h,'ID_PERSONA'), iRo=idx_(h,'ROL'), iPr=idx_(h,'PRINCIPAL');

  const existsPrincipal = r.some(x =>
    asText_(x[iCa])===idCauce && asText_(x[iRo]).toUpperCase()==='TOMERO' && /^SI$/i.test(asText_(x[iPr]))
  );
  if (principal && existsPrincipal && !force){
    return { needsConfirm:true, message:'Ya hay un TOMERO principal asignado. ¿Agregar otro principal?' };
  }
  let updated=false;
  r.forEach(x=>{
    if (asText_(x[iCa])===idCauce && asText_(x[iPe])===idPersona && asText_(x[iRo]).toUpperCase()==='TOMERO'){
      x[iPr] = principal?'SI':'NO'; updated=true;
    }
  });
  if (!updated){
    const idNew = nextId_('NEXT_ID_CAUCE_TOMERO', ID_RANGES.CAUCE_TOMERO);
    r.push([idNew, idCauce, idPersona, 'TOMERO', principal?'SI':'NO', '', '']);
  }
  writeTable_(sh,h,r);
  writeTomeroPrincipalToCauces_();
  return { ok:true };
}

/***********************
 * TELÉFONOS (concat → PERSONAS)
 ***********************/
function phonesConcatForPerson_(idPersona){
  const t = readTable_(getSh_(SH.PERSONA_TELEFONOS));
  const iP = idx_(t.headers,'ID_PERSONA');
  const iN = idx_(t.headers,'NUMERO');
  const iPr = idx_(t.headers,'PRINCIPAL');
  const iT = idx_(t.headers,'TIPO');

  const list = t.rows
    .filter(r=>asText_(r[iP])===idPersona && asText_(r[iN]))
    .sort((a,b)=>{
      const ap = /^TRUE$/i.test(asText_(a[iPr])) ? 0 : 1;
      const bp = /^TRUE$/i.test(asText_(b[iPr])) ? 0 : 1;
      if (ap!==bp) return ap-bp;
      return asText_(a[iT]).localeCompare(asText_(b[iT]));
    })
    .map(r=>asText_(r[iN]));
  return list.join(' / ');
}
function writePhonesToPersonas_(){
  ensureHeadersExist_();

  const tT = readTable_(getSh_(SH.PERSONA_TELEFONOS));
  const iPer = idx_(tT.headers,'ID_PERSONA');
  const iNum = idx_(tT.headers,'NUMERO');
  const iPri = idx_(tT.headers,'PRINCIPAL');

  const map = {};
  tT.rows.forEach(r=>{
    const p  = asText_(r[iPer]);
    const nu = asText_(r[iNum]);
    const pr = asText_(r[iPri]);
    if (!p || !nu) return;
    if (!map[p]) map[p] = { principal:[], otros:[] };
    if (/^TRUE$/i.test(pr)) map[p].principal.push(nu);
    else map[p].otros.push(nu);
  });

  const shP = getSh_(SH.PERSONAS);
  const tP  = readTable_(shP);
  const iId = idx_(tP.headers,'ID_PERSONA');
  const iTl = idx_(tP.headers,'TELEFONO');

  tP.rows.forEach(r=>{
    const id = asText_(r[iId]);
    const set = map[id] || {principal:[], otros:[]};
    r[iTl] = [].concat(set.principal, set.otros).join(' / ');
  });

  writeTable_(shP, tP.headers, tP.rows);
}
function updatePersonaPhonesFromDetail_(){
  ensureHeadersExist_();
  const shP = getSh_(SH.PERSONAS);
  const t = readTable_(shP);
  const iId = idx_(t.headers,'ID_PERSONA');
  const iTel = idx_(t.headers,'TELEFONO');

  t.rows.forEach(r=>{
    const id = asText_(r[iId]);
    r[iTel] = phonesConcatForPerson_(id);
  });
  writeTable_(shP, t.headers, t.rows);
  const ui = tryGetUi_(); if (ui) ui.alert('PERSONAS.TELEFONO sincronizado desde PERSONA_TELEFONOS ✓');
}

/***********************
 * ROLES (semillas + helpers)
 ***********************/
function rolesDefaultSeeds_(){
  return [
    ['980000','NO_ASIGNADO','No asignado',0,'FALSE','FALSE','FALSE','TRUE',''],
    ['980001','INSPECTOR','Inspector',10,'TRUE','FALSE','FALSE','TRUE',''],
    ['980002','GERENTE_TECNICO','Gerente Técnico',20,'FALSE','TRUE','FALSE','TRUE',''],
    ['980003','TOMERO','Tomero',30,'FALSE','FALSE','TRUE','TRUE',''],
    ['980004','SUBDELEGADO','Subdelegado',40,'FALSE','FALSE','FALSE','TRUE',''],
    ['980005','CONSEJERO','Consejero',50,'FALSE','FALSE','FALSE','TRUE',''],
    ['980006','TECNICO','Técnico',60,'FALSE','FALSE','FALSE','TRUE',''],
    ['980007','ADMINISTRATIVO','Administrativo',70,'FALSE','FALSE','FALSE','TRUE',''],
    ['980008','INGENIERO','Ingeniero',80,'FALSE','FALSE','FALSE','TRUE',''],
    ['980009','CONTADOR','Contador',90,'FALSE','FALSE','FALSE','TRUE',''],
    ['980010','PRESIDENTE','Presidente',100,'FALSE','FALSE','FALSE','TRUE','']
  ];
}
function rolesInit_(){
  ensureHeadersExist_();
  const sh = getSh_(SH.ROLES);
  const t = readTable_(sh);
  if (!t.rows.length){
    writeTable_(sh, HDR.ROLES, rolesDefaultSeeds_());
  }else{
    writeTable_(sh, HDR.ROLES, t.rows);
  }
  styleHeader_(sh);
  const ui = tryGetUi_(); if (ui) ui.alert('ROLES inicializado/verificado ✓');
}
function rolesSeedForce_(){
  ensureHeadersExist_();
  const sh = getSh_(SH.ROLES);
  const cur = indexBy_(SH.ROLES,'ROL_KEY');
  const seeds = rolesDefaultSeeds_();
  seeds.forEach(row=>{
    const key=row[1];
    if (!cur[key]){
      const t = readTable_(sh);
      t.rows.push(row);
      writeTable_(sh, t.headers, t.rows);
    }
  });
  const ui = tryGetUi_(); if (ui) ui.alert('Semillas de ROLES forzadas/aplicadas ✓');
}

/***********************
 * ENCABEZADOS: auditoría básica
 ***********************/
function auditHeaders(){
  ensureHeadersExist_();
  const ss = SpreadsheetApp.getActive();
  const repName = 'AUDIT_ENCABEZADOS';
  const old = ss.getSheetByName(repName);
  if (old) ss.deleteSheet(old);
  const rep = ss.insertSheet(repName).setTabColor('#EF6C00');
  rep.getRange(1,1,rep.getMaxRows(),rep.getMaxColumns()).clearDataValidations();

  const out = [['SHEET','OK','HEADERS_ACTUALES','HEADERS_ESPERADOS','FALTAN','SOBRAN']];
  Object.keys(HDR).forEach(name=>{
    const sh = ss.getSheetByName(name);
    if(!sh){
      out.push([name,'NO','','',HDR[name].join(' | '),'(no existe)']);
      return;
    }
    const current = sh.getRange(1,1,1,Math.max(1,sh.getLastColumn()))
                       .getValues()[0].map(v=>String(v||'').trim());
    const expected = HDR[name];
    const faltan = expected.filter(h=>!current.includes(h));
    const sobran = current.filter(h=>!expected.includes(h));
    out.push([
      name,
      (faltan.length===0 && sobran.length===0) ? 'SI' : 'NO',
      current.join(' | '),
      expected.join(' | '),
      faltan.join(' | '),
      sobran.join(' | ')
    ]);
  });
  writeTable_(rep, out[0], out.slice(1));
  styleHeader_(rep);
}

function fixHeadersSafe(){
  ensureHeadersExist_();
  const ss=SpreadsheetApp.getActive();
  Object.keys(HDR).forEach(name=>{
    const sh=ss.getSheetByName(name); if(!sh) return;
    const expected=HDR[name];
    const data=sh.getDataRange().getValues();
    const curr=data[0].map(v=>String(v||'').trim());
    const headers=curr.slice();
    const faltan=expected.filter(h=>!headers.includes(h));
    if (faltan.length){
      const extra=Array(faltan.length).fill('');
      data[0]=headers.concat(faltan);
      for(let i=1;i<data.length;i++) data[i]=data[i].concat(extra);
      writeTable_(sh, data[0], data.slice(1));
    }else{
      writeTable_(sh, headers, data.slice(1));
    }
    styleHeader_(sh);
  });
}

/**
 * Reordena INSPECCIONES al orden canónico (HDR.INSPECCIONES) en bloques.
 */
function reorderInspecciones(){
  return reorderInspeccionesFast_(2000);
}
function reorderInspeccionesFast_(CHUNK){
  const ss   = SpreadsheetApp.getActive();
  const src  = getSh_(SH.INSPECCIONES);
  const need = HDR.INSPECCIONES.slice();

  const lastCol = src.getLastColumn();
  const lastRow = src.getLastRow();
  if (lastCol === 0) return;

  const header = src.getRange(1,1,1,lastCol)
                     .getValues()[0]
                     .map(v => String(v||'').trim());

  let already = true;
  for (let i=0; i<need.length; i++){
    if ((header[i]||'') !== need[i]) { already = false; break; }
  }
  if (already) return;

  const map    = need.map(h => header.indexOf(h));
  const extraIdx = header.map((_,i)=>i).filter(i => !map.includes(i));
  const newHeader= need.concat(extraIdx.map(i => header[i]));

  const tmpName = SH.INSPECCIONES + '__NEW';
  const tmpOld  = ss.getSheetByName(tmpName);
  if (tmpOld) ss.deleteSheet(tmpOld);
  const tmp = ss.insertSheet(tmpName);
  tmp.getRange(1,1,1,newHeader.length).setValues([newHeader]);

  if (lastRow > 1){
    let srcRow = 2;
    let dstRow = 2;
    while (srcRow <= lastRow){
      const n = Math.min(CHUNK, lastRow - srcRow + 1);
      const chunk = src.getRange(srcRow, 1, n, lastCol).getValues();
      const newRows = chunk.map(row => {
        const base  = need.map((_,k)=> map[k] >= 0 ? row[ map[k] ] : '' );
        const extra = extraIdx.map(i => row[i]);
        return base.concat(extra);
      });
      tmp.getRange(dstRow, 1, n, newHeader.length).setValues(newRows);
      srcRow += n;
      dstRow += n;
      SpreadsheetApp.flush();
    }
  }

  const color = src.getTabColor();
  ss.deleteSheet(src);
  tmp.setName(SH.INSPECCIONES);
  if (color) tmp.setTabColor(color);
  styleHeader_(tmp);
}