// ═══════════════════════════════════════════════════════════════════════
//  SHARF · Devolución de Activos TI · Code.gs  v5b
//
//  FLUJO COMPLETO:
//  P1  Tipo de devolución + modo ingreso
//        → QR (auto-avanza a P2)
//        → DNI (busca en Sheet RH → Snipe)
//        → Manual (usuario no encontrado)
//
//  P2  Preview colaborador + activos + accesorios (desde custom fields de Snipe)
//        → Selección de accesorios a devolver
//        → Carrusel de imágenes del activo (desde pestaña Archivos de Snipe)
//        → Alerta si activo asignado a otra persona
//
//  P3  Estado físico: sin observaciones / con hallazgos + fotos de evidencia
//
//  P4  Procesamiento:
//        · Checkin Snipe-IT → Disponible
//        · PDF basado en formulario DEVOLUCION.pdf (Compras + Soporte TI)
//        · Guardar PDF en Drive /Usuario/Devoluciones/
//        · Correo al colaborador (personal) + CC jefe + BCC supervisoras
//        · Log en Sheet Devoluciones_TI
//        · Alerta interna si Snipe desactualizado
//
//  ACCESORIOS: se leen EXCLUSIVAMENTE desde custom_fields de Snipe-IT
//    · order_number NO se usa ni se modifica
//    · _parsearOrderNumber eliminado
//
//  FORMATO QR (slash-separated):
//    SERIAL/DNI/APELLIDOS,NOMBRES/CECO/SERIAL_CARGADOR
//
//  CORREO DEVOLUCIÓN:
//    Para:  correo personal colaborador (Sheet RH)
//    CC:    jefe inmediato + helpdesk@holasharf.com
//    Para:  correo personal del colaborador
//    CC:    sheily.rebaza@holasharf.com + melanie.galindo@holasharf.com + anais.chero@holasharf.com
//    Asunto: Devolución de <tipo_activo> por <tipo_devolución> asignada al <Apellidos, Nombres>
//
//  ALERTA SNIPE DESACTUALIZADO:
//    Para:  gabriel.helpdesk@holasharf.com + helpdesk@holasharf.com
//    Asunto: ALERTA SNIPE DESACTUALIZADO
// ═══════════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════════════
// 1. CONFIGURACIÓN GLOBAL
// ════════════════════════════════════════════════════════════════════════
var CFG = {
  SNIPE_BASE:          'https://scharff.snipe-it.io/api/v1',
  SNIPE_KEY_PROP:      'SNIPE_API_KEY',

  SHEET_RH_ID:         '14LNid3_E7deg9rHD65TA9P_h8Lmr2s0HvWoejHlYQow',
  SHEET_RH_HDR_ROW:    4,   // fila 4 = encabezados, datos desde fila 5

  // Pestañas donde buscar colaboradores (en orden de prioridad)
  // colDni: nombre EXACTO del encabezado de la columna DNI en esa pestaña
  SHEET_RH_TABS: [
    { nombre: 'BD_Personal',          gid: '1212628329', colDni: 'DNI/C.E.'        },
    { nombre: 'Cesados',              gid: '1515544815', colDni: ' DNI / C.E.'     },
    { nombre: 'Practicantes Cesados', gid: '1533896218', colDni: 'DNI'             }
  ],
  SHEET_LOG_NAME:      'Devoluciones_TI',

  DRIVE_ACTAS_ROOT_ID:         '1UxiNhppSYCS-fp1rF8taIBa2XreT9lOe',
  // ID de la carpeta raíz del sistema de ASIGNACIONES (diferente a la de devoluciones).
  // Si se deja vacío, la búsqueda usará DRIVE_ACTAS_ROOT_ID como fallback.
  // Configura aquí el ID de la carpeta donde el sistema de asignaciones organiza
  // los PDFs de actas por colaborador (subcarpetas con el nombre del colaborador).
  DRIVE_ASIGNACIONES_ROOT_ID:  '',
  DRIVE_FIRMAS_TEC_ID:         '1EXUIcq56A23yMmvjco56gYwKr5mFRP0v',
  DRIVE_SEGUIMIENTO_ID: '',          // se llena automáticamente al crear el Sheet
  DRIVE_SEGUIMIENTO_FOLDER: '0APknu_tBOg5SUk9PVA',  // carpeta log en Drive helpdesk

  EMAIL_SUPERVISORA:   'anais.chero@holasharf.com',
  EMAIL_SHEILY:        'sheily.rebaza@holasharf.com',
  EMAIL_MELANIE:       'melanie.galindo@holasharf.com',
  EMAIL_HELPDESK:      'helpdesk@holasharf.com',
  EMAIL_GABRIEL_HD:    'gabriel.helpdesk@holasharf.com',
  DOMINIO:             'holasharf.com',

  BRAND_COLOR:  '#ff6568',
  BRAND_DARK:   '#1a5276',
  TZ:           'America/Lima',

  // ── Modo Prueba ──────────────────────────────────────────────────
  // Cuando está activo, los correos se redirigen a los correos de prueba
  // Los destinatarios originales se preservan en el log pero no reciben el correo
  MODO_PRUEBA_PROP: 'SHARF_MODO_PRUEBA',

  TIPOS_DEVOLUCION: ['Por cese', 'Por cambio de renting', 'Por cambio de equipo'],

  // Empresas SHARF para el formulario de devolución
  EMPRESAS: [
    'SICCSA - Scharff Int. Courier & Cargo',
    'SLI - Scharff Logística Integrada',
    'SR - Scharff Representaciones',
    'SB - Scharff Bolivia'
  ],

  TECNICOS: [
    { nombre:'Gabriel García Diaz',     firmaKey:'GABRIEL',  dni:'49058626', email:'gabriel.helpdesk@holasharf.com', sede:'Callao'   },
    { nombre:'Michel Helpdesk',         firmaKey:'MICHEL',   dni:'',         email:'michael.helpdesk@holasharf.com',  sede:'Callao'   },
    { nombre:'Victor Gutiérrez Leiva',  firmaKey:'VICTOR',   dni:'25750736', email:'misael.helpdesk@holasharf.com', sede:'Callao'   },
    { nombre:'Ismael Gomez Sime',       firmaKey:'ISMAEL',   dni:'6311827',  email:'ismael.helpdesk@holasharf.com',     sede:'Callao'   },
    { nombre:'Perla Moreno Yarleque',   firmaKey:'PERLA',    dni:'76250380', email:'perla.helpdesk@holasharf.com',     sede:'Paita'    },
    { nombre:'Alfredo Rojas Guevara',   firmaKey:'ALFREDO',  dni:'7166818',  email:'alfredo.helpdesk@holasharf.com',    sede:'Callao'   },
    { nombre:'Rocio Helpdesk',          firmaKey:'ROCIO',    dni:'',         email:'rocio.helpdesk@holasharf.com',     sede:'Arequipa' },
    { nombre:'Alvarez Chora Jesus Miguel', firmaKey:'JESUS', dni:'71422426', email:'jesus.helpdesk@holasharf.com',    sede:'Callao'   }
  ]
};

// ════════════════════════════════════════════════════════════════════════
// 2. ENTRY POINT
// ════════════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════════════
// doGet — Punto de entrada de la Web App
//
// Selección de interfaz por parámetro URL:
//   ?v=mobile  → sirve Mobile.html  (técnicos de campo con celular)
//   sin param  → sirve Index.html   (técnicos de sede con PC)
//
// Cómo usar:
//   • URL base (sin parámetro) → PC/escritorio
//   • URL + ?v=mobile           → celular/tablet
//
// Los técnicos de campo guardan en favoritos la URL con ?v=mobile.
// Index.html muestra un banner en pantallas pequeñas con el enlace móvil.
// ════════════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════════════
// LISTA DE USUARIOS AUTORIZADOS
// Solo estos correos pueden acceder a la aplicación.
// Modificar aquí y re-desplegar para agregar o quitar acceso.
// ════════════════════════════════════════════════════════════════════════
var USUARIOS_AUTORIZADOS = [
  'ismael.helpdesk@holasharf.com',
  'anais.chero@holasharf.com',
  'eddie.fernandez@holasharf.com',
  'jesus.helpdesk@holasharf.com',
  'michael.helpdesk@holasharf.com',
  'gabriel.helpdesk@holasharf.com',
  'rocio.helpdesk@holasharf.com',
  'perla.helpdesk@holasharf.com',
  'alfredo.helpdesk@holasharf.com',
  'misael.helpdesk@holasharf.com'
];

function _esAutorizado(emailLower) {
  if (!emailLower) return false;
  for (var i = 0; i < USUARIOS_AUTORIZADOS.length; i++) {
    if (USUARIOS_AUTORIZADOS[i].toLowerCase() === emailLower) return true;
  }
  return false;
}

function _htmlAccesoDenegado(motivo) {
  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:"Segoe UI",Arial,sans-serif;background:#68002B;' +
    'display:flex;align-items:center;justify-content:center;' +
    'min-height:100vh;padding:20px}' +
    '.card{background:#fff;border-radius:16px;padding:36px 32px;max-width:400px;' +
    'width:100%;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.35)}' +
    '.logo{font-size:36px;font-weight:900;font-style:italic;color:#68002B;' +
    'letter-spacing:-2px;margin-bottom:4px}' +
    '.sub{font-size:11px;color:#b08090;letter-spacing:1px;margin-bottom:20px}' +
    'hr{border:none;border-top:2px solid #FF6568;margin:16px 0}' +
    '</style></head><body><div class="card">' +
    '<div class="logo">sharf</div>' +
    '<div class="sub">DEVOLUCIONES TI · MESA DE SERVICIO</div>' +
    '<hr>' +
    '<p style="color:#b91c1c;font-weight:700;font-size:14px;margin-bottom:10px">⛔ Acceso restringido</p>' +
    '<p style="color:#666;font-size:13px;line-height:1.6">' + motivo + '</p>' +
    '<p style="color:#999;font-size:11px;margin-top:16px">Si crees que deberías tener acceso,<br>contacta a Ismael o Anais.</p>' +
    '</div></body></html>'
  ).setTitle('SHARF · Acceso restringido')
   .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var vParam = (params.v || '').toLowerCase().trim();
  var archivoPagina = (vParam === 'mobile') ? 'Mobile' : 'Index';
  Logger.log('doGet -> ' + archivoPagina);
  return HtmlService.createTemplateFromFile(archivoPagina)
    .evaluate()
    .setTitle('SHARF · Devoluciones TI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getHtml(nombreArchivo) {
  var nombres = ['Index', 'Mobile'];
  if (nombres.indexOf(nombreArchivo) < 0) nombreArchivo = 'Index';
  var tpl = HtmlService.createTemplateFromFile(nombreArchivo);
  return tpl.evaluate().getContent();
}

/**
 * testSesion — ejecuta desde el editor para verificar qué correo detecta el script.
 */
function testSesion() {
  var email     = Session.getActiveUser().getEmail();
  var emailEfec = Session.getEffectiveUser().getEmail();
  Logger.log('═══════════════════════════════════');
  Logger.log('DIAGNÓSTICO DE SESIÓN');
  Logger.log('═══════════════════════════════════');
  Logger.log('getActiveUser():    ' + (email     || '(vacío)'));
  Logger.log('getEffectiveUser(): ' + (emailEfec || '(vacío)'));
  Logger.log('Dominio configurado: @holasharf.com');
  var ok = email && email.toLowerCase().endsWith('@holasharf.com');
  Logger.log('¿Tiene acceso?: ' + (ok ? 'SÍ ✅' : 'NO ❌'));
  if (!email) {
    Logger.log('');
    Logger.log('⚠️  getActiveUser() está vacío.');
    Logger.log('   → Republica la Web App con:');
    Logger.log('     Ejecutar como: Usuario que accede a la aplicación');
    Logger.log('     Acceso: Cualquier usuario de Google Workspace de holasharf.com');
  }
  Logger.log('═══════════════════════════════════');
}

// ════════════════════════════════════════════════════════════════════════
// 3. DISPATCHER CENTRAL
// ════════════════════════════════════════════════════════════════════════
function handleRequest(action, params) {
  try {
    switch (action) {
      case 'getSession':            return getUserSession();
      case 'parsearQR':             return parsearQR(params.qr);
      case 'buscarPorDNI':          return buscarPorDNI(params.dni);
      case 'buscarActivoPorSerial': return buscarActivoPorSerial(params.serial);
      case 'procesarDevolucion':    return procesarDevolucion(params);
      case 'buscarActaAsignacion':      return buscarActaAsignacion(params);
      case 'diagnosticarBusquedaActa':  return diagnosticarBusquedaActa(params);
      case 'getFirmaBase64':            return { ok: true, data: _imgDriveTec(params.firmaKey || '') };
      case 'getLogoBase64':             return { ok: true, data: _imgDrive('LOGO.png') };
      case 'autenticarTecnico':         return autenticarTecnico();
      case 'getAuditoriaData':          return getAuditoriaData(params);
      case 'getAuditoriaDetalle':       return getAuditoriaDetalle(params);
      case 'getCecosSnipe':             return getCecosSnipe();
      case 'getAuditoriaUrls':          return getAuditoriaUrls();
      case 'exportarAuditoriaXLSX':     return exportarAuditoriaXLSX(params);
      case 'getSnipeAssetImages':   return getSnipeAssetImages(params.assetId);
      case 'reenviarCorreo':        return reenviarCorreo(params);
      case 'getModoPrueba':         return getModoPrueba();
      case 'setModoPrueba':         return setModoPrueba(params);
      case 'getSnipeCatalogos':     return getSnipeCatalogos();
      case 'diagnosticarSheetRH':   return diagnosticarSheetRH(params.dni);
      case 'validarConexiones':     return validarConexiones();
      default: return { ok: false, error: 'Acción desconocida: ' + action };
    }
  } catch(e) {
    Logger.log('handleRequest ERROR [' + action + ']: ' + e.message + '\n' + e.stack);
    return { ok: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════════════════════
// 4. SESIÓN DEL TÉCNICO
// ════════════════════════════════════════════════════════════════════════
function getUserSession() {
  // Intentar getActiveUser primero, luego getEffectiveUser como fallback
  var email = '';
  try { email = Session.getActiveUser().getEmail()    || ''; } catch(ex) {}
  if (!email) {
    try { email = Session.getEffectiveUser().getEmail() || ''; } catch(ex) {}
  }
  var emailNorm = email.toLowerCase().trim();

  // Buscar en la lista de técnicos del CFG
  var tec = null;
  for (var i = 0; i < CFG.TECNICOS.length; i++) {
    if (CFG.TECNICOS[i].email && CFG.TECNICOS[i].email.toLowerCase() === emailNorm) {
      tec = CFG.TECNICOS[i]; break;
    }
  }

  // Si no está en TECNICOS pero está en USUARIOS_AUTORIZADOS (Anais, Eddie)
  if (!tec && _esAutorizado(emailNorm)) {
    var nombre = emailNorm.split('@')[0].replace(/[._]/g,' ')
                  .replace(/\b\w/g, function(c){ return c.toUpperCase(); });
    tec = { nombre: nombre, firmaKey: '', dni: '', email: emailNorm, sede: 'Lima' };
  }

  if (!tec) {
    return { ok: false, error: 'Sesión no reconocida: ' + email };
  }

  return {
    ok:       true,
    email:    emailNorm,
    emailTec: tec.email,
    nombre:   tec.nombre,
    firmaKey: tec.firmaKey,
    dni:      tec.dni,
    sede:     tec.sede
  };
}

// ════════════════════════════════════════════════════════════════════════
// 5. PARSEAR QR (formato slash v3)
//    SERIAL/DNI/APELLIDOS,NOMBRES/CECO/SERIAL_CARGADOR/MOCHILA/MOUSE/DOCKING
// ════════════════════════════════════════════════════════════════════════
function parsearQR(qrText) {
  if (!qrText || !qrText.trim()) return { ok: false, error: 'QR vacío' };

  // ── Separar campos del QR ────────────────────────────────────────────
  // Los seriales y nombres NUNCA contienen "-" ni "/", por lo tanto
  // cualquier "-" o "/" en el string es separador de campo.
  // Prioridad de detección del separador:
  //   1. " - "  (espacio-guión-espacio) ← pistola con espacios
  //   2. "/"    (slash)                 ← formato legacy
  //   3. "-"    (guión solo)            ← pistola sin espacios
  var txt    = qrText.trim();
  var partes = [];
  var sep    = '';

  if (txt.indexOf(' - ') >= 0) {
    partes = txt.split(' - ').map(function(p){ return p.trim(); });
    sep = ' - ';
  } else if (txt.indexOf('/') >= 0) {
    partes = txt.split('/').map(function(p){ return p.trim(); });
    sep = '/';
  } else if (txt.indexOf('-') >= 0) {
    partes = txt.split('-').map(function(p){ return p.trim(); });
    sep = '-';
  } else {
    // Sin separador → solo serial
    Logger.log('parsearQR: sin separador, tratando como serial: ' + txt);
    return { ok: true, soloSerial: true, serial: txt };
  }

  Logger.log('parsearQR: sep="' + sep + '" → ' + partes.length + ' partes: ' +
             partes.slice(0,5).join(' | '));

  if (partes.length < 2) {
    return { ok: true, soloSerial: true, serial: txt };
  }

  var qrData = {
    serial:         partes[0] || '',
    dni:            partes[1] || '',
    nombre:         partes[2] || '',
    centroCostos:   partes[3] || '',
    serialCargador: partes[4] || '',
    mochila:        partes[5] ? partes[5].toUpperCase() !== 'NO MOCHILA' : false,
    mochilaDesc:    partes[5] || '',
    mouse:          partes[6] ? partes[6].toUpperCase() !== 'NO MOUSE'   : false,
    mouseDesc:      partes[6] || '',
    docking:        partes[7] ? partes[7].toUpperCase() !== 'NO DOCKING' : false,
    dockingDesc:    partes[7] || ''
  };

  Logger.log('parsearQR OK: serial=' + qrData.serial +
             ' | dni=' + qrData.dni + ' | nombre=' + qrData.nombre);

  var resultSnipe = buscarActivosDeUsuarioPorDNI_(qrData.dni, qrData.serial);
  var resultRH    = _buscarEmpleadoPorDni(qrData.dni);

  return {
    ok:              true,
    qrData:          qrData,
    empleado:        resultRH.empleado || null,
    activos:         resultSnipe.activos || [],
    activoPrincipal: resultSnipe.activoPrincipal || null,
    snipeOk:         resultSnipe.ok,
    rhOk:            resultRH.ok
  };
}

// ════════════════════════════════════════════════════════════════════════
// 6. BUSCAR POR DNI (modo sin QR)
//    1. Busca en Sheet RH (3 pestañas)
//    2. Si lo encuentra, busca sus activos en Snipe-IT
//    3. Si no está en RH → modo manual
//    4. Si está en RH pero no en Snipe → continúa solo con datos RH
// ════════════════════════════════════════════════════════════════════════
function buscarPorDNI(dni) {
  if (!dni || !String(dni).trim()) return { ok: false, error: 'DNI vacío' };
  dni = String(dni).trim();

  // ── 1. Buscar en Sheet RH ─────────────────────────────────────────
  var resultRH = _buscarEmpleadoPorDni(dni);
  if (!resultRH.ok) {
    Logger.log('buscarPorDNI: DNI ' + dni + ' no encontrado en RH');
    return { ok: false, noEncontrado: true,
             error: resultRH.error || 'DNI ' + dni + ' no encontrado en el sistema.',
             modoManual: true };
  }

  Logger.log('buscarPorDNI: encontrado en RH (' + resultRH.tabOrigen + '): ' + resultRH.empleado.nombre);

  // ── 2. Buscar activos en Snipe-IT ─────────────────────────────────
  var resultSnipe = buscarActivosDeUsuarioPorDNI_(dni, null);

  if (!resultSnipe.ok) {
    Logger.log('buscarPorDNI: Snipe ERROR: ' + resultSnipe.error);
    return {
      ok:         true,
      empleado:   resultRH.empleado,
      tabOrigen:  resultRH.tabOrigen,
      activos:    [],
      snipeOk:    false,
      snipeError: resultSnipe.error || 'No se encontraron activos en Snipe-IT'
    };
  }

  Logger.log('buscarPorDNI: Snipe OK, activos: ' + resultSnipe.activos.length);

  // ── 3. Enriquecer datos del empleado con campos de Snipe ──────────
  // Si el Sheet RH no tiene algunos campos, completar con los del usuario Snipe
  // Mapeo: department (Snipe usuario) → area (RH)
  //        location   (Snipe usuario) → sede / Centro de Operaciones (RH)
  var emp = resultRH.empleado;
  if (resultSnipe.snipeArea      && !emp.area)    emp.area          = resultSnipe.snipeArea;
  if (resultSnipe.snipeUbicacion && !emp.sede)    emp.sede          = resultSnipe.snipeUbicacion;
  if (resultSnipe.snipeEmpresa   && !emp.empresa) emp.empresa       = resultSnipe.snipeEmpresa;
  // Guardar IDs Snipe del usuario para el PATCH post-checkin
  emp.snipeDeptName  = resultSnipe.snipeArea    || '';
  emp.snipeDeptId    = resultSnipe.snipeDeptId  || null;
  emp.snipeCompanyId = resultSnipe.snipeCompanyId || null;  // ← para heredar compañía al activo

  return {
    ok:          true,
    empleado:    emp,
    tabOrigen:   resultRH.tabOrigen,
    activos:     resultSnipe.activos || [],
    snipeOk:     true,
    snipeUserId: resultSnipe.snipeUserId || ''
  };
}

// ════════════════════════════════════════════════════════════════════════
// 7. BUSCAR ACTIVO POR SERIAL (modo manual o verificación)
// ════════════════════════════════════════════════════════════════════════
function buscarActivoPorSerial(serial) {
  if (!serial || !serial.trim()) return { ok: false, error: 'Serial vacío' };
  serial = serial.trim();
  try {
    var resp = snipeGET('/hardware?search=' + encodeURIComponent(serial) + '&limit=10');
    if (!resp || !resp.rows || !resp.rows.length) {
      return { ok: false, noEncontrado: true, error: 'Serial no encontrado en Snipe-IT' };
    }
    // Buscar SOLO coincidencia exacta — no aceptar similares
    var activo = null;
    for (var i = 0; i < resp.rows.length; i++) {
      if ((resp.rows[i].serial || '').trim().toUpperCase() === serial.toUpperCase()) {
        activo = resp.rows[i];
        break;
      }
    }
    if (!activo) {
      return { ok: false, noEncontrado: true,
               error: 'No se encontró un activo con serial exacto "' + serial + '" en Snipe-IT' };
    }

    var accesorios = _leerAccesoriosCustomFields(activo);
    var asignado   = activo.assigned_to && activo.assigned_to.id;

    return {
      ok:         true,
      activo:     _mapearActivo_(activo),
      accesorios: accesorios,
      asignado:   asignado ? {
        id:     activo.assigned_to.id,
        nombre: activo.assigned_to.name  || '',
        email:  activo.assigned_to.email || '',
        dni:    activo.assigned_to.username || ''
      } : null
    };
  } catch(e) {
    Logger.log('buscarActivoPorSerial: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════════════════════
// 8. BUSCAR ACTIVOS DE USUARIO POR DNI EN SNIPE-IT
//    Busca al usuario por: username, employee_num, y texto libre
//    Retorna todos sus activos asignados + el activo principal si hay serial
// ════════════════════════════════════════════════════════════════════════
// 8. BUSCAR ACTIVOS DE USUARIO POR DNI EN SNIPE-IT
//    DNI guardado en: employee_num (confirmado)
//    Activos: /hardware?assigned_to=ID&assigned_type=user
// ════════════════════════════════════════════════════════════════════════
function buscarActivosDeUsuarioPorDNI_(dni, serialPrincipal) {
  try {
    var dniStr = String(dni).trim();
    Logger.log('Snipe: buscando usuario employee_num=' + dniStr);

    // Buscar por employee_num (campo confirmado del DNI en SHARF)
    var snipeUser = null;
    var uResp = snipeGET('/users?employee_num=' + encodeURIComponent(dniStr) + '&limit=5');
    Logger.log('Snipe employee_num → total:' + (uResp.total||0) + ' rows:' + (uResp.rows ? uResp.rows.length : 0));

    if (uResp && uResp.rows && uResp.rows.length) {
      snipeUser = uResp.rows.filter(function(u) {
        return String(u.employee_num || '').trim() === dniStr;
      })[0] || uResp.rows[0];
    }

    // Fallback: search por si employee_num param no filtra en esta versión de Snipe
    if (!snipeUser) {
      Logger.log('Snipe: fallback a /users?search=' + dniStr);
      var sResp = snipeGET('/users?search=' + encodeURIComponent(dniStr) + '&limit=10');
      Logger.log('Snipe search → total:' + (sResp.total||0) + ' rows:' + (sResp.rows ? sResp.rows.length : 0));
      if (sResp && sResp.rows && sResp.rows.length) {
        snipeUser = sResp.rows.filter(function(u) {
          return String(u.employee_num || '').trim() === dniStr;
        })[0] || null;
        if (!snipeUser) {
          Logger.log('Sin coincidencia exacta. Usuarios devueltos:');
          sResp.rows.forEach(function(u) {
            Logger.log('  id:' + u.id + ' name:' + u.name +
                       ' employee_num:' + (u.employee_num||'-') + ' username:' + (u.username||'-'));
          });
        }
      }
    }

    if (!snipeUser) {
      Logger.log('Snipe: NO encontrado employee_num=' + dniStr);
      return { ok: false, activos: [], noEncontradoSnipe: true,
               error: 'Usuario no encontrado en Snipe-IT (employee_num: ' + dniStr + ')' };
    }

    Logger.log('Snipe usuario: id=' + snipeUser.id + ' name=' + snipeUser.name +
               ' employee_num=' + (snipeUser.employee_num||'-'));

    var snipeArea      = (snipeUser.department && snipeUser.department.name) ? snipeUser.department.name : '';
    var snipeDeptId    = (snipeUser.department && snipeUser.department.id)   ? snipeUser.department.id   : null;
    var snipeUbicacion = (snipeUser.location   && snipeUser.location.name)   ? snipeUser.location.name   : '';
    var snipeEmpresa   = (snipeUser.company    && snipeUser.company.name)    ? snipeUser.company.name    : '';
    // company_id del usuario → se hereda al activo en el PATCH post-checkin
    // Corresponde al campo "Seleccionar una compañía" del activo en Snipe-IT UI
    var snipeCompanyId = (snipeUser.company    && snipeUser.company.id)      ? snipeUser.company.id      : null;

    // ── Activos asignados al usuario ──────────────────────────────────
    // Usar /users/{id}/assets — endpoint específico que devuelve
    // SOLO los activos de ese usuario, sin riesgo de traer otros
    var activos = [];

    try {
      var aResp = snipeGET('/users/' + snipeUser.id + '/assets?limit=50');
      Logger.log('Snipe /users/' + snipeUser.id + '/assets → total:' + (aResp.total||0) +
                 ' rows:' + (aResp.rows ? aResp.rows.length : 0));

      if (aResp && aResp.rows && aResp.rows.length) {
        activos = aResp.rows.map(function(a) {
          var act = _mapearActivo_(a, _leerAccesoriosCustomFields(a));
          if (!act.area      && snipeArea)      act.area      = snipeArea;
          if (!act.ubicacion && snipeUbicacion) act.ubicacion = snipeUbicacion;
          if (!act.empresa   && snipeEmpresa)   act.empresa   = snipeEmpresa;
          // IDs para herencia en PATCH post-checkin
          act.snipeLocationId = (snipeUser.location && snipeUser.location.id) ? snipeUser.location.id : null;
          act.snipeCompanyId  = snipeCompanyId;   // ← herencia de compañía
          Logger.log('  ✅ ' + act.nombre + ' S/N:' + act.serial + ' [' + act.tipoActivo + '] ' + act.estado);
          return act;
        });
      } else {
        Logger.log('Snipe: usuario sin activos asignados.');
      }
    } catch(eAssets) {
      // Fallback: /hardware?assigned_to con verificación estricta del campo assigned_to
      Logger.log('Snipe /users/assets falló (' + eAssets.message + '), usando fallback con verificación...');
      try {
        var aResp2 = snipeGET('/hardware?assigned_to=' + snipeUser.id +
                              '&assigned_type=user&limit=50&sort=updated_at&order=desc');
        Logger.log('Snipe fallback → rows:' + (aResp2.rows ? aResp2.rows.length : 0));
        if (aResp2 && aResp2.rows && aResp2.rows.length) {
          // VERIFICACIÓN ESTRICTA: solo incluir activos cuyo assigned_to.id coincida
          activos = aResp2.rows.filter(function(a) {
            return a.assigned_to && String(a.assigned_to.id) === String(snipeUser.id);
          }).map(function(a) {
            var act = _mapearActivo_(a, _leerAccesoriosCustomFields(a));
            if (!act.area      && snipeArea)      act.area      = snipeArea;
            if (!act.ubicacion && snipeUbicacion) act.ubicacion = snipeUbicacion;
            if (!act.empresa   && snipeEmpresa)   act.empresa   = snipeEmpresa;
            // IDs para herencia en PATCH post-checkin
            act.snipeLocationId = (snipeUser.location && snipeUser.location.id) ? snipeUser.location.id : null;
            act.snipeCompanyId  = snipeCompanyId;   // ← herencia de compañía
            Logger.log('  ✅ ' + act.nombre + ' S/N:' + act.serial + ' [' + act.tipoActivo + ']');
            return act;
          });
          Logger.log('Snipe fallback verificado: ' + activos.length + ' activos del usuario');
        }
      } catch(eFallback) {
        Logger.log('Snipe fallback ERROR: ' + eFallback.message);
      }
    }

    var activoPrincipal = null;
    if (serialPrincipal) {
      activoPrincipal = activos.filter(function(a) {
        return (a.serial || '').toUpperCase() === serialPrincipal.toUpperCase();
      })[0] || null;
    }

    return {
      ok:             true,
      activos:        activos,
      activoPrincipal:activoPrincipal,
      snipeUserId:    snipeUser.id,
      snipeUserNombre:snipeUser.name     || '',
      snipeArea:      snipeArea,
      snipeDeptId:    snipeDeptId,
      snipeUbicacion: snipeUbicacion,
      snipeEmpresa:   snipeEmpresa,
      snipeCompanyId: snipeCompanyId     // ← ID de compañía del usuario
    };

  } catch(e) {
    Logger.log('buscarActivosDeUsuarioPorDNI_ ERROR: ' + e.message);
    if (e.message.indexOf('401') >= 0 || e.message.indexOf('Unauthorized') >= 0)
      return { ok: false, activos: [], error: 'Token Snipe-IT inválido o expirado.' };
    if (e.message.indexOf('403') >= 0)
      return { ok: false, activos: [], error: 'Sin permisos en Snipe-IT.' };
    return { ok: false, activos: [], error: 'Snipe-IT: ' + e.message };
  }
}


// ════════════════════════════════════════════════════════════════════════
// 9. PROCESAR DEVOLUCIÓN COMPLETA
// ════════════════════════════════════════════════════════════════════════
function procesarDevolucion(datos) {
  Logger.log('INICIO procesarDevolucion — ' + (datos.empleado?datos.empleado.nombre:'?') + ' | ' + (datos.activo?datos.activo.serial:'?'));
  var resultado = {
    snipeOk: false, pdfOk: false, driveOk: false,
    emailOk: false, logOk:  false, alertaOk: false,
    devId: '', driveUrl: '', pdfDriveId: '',
    errors: [],
    // Detalle granular por paso (para mostrar en el frontend)
    pasos: {
      snipe: { ok: false, msg: '' },
      pdf:   { ok: false, msg: '' },
      drive: { ok: false, msg: '' },
      email: { ok: false, msg: '', errorCorreo: false, correoInvalido: '' },
      log:   { ok: false, msg: '' },
      alerta:{ ok: false, msg: '' }
    },
    // Datos para reintento de correo
    payloadCorreo: null
  };

  var ts        = Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy HH:mm:ss');
  var devId     = 'DEV-' + Utilities.formatDate(new Date(), CFG.TZ, 'yyyyMMddHHmmss');
  resultado.devId = devId;

  var nombre     = datos.empleado ? (datos.empleado.nombre || '').trim() : '(Sin nombre)';
  var serial     = datos.activo   ? (datos.activo.serial   || '').trim() : '(Sin serial)';
  var tipoActivo = _clasificarTipoActivo(datos.activo ? datos.activo.nombre + ' ' + (datos.activo.categoria||'') : '');

  // ── HERENCIA: activo hereda área/ceco/ubicación del empleado ────────
  // Fuente primaria: Sheet RH → campos AREA, Código, CENTRO DE OPERACIONES
  // Fuente secundaria: formulario manual (manArea, manSede, manCeco)
  if (datos.activo && datos.empleado) {
    if (!datos.activo.area      && datos.empleado.area) datos.activo.area      = datos.empleado.area;
    if (!datos.activo.ceco      && datos.empleado.ceco) datos.activo.ceco      = datos.empleado.ceco;
    if (!datos.activo.ubicacion && datos.empleado.sede) datos.activo.ubicacion = datos.empleado.sede;
  }

  // ── 1. CHECKIN SNIPE-IT + HERENCIA ÁREA/CECO/UBICACIÓN ──────────────
  //
  //   Campos custom confirmados en Snipe-IT SHARF:
  //     Área:              _snipeit_area_14              (field_id = 14)
  //     Centro de Costos:  _snipeit_centro_de_costos_15  (field_id = 15)
  //   Campo estándar:
  //     Ubicación:         location_id  (buscar el ID por nombre)
  //
  //   Fuente de datos:
  //     área     → emp.area  (Sheet RH: columna AREA  / Snipe: department.name)
  //     ceco     → emp.ceco  (Sheet RH: columna Código)
  //     ubicación→ emp.sede  (Sheet RH: columna CENTRO DE OPERACIONES / Snipe: location.name)
  //
  if (datos.activo && datos.activo.id) {
    try {
      var checkinResp = snipePOST('/hardware/' + datos.activo.id + '/checkin', {
        status_id: 7,  // 7 = Disponible (confirmado catálogo Snipe-IT SHARF)
        note: 'Devolucion ' + datos.tipoDev + ' | ' + ts + ' | Tecnico: ' + (datos.tecnico ? datos.tecnico.nombre : '')
      });
      resultado.snipeOk = (checkinResp && checkinResp.status === 'success');
      resultado.pasos.snipe = {
        ok:  resultado.snipeOk,
        msg: resultado.snipeOk
          ? 'Activo desvinculado → estado Disponible'
          : 'Respuesta Snipe: ' + JSON.stringify(checkinResp).slice(0, 120)
      };
    } catch(eSnipe) {
      resultado.pasos.snipe = { ok: false, msg: eSnipe.message };
      resultado.errors.push('Snipe checkin: ' + eSnipe.message);
    }

    // ── PATCH: actualizar campos del activo con datos del empleado ───────
    // IMPORTANTE: el checkin puede blanquear los custom_fields del activo.
    // Hay que re-aplicar área, CECO y ubicación SIEMPRE después del checkin.
    // Fuente primaria del área: department del usuario en Snipe (select2-department_select-container)
    try {
      var emp2       = datos.empleado || {};
      // Área: prioridad → Snipe department.name del usuario → Sheet RH → activo
      var areaVal    = (emp2.snipeDeptName  || emp2.area  || datos.activo.area      || '').trim();
      var cecoVal    = (emp2.ceco  || datos.activo.ceco      || '').trim();
      var sedeVal    = (emp2.sede  || datos.activo.ubicacion || '').trim();

      var patchPayload = {};
      var patchCampos  = [];

      // ── Compañía: herencia desde el usuario Snipe ─────────────────────
      // El activo hereda la company_id del ÚLTIMO usuario que lo tuvo asignado.
      // Fuente: activo.snipeCompanyId (guardado al listar activos del usuario)
      //         o emp2.snipeCompanyId (guardado al enriquecer empleado en buscarPorDNI)
      var companyId = (datos.activo && datos.activo.snipeCompanyId)
                    || emp2.snipeCompanyId
                    || null;
      if (companyId) {
        patchPayload['company_id'] = companyId;
        patchCampos.push('company_id=' + companyId + '(' + (emp2.empresa || 'empresa usuario') + ')');
      } else {
        Logger.log('PATCH herencia: company_id no disponible (usuario sin compañía en Snipe)');
      }

      // Custom field Área (field_id 14 → clave _snipeit_area_14)
      // Corresponde al campo "select2-selection__rendered" del activo en Snipe UI
      if (areaVal) {
        patchPayload['_snipeit_area_14'] = areaVal;
        patchCampos.push('Area=' + areaVal);
      }

      // department_id del activo (opcional — si tenemos el ID del dept del usuario)
      if (emp2.snipeDeptId) {
        patchPayload['department_id'] = emp2.snipeDeptId;
        patchCampos.push('department_id=' + emp2.snipeDeptId);
      }

      // Custom field Centro de Costos (field_id 15 → clave _snipeit_centro_de_costos_15)
      // Solo el código antes del primer guion: "PE121ST229 - PE10 LIM SOPORTE" → "PE121ST229"
      if (cecoVal) {
        var cecoSnipe = cecoVal.indexOf('-') >= 0
          ? cecoVal.split('-')[0].trim()
          : cecoVal.trim();
        patchPayload['_snipeit_centro_de_costos_15'] = cecoSnipe;
        patchCampos.push('CECO=' + cecoSnipe);
      }

      // Ubicación estándar:
      // rtd_location_id del activo = location_id del usuario (ubicación predeterminada)
      // Primero intentar usar el ID directo del usuario Snipe si lo tenemos,
      // si no, buscar por nombre de sede del Sheet RH
      var locationIdHerencia = null;

      // Opción A: el activo ya trae location del usuario (se guardó en activo.snipeLocationId)
      if (datos.activo.snipeLocationId) {
        locationIdHerencia = datos.activo.snipeLocationId;
        patchCampos.push('rtd_location_id=' + locationIdHerencia + '(del usuario Snipe)');
      }
      // Opción B: buscar por nombre de sede (Sheet RH: CENTRO DE OPERACIONES)
      else if (sedeVal) {
        try {
          var locResp = snipeGET('/locations?search=' + encodeURIComponent(sedeVal) + '&limit=5');
          if (locResp && locResp.rows && locResp.rows.length) {
            var loc = locResp.rows.filter(function(l) {
              return l.name.toLowerCase() === sedeVal.toLowerCase();
            })[0] || locResp.rows[0];
            locationIdHerencia = loc.id;
            patchCampos.push('rtd_location_id=' + loc.name + '(id:' + loc.id + ')');
          }
        } catch(eLoc) {
          Logger.log('Buscar location para rtd_location_id "' + sedeVal + '": ' + eLoc.message);
        }
      }

      if (locationIdHerencia) {
        patchPayload['rtd_location_id'] = locationIdHerencia;
      }

      if (Object.keys(patchPayload).length > 0) {
        var patchResp = snipePATCH('/hardware/' + datos.activo.id, patchPayload);
        var patchOk   = patchResp && patchResp.status === 'success';
        Logger.log('PATCH herencia activo [' + patchCampos.join(', ') + ']: ' +
                   (patchOk ? 'OK' : JSON.stringify(patchResp).slice(0, 150)));
        if (resultado.pasos.snipe.ok) {
          resultado.pasos.snipe.msg += (patchOk
            ? ' · Campos heredados: ' + patchCampos.join(', ')
            : ' · PATCH campos falló: ' + (patchResp && patchResp.messages ? JSON.stringify(patchResp.messages) : 'sin detalle'));
        }
      } else {
        Logger.log('PATCH herencia: sin datos de area/ceco/ubicacion para actualizar');
      }

      // ── PATCH accesorios: actualiza el campo checkbox múltiple en Snipe ──
      // El campo "Accesorios" (_snipeit_accesorios_37[]) indica qué accesorios
      // están DISPONIBLES con el equipo (para la próxima asignación).
      //
      // Lógica al devolver:
      //   · Accesorios DEVUELTOS     → se MARCAN   (regresan al inventario con el equipo)
      //   · Accesorios NO devueltos  → se DESMARCAN (el colaborador los retiene o se perdieron)
      //   · Accesorios nuevos marcados por técnico → se MARCAN también
      //
      // Resultado: el campo queda con exactamente los accesorios que se recibieron hoy.
      var accesoriosDevueltos = datos.accesoriosDevueltos || {};
      var accesoriosNuevos    = datos.accesoriosNuevos    || {};
      var debeActualizarAcc   = datos.activo && datos.activo.id;

      if (debeActualizarAcc) {
        try {
          var cfActual = (datos.activo && datos.activo.custom_fields) ? datos.activo.custom_fields : {};

          // Encontrar la db_column_name real del campo en los custom_fields del activo
          var claveAccField = '_snipeit_accesorios_37'; // default
          var camposPosibs  = ['Accesorios','accesorios','ACCESORIOS','_snipeit_accesorios_37','accesorios_37'];
          for (var ci = 0; ci < camposPosibs.length; ci++) {
            if (cfActual[camposPosibs[ci]] !== undefined) {
              var fieldMeta = cfActual[camposPosibs[ci]];
              if (fieldMeta && fieldMeta.field) claveAccField = fieldMeta.field;
              break;
            }
          }

          // Valor actual del campo para log
          var valorActual = '';
          for (var ci2 = 0; ci2 < camposPosibs.length; ci2++) {
            if (cfActual[camposPosibs[ci2]] && cfActual[camposPosibs[ci2]].value != null) {
              valorActual = String(cfActual[camposPosibs[ci2]].value || '').trim();
              break;
            }
          }

          // Mapa: key interna → valor exacto en Snipe
          var mapaValores = {
            'mouse':   'Mouse',
            'mochila': 'Mochila',
            'docking': 'Docking',
            'teclado': 'Teclado'
          };

          // Construir el nuevo valor del campo:
          // Se incluyen SOLO los accesorios que se devuelven HOY
          // (los que el técnico marcó con toggle ON en el formulario)
          var nuevosValores = [];

          Object.keys(mapaValores).forEach(function(key) {
            var val    = mapaValores[key];
            var devuelto = !!(accesoriosDevueltos[key]);
            var nuevo    = !!(accesoriosNuevos[key]);
            // Marcar si se devuelve ahora O si el técnico lo añadió manualmente
            if (devuelto || nuevo) {
              nuevosValores.push(val);
            }
          });

          var patchAccPayload = {};
          patchAccPayload[claveAccField] = nuevosValores;

          var patchAccResp = snipePATCH('/hardware/' + datos.activo.id, patchAccPayload);
          Logger.log('PATCH accesorios: ' + claveAccField + ' = ' + JSON.stringify(nuevosValores) +
                     ' (antes: "' + valorActual + '") → ' +
                     (patchAccResp && patchAccResp.status === 'success' ? '✅ OK' : JSON.stringify(patchAccResp).slice(0,120)));

        } catch(eAccPatch) {
          Logger.log('PATCH accesorios ERROR (no crítico): ' + eAccPatch.message);
        }
      }
    } catch(ePatch) {
      // No es error crítico — el checkin ya se hizo
      Logger.log('PATCH herencia ERROR (no crítico): ' + ePatch.message);
      if (resultado.pasos.snipe.ok) {
        resultado.pasos.snipe.msg += ' · No se pudo actualizar campos (permisos o config Snipe)';
      }
    }

  } else {
    resultado.pasos.snipe = { ok: false, msg: 'Sin ID Snipe — devolucion manual registrada igualmente' };
  }

  // ── 2. ALERTA SNIPE DESACTUALIZADO ───────────────────────────────────
  if (datos.snipeDesactualizado && datos.snipeViejoPropietario) {
    try {
      _enviarAlertaSnipeDesactualizado(datos, ts, nombre, serial, tipoActivo);
      resultado.alertaOk = true;
      resultado.pasos.alerta = { ok: true, msg: 'Alerta enviada a gabriel.helpdesk + helpdesk' };
    } catch(eAl) {
      resultado.pasos.alerta = { ok: false, msg: eAl.message };
    }
  }

  // ── 3. LOGO + FIRMA ──────────────────────────────────────────────────
  var logoData     = _imgDrive('LOGO.png');
  var firmaTecData = null;
  try { firmaTecData = _imgDriveTec(datos.tecnico ? datos.tecnico.firmaKey : ''); } catch(ef) {}

  // ── 4. PDF ───────────────────────────────────────────────────────────
  var pdfBlob    = null;
  var pdfDriveId = '';
  try {
    var carpeta   = _carpetaDevolucionesUsuario(datos.empleado);
    var pdfResult = _generarPDFDevolucion(datos, ts, nombre, serial, devId, tipoActivo, carpeta.getId());
    resultado.pdfOk   = pdfResult.ok;
    resultado.driveOk = pdfResult.ok;
    resultado.driveUrl  = pdfResult.pdfUrl || '';
    pdfDriveId = pdfResult.pdfId || '';
    resultado.pdfDriveId = pdfDriveId;
    resultado.pasos.pdf   = { ok: pdfResult.ok, msg: pdfResult.ok ? 'PDF generado con firma técnico' : (pdfResult.error || 'Error generando PDF') };
    resultado.pasos.drive = { ok: pdfResult.ok, msg: pdfResult.ok ? 'Guardado en Drive: ' + (pdfResult.pdfUrl || '') : 'No se guardó en Drive' };

    if (pdfDriveId) {
      // Obtener el nombre exacto del archivo guardado en Drive
      var pdfNombreReal = 'Devolucion_' + _slug(nombre) + '_' + serial + '.pdf';
      try { pdfNombreReal = DriveApp.getFileById(pdfDriveId).getName(); } catch(eN) {}

      try {
        var pdfFetch = UrlFetchApp.fetch(
          'https://www.googleapis.com/drive/v3/files/' + pdfDriveId + '?alt=media&supportsAllDrives=true',
          { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
        );
        if (pdfFetch.getResponseCode() === 200)
          pdfBlob = pdfFetch.getBlob().setName(pdfNombreReal);
      } catch(eFetch) { Logger.log('PDF fetch blob: ' + eFetch.message); }

      // ── Subir PDF a Archivos del activo en Snipe-IT ───────────────
      // Mismo nombre que el archivo en Drive
      if (pdfBlob && datos.activo && datos.activo.id) {
        try {
          var snipeKey   = _snipeKey();
          var uploadBlob = null;
          try { uploadBlob = pdfBlob.getAs('application/pdf'); } catch(e) { uploadBlob = pdfBlob; }
          // Usar el nombre exacto del archivo en Drive
          uploadBlob.setName(pdfNombreReal);

          var snipeResp = UrlFetchApp.fetch(
            CFG.SNIPE_BASE + '/hardware/' + datos.activo.id + '/files',
            {
              method:             'POST',
              headers:            { 'Authorization': 'Bearer ' + snipeKey, 'Accept': 'application/json' },
              payload:            { 'file[]': uploadBlob },
              // NO contentType — Apps Script genera el boundary automáticamente
              muteHttpExceptions: true
            }
          );
          var snipeCode = snipeResp.getResponseCode();
          var snipeText = snipeResp.getContentText();
          Logger.log('Snipe upload PDF activo[' + datos.activo.id + '] → HTTP ' + snipeCode + ': ' + snipeText.slice(0,200));

          var snipeOk = false;
          try {
            var sb = JSON.parse(snipeText);
            snipeOk = snipeCode === 200 && sb.status === 'success';
          } catch(ep) { snipeOk = snipeCode === 200; }

          if (snipeOk) {
            resultado.pasos.drive.msg += ' · PDF adjunto en Snipe-IT';
            Logger.log('Snipe upload PDF OK');
          } else {
            Logger.log('Snipe upload PDF no exitoso (no critico): HTTP ' + snipeCode);
          }
        } catch(eSnipeFile) {
          Logger.log('Snipe upload PDF error (no critico): ' + eSnipeFile.message);
        }
      }
    }
  } catch(ePDF) {
    resultado.pasos.pdf   = { ok: false, msg: ePDF.message };
    resultado.pasos.drive = { ok: false, msg: 'No generado' };
    resultado.errors.push('PDF: ' + ePDF.message);
  }

  // ── 5. LOG ───────────────────────────────────────────────────────────
  try {
    _registrarLogDevolucion(devId, datos, ts, nombre, serial, tipoActivo, resultado, pdfDriveId);
    resultado.logOk = true;
    resultado.pasos.log = { ok: true, msg: 'Registrado en hoja Devoluciones_TI' };
  } catch(eLog) {
    resultado.pasos.log = { ok: false, msg: eLog.message };
    resultado.errors.push('Log: ' + eLog.message);
  }

  // ── 6. CORREO ────────────────────────────────────────────────────────
  // Guardar payload para posible reintento
  resultado.payloadCorreo = {
    datos: datos, ts: ts, nombre: nombre, serial: serial,
    devId: devId, tipoActivo: tipoActivo, pdfDriveId: pdfDriveId
  };
  try {
    var emailRes = _enviarCorreoDevolucion(datos, ts, nombre, serial, devId, tipoActivo, logoData, firmaTecData, pdfBlob);
    resultado.emailOk  = emailRes.ok;
    resultado.emailInfo = emailRes;
    resultado.pasos.email = {
      ok: emailRes.ok,
      msg: emailRes.ok
        ? 'Enviado a: ' + (emailRes.para || '') + (emailRes.cc ? ' · CC: ' + emailRes.cc : '')
        : (emailRes.error || 'Error desconocido'),
      errorCorreo:     !emailRes.ok && !!emailRes.faltanCorreos,
      correoProblema:  !emailRes.ok ? emailRes.correoInvalido || '' : '',
      faltanCorreos:   emailRes.faltanCorreos || false
    };
  } catch(eEmail) {
    var esCorreoMal = eEmail.message.indexOf('Invalid email') >= 0 || eEmail.message.indexOf('recipient') >= 0;
    var correoMal = '';
    // Intentar identificar qué correo falló
    if (esCorreoMal) {
      var ep = datos.empleado || {};
      correoMal = ep.emailPersonal || ep.emailJefe || '';
    }
    resultado.pasos.email = {
      ok: false,
      msg: eEmail.message,
      errorCorreo: esCorreoMal,
      correoProblema: correoMal,
      faltanCorreos: !datos.empleado || (!datos.empleado.emailPersonal && !datos.empleado.emailJefe)
    };
    resultado.errors.push('Email: ' + eEmail.message);
  }

  resultado.ok = resultado.snipeOk || resultado.pdfOk;

  // ── 7. CORREO DE COTIZACIÓN (si hay accesorios para cotizar) ─────────
  var accCotizar = datos.accesoriosCotizar || {};
  var itemsACotizar = Object.keys(accCotizar).filter(function(k){ return accCotizar[k]; });
  if (itemsACotizar.length > 0) {
    try {
      var cotRes = _enviarCorreoCotizacion(datos, ts, nombre, tipoActivo, itemsACotizar, logoData);
      resultado.pasos.cotizacion = {
        ok:  cotRes.ok,
        msg: cotRes.ok ? '✅ Solicitud de cotización enviada a Logística' : ('Error: ' + (cotRes.error || ''))
      };
      Logger.log('Correo cotización: ' + (cotRes.ok ? 'OK' : cotRes.error));
    } catch(eCot) {
      resultado.pasos.cotizacion = { ok: false, msg: 'Error cotización: ' + eCot.message };
      Logger.log('Correo cotización ERROR: ' + eCot.message);
    }
  }

  return resultado;
}

// ════════════════════════════════════════════════════════════════════════
// 10. GENERADOR DE PDF — ESTILO FORMULARIO DEVOLUCION.pdf
//     DocumentApp: dos secciones (Compras + Soporte TI) + fotos si hay
// ════════════════════════════════════════════════════════════════════════
// ── Helper PDF: construye una fila de la tabla Soporte TI ─────────────
// devuelto: bool — si el accesorio se devuelve en esta oportunidad
// bueno:    bool — estado individual (true=bueno, false=con observaciones)
// obs:      string — texto de observación
function _filaPDF(n, nombre, devuelto, bueno, obs, costo) {
  var TD = 'padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:7.5pt;color:#000000;vertical-align:middle;text-align:center';
  var TL = 'padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:7.5pt;color:#000000;vertical-align:middle;text-align:left';
  var OB = 'padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:6.5pt;color:#000000;vertical-align:middle;text-align:left';
  function cb(v){ return v ? '&#9746;' : '&#9744;'; }
  function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  var costoVal = (costo && costo > 0) ? 'S/. ' + parseFloat(costo).toFixed(2) : '&mdash;';
  return '<tr>' +
    '<td style="' + TD + '">' + n + '</td>' +
    '<td style="' + TL + '">' + esc(nombre) + '</td>' +
    '<td style="' + TD + '">1</td>' +
    '<td style="' + TD + '">' + cb(devuelto)    + '</td>' +
    '<td style="' + TD + '">' + cb(!devuelto)   + '</td>' +
    '<td style="' + TD + '">&#9744;</td>' +
    '<td style="' + TD + '">' + cb(devuelto && bueno)    + '</td>' +
    '<td style="' + TD + '">' + cb(devuelto && !bueno)   + '</td>' +
    '<td style="' + TD + ';font-size:6.5pt">' + costoVal + '</td>' +
    '<td style="' + OB + '">' + esc(obs) + '</td>' +
    '</tr>';
}

// ════════════════════════════════════════════════════════════════════════
// _generarPDFDevolucion
//   Recibe un PNG del acta (generado por el cliente con html2canvas)
//   y lo convierte a PDF usando el motor de GAS.
//   Si no llega el PNG, cae al método HTML legacy.
//
//   El PNG ocupa toda la hoja A4 con márgenes mínimos.
//   Las fotos de evidencia van en páginas adicionales.
// ════════════════════════════════════════════════════════════════════════
function _generarPDFDevolucion(datos, ts, nombre, serial, devId, tipoActivo, carpetaDestId) {
  try {
    var fechaFile = Utilities.formatDate(new Date(), CFG.TZ, 'yyyyMMdd_HHmm');
    var nomPDF    = 'Devolucion_' + _slug(nombre) + '_' + serial + '_' + fechaFile + '.pdf';

    // ── Fotos de evidencia (para página 2+) ──────────────────────────────
    var fotosDataUrls = [];
    if (datos.photos && datos.photos.length) {
      datos.photos.forEach(function(b64) {
        try {
          if (b64 && b64.indexOf('data:') === 0) fotosDataUrls.push(b64);
          else if (b64 && b64.length > 100)      fotosDataUrls.push('data:image/jpeg;base64,' + b64.split(',').pop());
        } catch(ef) {}
      });
    }

    // ── Ruta A: PNG del cliente recibido → insertar como imagen en PDF ────
    var pngActa = datos.pngActa || '';   // base64 del PNG renderizado en el cliente
    if (pngActa && pngActa.length > 100) {
      Logger.log('_generarPDFDevolucion: usando PNG del cliente (' + Math.round(pngActa.length/1024) + ' KB)');

      // HTML mínimo: una imagen por página, márgenes 8mm
      // El motor de GAS maneja <img> perfectamente
      var S = '#FF6568';
      var paginaActa =
        '<div style="width:100%;text-align:center;page-break-after:' + (fotosDataUrls.length > 0 ? 'always' : 'auto') + '">' +
          '<img src="' + pngActa + '" ' +
               'style="width:100%;height:auto;display:block;margin:0 auto">' +
        '</div>';

      var paginasFotos = fotosDataUrls.length > 0
        ? '<div style="text-align:center;padding-top:8px">' +
            '<div style="font-family:Arial;font-size:12pt;font-weight:bold;color:' + S + ';text-align:center;margin-bottom:4px">' +
              'EVIDENCIA FOTOGR\u00c1FICA \u2014 HALLAZGOS' +
            '</div>' +
            '<div style="font-family:Arial;font-size:7.5pt;text-align:center;color:#555555;margin-bottom:10px">' +
              _escHTML(nombre) + ' \u00b7 S/N: ' + _escHTML(serial) + ' \u00b7 ' + _escHTML(ts) +
            '</div>' +
            fotosDataUrls.map(function(du, i) {
              return '<div style="text-align:center;margin-bottom:16px">' +
                     '<img src="' + du + '" style="max-width:100%;max-height:270px;display:block;margin:0 auto">' +
                     '<div style="font-family:Arial;font-size:7pt;color:#555555;font-style:italic;margin-top:3px">Hallazgo ' + (i+1) + '</div>' +
                     '</div>';
            }).join('') +
          '</div>'
        : '';

      var htmlFinal =
        '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
        '<style>@page{size:A4;margin:8mm 8mm 8mm 8mm}body{margin:0;padding:0}</style>' +
        '</head><body>' +
        paginaActa + paginasFotos +
        '</body></html>';

      var htmlBlob  = Utilities.newBlob(htmlFinal, 'text/html', 'temp.html');
      var driveTemp = DriveApp.createFile(htmlBlob);
      var pdfBlob   = driveTemp.getAs('application/pdf');
      pdfBlob.setName(nomPDF);
      var carpetaDest = DriveApp.getFolderById(carpetaDestId);
      var pdfFile = carpetaDest.createFile(pdfBlob);
      _driveSetPublicReader(pdfFile.getId());
      try { driveTemp.setTrashed(true); } catch(e) {}
      Logger.log('_generarPDFDevolucion (PNG): OK → ' + pdfFile.getUrl());
      return { ok: true, pdfId: pdfFile.getId(), pdfUrl: pdfFile.getUrl() };
    }

    // ── Ruta B: fallback HTML legacy (si no llega PNG) ────────────────────
    Logger.log('_generarPDFDevolucion: sin PNG del cliente — usando HTML legacy');
    return _generarPDFLegacy(datos, ts, nombre, serial, devId, tipoActivo, carpetaDestId, fotosDataUrls, nomPDF);

  } catch(e) {
    Logger.log('_generarPDFDevolucion ERROR: ' + e.message + '\n' + e.stack);
    return { ok: false, error: e.message };
  }
}

// Helper HTML escape para el PDF
function _escHTML(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function _generarPDFLegacy(datos, ts, nombre, serial, devId, tipoActivo, carpetaDestId, fotosDataUrls, nomPDF) {
  try {
    var emp     = datos.empleado            || {};
    var activo  = datos.activo              || {};
    var acc     = datos.accesoriosDevueltos || {};
    var accData = datos.accesorios          || {};
    var accObs  = datos.accesoriosObs       || {};
    var accEst  = datos.accesoriosEstado    || {};
    var accNuev = datos.accesoriosNuevos    || {};
    var accCos  = datos.accesoriosCosto     || {};
    var tec     = datos.tecnico             || {};
    var esObs   = datos.hayObservaciones;
    var hayActivo = !!activo.serial;

    function accDevuelto(key) { return acc[key] !== false && acc[key] !== undefined ? !!acc[key] : false; }
    function accBueno(key) { return (accEst[key] || 'bueno') === 'bueno'; }

    function cb(val)  { return val ? '&#9746;' : '&#9744;'; }
    function esc(s)   { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
    function blobToDataUrl(blob) {
      if (!blob) return '';
      try { return 'data:' + (blob.getContentType()||'image/png') + ';base64,' + Utilities.base64Encode(blob.getBytes()); }
      catch(e2) { return ''; }
    }

    var logoBlob = null;
    try {
      var carp = DriveApp.getFolderById(CFG.DRIVE_FIRMAS_TEC_ID);
      var logoNames = ['LOGO.png', 'Vive_Sharf.png', 'logo.jpg', 'LOGO.jpg'];
      for (var li2 = 0; li2 < logoNames.length; li2++) {
        var lf = carp.getFilesByName(logoNames[li2]);
        if (lf.hasNext()) { logoBlob = lf.next().getBlob(); break; }
      }
    } catch(e) {}

    var versionBlob = null;
    try {
      var vf = DriveApp.getFolderById(CFG.DRIVE_FIRMAS_TEC_ID).getFilesByName('VersionDev.jpg');
      if (vf.hasNext()) versionBlob = vf.next().getBlob();
      else { var vf2 = DriveApp.getFolderById(CFG.DRIVE_FIRMAS_TEC_ID).getFilesByName('VERSION.png'); if (vf2.hasNext()) versionBlob = vf2.next().getBlob(); }
    } catch(e) {}

    var firmaTecBlob = null;
    try { firmaTecBlob = _obtenerFirmaTecBlob(tec.firmaKey || ''); } catch(e) {}

    var logoDataUrl    = blobToDataUrl(logoBlob);
    var versionDataUrl = blobToDataUrl(versionBlob);
    var firmaDataUrl   = blobToDataUrl(firmaTecBlob);

    // ── Fotos ───────────────────────────────────────────────────
    var fotosDataUrls = [];
    if (datos.photos && datos.photos.length) {
      datos.photos.forEach(function(b64, idx) {
        try {
          if (b64 && b64.indexOf('data:') === 0) {
            fotosDataUrls.push(b64);
          } else if (b64 && b64.length > 100) {
            fotosDataUrls.push('data:image/jpeg;base64,' + b64.split(',').pop());
          }
        } catch(ef) { Logger.log('Foto ' + (idx+1) + ' error: ' + ef.message); }
      });
    }
    Logger.log('Fotos para PDF: ' + fotosDataUrls.length);

    // ── Variables de contenido ──────────────────────────────────
    var fecha = Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy');

    // Logo HTML
    var logoImgHtml = logoDataUrl
      ? '<img src="' + logoDataUrl + '" style="height:52px;display:block">'
      : '<span style="font-family:Manrope,Arial,sans-serif;font-size:28pt;font-weight:900;font-style:italic;color:#68002B">sharf</span>';

    // Versión HTML
    var verHtml = versionDataUrl
      ? '<img src="' + versionDataUrl + '" style="width:58px;height:36px">'
      : '<span style="font-size:6pt;line-height:1.5">T&amp;C-RG-03<br>V.10<br>Pág.: 1 de 1</span>';

    // Firma técnico HTML
    var firmaImgHtml = firmaDataUrl
      ? '<img src="' + firmaDataUrl + '" style="height:42px;display:block;margin:0 auto 4px">'
      : '<div style="height:46px"></div>';

    // Empresas checkboxes
    var empActual = (emp.empresa || '').toUpperCase();
    var empHtml = '<div style="font-weight:bold;font-size:7.5pt;margin-bottom:3px">Empresa donde laboró (marque con X):</div>';
    [['SICCSA','SICCSA - Scharff Int. Courier &amp; Cargo'],
     ['SLI',   'SLI - Scharff Log&iacute;stica Integrada'],
     ['SR',    'SR - Scharff Representaciones'],
     ['SB',    'SB - Scharff Bolivia']
    ].forEach(function(e) {
      var marcado = empActual.indexOf(e[0]) >= 0;
      empHtml += '<div style="font-size:7.5pt;margin:1px 0' + (marcado ? ';font-weight:bold' : '') + '">' +
                 (marcado ? '&#9746;' : '&#9744;') + ' ' + e[1] + '</div>';
    });

    // Filas tabla Soporte TI
    var filasIT = '';
    var noIt = 1;

    if (hayActivo) {
      var descActivo = {'Laptop':'Laptop','PC':'PC','All-in-One':'All-in-One',
                        'Minidesktop':'Minidesktop','Monitor':'Monitor'}[tipoActivo]
                      || esc(activo.nombre || tipoActivo);
      var obsAct = 'S/N: ' + esc(serial);
      if (datos.observacionDesc && datos.observacionDesc.trim()) obsAct += '. ' + esc(datos.observacionDesc.trim());
      else if (esObs) obsAct += '. Ver fotos p&aacute;g. 2.';
      else obsAct += '. Sin observaciones.';
      // equipoBueno: true=bueno, false=malo — viene del formulario (botón Bueno/Con obs.)
      var equipoBueno = datos.equipoBueno !== false;  // default true si no viene
      filasIT += _filaPDF(noIt++, descActivo, true, equipoBueno, obsAct, accCos['activo']||0);
    }

    if (accData.cargadorSerial) {
      var devC   = accDevuelto('cargador');
      var buenoC = accBueno('cargador');
      var obsC   = 'S/N: ' + esc(accData.cargadorSerial) + ((accObs.cargador||'').trim() ? '. ' + esc(accObs.cargador.trim()) : '');
      filasIT += _filaPDF(noIt++, 'Cargador de Laptop', devC, buenoC, obsC, accCos.cargador||0);
    } else if (accNuev.cargador && accDevuelto('cargador')) {
      var obsC2 = (accObs.cargador||'').trim() || 'Registrado en esta devolución';
      filasIT += _filaPDF(noIt++, 'Cargador de Laptop', true, accBueno('cargador'), obsC2, accCos.cargador||0);
    }

    if (accData.mouseDesc) {
      var devM   = accDevuelto('mouse');
      var buenoM = accBueno('mouse');
      var obsM   = (accObs.mouse||'').trim();
      filasIT += _filaPDF(noIt++, 'Mouse', devM, buenoM, obsM, accCos.mouse||0);
    } else if (accNuev.mouse && accDevuelto('mouse')) {
      filasIT += _filaPDF(noIt++, 'Mouse', true, accBueno('mouse'), (accObs.mouse||'').trim() || 'Registrado en esta devolución', accCos.mouse||0);
    }

    if (accData.mochilaDesc) {
      var devMo   = accDevuelto('mochila');
      var buenoMo = accBueno('mochila');
      var obsMo   = (accObs.mochila||'').trim();
      var fDif = '';
      if (!devMo && datos.compromisosDiferidos) datos.compromisosDiferidos.forEach(function(cd) { if (cd.label==='Mochila' && cd.fechaCompromiso) fDif = cd.fechaCompromiso; });
      if (fDif) obsMo = 'Pendiente: ' + fDif + (obsMo ? '. ' + obsMo : '');
      filasIT += _filaPDF(noIt++, 'Mochila', devMo, buenoMo, obsMo, accCos.mochila||0);
    } else if (accNuev.mochila && accDevuelto('mochila')) {
      filasIT += _filaPDF(noIt++, 'Mochila', true, accBueno('mochila'), (accObs.mochila||'').trim() || 'Registrado en esta devolución', accCos.mochila||0);
    }

    if (accData.dockingDesc) {
      var devD   = accDevuelto('docking');
      var buenoD = accBueno('docking');
      var obsD   = (accObs.docking||'').trim();
      filasIT += _filaPDF(noIt++, 'Docking Station', devD, buenoD, obsD, accCos.docking||0);
    } else if (accNuev.docking && accDevuelto('docking')) {
      filasIT += _filaPDF(noIt++, 'Docking Station', true, accBueno('docking'), (accObs.docking||'').trim() || 'Registrado en esta devolución', accCos.docking||0);
    }

    if (accData.tecladoDesc || accData.teclado) {
      var devT   = accDevuelto('teclado');
      var buenoT = accBueno('teclado');
      var obsT   = (accObs.teclado||'').trim();
      filasIT += _filaPDF(noIt++, 'Teclado', devT, buenoT, obsT, accCos.teclado||0);
    } else if (accNuev.teclado && accDevuelto('teclado')) {
      filasIT += _filaPDF(noIt++, 'Teclado', true, accBueno('teclado'), (accObs.teclado||'').trim() || 'Registrado en esta devolución', accCos.teclado||0);
    }

    ['Backup del Equipo','&iquest;Se cancel&oacute; Cuenta de Correo?','&iquest;Se cancel&oacute; Cuenta de Legacy?',
     '&iquest;Se cancel&oacute; Cuenta de Sintad?','&iquest;Se cancel&oacute; Cuenta de Dominio?'].forEach(function(cta) {
      filasIT += '<tr><td style="padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:7.5pt;color:#000000;vertical-align:middle;text-align:center">' + noIt++ + '</td><td style="padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:7.5pt;color:#000000;vertical-align:middle;text-align:left">' + cta + '</td>' +
        '<td style="padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:7.5pt;color:#000000;vertical-align:middle;text-align:center">&mdash;</td>' +
        '<td style="padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:7.5pt;color:#000000;vertical-align:middle;text-align:center">&#9744;</td><td style="padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:7.5pt;color:#000000;vertical-align:middle;text-align:center">&#9744;</td><td style="padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:7.5pt;color:#000000;vertical-align:middle;text-align:center">&mdash;</td>' +
        '<td style="padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:7.5pt;color:#000000;vertical-align:middle;text-align:center">&mdash;</td><td style="padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:7.5pt;color:#000000;vertical-align:middle;text-align:center">&mdash;</td>' +
        '<td style="padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:6.5pt;color:#000000;vertical-align:middle;text-align:center">&mdash;</td>' +
        '<td style="padding:3px 5px;border:0.4pt solid #000000;font-family:Arial;font-size:7.5pt;color:#000000;vertical-align:middle;text-align:left;font-size:6.5pt"></td></tr>';
    });

    // Comentarios
    var obsTexto = 'Equipo S/N: ' + esc(serial);
    if (datos.observacionDesc && datos.observacionDesc.trim()) obsTexto += ' &middot; ' + esc(datos.observacionDesc.trim());
    else if (esObs) obsTexto += '. Hallazgos registrados. Ver fotos en p&aacute;gina siguiente.';
    else obsTexto += '. Sin observaciones adicionales.';
    if (datos.compromisosDiferidos && datos.compromisosDiferidos.length) {
      datos.compromisosDiferidos.forEach(function(cd) {
        obsTexto += ' ' + esc(cd.label) + ' pendiente de entrega';
        if (cd.fechaCompromiso) obsTexto += ' al ' + esc(cd.fechaCompromiso);
        obsTexto += '.';
      });
    }
    obsTexto += ' Tipo de devoluci&oacute;n: ' + esc(datos.tipoDev || '');

    // ── HTML COMPLETO — estilos 100% inline, compatible motor PDF de GAS ──
    // SIN: @import, variables CSS, flex, grid, nth-child, clases complejas
    // SOLO: font-family:Arial, colores HEX, table/td/th con style inline
    var S  = '#FF6568';
    var BK = '#000000';
    var WH = '#ffffff';
    var RO = '#fff5f5';

    var TD  = 'padding:3px 5px;border:0.4pt solid ' + BK + ';font-family:Arial;font-size:7.5pt;color:' + BK + ';vertical-align:middle';
    var TH  = 'padding:3px 5px;border:0.4pt solid ' + BK + ';font-family:Arial;font-size:6.5pt;font-weight:bold;text-align:center;background:' + S + ';color:' + WH;
    var THL = 'padding:3px 5px;border:0.4pt solid ' + BK + ';font-family:Arial;font-size:6.5pt;font-weight:bold;text-align:left;background:' + S + ';color:' + WH;
    var LBL = 'padding:3px 5px;border:0.4pt solid ' + BK + ';font-family:Arial;font-size:7.5pt;font-weight:bold;color:' + BK + ';background:' + WH;

    var html =
'<!DOCTYPE html><html><head><meta charset="UTF-8">' +
'<style>@page{size:A4;margin:14mm 12mm 14mm 12mm}body{margin:0;padding:0}</style>' +
'</head>' +
'<body style="font-family:Arial;font-size:7.5pt;color:' + BK + ';background:' + WH + '">' +

// ── CABECERA ──
'<table style="width:100%;border-collapse:collapse;margin-bottom:3px"><tr>' +
'<td style="border:none;width:130pt;padding:2px 4px 2px 0;vertical-align:middle">' + logoImgHtml + '</td>' +
'<td style="border:none;text-align:center;padding:2px 6px;vertical-align:middle">' +
  '<div style="font-family:Arial;font-size:13pt;font-weight:bold;margin-bottom:2px;color:' + BK + '">Devoluci&oacute;n de Equipos y Materiales</div>' +
  '<div style="font-family:Arial;font-size:6.5pt;font-style:italic;color:' + BK + '">(Documento imprescindible a efectos de la Liquidaci&oacute;n de Beneficios Sociales)</div>' +
  '<div style="font-family:Arial;font-size:5.5pt;color:#555555;margin-top:2px">T&amp;C-RG-03 &middot; V.10 &middot; ID: ' + esc(devId) + ' &middot; ' + esc(ts) + '</div>' +
'</td>' +
'<td style="border:0.4pt solid ' + BK + ';width:60pt;text-align:right;padding:4px;vertical-align:middle">' + verHtml + '</td>' +
'</tr></table>' +

// Separador rojo
'<table style="width:100%;border-collapse:collapse;margin:4px 0"><tr>' +
'<td style="height:2pt;background:' + S + ';border:none;padding:0;font-size:0">&nbsp;</td>' +
'</tr></table>' +

// ── TABLA COLABORADOR — negro, sin fondo ──
'<table style="width:100%;border-collapse:collapse;margin-bottom:2px">' +
'<tr>' +
'<td style="' + LBL + ';width:25%">Nombres y Apellidos:</td>' +
'<td style="' + TD  + ';width:43%">' + esc(nombre) + '</td>' +
'<td style="' + LBL + ';width:13%">DNI / CI:</td>' +
'<td style="' + TD  + ';width:19%">' + esc(emp.dni||'') + '</td>' +
'</tr>' +
'<tr>' +
'<td style="' + LBL + '">&Aacute;rea:</td>' +
'<td style="' + TD  + '">' + esc(emp.area||'') + '</td>' +
'<td style="' + LBL + '">Cargo:</td>' +
'<td style="' + TD  + '">' + esc(emp.cargo||'') + '</td>' +
'</tr>' +
'<tr>' +
'<td style="' + LBL + '">Sede:</td>' +
'<td style="' + TD  + '">' + esc(emp.sede||'') + '</td>' +
'<td style="' + LBL + '">Fecha:</td>' +
'<td style="' + TD  + '">' + fecha + '</td>' +
'</tr>' +
'</table>' +

// ── EMPRESA ──
'<table style="width:100%;border-collapse:collapse;margin-bottom:4px"><tr>' +
'<td style="' + LBL + ';width:25%">Empresa:</td>' +
'<td style="' + TD  + '">' + empHtml + '</td>' +
'</tr></table>' +

// ── TÍTULO SOPORTE TI ──
'<table style="width:100%;border-collapse:collapse;margin:5px 0 0"><tr>' +
'<td style="background:' + S + ';text-align:center;font-family:Arial;font-size:8pt;font-weight:bold;' +
    'padding:4px;border:0.4pt solid ' + BK + ';color:' + WH + '">Soporte TI</td>' +
'</tr></table>' +

// ── TABLA SOPORTE TI — th 100% inline ──
'<table style="width:100%;border-collapse:collapse"><thead>' +
'<tr>' +
'<th style="' + TH  + ';width:4%"   rowspan="2">No.</th>' +
'<th style="' + THL + ';width:24%"  rowspan="2">Descripci&oacute;n</th>' +
'<th style="' + TH  + ';width:5%"   rowspan="2">Cant.</th>' +
'<th style="' + TH  + '"            colspan="3">Recibido</th>' +
'<th style="' + TH  + '"            colspan="2">Estado</th>' +
'<th style="' + TH  + ';width:9%"   rowspan="2">Costo S/.<br>(trabajador)</th>' +
'<th style="' + THL + '"            rowspan="2">Observaciones</th>' +
'</tr>' +
'<tr>' +
'<th style="' + TH + ';width:5%">SI</th>' +
'<th style="' + TH + ';width:5%">NO</th>' +
'<th style="' + TH + ';width:5%">NA</th>' +
'<th style="' + TH + ';width:6%">Bueno</th>' +
'<th style="' + TH + ';width:6%">Malo</th>' +
'</tr></thead><tbody>' + filasIT +
'<tr>' +
'<td colspan="9" style="' + TD + ';text-align:right;font-weight:bold;color:' + S + '"">(*) Total:</td>' +
'<td style="background:' + S + ';border:0.4pt solid ' + BK + ';padding:3px 5px"></td>' +
'</tr>' +
'</tbody></table>' +

// ── COMENTARIOS ──
'<div style="border:0.4pt solid ' + BK + ';padding:5px 7px;margin:5px 0;font-family:Arial;font-size:7.5pt;color:' + BK + '">' +
'<span style="font-weight:bold">Comentarios de Soporte TI:</span><br>' + obsTexto +
'</div>' +

// ── FIRMA TÉCNICO — tabla pura, sin display:table-cell ──
'<table style="width:100%;border-collapse:collapse;margin-top:8px"><tr>' +
'<td style="width:50%;border:none;padding:0"></td>' +
'<td style="width:50%;border:none;padding:0;vertical-align:top">' +
  '<table style="width:100%;border-collapse:collapse">' +
  '<tr><td style="background:' + S + ';text-align:center;font-family:Arial;font-weight:bold;font-size:8pt;' +
      'padding:4px 5px;border:0.4pt solid ' + BK + ';color:' + WH + '">VB del &Aacute;rea de Soporte TI</td></tr>' +
  '<tr><td style="border:0.4pt solid ' + BK + ';border-top:none;padding:10px 12px 14px;text-align:center">' +
    firmaImgHtml +
    '<table style="width:100%;border-collapse:collapse;margin-top:6px"><tr>' +
    '<td style="border:none;border-top:0.5pt solid #aaaaaa;padding-top:4px;text-align:center">' +
      '<div style="font-family:Arial;font-weight:bold;font-size:8.5pt;color:' + BK + '">' + esc(tec.nombre||'') + '</div>' +
      '<div style="font-family:Arial;font-size:7pt;font-style:italic;color:' + BK + ';margin-top:2px">VB del &Aacute;rea de Soporte TI</div>' +
      '<div style="font-family:Arial;font-size:7pt;color:#555555;margin-top:2px">Fecha: ' + esc(ts) + '</div>' +
    '</td></tr></table>' +
  '</td></tr>' +
  '</table>' +
'</td>' +
'</tr></table>' +

// ── PIE ──
'<table style="width:100%;border-collapse:collapse;margin-top:6px"><tr>' +
'<td style="border:none;border-top:0.4pt solid ' + BK + ';padding-top:4px;font-family:Arial;font-size:6pt;font-style:italic;color:' + BK + '">' +
'De existir observaciones con lo indicado en el presente documento, s&iacute;rvase proporcionar el formato de ' +
'descuento correspondiente debidamente firmado para hacerse responsable por los equipos y/o materiales ' +
'no entregados, y/o los devueltos en mal estado, para proceder con el descuento correspondiente en su ' +
'liquidaci&oacute;n de beneficios sociales.' +
'</td></tr></table>' +

// ── FOTOS ──
(fotosDataUrls.length > 0
  ? '<div style="page-break-before:always;padding-top:8px">' +
    '<div style="font-family:Arial;font-size:12pt;font-weight:bold;color:' + S + ';text-align:center;margin-bottom:4px">EVIDENCIA FOTOGR&Aacute;FICA &mdash; HALLAZGOS</div>' +
    '<div style="font-family:Arial;font-size:7.5pt;text-align:center;color:#555555;margin-bottom:10px">' + esc(nombre) + ' &middot; S/N: ' + esc(serial) + ' &middot; ' + esc(ts) + '</div>' +
    fotosDataUrls.map(function(du, i) {
      return '<div style="text-align:center;margin-bottom:16px">' +
             '<img src="' + du + '" style="max-width:100%;max-height:270px;display:block;margin:0 auto">' +
             '<div style="font-family:Arial;font-size:7pt;color:#555555;font-style:italic;margin-top:3px">Hallazgo ' + (i+1) + '</div>' +
             '</div>';
    }).join('') +
    '</div>'
  : '') +

'</body></html>';

    // ── GENERAR PDF ─────────────────────────────────────────────
    var htmlBlob  = Utilities.newBlob(html, 'text/html', 'temp.html');
    var driveTemp = DriveApp.createFile(htmlBlob);
    var pdfBlob2  = driveTemp.getAs('application/pdf');
    var fechaFile = Utilities.formatDate(new Date(), CFG.TZ, 'yyyyMMdd_HHmm');
    pdfBlob2.setName('Devolucion_' + _slug(nombre) + '_' + serial + '_' + fechaFile + '.pdf');

    var carpetaDest = DriveApp.getFolderById(carpetaDestId);
    var pdfFile     = carpetaDest.createFile(pdfBlob2);
    _driveSetPublicReader(pdfFile.getId());
    try { driveTemp.setTrashed(true); } catch(e) {}

    return { ok: true, pdfId: pdfFile.getId(), pdfUrl: pdfFile.getUrl() };

  } catch(e) {
    Logger.log('_generarPDFLegacy ERROR: ' + e.message + '\n' + e.stack);
    return { ok: false, error: e.message };
  }
}
// ════════════════════════════════════════════════════════════════════════
// 11. ENVIAR CORREO DE DEVOLUCIÓN
//     Diseño: header oscuro (#1a5276) + logo blanco · sin imagen de firma
//     Incluye: obs por artículo · artículos no devueltos · herencia área/ceco
// ════════════════════════════════════════════════════════════════════════
function _enviarCorreoDevolucion(datos, ts, nombre, serial, devId, tipoActivo, logoData, firmaTecData, pdfBlob) {
  var emp     = datos.empleado           || {};
  var activo  = datos.activo             || {};
  var tec     = datos.tecnico            || {};
  var acc     = datos.accesoriosDevueltos || {};
  var accData = datos.accesorios         || {};
  var accObs  = datos.accesoriosObs      || {};

  var emailDest = emp.emailPersonal || '';
  var emailJefe = emp.emailJefe     || '';
  var emailCorp = emp.emailCorp     || '';   // "Email Corp" del Sheet RH
  if (!emailDest && !emailCorp && !emailJefe) {
    return { ok: false, error: 'No hay correos disponibles', faltanCorreos: true };
  }

  // Asunto sin caracteres especiales para evitar encoding issues en GmailApp
  function _ascii(s) {
    return (s || '').replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i')
      .replace(/ó/g,'o').replace(/ú/g,'u').replace(/ü/g,'u')
      .replace(/Á/g,'A').replace(/É/g,'E').replace(/Í/g,'I')
      .replace(/Ó/g,'O').replace(/Ú/g,'U').replace(/Ü/g,'U')
      .replace(/ñ/g,'n').replace(/Ñ/g,'N')
      .replace(/·/g,'-').replace(/[^\x00-\x7F]/g,'');
  }
  var asunto = _ascii('Devolucion de ' + tipoActivo + ' - ' + datos.tipoDev + ' - ' + nombre);

  // ── Colores del diseño corporativo SHARF — Paleta oficial ────────────
  var C_DARK   = '#68002B';   // Rojo Intenso PMS 207 — header
  var C_BRAND  = '#FF6568';   // Rojo Bold PMS 178
  var C_PALE   = '#FFC1C2';   // Rojo Pálido (40% Bold)
  var C_HDR_BG = '#68002B';   // Rojo Intenso para sub-headers
  var C_TEXT   = '#1a0a0e';   // casi negro con tinte vino
  var C_MUTED  = '#5c3040';   // vino suave
  var C_BORDER = '#f0d6d8';   // borde rosado suave
  var C_ALT    = '#fff8f8';   // fondo alternado
  // Fuente Manrope (con fallback web-safe)
  var FONT     = "'Manrope','Segoe UI',Arial,sans-serif";

  // ── Logo en header ────────────────────────────────────────────────────
  var logoHeaderHtml = logoData
    ? '<img src="' + logoData + '" style="height:32px;vertical-align:middle;margin-right:14px">'
    : '<span style="font-family:Manrope,Arial,sans-serif;font-weight:900;font-style:italic;font-size:22px;color:#ffffff;letter-spacing:-1px">sharf</span>';

  // ── Saludo ────────────────────────────────────────────────────────────
  var primerNombre = nombre.split(',')[1]
    ? nombre.split(',')[1].trim().split(' ')[0]
    : nombre.split(' ')[0];

  // ── Función helper fila de tabla ──────────────────────────────────────
  function tr(label, val, alt, color) {
    var bg = color ? color : (alt ? C_ALT : '#ffffff');
    return '<tr>' +
      '<td style="padding:7px 12px;font-weight:600;color:' + C_TEXT + ';background:' + bg + ';width:38%;border-bottom:1px solid ' + C_BORDER + '">' + label + '</td>' +
      '<td style="padding:7px 12px;color:' + C_TEXT + ';background:' + bg + ';border-bottom:1px solid ' + C_BORDER + '">' + (val || '—') + '</td>' +
    '</tr>';
  }
  function thRow(t1, t2, t3) {
    var style = 'padding:8px 12px;font-weight:700;color:#ffffff;background:' + C_HDR_BG + ';text-align:left';
    return '<tr><th style="' + style + '">' + t1 + '</th><th style="' + style + '">' + t2 + '</th>' +
      (t3 !== undefined ? '<th style="' + style + '">' + t3 + '</th>' : '') + '</tr>';
  }
  function tbl(content) {
    return '<table style="width:100%;border-collapse:collapse;font-size:9.5pt;border:1px solid ' + C_BORDER + ';margin:12px 0">' + content + '</table>';
  }
  function secTitle(txt) {
    return '<h4 style="font-size:10pt;color:' + C_TEXT + ';margin:18px 0 6px;' +
      'border-bottom:2.5px solid ' + C_BRAND + ';padding-bottom:5px;font-weight:700">' + txt + '</h4>';
  }

  // ── Tabla: Datos del colaborador ──────────────────────────────────────
  var tablaDatos = tbl(
    '<thead>' + thRow('Campo', 'Dato') + '</thead><tbody>' +
    tr('Nombre completo',   nombre,              false) +
    tr('DNI / CI',          emp.dni || '—',      true)  +
    tr('Empresa',           emp.empresa || '—',  false) +
    tr('Área',              emp.area    || '—',  true)  +
    tr('Cargo',             emp.cargo   || '—',  false) +
    tr('Sede / Ubicación',  emp.sede    || '—',  true)  +
    tr('Centro de Costos',  emp.ceco    || '—',  false) +
    tr('Tipo de devolución',datos.tipoDev || '—',true)  +
    '</tbody>'
  );

  // ── Tabla: Equipo devuelto ────────────────────────────────────────────
  var estadoEquipo = datos.hayObservaciones
    ? '<span style="color:#b45309;font-weight:700">⚠️ Con observaciones</span>'
    : '<span style="color:#16a34a;font-weight:700">✅ Sin observaciones</span>';
  var obsEquipo = datos.observacionDesc
    ? datos.observacionDesc
    : (datos.hayObservaciones ? 'Ver fotos en documento adjunto' : 'Sin observaciones');

  var tablaEquipo = tbl(
    '<thead>' + thRow('Artículo', 'Detalle', 'Observaciones') + '</thead><tbody>' +
    '<tr>' +
      '<td style="padding:7px 12px;color:' + C_TEXT + ';border-bottom:1px solid ' + C_BORDER + '">' + tipoActivo + '</td>' +
      '<td style="padding:7px 12px;color:' + C_TEXT + ';border-bottom:1px solid ' + C_BORDER + '">' +
        '<strong>S/N:</strong> ' + serial + '<br>' +
        '<span style="font-size:9pt;color:' + C_MUTED + '">' + (activo.nombre || '') + '</span>' +
      '</td>' +
      '<td style="padding:7px 12px;color:' + C_TEXT + ';border-bottom:1px solid ' + C_BORDER + '">' +
        estadoEquipo + '<br><span style="font-size:9pt;color:' + C_MUTED + '">' + obsEquipo + '</span>' +
      '</td>' +
    '</tr>' +
    '</tbody>'
  );

  // ── Tabla: Accesorios devueltos + estado + observaciones ─────────────
  // Incluye: accesorios de Snipe + accesorios nuevos marcados por el técnico
  var accEst  = datos.accesoriosEstado   || {};
  var accNuev = datos.accesoriosNuevos   || {};
  var accRows = '';
  var accItems = [
    { key:'cargador', label:'Cargador de Laptop', detalle: accData.cargadorSerial || '' },
    { key:'mouse',    label:'Mouse',              detalle: accData.mouseDesc       || '' },
    { key:'mochila',  label:'Mochila',            detalle: accData.mochilaDesc     || '' },
    { key:'docking',  label:'Docking Station',    detalle: accData.dockingDesc     || '' },
    { key:'teclado',  label:'Teclado',            detalle: accData.tecladoDesc     || '' }
  ];

  var rowIdx = 0;
  accItems.forEach(function(it) {
    var tieneEnSnipe = it.key === 'cargador' ? !!accData.cargadorSerial
                     : !!(accData[it.key + 'Desc'] || accData[it.key]);
    var esNuevo      = !!accNuev[it.key];
    var devuelto     = !!acc[it.key];

    // Solo mostrar si estaba en Snipe O si el técnico lo marcó explícitamente
    if (!tieneEnSnipe && !esNuevo && !devuelto) return;

    var obs      = (accObs[it.key] || '').trim();
    var estadoAcc = (accEst[it.key] || 'bueno') === 'bueno';
    var detalle  = it.detalle || (esNuevo ? 'Registrado en esta devolución' : '—');

    // Buscar compromiso diferido para este accesorio
    var fechaCompromiso = '';
    if (!devuelto && datos.compromisosDiferidos) {
      datos.compromisosDiferidos.forEach(function(cd) {
        if (cd.label && cd.label.toUpperCase().indexOf(it.label.toUpperCase().split(' ')[0]) >= 0) {
          fechaCompromiso = cd.fechaCompromiso || '';
        }
      });
    }

    var estadoHtml;
    if (!devuelto && fechaCompromiso) {
      estadoHtml = '<span style="color:#c2410c;font-weight:700">⏰ Entrega pendiente</span>' +
                   '<br><span style="font-size:9pt;color:#c2410c">Compromiso: ' + fechaCompromiso + '</span>';
    } else if (!devuelto) {
      estadoHtml = '<span style="color:#dc2626;font-weight:700">❌ No entregado</span>';
    } else if (estadoAcc) {
      estadoHtml = '<span style="color:#16a34a;font-weight:700">✅ Entregado · Buen estado</span>';
    } else {
      estadoHtml = '<span style="color:#b45309;font-weight:700">⚠️ Entregado · Con observaciones</span>';
    }

    if (obs) {
      estadoHtml += '<br><span style="font-size:9pt;color:' + C_MUTED + '">' + obs + '</span>';
    }

    var bg = rowIdx % 2 === 0 ? '#ffffff' : C_ALT;
    rowIdx++;
    accRows +=
      '<tr>' +
        '<td style="padding:7px 12px;font-weight:600;color:' + C_TEXT + ';background:' + bg + ';border-bottom:1px solid ' + C_BORDER + '">' + it.label + '</td>' +
        '<td style="padding:7px 12px;color:' + C_MUTED + ';background:' + bg + ';border-bottom:1px solid ' + C_BORDER + ';font-size:9.5pt">' + detalle + '</td>' +
        '<td style="padding:7px 12px;background:' + bg + ';border-bottom:1px solid ' + C_BORDER + '">' + estadoHtml + '</td>' +
      '</tr>';
  });

  var hayAccRows = rowIdx > 0;

  // ── Bloque: artículos NO devueltos hoy (compromisos) ──────────────────
  var compromisoHtml = '';
  if (datos.compromisosDiferidos && datos.compromisosDiferidos.length) {
    var compRows = datos.compromisosDiferidos.map(function(cd) {
      return '<li style="margin:4px 0">' +
        '<strong>' + cd.label + '</strong>' +
        (cd.serial ? ' — S/N: ' + cd.serial : '') +
        (cd.fechaCompromiso ? ' — <strong>Fecha compromiso: ' + cd.fechaCompromiso + '</strong>' : '') +
      '</li>';
    }).join('');
    compromisoHtml =
      '<div style="background:#fff7ed;border:1.5px solid #fed7aa;border-radius:8px;padding:14px 16px;margin:14px 0">' +
        '<p style="margin:0 0 8px;font-weight:700;color:#c2410c;font-size:10pt">⏰ Artículos pendientes de entrega posterior</p>' +
        '<ul style="margin:0;padding-left:18px;font-size:10pt;color:' + C_TEXT + '">' + compRows + '</ul>' +
        '<p style="margin:8px 0 0;font-size:9pt;color:' + C_MUTED + '">' +
          'El colaborador se comprometió a entregar los artículos indicados en las fechas señaladas.' +
        '</p>' +
      '</div>';
  }

  // ── Observaciones generales (si las hay) ──────────────────────────────
  var obsHtml = '';
  if (datos.hayObservaciones && datos.observacionDesc) {
    obsHtml =
      '<div style="background:#fef2f2;border:1.5px solid #fecaca;border-radius:8px;padding:14px 16px;margin:14px 0">' +
        '<p style="margin:0 0 6px;font-weight:700;color:#b91c1c;font-size:10pt">⚠️ Observaciones registradas durante la recepción</p>' +
        '<p style="margin:0;font-size:10pt;color:' + C_TEXT + '">' + datos.observacionDesc + '</p>' +
        (datos.photos && datos.photos.length ? '<p style="margin:8px 0 0;font-size:9pt;color:' + C_MUTED + '">Se adjuntan ' + datos.photos.length + ' foto(s) de evidencia en el documento adjunto.</p>' : '') +
      '</div>';
  }

  // ── Firma técnico (solo nombre — sin imagen) ──────────────────────────
  var firmaHtml =
    '<div style="margin-top:24px;padding-top:14px;border-top:1px solid ' + C_BORDER + ';font-family:' + FONT + ';font-size:9pt;color:' + C_MUTED + '">' +
      '<strong style="font-size:11pt;color:' + C_TEXT + '">' + (tec.nombre || 'Equipo de Soporte TI') + '</strong><br>' +
      '<span style="font-style:italic">Soporte TI · Mesa de Servicio SHARF</span><br>' +
      '<a href="https://www.holasharf.com" style="color:' + C_BRAND + ';text-decoration:none">www.holasharf.com</a>' +
    '</div>';

  // ── Armado del correo ─────────────────────────────────────────────────
  var htmlBody =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:0 auto;background:#f3f4f6;padding:12px">' +

    // HEADER — solo logo sobre fondo blanco, sin tagline
    '<div style="background:#ffffff;border-radius:10px 10px 0 0;padding:16px 24px;border:1px solid ' + C_BORDER + ';border-bottom:none">' +
      '<div style="display:table-cell;vertical-align:middle">' + logoHeaderHtml + '</div>' +
    '</div>' +

    // CUERPO
    '<div style="background:#ffffff;border:1px solid ' + C_BORDER + ';border-top:none;border-radius:0 0 10px 10px;padding:26px 28px">' +

      // Saludo
      '<p style="font-size:11pt;color:' + C_TEXT + ';margin:0 0 6px"><strong>Estimado/a ' + primerNombre + ',</strong></p>' +
      '<p style="font-size:10pt;color:' + C_TEXT + ';line-height:1.6;margin:0 0 14px">' +
        'A continuación encontrará el cargo de devolución de sus activos de TI. ' +
        'Este documento queda registrado en nuestro sistema.' +
      '</p>' +

      // Datos del colaborador
      secTitle('Datos del colaborador') + tablaDatos +

      // Equipo devuelto
      secTitle('Equipo devuelto') + tablaEquipo +

      // Accesorios
      (hayAccRows ? secTitle('Accesorios entregados') + tbl('<thead>' + thRow('Accesorio','Detalle','Estado / Observaciones') + '</thead><tbody>' + accRows + '</tbody>') : '') +

      // Pendientes de entrega
      compromisoHtml +

      // Observaciones generales
      obsHtml +

      // ID devolución
      '<p style="font-size:8pt;color:' + C_MUTED + ';margin-top:20px;border-top:1px solid ' + C_BORDER + ';padding-top:10px">' +
        'ID Devolución: <strong>' + devId + '</strong> &nbsp;·&nbsp; ' + ts +
      '</p>' +

      firmaHtml +
    '</div>' +

    // Footer
    '<div style="text-align:center;padding:12px;font-size:8pt;color:#9ca3af">' +
      'SHARF — Sistema de Gestión de Activos TI' +
    '</div>' +

  '</div>';

  // ── Modo Prueba ───────────────────────────────────────────────────────
  var mp = _getModoPrueba_();
  var esModoPrueba = mp.activo;

  // ── Destinatarios por tipo de devolución ──────────────────────────────
  //
  //  Por cese           → Para: personal (editable S3B) | CC: jefe + sheily + melanie + anais
  //  Cambio de renting  → Para: corp     (editable S3B) | CC: jefe + gabriel + anais
  //  Cambio de equipo   → Para: corp     (editable S3B) | CC: jefe + anais
  //  No usará           → Para: corp     (editable S3B) | CC: jefe + gabriel + anais
  //
  var tipoDvLower = (datos.tipoDev || '').toLowerCase();
  var esCese      = tipoDvLower.indexOf('cese')     >= 0;
  var esRenting   = tipoDvLower.indexOf('renting')  >= 0;
  var esNoUsara   = tipoDvLower.indexOf('no usar')  >= 0;
  // Cambio de equipo = cualquier otro tipo que no sea los anteriores
  var esCambioEq  = !esCese && !esRenting && !esNoUsara;

  var emailCorp   = emp.emailCorp || '';              // "Email Corp" del Sheet RH
  var emailPers   = emp.emailPersonal || '';           // correo personal
  var destFinalOriginal, ccOriginal;

  if (esCese) {
    // Para: correo personal (editable en S3B)
    destFinalOriginal = emailPers || emailJefe;
    ccOriginal = [];
    if (emailJefe && emailJefe !== destFinalOriginal) ccOriginal.push(emailJefe);
    ccOriginal.push(CFG.EMAIL_SHEILY);
    ccOriginal.push(CFG.EMAIL_MELANIE);
    ccOriginal.push(CFG.EMAIL_SUPERVISORA);

  } else if (esRenting) {
    // Para: correo corporativo "Email Corp" (editable en S3B)
    destFinalOriginal = emailCorp || emailPers || emailJefe;
    ccOriginal = [];
    if (emailJefe && emailJefe !== destFinalOriginal) ccOriginal.push(emailJefe);
    ccOriginal.push(CFG.EMAIL_GABRIEL_HD);   // reemplaza a Sheily y Melanie
    ccOriginal.push(CFG.EMAIL_SUPERVISORA);

  } else if (esNoUsara) {
    // Para: correo corporativo "Email Corp" (editable en S3B)
    destFinalOriginal = emailCorp || emailPers || emailJefe;
    ccOriginal = [];
    if (emailJefe && emailJefe !== destFinalOriginal) ccOriginal.push(emailJefe);
    ccOriginal.push(CFG.EMAIL_GABRIEL_HD);
    ccOriginal.push(CFG.EMAIL_SUPERVISORA);

  } else {
    // Cambio de equipo y otros
    // Para: correo corporativo "Email Corp" (editable en S3B)
    destFinalOriginal = emailCorp || emailPers || emailJefe;
    ccOriginal = [];
    if (emailJefe && emailJefe !== destFinalOriginal) ccOriginal.push(emailJefe);
    ccOriginal.push(CFG.EMAIL_SUPERVISORA);
  }

  var bccOriginal = [];

  var destFinal, ccFinal, bccFinal;
  if (esModoPrueba) {
    destFinal = mp.emailPara || destFinalOriginal;
    ccFinal   = mp.emailCC   || '';
    bccFinal  = mp.emailBCC  || '';
    Logger.log('⚠️ MODO PRUEBA: Para=' + destFinal + ' CC=' + ccFinal);
  } else if (datos._overridePara) {
    // Usar destinatarios confirmados/editados en la pantalla S3B del frontend
    destFinal = datos._overridePara;
    ccFinal   = datos._overrideCC   || ccOriginal.filter(Boolean).join(',');
    bccFinal  = '';
    Logger.log('Override S3B: Para=' + destFinal + ' CC=' + ccFinal);
  } else {
    destFinal = destFinalOriginal;
    ccFinal   = ccOriginal.filter(Boolean).join(',');
    bccFinal  = bccOriginal.join(',');
  }

  Logger.log('Email destinatarios [' + datos.tipoDev + ']: Para=' + destFinal + ' CC=' + ccFinal);

  var opciones = {
    htmlBody: htmlBody,
    name:     'SHARF - Soporte TI' + (esModoPrueba ? ' [PRUEBA]' : ''),
    replyTo:  CFG.EMAIL_HELPDESK
  };
  if (pdfBlob)  opciones.attachments = [pdfBlob];
  if (ccFinal)  opciones.cc  = ccFinal;
  if (bccFinal) opciones.bcc = bccFinal;

  GmailApp.sendEmail(destFinal, asunto + (esModoPrueba ? ' [PRUEBA]' : ''), '', opciones);

  return {
    ok:           true,
    modoPrueba:   esModoPrueba,
    para:         destFinal,
    paraOriginal: destFinalOriginal,
    cc:           ccFinal,
    ccOriginal:   ccOriginal.join(','),
    bcc:          bccFinal,
    bccOriginal:  bccOriginal.join(',')
  };
}

// ════════════════════════════════════════════════════════════════════════
// 12. CORREO DE SOLICITUD DE COTIZACIÓN DE ACCESORIOS
//   Se envía cuando hay accesorios marcados con "Se cotizará".
//   Para: alex.cano@holasharf.com
//   CC:   helpdesk@holasharf.com, anais.chero@holasharf.com,
//         gabriel.helpdesk@holasharf.com
//   Respeta modo prueba.
// ════════════════════════════════════════════════════════════════════════
function _enviarCorreoCotizacion(datos, ts, nombre, tipoActivo, itemsACotizar, logoData) {
  var emp    = datos.empleado || {};
  var activo = datos.activo   || {};

  var LABELS = {
    cargador: 'Cargador de Laptop',
    mouse:    'Mouse',
    mochila:  'Mochila',
    docking:  'Docking Station',
    teclado:  'Teclado'
  };

  var FONT   = "'Manrope','Segoe UI',Arial,sans-serif";
  var C_DARK = '#68002B';
  var C_BRAND= '#FF6568';
  var C_PALE = '#FFC1C2';
  var C_TEXT = '#1a0a0e';
  var C_MUTED= '#5c3040';
  var C_BORDER='#f0d6d8';

  // Asunto
  var asunto = 'SOLICITUD DE COTIZACION DE ACCESORIOS - ' +
               (datos.tipoDev || tipoActivo || 'DEVOLUCION') + ' - ' + nombre;

  // Lista de accesorios a cotizar en HTML
  var listaItems = itemsACotizar.map(function(k) {
    return '<li style="margin:6px 0;font-size:10pt;color:' + C_TEXT + '"><strong>' +
           (LABELS[k] || k) + '</strong></li>';
  }).join('');

  var listaTexto = itemsACotizar.map(function(k){ return LABELS[k] || k; }).join(', ');

  // Logo
  var logoHtml = logoData
    ? '<img src="' + logoData + '" style="height:30px;vertical-align:middle;margin-right:12px">'
    : '<span style="font-family:Manrope,Arial;font-weight:900;font-style:italic;font-size:20px;color:#fff;letter-spacing:-1px">sharf</span>';

  var htmlBody =
    '<div style="font-family:' + FONT + ';max-width:620px;margin:0 auto;background:#f3f4f6;padding:12px">' +

    // Header
    '<div style="background:' + C_DARK + ';border-radius:10px 10px 0 0;padding:14px 24px;' +
      'display:flex;align-items:center;border-bottom:3px solid ' + C_BRAND + '">' +
      logoHtml +
      '<div>' +
        '<div style="color:#fff;font-size:11pt;font-weight:700">SHARF — Soporte TI</div>' +
        '<div style="color:' + C_PALE + ';font-size:8pt">Sistema de Gestión de Activos</div>' +
      '</div>' +
    '</div>' +

    // Cuerpo
    '<div style="background:#fff;border:1px solid ' + C_BORDER + ';border-top:none;border-radius:0 0 10px 10px;padding:26px 28px">' +

      '<p style="font-size:11pt;color:' + C_TEXT + ';margin:0 0 6px">' +
        '<strong>Estimado equipo de Logística / Administración,</strong>' +
      '</p>' +

      '<p style="font-size:10pt;color:' + C_TEXT + ';line-height:1.7;margin:0 0 16px">' +
        'Por medio del presente, el área de <strong>Soporte TI</strong> solicita formalmente la cotización ' +
        'de los accesorios de cómputo que se detallan a continuación, los cuales no fueron devueltos ' +
        'en el proceso de devolución de activos registrado en el sistema.' +
      '</p>' +

      // Datos del colaborador
      '<h4 style="font-size:10pt;color:' + C_TEXT + ';border-bottom:2px solid ' + C_BRAND + ';' +
        'padding-bottom:5px;margin:0 0 10px;font-weight:700">Datos del colaborador</h4>' +
      '<table style="width:100%;border-collapse:collapse;font-size:9.5pt;margin-bottom:16px">' +
        '<tr><td style="padding:5px 10px;font-weight:600;color:' + C_TEXT + ';background:#fff8f8;' +
          'border:1px solid ' + C_BORDER + ';width:35%">Nombre completo</td>' +
          '<td style="padding:5px 10px;border:1px solid ' + C_BORDER + '">' + nombre + '</td></tr>' +
        '<tr><td style="padding:5px 10px;font-weight:600;color:' + C_TEXT + ';background:#fff;' +
          'border:1px solid ' + C_BORDER + '">DNI / CI</td>' +
          '<td style="padding:5px 10px;border:1px solid ' + C_BORDER + '">' + (emp.dni||'—') + '</td></tr>' +
        '<tr><td style="padding:5px 10px;font-weight:600;color:' + C_TEXT + ';background:#fff8f8;' +
          'border:1px solid ' + C_BORDER + '">Área / Cargo</td>' +
          '<td style="padding:5px 10px;border:1px solid ' + C_BORDER + '">' + (emp.area||'—') + ' / ' + (emp.cargo||'—') + '</td></tr>' +
        '<tr><td style="padding:5px 10px;font-weight:600;color:' + C_TEXT + ';background:#fff;' +
          'border:1px solid ' + C_BORDER + '">Tipo de devolución</td>' +
          '<td style="padding:5px 10px;border:1px solid ' + C_BORDER + '">' + (datos.tipoDev||'—') + '</td></tr>' +
        '<tr><td style="padding:5px 10px;font-weight:600;color:' + C_TEXT + ';background:#fff8f8;' +
          'border:1px solid ' + C_BORDER + '">Equipo devuelto</td>' +
          '<td style="padding:5px 10px;border:1px solid ' + C_BORDER + '">' + tipoActivo + ' — S/N: ' + (activo.serial||'—') + '</td></tr>' +
        '<tr><td style="padding:5px 10px;font-weight:600;color:' + C_TEXT + ';background:#fff;' +
          'border:1px solid ' + C_BORDER + '">Fecha de registro</td>' +
          '<td style="padding:5px 10px;border:1px solid ' + C_BORDER + '">' + ts + '</td></tr>' +
      '</table>' +

      // Accesorios a cotizar
      '<h4 style="font-size:10pt;color:' + C_TEXT + ';border-bottom:2px solid ' + C_BRAND + ';' +
        'padding-bottom:5px;margin:0 0 10px;font-weight:700">Accesorios que requieren cotización</h4>' +
      '<div style="background:#fff7ed;border:1.5px solid #fed7aa;border-radius:8px;padding:14px 18px;margin-bottom:16px">' +
        '<p style="margin:0 0 8px;font-weight:700;color:#c2410c;font-size:10pt">💰 Artículos no devueltos / con daño total:</p>' +
        '<ul style="margin:0;padding-left:20px">' + listaItems + '</ul>' +
        '<p style="margin:10px 0 0;font-size:9pt;color:' + C_MUTED + '">' +
          'Se solicita cotización de reposición por unidad para cada accesorio indicado.' +
        '</p>' +
      '</div>' +

      // Solicitud formal
      '<p style="font-size:10pt;color:' + C_TEXT + ';line-height:1.7;margin:0 0 16px">' +
        'Se solicita amablemente que la cotización sea remitida al área de Soporte TI ' +
        'a la brevedad posible, indicando precio unitario, disponibilidad y tiempo de entrega ' +
        'para cada artículo listado, a fin de proceder con el descuento correspondiente en la ' +
        'liquidación de beneficios del colaborador.' +
      '</p>' +

      '<p style="font-size:10pt;color:' + C_TEXT + ';line-height:1.7;margin:0 0 20px">' +
        'Agradecemos su atención y quedamos a disposición para cualquier consulta.' +
      '</p>' +

      // Firma
      '<div style="border-top:1px solid ' + C_BORDER + ';padding-top:14px;font-size:9pt;color:' + C_MUTED + '">' +
        '<strong style="font-size:11pt;color:' + C_TEXT + '">Área de Soporte TI</strong><br>' +
        'Mesa de Servicio SHARF<br>' +
        '<a href="https://www.holasharf.com" style="color:' + C_BRAND + ';text-decoration:none">www.holasharf.com</a>' +
      '</div>' +

    '</div>' +
    '<div style="text-align:center;padding:10px;font-size:8pt;color:#9ca3af">' +
      'SHARF — Sistema de Gestión de Activos TI · ' + ts +
    '</div>' +
    '</div>';

  // ── Destinatarios fijos ───────────────────────────────────────────────
  var DEST_COT = 'alex.cano@holasharf.com';
  var CC_COT   = [
    CFG.EMAIL_HELPDESK,          // helpdesk@holasharf.com
    CFG.EMAIL_SUPERVISORA,       // anais.chero@holasharf.com
    CFG.EMAIL_GABRIEL_HD         // gabriel.helpdesk@holasharf.com
  ].filter(Boolean).join(',');

  // ── Aplicar modo prueba si está activo ───────────────────────────────
  var mp = _getModoPrueba_();
  var destCot, ccCot;
  if (mp.activo) {
    destCot = mp.emailPara || DEST_COT;
    ccCot   = mp.emailCC   || '';
    Logger.log('⚠️ Cotización MODO PRUEBA → Para=' + destCot);
  } else {
    destCot = DEST_COT;
    ccCot   = CC_COT;
  }

  var opCot = {
    htmlBody: htmlBody,
    name:     'SHARF - Soporte TI' + (mp.activo ? ' [PRUEBA]' : ''),
    replyTo:  CFG.EMAIL_HELPDESK
  };
  if (ccCot) opCot.cc = ccCot;

  GmailApp.sendEmail(destCot, asunto + (mp.activo ? ' [PRUEBA]' : ''), '', opCot);

  Logger.log('Correo cotización enviado: Para=' + destCot + ' CC=' + ccCot +
             ' · Accesorios: ' + listaTexto);

  return { ok: true, para: destCot, cc: ccCot, accesorios: listaTexto };
}

// ════════════════════════════════════════════════════════════════════════
// 12. ALERTA SNIPE DESACTUALIZADO
// ════════════════════════════════════════════════════════════════════════
function _enviarAlertaSnipeDesactualizado(datos, ts, nombreReal, serial, tipoActivo) {
  var viejoProp = datos.snipeViejoPropietario || {};
  var tec       = datos.tecnico               || {};
  var emp       = datos.empleado              || {};

  // Propietario real en Snipe (a quien estaba asignado)
  var propNombre = viejoProp.nombre || '—';
  var propEmail  = viejoProp.email  || '—';
  var propDni    = viejoProp.dni    || '—';

  // Quien devuelve según el técnico que recibe
  var devNombre  = nombreReal               || '—';
  var devDni     = emp.dni                  || '—';
  var devEmail   = emp.emailCorp || emp.emailPersonal || '—';
  var devArea    = emp.area                 || '—';

  // Técnico que procesó la devolución
  var tecNombre  = tec.nombre || '—';
  var tecEmail   = tec.email  || '—';
  var tecSede    = tec.sede   || '—';

  var asunto = 'ALERTA SNIPE DESACTUALIZADO: Equipo devuelto por persona no asignada — ' + serial;

  var body =
    '<div style="font-family:Manrope,Arial,sans-serif;max-width:620px;margin:0 auto">' +

    // Header rojo alerta
    '<div style="background:#68002B;padding:18px 24px;border-radius:8px 8px 0 0;display:flex;align-items:center;gap:12px">' +
      '<div style="font-size:28px">🚨</div>' +
      '<div>' +
        '<h2 style="color:#fff;margin:0;font-size:14pt;font-weight:800">ALERTA: Inconsistencia en Inventario Snipe-IT</h2>' +
        '<div style="color:#FFC1C2;font-size:9.5pt;margin-top:3px">Equipo devuelto por persona diferente a la asignada en el sistema</div>' +
      '</div>' +
    '</div>' +

    '<div style="background:#fff;border:1.5px solid #f0d6d8;border-top:none;padding:22px;border-radius:0 0 8px 8px">' +

    // Bloque rojo: equipo
    '<div style="background:#fff5f5;border-left:4px solid #FF6568;padding:12px 16px;border-radius:4px;margin-bottom:16px">' +
      '<div style="font-size:9pt;color:#9ca3af;text-transform:uppercase;font-weight:700;margin-bottom:6px">Activo involucrado</div>' +
      '<table style="width:100%;border-collapse:collapse;font-size:10pt">' +
      '<tr><td style="padding:4px 8px 4px 0;color:#5c3040;font-weight:600;width:40%">Serial:</td>' +
          '<td style="padding:4px 0;font-family:monospace;font-weight:700;color:#68002B">' + serial + '</td></tr>' +
      '<tr><td style="padding:4px 8px 4px 0;color:#5c3040;font-weight:600">Tipo de equipo:</td>' +
          '<td style="padding:4px 0">' + tipoActivo + '</td></tr>' +
      '<tr><td style="padding:4px 8px 4px 0;color:#5c3040;font-weight:600">Tipo de devolución:</td>' +
          '<td style="padding:4px 0">' + (datos.tipoDev || '—') + '</td></tr>' +
      '</table>' +
    '</div>' +

    // Tabla de involucrados
    '<div style="font-size:9pt;color:#9ca3af;text-transform:uppercase;font-weight:700;margin-bottom:8px">Personas involucradas</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:10pt;border:1px solid #f0d6d8;margin-bottom:16px">' +
    '<thead><tr style="background:#68002B">' +
      '<th style="padding:8px 12px;color:#fff;text-align:left;font-weight:700">Rol</th>' +
      '<th style="padding:8px 12px;color:#fff;text-align:left;font-weight:700">Nombre</th>' +
      '<th style="padding:8px 12px;color:#fff;text-align:left;font-weight:700">DNI / Email</th>' +
    '</tr></thead><tbody>' +

    '<tr style="background:#fff5f5">' +
      '<td style="padding:8px 12px;border-bottom:1px solid #f0d6d8;font-weight:700;color:#b91c1c">⚠️ Propietario en Snipe<br><span style="font-size:8.5pt;font-weight:400;color:#9ca3af">(a quien estaba asignado)</span></td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #f0d6d8;font-weight:700;color:#b91c1c">' + propNombre + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #f0d6d8;color:#b91c1c">' + propDni + '<br><span style="font-size:9pt">' + propEmail + '</span></td>' +
    '</tr>' +

    '<tr>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #f0d6d8;font-weight:700;color:#1a0a0e">📦 Quien devuelve el equipo<br><span style="font-size:8.5pt;font-weight:400;color:#9ca3af">(persona que entrega físicamente)</span></td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #f0d6d8;font-weight:700">' + devNombre + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #f0d6d8">' + devDni + ' · Área: ' + devArea + '<br><span style="font-size:9pt">' + devEmail + '</span></td>' +
    '</tr>' +

    '<tr style="background:#fdf5f6">' +
      '<td style="padding:8px 12px;font-weight:700;color:#1a0a0e">🔧 Técnico TI que recibe<br><span style="font-size:8.5pt;font-weight:400;color:#9ca3af">(procesó la devolución)</span></td>' +
      '<td style="padding:8px 12px;font-weight:700">' + tecNombre + '</td>' +
      '<td style="padding:8px 12px">' + tecSede + '<br><span style="font-size:9pt">' + tecEmail + '</span></td>' +
    '</tr>' +

    '</tbody></table>' +

    // Acción automática
    '<div style="background:#fef3c7;border:1.5px solid #fcd34d;border-radius:6px;padding:14px 16px;margin-bottom:14px">' +
      '<div style="font-size:9pt;color:#78350f;font-weight:700;margin-bottom:6px">⚡ Acciones ejecutadas automáticamente</div>' +
      '<ul style="margin:0;padding-left:18px;font-size:9.5pt;color:#78350f;line-height:1.7">' +
        '<li>El activo fue <strong>desvinculado</strong> del propietario anterior en Snipe-IT</li>' +
        '<li>El estado del activo fue cambiado a <strong>Disponible</strong></li>' +
        '<li>Se generó el acta de devolución y se registró en Drive</li>' +
      '</ul>' +
    '</div>' +

    // Llamado a acción
    '<div style="background:#fee2e2;border:1.5px solid #fca5a5;border-radius:6px;padding:14px 16px;margin-bottom:14px">' +
      '<div style="font-size:10pt;color:#991b1b;font-weight:800;margin-bottom:6px">🚨 ACCIÓN REQUERIDA</div>' +
      '<div style="font-size:9.5pt;color:#7f1d1d;line-height:1.7">' +
        'Este equipo fue devuelto por <strong>' + devNombre + '</strong> pero en Snipe-IT estaba registrado a nombre de <strong>' + propNombre + '</strong>. ' +
        'Esto indica una <strong>desactualización del sistema de inventario</strong>.<br><br>' +
        'Por favor:<br>' +
        '1. Verificar con <strong>' + devNombre + ' (' + devArea + ')</strong> por qué tenía el equipo.<br>' +
        '2. Contactar a <strong>' + propNombre + '</strong> para confirmar su situación con el activo.<br>' +
        '3. Actualizar el inventario en Snipe-IT según corresponda.' +
      '</div>' +
    '</div>' +

    '<p style="font-size:8pt;color:#9ca3af;margin:0">Generado automáticamente por SHARF · Sistema de Devoluciones TI · ' + ts + ' · ID Técnico: ' + tecNombre + '</p>' +

    '</div></div>';

  GmailApp.sendEmail(
    CFG.EMAIL_GABRIEL_HD,
    asunto, '',
    {
      htmlBody: body,
      name:     'SHARF · Sistema de Devoluciones TI',
      cc:       CFG.EMAIL_HELPDESK + ',' + CFG.EMAIL_SUPERVISORA
    }
  );
}

// ════════════════════════════════════════════════════════════════════════
// 13. LOG EN SHEET
// ════════════════════════════════════════════════════════════════════════
function _registrarLogDevolucion(devId, datos, ts, nombre, serial, tipoActivo, resultado, pdfId) {
  try {
    var hoja   = _sheetLog();
    var emp    = datos.empleado            || {};
    var activo = datos.activo              || {};
    var tec    = datos.tecnico             || {};
    var acc    = datos.accesoriosDevueltos || {};
    var accEst = datos.accesoriosEstado    || {};
    var accCot = datos.accesoriosCotizar   || {};
    var accObs = datos.accesoriosObs       || {};

    // Accesorios NO devueltos
    var noDevueltos = [];
    ['cargador','mouse','mochila','docking','teclado'].forEach(function(k){
      if(acc[k] === false || acc[k] === undefined && !(datos.accesorios && datos.accesorios[k])) return;
      if(!acc[k]) noDevueltos.push(k);
    });
    // Accesorios a cotizar
    var cotizados = Object.keys(accCot).filter(function(k){ return accCot[k]; });
    // Estado individual accesorios
    var estadoAcc = ['cargador','mouse','mochila','docking','teclado'].map(function(k){
      return (accEst[k] || 'bueno').toUpperCase();
    }).join(', ');

    var fila = [
      devId,
      ts,
      datos.tipoDev        || '',
      datos.modoIngreso    || '',
      nombre,
      emp.dni              || '',
      emp.empresa          || '',
      emp.area             || '',
      emp.cargo            || '',
      emp.sede             || '',
      tipoActivo,
      activo.nombre        || '',
      serial,
      activo.tag           || '',
      activo.id            || '',
      // Accesorios devueltos
      acc.cargador  ? 'SI' : 'NO',
      acc.mochila   ? 'SI' : 'NO',
      acc.mouse     ? 'SI' : 'NO',
      acc.docking   ? 'SI' : 'NO',
      acc.teclado   ? 'SI' : 'NO',
      // Estado equipo
      datos.hayObservaciones ? 'CON OBSERVACIONES' : 'SIN OBSERVACIONES',
      datos.observacionDesc  || '',
      datos.equipoBueno !== false ? 'BUENO' : 'CON DAÑO',
      // Estado individual accesorios
      estadoAcc,
      // No devueltos
      noDevueltos.length > 0 ? noDevueltos.join(', ') : 'Ninguno',
      noDevueltos.length,
      // Cotizaciones
      cotizados.length > 0 ? cotizados.join(', ') : 'Ninguna',
      cotizados.length,
      // Observaciones por accesorio
      JSON.stringify(accObs),
      // Técnico
      tec.nombre   || '',
      tec.email    || '',
      tec.sede     || '',
      // Resultados proceso
      resultado.snipeOk ? 'OK' : 'ERROR',
      resultado.pdfOk   ? 'OK' : 'ERROR',
      resultado.driveOk ? 'OK' : 'ERROR',
      resultado.emailOk ? 'OK' : 'ERROR',
      resultado.driveUrl || '',
      pdfId              || '',
      datos.snipeDesactualizado ? 'SÍ - ALERTA ENVIADA' : 'No'
    ];

    hoja.appendRow(fila);

    var ultima = hoja.getLastRow();
    if (datos.snipeDesactualizado) {
      hoja.getRange(ultima, 1, 1, fila.length).setBackground('#fee2e2');
    } else if (cotizados.length > 0) {
      hoja.getRange(ultima, 1, 1, fila.length).setBackground('#fff7ed');  // naranja si hay cotizaciones
    } else if (datos.hayObservaciones) {
      hoja.getRange(ultima, 1, 1, fila.length).setBackground('#fef3c7');
    }
  } catch(e) {
    Logger.log('_registrarLogDevolucion: ' + e.message);
    throw e;
  }
}

// ════════════════════════════════════════════════════════════════════════
// 14. HELPERS SHEET RH
//     Estructura confirmada por el cliente:
//       Fila 4 = encabezados, datos desde fila 5
//       Columnas (1-based):
//         1=DNI  2=APELLIDOS Y NOMBRES  3=EMPRESA  4=CARGO  5=ÁREA
//         6=SUB ÁREA  7=TELEFONO  8=CEL CORP  9=CENTRO DE OPERACIONES
//         10=CENTRO DE COSTO  11=DENOMINACION  12=email corp
//         13=email personal  14=FECHA DE CESE
// ════════════════════════════════════════════════════════════════════════

// Posiciones fijas (1-based) — fuente de verdad si el mapeo dinámico falla
var RH_COL = {
  DNI:          1,
  NOMBRE:       2,
  EMPRESA:      3,
  CARGO:        4,
  AREA:         5,
  SUB_AREA:     6,
  TELEFONO:     7,
  CEL_CORP:     8,
  CENTRO_OPS:   9,   // CENTRO DE OPERACIONES → sede
  CECO_COD:     10,  // CENTRO DE COSTO
  DENOMINACION: 11,
  EMAIL_CORP:   12,
  EMAIL_PERS:   13,
  FECHA_CESE:   14
};

function _buscarEmpleadoPorDni(dni) {
  if (!dni || !String(dni).trim()) return { ok: false, error: 'DNI vacío' };
  var dniBusq = String(dni).trim();

  var ss;
  try {
    ss = SpreadsheetApp.openById(CFG.SHEET_RH_ID);
  } catch(e) {
    Logger.log('_buscarEmpleadoPorDni: no se pudo abrir Sheet RH ' + CFG.SHEET_RH_ID + ': ' + e.message);
    return { ok: false, error: 'Sin acceso al Sheet RH (' + CFG.SHEET_RH_ID + '): ' + e.message };
  }

  for (var t = 0; t < CFG.SHEET_RH_TABS.length; t++) {
    var tabCfg = CFG.SHEET_RH_TABS[t];
    var hoja   = null;

    // Buscar pestaña por nombre, luego por GID
    try {
      hoja = ss.getSheetByName(tabCfg.nombre);
      if (!hoja && tabCfg.gid) {
        var todas = ss.getSheets();
        for (var s = 0; s < todas.length; s++) {
          if (String(todas[s].getSheetId()) === String(tabCfg.gid)) { hoja = todas[s]; break; }
        }
      }
    } catch(eTab) {
      Logger.log('Tab ' + tabCfg.nombre + ': ' + eTab.message);
      continue;
    }
    if (!hoja) {
      Logger.log('Tab no encontrada: ' + tabCfg.nombre + ' (GID ' + tabCfg.gid + ')');
      continue;
    }

    var lastRow   = hoja.getLastRow();
    var lastCol   = Math.max(hoja.getLastColumn(), RH_COL.FECHA_CESE);
    var dataStart = CFG.SHEET_RH_HDR_ROW + 1;  // fila 5
    if (lastRow < dataStart) {
      Logger.log('Tab ' + tabCfg.nombre + ': sin datos (lastRow=' + lastRow + ')');
      continue;
    }

    // ── Leer encabezados fila 4 ──────────────────────────────────────
    var hdrRow = hoja.getRange(CFG.SHEET_RH_HDR_ROW, 1, 1, lastCol).getValues()[0];
    var hdrs   = _hdrs(hdrRow);

    // ── Determinar columna DNI — tres niveles de prioridad ───────────
    // 1. Coincidencia exacta con el nombre declarado en tabCfg.colDni
    // 2. Detección dinámica por alias (_hdrs)
    // 3. Posición fija col 1 (RH_COL.DNI)
    var colDni = RH_COL.DNI - 1;  // fallback final
    if (tabCfg.colDni) {
      var colDniExacto = -1;
      for (var hi = 0; hi < hdrRow.length; hi++) {
        if (String(hdrRow[hi]).trim() === tabCfg.colDni.trim()) { colDniExacto = hi; break; }
      }
      if (colDniExacto >= 0) {
        colDni = colDniExacto;
        Logger.log('Tab ' + tabCfg.nombre + ': colDni por nombre exacto "' + tabCfg.colDni + '" → col ' + (colDni+1));
      } else if (hdrs['dni'] !== undefined) {
        colDni = hdrs['dni'];
        Logger.log('Tab ' + tabCfg.nombre + ': colDni por alias → col ' + (colDni+1) + ' "' + hdrRow[colDni] + '"');
      } else {
        Logger.log('Tab ' + tabCfg.nombre + ': colDni NO detectado — usando posición fija col ' + RH_COL.DNI);
        Logger.log('  Encabezado buscado: "' + tabCfg.colDni + '"');
        Logger.log('  Encabezados reales: ' + hdrRow.slice(0,15).map(function(h,i){return (i+1)+'="'+h+'"';}).join(', '));
      }
    } else if (hdrs['dni'] !== undefined) {
      colDni = hdrs['dni'];
    }

    Logger.log('Tab ' + tabCfg.nombre + ' | lastRow=' + lastRow +
               ' | colDni=' + (colDni+1) + ' | hdr="' + (hdrRow[colDni]||'') + '"');

    // ── Leer datos desde fila 5 ──────────────────────────────────────
    var numRows = lastRow - dataStart + 1;
    var data    = hoja.getRange(dataStart, 1, numRows, lastCol).getValues();

    for (var i = 0; i < data.length; i++) {
      var fila    = data[i];
      var filaDni = String(fila[colDni] || '').trim();

      // Coincidencia flexible: exacta o ignorando ceros a la izquierda
      if (filaDni === dniBusq ||
          filaDni === dniBusq.replace(/^0+/, '') ||
          dniBusq === filaDni.replace(/^0+/, '')) {

        Logger.log('DNI ' + dniBusq + ' encontrado en tab ' + tabCfg.nombre + ' fila ' + (dataStart + i));

        // ── Mapear fila: dinámica primero, posición fija como respaldo ──
        var emp = _mapearEmpleadoRH_(fila, hdrs);

        // Si el mapeo dinámico dejó campos vacíos, rellenar por posición fija
        if (!emp.dni)          emp.dni          = String(fila[RH_COL.DNI          - 1] || '').trim();
        if (!emp.nombre)       emp.nombre       = String(fila[RH_COL.NOMBRE       - 1] || '').trim();
        if (!emp.empresa)      emp.empresa      = String(fila[RH_COL.EMPRESA      - 1] || '').trim();
        if (!emp.cargo)        emp.cargo        = String(fila[RH_COL.CARGO        - 1] || '').trim();
        if (!emp.area)         emp.area         = String(fila[RH_COL.AREA         - 1] || '').trim();
        if (!emp.sede)         emp.sede         = String(fila[RH_COL.CENTRO_OPS   - 1] || '').trim();
        if (!emp.emailCorp)    emp.emailCorp    = String(fila[RH_COL.EMAIL_CORP   - 1] || '').trim();
        if (!emp.emailPersonal)emp.emailPersonal= String(fila[RH_COL.EMAIL_PERS   - 1] || '').trim();
        if (!emp.ceco) {
          var cecoPos = String(fila[RH_COL.CECO_COD     - 1] || '').trim();
          var denomPos= String(fila[RH_COL.DENOMINACION - 1] || '').trim();
          emp.ceco    = cecoPos && denomPos ? cecoPos + ' - ' + denomPos : cecoPos || denomPos;
        }

        emp.tabOrigen = tabCfg.nombre;
        return { ok: true, empleado: emp, tabOrigen: tabCfg.nombre };
      }
    }

    Logger.log('DNI ' + dniBusq + ' no encontrado en tab ' + tabCfg.nombre);
  }

  return { ok: false, noEncontrado: true,
           error: 'DNI ' + dniBusq + ' no encontrado en ninguna pestaña del Sheet RH.' };
}

// ── DIAGNÓSTICO DESDE EL FRONTEND ────────────────────────────────────────
// Llamar desde el app con handleRequest('diagnosticarSheetRH', {dni:'12345678'})
// Retorna un informe detallado de qué encontró o no en el Sheet.
function diagnosticarSheetRH(dniPrueba) {
  var R = { ok: true, pasos: [], sheetAccesible: false, pestanas: [], dniEncontrado: false };

  function log(msg) { R.pasos.push(msg); Logger.log(msg); }

  // 1. Abrir Sheet
  var ss;
  try {
    ss = SpreadsheetApp.openById(CFG.SHEET_RH_ID);
    R.sheetAccesible = true;
    log('✅ Sheet RH accesible: "' + ss.getName() + '"  ID: ' + CFG.SHEET_RH_ID);
  } catch(e) {
    log('❌ No se pudo abrir Sheet RH (' + CFG.SHEET_RH_ID + '): ' + e.message);
    log('   → Asegúrate de compartir el Sheet con la cuenta que ejecuta este script.');
    R.ok = false;
    return R;
  }

  // 2. Listar pestañas reales
  var todasHojas = ss.getSheets();
  log('Pestañas disponibles en el Sheet:');
  todasHojas.forEach(function(h) {
    log('  · "' + h.getName() + '"  GID:' + h.getSheetId() + '  filas:' + h.getLastRow());
    R.pestanas.push({ nombre: h.getName(), gid: h.getSheetId(), filas: h.getLastRow() });
  });

  // 3. Verificar cada pestaña configurada
  CFG.SHEET_RH_TABS.forEach(function(tabCfg) {
    log('');
    log('── Pestaña configurada: "' + tabCfg.nombre + '" (GID:' + tabCfg.gid + ')');

    var hoja = ss.getSheetByName(tabCfg.nombre);
    if (!hoja) {
      todasHojas.forEach(function(h) {
        if (String(h.getSheetId()) === String(tabCfg.gid)) hoja = h;
      });
    }
    if (!hoja) {
      log('  ❌ Pestaña no encontrada por nombre ni por GID.');
      log('  ¿Está en la lista de pestañas? ' + todasHojas.map(function(h){return '"'+h.getName()+'"';}).join(', '));
      return;
    }
    log('  ✅ Pestaña encontrada: "' + hoja.getName() + '"');

    var lastRow = hoja.getLastRow();
    var lastCol = hoja.getLastColumn();
    log('  Filas: ' + lastRow + '  Columnas: ' + lastCol);

    if (lastRow < CFG.SHEET_RH_HDR_ROW) {
      log('  ❌ Sin encabezados en fila ' + CFG.SHEET_RH_HDR_ROW);
      return;
    }

    // Leer encabezados
    var hdrRow = hoja.getRange(CFG.SHEET_RH_HDR_ROW, 1, 1, lastCol).getValues()[0];
    log('  Encabezados fila ' + CFG.SHEET_RH_HDR_ROW + ':');
    hdrRow.forEach(function(h, i) {
      if (h) log('    col ' + (i+1) + ': "' + h + '"');
    });

    var hdrs = _hdrs(hdrRow);
    var colDni = hdrs['dni'];
    if (colDni === undefined) {
      log('  ⚠️  No se detectó columna DNI dinámicamente.');
      log('  → Usando posición fija: col ' + RH_COL.DNI + ' = "' + (hdrRow[RH_COL.DNI-1]||'(vacío)') + '"');
      colDni = RH_COL.DNI - 1;
    } else {
      log('  ✅ Columna DNI detectada: col ' + (colDni+1) + ' = "' + hdrRow[colDni] + '"');
    }

    // Buscar el DNI de prueba
    if (!dniPrueba) { log('  (sin DNI de prueba — solo diagnóstico de estructura)'); return; }

    var dataStart = CFG.SHEET_RH_HDR_ROW + 1;
    if (lastRow < dataStart) { log('  Sin datos desde fila ' + dataStart); return; }

    var numRows = Math.min(lastRow - dataStart + 1, 5000);
    var data    = hoja.getRange(dataStart, 1, numRows, lastCol).getValues();

    var encontrado = false;
    // Mostrar los primeros 3 DNIs para diagnóstico de formato
    log('  Primeros 5 valores en col DNI (fila ' + dataStart + ' en adelante):');
    for (var k = 0; k < Math.min(5, data.length); k++) {
      log('    fila ' + (dataStart+k) + ': "' + String(data[k][colDni]||'') + '"');
    }

    for (var i = 0; i < data.length; i++) {
      var filaDni = String(data[i][colDni] || '').trim();
      if (filaDni === dniPrueba ||
          filaDni === dniPrueba.replace(/^0+/,'') ||
          dniPrueba === filaDni.replace(/^0+/,'')) {
        log('  ✅ DNI "' + dniPrueba + '" ENCONTRADO en fila ' + (dataStart+i));
        log('     Nombre: "' + String(data[i][RH_COL.NOMBRE-1]||'') + '"');
        log('     Email corp: "' + String(data[i][RH_COL.EMAIL_CORP-1]||'') + '"');
        log('     Email pers: "' + String(data[i][RH_COL.EMAIL_PERS-1]||'') + '"');
        encontrado = true;
        R.dniEncontrado = true;
        break;
      }
    }
    if (!encontrado) {
      log('  ❌ DNI "' + dniPrueba + '" NO encontrado en esta pestaña (' + numRows + ' filas revisadas)');
    }
  });

  log('');
  log(R.dniEncontrado
    ? '✅ RESULTADO: DNI encontrado en el Sheet RH'
    : '❌ RESULTADO: DNI no encontrado en ninguna pestaña');

  return R;
}

// ════════════════════════════════════════════════════════════════════════
// VALIDAR CONEXIONES — Sheet RH + Snipe-IT
// Llamar desde el frontend: handleRequest('validarConexiones', {})
// También ejecutable desde el editor GAS como función standalone.
// Retorna un objeto con el estado de cada conexión y un resumen legible.
// ════════════════════════════════════════════════════════════════════════
// AUTENTICACIÓN DE TÉCNICO POR CUENTA GOOGLE @holasharf.com
//
//   Funciona con el script en modo "Ejecutar como: Propietario" y
//   "Quién puede acceder: Cualquier usuario de Google".
//   Session.getActiveUser().getEmail() devuelve el email del técnico
//   cuando está logueado con su cuenta Google en el mismo navegador.
// ════════════════════════════════════════════════════════════════════════
function autenticarTecnico() {
  try {
    var email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch(ex) {}

    // ── Sin sesión Google activa ──────────────────────────────────────────
    if (!email) {
      var appUrl = '';
      try { appUrl = ScriptApp.getService().getUrl(); } catch(eu) {}
      return {
        ok:        false,
        sinSesion: true,
        appUrl:    appUrl,
        error:     'No hay sesión de Google activa. Inicia sesión con tu ' +
                   'cuenta @holasharf.com en Google y vuelve a abrir la aplicación.'
      };
    }

    var emailNorm = email.toLowerCase().trim();
    Logger.log('autenticarTecnico: ' + emailNorm);

    // ── Verificar dominio @holasharf.com ──────────────────────────────────
    if (!emailNorm.endsWith('@holasharf.com')) {
      return {
        ok: false,
        error: 'Solo cuentas @holasharf.com tienen acceso. Cuenta detectada: ' + email
      };
    }

    // ── Verificar que el usuario está en la lista autorizada ─────────────
    if (!_esAutorizado(emailNorm)) {
      return {
        ok: false, noRegistrado: true,
        error: 'La cuenta ' + email + ' no está autorizada. ' +
               'Contacta a Ismael o Anais para solicitar acceso.'
      };
    }

    // ── Buscar técnico registrado ─────────────────────────────────────────
    var TECNICOS = [
      { nombre:'Gabriel García Diaz',        firmaKey:'GABRIEL',  dni:'49058626', email:'gabriel.helpdesk@holasharf.com', sede:'Callao'   },
      { nombre:'Mauricio Sotelo Nemecio',    firmaKey:'MAURICIO', dni:'42902629', email:'michael.helpdesk@holasharf.com',  sede:'Callao'   },
      { nombre:'Victor Gutiérrez Leiva',     firmaKey:'VICTOR',   dni:'25750736', email:'misael.helpdesk@holasharf.com', sede:'Callao'   },
      { nombre:'Ismael Gomez Sime',          firmaKey:'ISMAEL',   dni:'6311827',  email:'ismael.helpdesk@holasharf.com',  sede:'Callao'   },
      { nombre:'Perla Moreno Yarleque',      firmaKey:'PERLA',    dni:'76250380', email:'perla.helpdesk@holasharf.com',     sede:'Paita'    },
      { nombre:'Alfredo Rojas Guevara',      firmaKey:'ALFREDO',  dni:'7166818',  email:'alfredo.helpdesk@holasharf.com',    sede:'Callao'   },
      { nombre:'Doris Quispe',               firmaKey:'DORIS',    dni:'77014630', email:'rocio.helpdesk@holasharf.com',     sede:'Arequipa' },
      { nombre:'Alvarez Chora Jesus Miguel', firmaKey:'JESUS',    dni:'71422426', email:'jesus.helpdesk@holasharf.com',    sede:'Callao'   },
      { nombre:'Eddie Fernandez',             firmaKey:'EDDIE',    dni:'',         email:'eddie.fernandez@holasharf.com', sede:'Callao'   },
      { nombre:'Anais Chero Benites',         firmaKey:'ANAIS',    dni:'',         email:'anais.chero@holasharf.com',    sede:'Callao'   }
    ];

    var tecnico = null;
    for (var i = 0; i < TECNICOS.length; i++) {
      if (TECNICOS[i].email.toLowerCase() === emailNorm) {
        tecnico = TECNICOS[i]; break;
      }
    }

    if (!tecnico) {
      return {
        ok: false, noRegistrado: true,
        error: 'Cuenta reconocida (' + email + ') pero no registrada como técnico. ' +
               'Contacta al administrador.'
      };
    }

    var firmaData = '';
    try { firmaData = _imgDriveTec(tecnico.firmaKey) || ''; } catch(ef) {}

    Logger.log('autenticarTecnico OK: ' + tecnico.nombre);
    return { ok: true, tecnico: tecnico, firmaData: firmaData, email: email };

  } catch(e) {
    Logger.log('autenticarTecnico ERROR: ' + e.message);
    return { ok: false, error: 'Error: ' + e.message };
  }
}

function validarConexiones() {
  var R = {
    ok:         true,
    timestamp:  Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy HH:mm:ss'),
    pasos:      [],
    sheetRH:    { ok: false, pestanas: [] },
    snipe:      { ok: false, apiKeyOk: false, totalActivos: 0, totalUsuarios: 0 },
    drive:      { actas: false, firmas: false, log: false },
    resumen:    []
  };

  function log(msg)    { R.pasos.push(msg); Logger.log(msg); }
  function ok(label)   { R.resumen.push('✅ ' + label); log('✅ ' + label); }
  function err(label)  { R.resumen.push('❌ ' + label); log('❌ ' + label); R.ok = false; }
  function warn(label) { R.resumen.push('⚠️  ' + label); log('⚠️  ' + label); }

  log('═══ VALIDACIÓN DE CONEXIONES — SHARF Devoluciones TI ═══');
  log('Fecha: ' + R.timestamp);
  log('Script ejecutado por: ' + Session.getActiveUser().getEmail());
  log('');

  // ── 1. SHEET RH ────────────────────────────────────────────────────
  log('── 1. SHEET RH COMPARTIDO ──');
  log('ID configurado: ' + CFG.SHEET_RH_ID);
  try {
    var ss = SpreadsheetApp.openById(CFG.SHEET_RH_ID);
    ok('Sheet RH accesible: "' + ss.getName() + '"');
    R.sheetRH.ok     = true;
    R.sheetRH.nombre = ss.getName();

    var todasHojas = ss.getSheets();
    log('Pestañas en el archivo: ' + todasHojas.map(function(h){return '"'+h.getName()+'"';}).join(', '));

    CFG.SHEET_RH_TABS.forEach(function(tabCfg) {
      var hoja = ss.getSheetByName(tabCfg.nombre);
      if (!hoja) {
        todasHojas.forEach(function(h) {
          if (String(h.getSheetId()) === String(tabCfg.gid)) hoja = h;
        });
      }
      if (!hoja) {
        err('Pestaña "' + tabCfg.nombre + '" — NO encontrada (ni por nombre ni por GID ' + tabCfg.gid + ')');
        R.sheetRH.pestanas.push({ nombre: tabCfg.nombre, ok: false });
        return;
      }

      var lastRow = hoja.getLastRow();
      var lastCol = hoja.getLastColumn();
      var filasDatos = Math.max(0, lastRow - CFG.SHEET_RH_HDR_ROW);
      log('Pestaña "' + hoja.getName() + '": ' + filasDatos + ' registros, ' + lastCol + ' columnas');

      // Verificar columna DNI
      var hdrRow  = lastRow >= CFG.SHEET_RH_HDR_ROW
        ? hoja.getRange(CFG.SHEET_RH_HDR_ROW, 1, 1, lastCol).getValues()[0]
        : [];
      var colDniIdx = -1;
      // Buscar por nombre exacto configurado
      for (var hi = 0; hi < hdrRow.length; hi++) {
        if (String(hdrRow[hi]).trim() === tabCfg.colDni.trim()) { colDniIdx = hi; break; }
      }
      // Fallback: alias dinámico
      if (colDniIdx < 0) {
        var hdrs = _hdrs(hdrRow);
        if (hdrs['dni'] !== undefined) colDniIdx = hdrs['dni'];
      }

      if (colDniIdx >= 0) {
        ok('Pestaña "' + hoja.getName() + '" → col DNI: ' + (colDniIdx+1) + ' "' + hdrRow[colDniIdx] + '" (' + filasDatos + ' registros)');
        R.sheetRH.pestanas.push({ nombre: hoja.getName(), ok: true, colDni: colDniIdx+1, registros: filasDatos });
      } else {
        err('Pestaña "' + hoja.getName() + '" → columna DNI no encontrada. Buscado: "' + tabCfg.colDni + '"');
        log('  Encabezados encontrados: ' + hdrRow.slice(0,10).map(function(h,i){return (i+1)+'="'+h+'"';}).join(', '));
        R.sheetRH.pestanas.push({ nombre: hoja.getName(), ok: false, error: 'Col DNI no encontrada' });
      }
    });
  } catch(e) {
    err('Sheet RH NO accesible (' + CFG.SHEET_RH_ID + '): ' + e.message);
    log('  → Comparte el Sheet con la cuenta: ' + Session.getActiveUser().getEmail());
    R.sheetRH.error = e.message;
  }
  log('');

  // ── 2. SNIPE-IT ────────────────────────────────────────────────────
  log('── 2. SNIPE-IT ──');
  log('URL: ' + CFG.SNIPE_BASE);
  var snipeKey = _snipeKey();
  if (!snipeKey || snipeKey.length < 20) {
    err('API Key de Snipe-IT NO configurada — ejecuta setKeyManual()');
  } else {
    ok('API Key presente (' + snipeKey.length + ' chars)');
    R.snipe.apiKeyOk = true;

    // Probar conexión con /hardware?limit=1
    try {
      var rHW = snipeGET('/hardware?limit=1');
      ok('Snipe-IT /hardware: OK — ' + (rHW.total || 0) + ' activos totales');
      R.snipe.ok           = true;
      R.snipe.totalActivos = rHW.total || 0;
    } catch(eHW) {
      err('Snipe-IT /hardware: ' + eHW.message);
      R.snipe.errorHardware = eHW.message;
    }

    // Probar /users?limit=1
    try {
      var rU = snipeGET('/users?limit=1');
      ok('Snipe-IT /users: OK — ' + (rU.total || 0) + ' usuarios totales');
      R.snipe.totalUsuarios = rU.total || 0;
    } catch(eU) {
      err('Snipe-IT /users: ' + eU.message);
      R.snipe.errorUsers = eU.message;
    }

    // Probar búsqueda por employee_num (campo donde está el DNI)
    try {
      var rEmp = snipeGET('/users?employee_num=99999999&limit=1');
      ok('Snipe-IT búsqueda por employee_num: soportada');
    } catch(eEmp) {
      warn('Snipe-IT employee_num search: ' + eEmp.message + ' (se usará search= como fallback)');
    }
  }
  log('');

  // ── 3. DRIVE ───────────────────────────────────────────────────────
  log('── 3. DRIVE ──');
  try {
    var fActas = DriveApp.getFolderById(CFG.DRIVE_ACTAS_ROOT_ID);
    ok('Drive Actas raíz: "' + fActas.getName() + '"');
    R.drive.actas = true;
  } catch(e) { err('Drive Actas: ' + e.message); }

  try {
    var fFirmas = DriveApp.getFolderById(CFG.DRIVE_FIRMAS_TEC_ID);
    ok('Drive Firmas/Logo: "' + fFirmas.getName() + '"');
    R.drive.firmas = true;
    var logoIter = fFirmas.getFilesByName('LOGO.png');
    logoIter.hasNext()
      ? ok('LOGO.png: encontrado en carpeta firmas')
      : warn('LOGO.png: NO encontrado en carpeta firmas (el acta saldrá sin logo)');
  } catch(e) { err('Drive Firmas: ' + e.message); }

  try {
    var fLog = DriveApp.getFolderById(CFG.DRIVE_SEGUIMIENTO_FOLDER);
    ok('Drive carpeta log: "' + fLog.getName() + '"');
    R.drive.log = true;
  } catch(e) { err('Drive carpeta log: ' + e.message); }

  // Sheet Log (PropertiesService)
  var ssLogId = PropertiesService.getScriptProperties().getProperty('SHEET_LOG_ID');
  if (ssLogId) {
    try {
      var ssLog = SpreadsheetApp.openById(ssLogId);
      ok('Sheet Log: "' + ssLog.getName() + '" (' + ssLogId + ')');
    } catch(e) {
      warn('Sheet Log ID guardado pero no accesible — ejecuta inicializarSheetLog()');
    }
  } else {
    warn('Sheet Log: no inicializado — ejecuta inicializarSheetLog() una vez');
  }
  log('');

  // ── RESUMEN ────────────────────────────────────────────────────────
  log('═══ RESUMEN ═══');
  R.resumen.forEach(function(l) { log(l); });
  var totalOk   = R.resumen.filter(function(l){return l.indexOf('✅')>=0;}).length;
  var totalErr  = R.resumen.filter(function(l){return l.indexOf('❌')>=0;}).length;
  var totalWarn = R.resumen.filter(function(l){return l.indexOf('⚠️')>=0;}).length;
  log('');
  log(totalOk + ' OK  |  ' + totalWarn + ' avisos  |  ' + totalErr + ' errores');
  if (totalErr === 0) log('🚀 Todo OK — el sistema está listo para usar');
  else                log('⚠️  Revisa los errores antes de usar el aplicativo');

  return R;
}

/**
 * testBusquedaDNI — diagnóstico completo.
 *
 * CÓMO VER LOS RESULTADOS:
 *   Después de ejecutar: Ver → Registros de ejecución (panel inferior)
 *   O: Ctrl+Enter sobre la función → panel "Ejecución" abajo a la izquierda
 *
 * ANTES DE EJECUTAR:
 *   Reemplaza DNI_PRUEBA con un DNI real del Sheet.
 */
function testBusquedaDNI() {
  var DNI_PRUEBA = '70406659';  // DNI de prueba SHARF

  var resultado = {
    dni:          DNI_PRUEBA,
    sheetRH:      { accesible: false, pestanas: [], encontrado: false, empleado: null, error: '' },
    snipe:        { apiKeyOk: false, encontrado: false, activos: [], error: '' },
    resumen:      ''
  };

  // ── 1. Verificar Sheet RH ─────────────────────────────────────────
  try {
    var ss = SpreadsheetApp.openById(CFG.SHEET_RH_ID);
    resultado.sheetRH.accesible = true;
    resultado.sheetRH.pestanas  = ss.getSheets().map(function(h) {
      return { nombre: h.getName(), gid: h.getSheetId(), filas: h.getLastRow() };
    });
  } catch(e) {
    resultado.sheetRH.error   = e.message;
    resultado.sheetRH.accesible = false;
  }

  // ── 2. Buscar DNI en las 3 pestañas ───────────────────────────────
  if (resultado.sheetRH.accesible) {
    var rh = _buscarEmpleadoPorDni(DNI_PRUEBA);
    if (rh.ok) {
      resultado.sheetRH.encontrado = true;
      resultado.sheetRH.pestanaEncontrada = rh.tabOrigen;
      resultado.sheetRH.empleado = {
        nombre:        rh.empleado.nombre,
        dni:           rh.empleado.dni,
        empresa:       rh.empleado.empresa,
        area:          rh.empleado.area,
        cargo:         rh.empleado.cargo,
        sede:          rh.empleado.sede,
        ceco:          rh.empleado.ceco,
        emailPersonal: rh.empleado.emailPersonal,
        emailCorp:     rh.empleado.emailCorp,
        emailJefe:     rh.empleado.emailJefe
      };
    } else {
      resultado.sheetRH.error = rh.error;
    }
  }

  // ── 3. Verificar token Snipe ──────────────────────────────────────
  var key = PropertiesService.getScriptProperties().getProperty(CFG.SNIPE_KEY_PROP);
  resultado.snipe.apiKeyOk = !!(key && key.length > 5);

  // ── 4. Buscar en Snipe-IT ─────────────────────────────────────────
  if (resultado.snipe.apiKeyOk) {
    var rs = buscarActivosDeUsuarioPorDNI_(DNI_PRUEBA, null);
    if (rs.ok) {
      resultado.snipe.encontrado = true;
      resultado.snipe.activos    = rs.activos.map(function(a) {
        return { nombre: a.nombre, serial: a.serial, tipo: a.tipoActivo, estado: a.estado };
      });
    } else {
      resultado.snipe.error = rs.error;
    }
  } else {
    resultado.snipe.error = 'SNIPE_API_KEY no configurada';
  }

  // ── 5. Resumen legible ────────────────────────────────────────────
  var lineas = [
    '══════════════════════════════════',
    'TEST BÚSQUEDA DNI: ' + DNI_PRUEBA,
    '══════════════════════════════════',
    '',
    '[ SHEET RH ]',
    '  Accesible:  ' + (resultado.sheetRH.accesible ? '✅ SÍ' : '❌ NO — ' + resultado.sheetRH.error),
  ];
  if (resultado.sheetRH.accesible) {
    lineas.push('  Pestañas disponibles:');
    resultado.sheetRH.pestanas.forEach(function(p) {
      var enConfig = CFG.SHEET_RH_TABS.some(function(t){ return String(t.gid)===String(p.gid)||t.nombre===p.nombre; });
      lineas.push('    ' + (enConfig?'→':'  ') + ' ' + p.nombre + '  GID:' + p.gid + '  Filas:' + p.filas);
    });
    lineas.push('  Encontrado:  ' + (resultado.sheetRH.encontrado
      ? '✅ SÍ — pestaña: ' + resultado.sheetRH.pestanaEncontrada
      : '❌ NO — ' + resultado.sheetRH.error));
    if (resultado.sheetRH.encontrado && resultado.sheetRH.empleado) {
      var e = resultado.sheetRH.empleado;
      lineas.push('  Nombre:         ' + (e.nombre        || '(vacío)'));
      lineas.push('  Empresa:        ' + (e.empresa       || '(vacío)'));
      lineas.push('  Área:           ' + (e.area          || '(vacío)'));
      lineas.push('  Cargo:          ' + (e.cargo         || '(vacío)'));
      lineas.push('  Sede:           ' + (e.sede          || '(vacío)'));
      lineas.push('  Email personal: ' + (e.emailPersonal || '(vacío)'));
      lineas.push('  Email corp:     ' + (e.emailCorp     || '(vacío)'));
      lineas.push('  NOTA: emailJefe se solicita al técnico en pantalla (no existe en el Sheet)');
    }
  }
  lineas.push('');
  lineas.push('[ SNIPE-IT ]');
  lineas.push('  API Key:     ' + (resultado.snipe.apiKeyOk ? '✅ Configurada' : '❌ NO configurada'));
  if (resultado.snipe.apiKeyOk) {
    lineas.push('  Encontrado:  ' + (resultado.snipe.encontrado
      ? '✅ SÍ — ' + resultado.snipe.activos.length + ' activo(s)'
      : '❌ NO — ' + resultado.snipe.error));
    resultado.snipe.activos.forEach(function(a) {
      lineas.push('    → ' + a.nombre + '  S/N:' + a.serial + '  Tipo:' + a.tipo + '  Estado:' + a.estado);
    });
  }
  lineas.push('');
  lineas.push('══════════════════════════════════');

  resultado.resumen = lineas.join('\n');
  Logger.log(resultado.resumen);

  // Retornar el objeto completo — visible en el panel de ejecución
  return resultado;
}

function _mapearEmpleadoRH_(fila, hdrs) {
  function col() {
    for (var a = 0; a < arguments.length; a++) {
      if (hdrs[arguments[a]] !== undefined) {
        return String(fila[hdrs[arguments[a]]] || '').trim();
      }
    }
    return '';
  }
  // CECO: columna "Código" + descripción "Denominación"
  var cecoVal = col('ceco', 'centrodecostos', 'centrodecosto', 'cc', 'codigoceco', 'codigo');
  var cecoNom = col('ceconom', 'denominacion', 'denominaci');
  if (cecoVal && cecoNom) cecoVal = cecoVal + ' - ' + cecoNom;

  return {
    dni:           col('dni', 'dnice', 'ce', 'documento'),
    nombre:        col('nombre', 'apellidosynombres', 'apellidosnombres', 'nombresyapellidos'),
    empresa:       col('empresa', 'compania', 'company'),
    area:          col('area', 'gerencia', 'division'),        // columna "AREA"
    subarea:       col('subarea', 'subarrea', 'departamento'),
    cargo:         col('cargo', 'puesto', 'posicion'),
    sede:          col('centroops', 'centrodeoperaciones', 'sede', 'sucursal', 'ubicacion'), // "CENTRO DE OPERACIONES"
    ceco:          cecoVal,                                    // columna "Código"
    emailPersonal: col('emailpers', 'correopersonal', 'emailpersonal', 'emailp'),
    emailCorp:     col('emailcorp', 'correocorporativo', 'emailcorporativo', 'emailc', 'email'),
    // emailJefe NO existe en ningún Sheet ni en Snipe — se pide al técnico en pantalla
    emailJefe:     ''
  };
}

function _hdrs(headerRow) {
  var map = {};
  if (!headerRow) return map;
  headerRow.forEach(function(h, i) {
    if (h === null || h === undefined) return;
    var raw = String(h).trim();
    if (!raw) return;

    // Clave normalizada: sin tildes, sin especiales, minúsculas
    var k = raw.toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]/g, '');
    if (k) map[k] = i;

    // ── Alias explícitos por contenido del encabezado ─────────────────
    var UP = raw.toUpperCase();

    // DNI — captura todas las variantes confirmadas:
    //   "DNI/C.E."  (BD_Personal)
    //   " DNI / C.E."  (Cesados — con espacios)
    //   "DNI"  (Practicantes Cesados)
    //   y cualquier encabezado que contenga "DNI"
    if (UP.replace(/\s/g,'').indexOf('DNI') >= 0 ||
        UP.indexOf('C.E') >= 0 ||
        UP === 'DOCUMENTO') {
      map['dni'] = i;
    }
    // Apellidos y Nombres → 'nombre'
    if ((UP.indexOf('APELLIDO') >= 0 || UP.indexOf('NOMBRE') >= 0) && !map['nombre']) {
      map['nombre'] = i;
    }
    // Email Personal → 'emailpers'
    if ((UP.indexOf('CORREO') >= 0 || UP.indexOf('EMAIL') >= 0) &&
         UP.indexOf('PERSONAL') >= 0) {
      map['emailpers'] = i;
    }
    // Email Corp → 'emailcorp'
    if ((UP.indexOf('CORREO') >= 0 || UP.indexOf('EMAIL') >= 0) &&
        (UP.indexOf('CORP') >= 0 || UP.indexOf('EMPR') >= 0 ||
         UP.indexOf('TRAB') >= 0 || UP.indexOf('INST') >= 0)) {
      map['emailcorp'] = i;
    }
    // JEFATURA / JEFE → ya no se mapea (emailJefe no existe en el Sheet)
    // Cargo / Puesto → 'cargo'
    if (UP.indexOf('CARGO') >= 0 || UP.indexOf('PUESTO') >= 0) {
      if (!map['cargo'])  map['cargo']  = i;
      if (!map['puesto']) map['puesto'] = i;
    }
    // Centro de Operaciones → 'centroops'
    if ((UP.indexOf('CENTRO') >= 0 && UP.indexOf('OPER') >= 0) ||
         UP.indexOf('SEDE') >= 0) {
      if (!map['centroops']) map['centroops'] = i;
    }
    // CECO / Centro de costo / Código / Denominación → 'ceco'
    // En BD_Personal la columna CECO se llama "Código" (col 10) y "Denominación" (col 11)
    // Usamos "Código" como el código del CECO
    if (UP.indexOf('CECO') >= 0 ||
        (UP.indexOf('CENTRO') >= 0 && UP.indexOf('COSTO') >= 0) ||
        UP === 'CÓDIGO' || UP === 'CODIGO' ||
        UP.indexOf('CENTRO DE COSTO') >= 0) {
      if (!map['ceco']) map['ceco'] = i;
    }
    // Denominación del CECO → 'ceconom'
    if (UP.indexOf('DENOMINACI') >= 0) {
      if (!map['ceconom']) map['ceconom'] = i;
    }
  });
  return map;
}

// ════════════════════════════════════════════════════════════════════════
// 15. HELPERS SNIPE-IT
// ════════════════════════════════════════════════════════════════════════
// Token Snipe-IT — guardado en PropertiesService por setKeyManual()
// Si PropertiesService está vacío, usa este fallback directo
var SNIPE_TOKEN_FALLBACK = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiNjZlZmMyOTllOTMyMzVmMzJmMTg5NDg3MGMyYzNhYzQ0ODhmNTZkNTJlNjEwNjQ2NWM1N2FmNmJkYjdjNDE4YjNlNmI4NmQ2Y2I4ZWQ0OTUiLCJpYXQiOjE3NzI5MzM2NjguMzM2MzA0LCJuYmYiOjE3NzI5MzM2NjguMzM2MzA5LCJleHAiOjI0MDQwODU2NjguMzE2Mjg2LCJzdWIiOiI3Iiwic2NvcGVzIjpbXX0.KhW_WJ6Y4hTtlL6bNC9wXlLDiiHdsptWSTfa2NXn9iLiKAgLnTgJBzbGGhi83hjz7y3adyc5slRS-YEvDqdlVY-VERHtlU6JAIRuYcrzFZlbaEBD-iuuaalA6OnfphG4C-5OLZd1Y0BZWXna37f6yNns0EPwRnhdINL2QqV7Y-pehzAYraZfmZg6aOL5HENH-WiWRt_kP305u3mECOZhUMjsiNZIygqG1MipcXB6NMGtfvaD6JmkhHUNFpwnLZdgfHjWsIekAxVCZjeRvf2oRRg6z9T3U3QGks3lHF46yhb1qo9vLrA9wMngjiEkWi80pbi3005hPCvePm01b78lZNJ1Ckxlxq45A0vtg7OPBfQWzRbFG_NZp6lOLDfEotxEddbfUy0EiWQUPAs7oNK8tKh5jPYZvvwhSWGC2EpD3BTMpfxlmU_kCRJ3PobOGk9o_4GQuYkD3n52vkXoAzrJcYikJCcRdNN2isn7JPWH386VeVE8mQxNieS0H2g6bzOuFKMghKcp_uHf1eaV6hDq98Tv_6kmAK1PXISY0lzD7l_7S849iKpoYQscy2xHtwRXYIgCWSGY339aUG3iuqbg4MDynisnfurABSj9zEUgYy2DKKLgDnugq2IXtJP3dlV0aB6qUQ89d6mzAUfRtPd39SlvVvtN8piVSFt2YAgMS78';

function _snipeKey() {
  return PropertiesService.getScriptProperties().getProperty(CFG.SNIPE_KEY_PROP)
         || SNIPE_TOKEN_FALLBACK;
}

function snipeGET(endpoint) {
  var key = _snipeKey();
  var resp = UrlFetchApp.fetch(CFG.SNIPE_BASE + endpoint, {
    method: 'GET',
    headers: { 'Authorization': 'Bearer ' + key, 'Accept': 'application/json', 'Content-Type': 'application/json' },
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code !== 200) throw new Error('Snipe GET ' + endpoint + ' → HTTP ' + code + ': ' + resp.getContentText().slice(0,200));
  return JSON.parse(resp.getContentText());
}

function snipePOST(endpoint, payload) {
  var key = _snipeKey();
  var resp = UrlFetchApp.fetch(CFG.SNIPE_BASE + endpoint, {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + key, 'Accept': 'application/json', 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  return JSON.parse(resp.getContentText());
}

function snipePATCH(endpoint, payload) {
  var key  = _snipeKey();
  var resp = UrlFetchApp.fetch(CFG.SNIPE_BASE + endpoint, {
    method:  'PATCH',
    headers: { 'Authorization': 'Bearer ' + key, 'Accept': 'application/json', 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  var body = JSON.parse(resp.getContentText());
  if (code !== 200) throw new Error('Snipe PATCH ' + endpoint + ' → HTTP ' + code + ': ' + resp.getContentText().slice(0,200));
  return body;
}

function _mapearActivo_(a, accesorios) {
  // Accesorios: prioridad → campos custom Snipe (Mouse/Mochila/Docking/Cargador)
  //             fallback  → order_number legacy (SERIAL_CARGADOR/MOCHILA/MOUSE/DOCKING)
  var acc = accesorios || _leerAccesoriosCustomFields(a);
  var cat = (a.category && a.category.name) ? a.category.name : '';

  // ── Mapeo de campos Snipe-IT → campos internos ────────────────────
  // CECO:               campo "custom_fields.Centro de Costos" del activo
  // Ubicación:          campo "location.name"  = Centro de Operaciones (RH)
  // Área del activo:    campo "custom_fields.Area" del activo
  // Área del usuario:   campo "department.name" en el objeto usuario de Snipe

  var ceco      = '';
  var areaActivo = '';
  if (a.custom_fields) {
    // Los custom_fields vienen como objeto: { "Centro de Costos": { value: "..." }, "Area": { value: "..." } }
    var cf = a.custom_fields;
    // Centro de Costos (CECO)
    if (cf['Centro de Costos'] && cf['Centro de Costos'].value) {
      ceco = String(cf['Centro de Costos'].value).trim();
    } else if (cf['centro_de_costos'] && cf['centro_de_costos'].value) {
      ceco = String(cf['centro_de_costos'].value).trim();
    } else if (cf['CECO'] && cf['CECO'].value) {
      ceco = String(cf['CECO'].value).trim();
    }
    // Área del activo
    if (cf['Area'] && cf['Area'].value) {
      areaActivo = String(cf['Area'].value).trim();
    } else if (cf['AREA'] && cf['AREA'].value) {
      areaActivo = String(cf['AREA'].value).trim();
    }
  }

  // Ubicación del activo = Centro de Operaciones
  var ubicacion = '';
  if (a.location && a.location.name) {
    ubicacion = String(a.location.name).trim();
  } else if (a.rtd_location && a.rtd_location.name) {
    ubicacion = String(a.rtd_location.name).trim();
  }

  return {
    id:            a.id,
    nombre:        a.name || '',
    modelo:        (a.model && a.model.name) ? a.model.name : '',
    serial:        a.serial || '',
    tag:           a.asset_tag || '',
    categoria:     cat,
    empresa:       (a.company && a.company.name) ? a.company.name : '',
    estado:        (a.status_label && a.status_label.name) ? a.status_label.name : '',
    orderNumber:   '',   // eliminado — accesorios vienen de custom fields
    accesorios:    acc,
    tipoActivo:    _clasificarTipoActivo(a.name + ' ' + cat),
    // Campos adicionales con mapeo correcto
    ceco:          ceco,       // custom_field "Centro de Costos"
    ubicacion:     ubicacion,  // location.name = Centro de Operaciones (RH)
    area:          areaActivo, // custom_field "Area"
    // company_id del activo actual — usado como referencia en modo manual
    // (en flujo DNI se sobreescribe con snipeCompanyId del usuario)
    snipeCompanyId: (a.company && a.company.id) ? a.company.id : null
  };
}

function _leerAccesoriosCustomFields(a) {
  var cf = (a && a.custom_fields) ? a.custom_fields : {};

  // ── Helper: leer valor de un campo por posibles nombres ──────────────
  function cfVal(nombres) {
    for (var i = 0; i < nombres.length; i++) {
      var key = nombres[i];
      if (cf[key] && cf[key].value !== null && cf[key].value !== undefined) {
        return String(cf[key].value || '').trim();
      }
    }
    return '';
  }

  // ── Helper: verificar si un valor indica "tiene accesorio" ────────────
  function tieneAcc(val) {
    if (!val) return false;
    var v = val.toUpperCase().trim();
    return v !== '' && v !== 'NO' && v !== 'N/A' && v !== 'NINGUNO' &&
           v !== '0' && v !== 'SIN' &&
           v.indexOf('NO ') !== 0 && v.indexOf('SIN ') !== 0;
  }

  // ── Leer campo cargador (serial) ──────────────────────────────────────
  var cargador = cfVal(['Cargador','cargador','CARGADOR','Serial Cargador','serial_cargador']);

  // ── Leer campo accesorios múltiple: _snipeit_accesorios_37 ───────────
  // En la API de Snipe-IT, los campos checkbox múltiple llegan como:
  //   custom_fields["Accesorios"].value = "Mouse, Mochila"  (string separado por comas)
  // o también pueden llegar como campo individual por nombre de accesorio
  var accesoriosStr = cfVal([
    'Accesorios', 'accesorios', 'ACCESORIOS',
    '_snipeit_accesorios_37', 'accesorios_37'
  ]);

  // Parsear el string de accesorios (ej: "Mouse, Mochila, Docking")
  var accList = accesoriosStr
    ? accesoriosStr.split(/[,;|]/).map(function(s){ return s.trim().toUpperCase(); })
    : [];

  function enLista(nombre) {
    return accList.indexOf(nombre.toUpperCase()) >= 0;
  }

  // ── Leer cada accesorio individualmente (campos separados o lista) ────
  var mouseVal   = cfVal(['Mouse','mouse','MOUSE'])   || (enLista('Mouse')   ? 'Mouse'   : '');
  var mochilaVal = cfVal(['Mochila','mochila','MOCHILA']) || (enLista('Mochila') ? 'Mochila' : '');
  var dockingVal = cfVal(['Docking','docking','DOCKING','Docking Station']) || (enLista('Docking') ? 'Docking' : '');
  var tecladoVal = cfVal(['Teclado','teclado','TECLADO','Keyboard']) || (enLista('Teclado') ? 'Teclado' : '');

  var acc = {
    cargadorSerial: cargador,
    mouse:          tieneAcc(mouseVal),   mouseDesc:   mouseVal,
    mochila:        tieneAcc(mochilaVal), mochilaDesc: mochilaVal,
    docking:        tieneAcc(dockingVal), dockingDesc: dockingVal,
    teclado:        tieneAcc(tecladoVal), tecladoDesc: tecladoVal
  };

  // NOTA: fallback a order_number eliminado intencionalmente.
  // Los accesorios (incluido "Cargador") se leen exclusivamente
  // desde campos custom de Snipe-IT — nunca desde order_number.

  return acc;
}

// _parsearOrderNumber eliminado — ya no se usa ni se guarda nada en order_number.
// Los accesorios se leen exclusivamente desde custom_fields de Snipe-IT.

// ════════════════════════════════════════════════════════════════════════
// getSnipeAssetImages — Lee imágenes de la pestaña "Archivos" de un activo
//   Filtra solo archivos de imagen (jpg, jpeg, png, gif, webp, bmp)
//   Retorna: { ok, imagenes:[{nombre, url, miniatura}], total }
//
//   CÓMO COMPROBAR si tu Snipe soporta esto:
//     1. En el editor GAS ejecuta: testGetSnipeImages()
//     2. Abre el Log (Ctrl+Enter) y verifica la respuesta
//     3. Si ves rows con filename → funciona
//     4. Si ves HTTP 404 → tu versión de Snipe-IT no tiene el endpoint
// ════════════════════════════════════════════════════════════════════════
function getSnipeAssetImages(assetId) {
  if (!assetId) return { ok: false, error: 'assetId requerido', imagenes: [] };
  try {
    var r = snipeGET('/hardware/' + assetId + '/files');
    if (!r) return { ok: false, error: 'Sin respuesta de Snipe-IT', imagenes: [] };

    var rows = r.rows || r.payload || [];
    if (!Array.isArray(rows)) return { ok: false, error: 'Respuesta inesperada', imagenes: [] };

    // Log completo del primer archivo para diagnóstico
    if (rows.length > 0) {
      Logger.log('getSnipeAssetImages — primer archivo raw: ' + JSON.stringify(rows[0]).slice(0, 400));
    }

    var EXTS_IMAGEN = /\.(jpg|jpeg|png|gif|webp|bmp|svg)$/i;
    var imagenes    = [];
    var key         = _snipeKey();
    var baseUrl     = CFG.SNIPE_BASE.replace('/api/v1', '');

    rows.forEach(function(f) {
      var nombre = f.filename || f.name || '';
      if (!EXTS_IMAGEN.test(nombre)) return;

      // Probar todas las URLs posibles según la versión de Snipe-IT
      // v6+: download_url (ya lleva el token?), v5: url relativa, otros: file_url
      var urls = [
        f.download_url,
        f.url,
        f.file_url,
        f.preview_url,
        // Construir manualmente si tenemos el ID del archivo
        f.id ? (baseUrl + '/hardware/' + assetId + '/files/' + f.id) : null,
        f.id ? (baseUrl + '/api/v1/hardware/' + assetId + '/files/' + f.id) : null
      ].filter(Boolean);

      // Completar URLs relativas
      urls = urls.map(function(u) {
        if (u && !u.startsWith('http')) return baseUrl + u;
        return u;
      });

      Logger.log('getSnipeAssetImages: ' + nombre + ' → URLs a probar: ' + JSON.stringify(urls));

      var dataUrl = '';
      for (var ui = 0; ui < urls.length; ui++) {
        if (!urls[ui]) continue;
        try {
          var resp = UrlFetchApp.fetch(urls[ui], {
            method: 'GET',
            headers: { 'Authorization': 'Bearer ' + key, 'Accept': 'image/*,*/*' },
            muteHttpExceptions: true,
            followRedirects: true
          });
          var code = resp.getResponseCode();
          Logger.log('  URL[' + ui + '] HTTP ' + code + ': ' + urls[ui]);
          if (code === 200) {
            var blob = resp.getBlob();
            var mime = blob.getContentType() || 'image/jpeg';
            // Solo proceder si realmente es una imagen
            if (mime.indexOf('image') >= 0 || mime.indexOf('octet') >= 0) {
              dataUrl = 'data:image/jpeg;base64,' + Utilities.base64Encode(blob.getBytes());
              Logger.log('  ✅ Imagen cargada: ' + blob.getBytes().length + ' bytes');
              break;
            } else {
              Logger.log('  ⚠️ Respuesta no es imagen: ' + mime);
            }
          }
        } catch(eFetch) {
          Logger.log('  ❌ Error fetch URL[' + ui + ']: ' + eFetch.message);
        }
      }

      if (!dataUrl) Logger.log('  ❌ No se pudo obtener imagen: ' + nombre);

      imagenes.push({
        nombre:  nombre,
        dataUrl: dataUrl,
        url:     urls[0] || '',
        id:      f.id || ''
      });
    });

    Logger.log('getSnipeAssetImages[' + assetId + ']: ' + rows.length + ' archivos, ' +
               imagenes.length + ' imágenes, ' +
               imagenes.filter(function(i){return !!i.dataUrl;}).length + ' cargadas como base64');

    return { ok: true, imagenes: imagenes, total: imagenes.length, totalArchivos: rows.length };

  } catch(e) {
    Logger.log('getSnipeAssetImages ERROR: ' + e.message);
    return { ok: false, error: e.message, imagenes: [] };
  }
}

// ── Función de prueba — ejecutar desde el editor GAS ──────────────────
function testGetSnipeImages() {
  // Reemplaza con el ID de un activo real que tenga imágenes en Snipe
  var ASSET_ID_PRUEBA = 1; // ← CAMBIAR A UN ID REAL
  Logger.log('=== TEST getSnipeAssetImages ===');
  Logger.log('Asset ID: ' + ASSET_ID_PRUEBA);
  var r = getSnipeAssetImages(ASSET_ID_PRUEBA);
  Logger.log('Resultado: ' + JSON.stringify(r));
  if (r.ok) {
    Logger.log('✅ ' + r.total + ' imágenes encontradas de ' + r.totalArchivos + ' archivos totales');
    r.imagenes.forEach(function(img, i) {
      Logger.log('  Img ' + (i+1) + ': ' + img.nombre + ' → ' + img.url);
    });
  } else {
    Logger.log('❌ Error: ' + r.error);
  }
}

// ════════════════════════════════════════════════════════════════════════
// testGenerarPDF — Diagnóstico de generación de PDF
//   Ejecutar desde el editor GAS para verificar por qué el PDF no se genera
//   Crea un PDF de prueba mínimo y lo sube a la carpeta raíz de actas
// ════════════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════════════
// diagnosticarBusquedaActa — Diagnóstico completo de "Ver última acta"
//
// Ejecutar desde el editor GAS:
//   1. Cambia NOMBRE_USUARIO por el nombre exacto de un colaborador
//   2. Ejecuta la función
//   3. Abre el Log (Ctrl+Enter) — verás el diagnóstico paso a paso
//
// También expuesta al frontend via handleRequest('diagnosticarBusquedaActa')
// para que el técnico pueda ver el reporte sin acceder al editor GAS.
// ════════════════════════════════════════════════════════════════════════
function diagnosticarBusquedaActa(params) {
  var nombreUsuario = (params && params.nombre) ? params.nombre : '';
  var dniUsuario    = (params && params.dni)    ? params.dni    : '';
  // Si se llama desde el editor sin parámetros, usar un nombre de prueba
  if (!nombreUsuario && !dniUsuario) {
    nombreUsuario = 'NOMBRE APELLIDO';  // ← CAMBIAR ANTES DE EJECUTAR
  }

  var rep = {
    ok:     false,
    pasos:  [],
    actas:  [],
    config: {}
  };

  function paso(num, desc, ok, detalle) {
    var p = { num: num, desc: desc, ok: ok, detalle: detalle || '' };
    rep.pasos.push(p);
    Logger.log((ok ? '✅' : '❌') + ' PASO ' + num + ': ' + desc + (detalle ? ' → ' + detalle : ''));
    return ok;
  }

  Logger.log('════════════════════════════════════════════════');
  Logger.log('DIAGNÓSTICO: Ver última acta de asignación');
  Logger.log('Buscando para: "' + nombreUsuario + '"' + (dniUsuario ? ' / DNI: ' + dniUsuario : ''));
  Logger.log('Fecha: ' + Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('════════════════════════════════════════════════');

  // ── PASO 1: Verificar DRIVE_ACTAS_ROOT_ID ─────────────────────────────
  var raizActas = null;
  try {
    raizActas = DriveApp.getFolderById(CFG.DRIVE_ACTAS_ROOT_ID);
    paso(1, 'Acceso a DRIVE_ACTAS_ROOT_ID', true,
      'ID: ' + CFG.DRIVE_ACTAS_ROOT_ID + ' → "' + raizActas.getName() + '"');
    rep.config.raizActasNombre = raizActas.getName();
    rep.config.raizActasUrl    = raizActas.getUrl();
  } catch(e) {
    paso(1, 'Acceso a DRIVE_ACTAS_ROOT_ID', false,
      'ERROR: ' + e.message + ' · ID configurado: ' + CFG.DRIVE_ACTAS_ROOT_ID);
    rep.error = 'DRIVE_ACTAS_ROOT_ID inválido o sin permisos: ' + e.message;
    return rep;
  }

  // ── PASO 2: Verificar DRIVE_ASIGNACIONES_ROOT_ID (opcional) ───────────
  var raizAsig = null;
  if (CFG.DRIVE_ASIGNACIONES_ROOT_ID) {
    try {
      raizAsig = DriveApp.getFolderById(CFG.DRIVE_ASIGNACIONES_ROOT_ID);
      paso(2, 'Acceso a DRIVE_ASIGNACIONES_ROOT_ID', true,
        'ID: ' + CFG.DRIVE_ASIGNACIONES_ROOT_ID + ' → "' + raizAsig.getName() + '"');
      rep.config.raizAsigNombre = raizAsig.getName();
    } catch(e) {
      paso(2, 'Acceso a DRIVE_ASIGNACIONES_ROOT_ID', false,
        'ERROR: ' + e.message + ' (no crítico — usará DRIVE_ACTAS_ROOT_ID como fallback)');
    }
  } else {
    paso(2, 'DRIVE_ASIGNACIONES_ROOT_ID', true, 'No configurado — se usará DRIVE_ACTAS_ROOT_ID');
  }

  // ── PASO 3: Listar carpetas en la raíz y buscar al usuario ────────────
  Logger.log('');
  Logger.log('── Carpetas encontradas en DRIVE_ACTAS_ROOT_ID ──');
  var todasLasCarpetas = [];
  try {
    var itC = raizActas.getFolders();
    while (itC.hasNext()) {
      var fc = itC.next();
      todasLasCarpetas.push(fc.getName());
      Logger.log('  📁 ' + fc.getName());
    }
    paso(3, 'Listar carpetas en raíz ACTAS', true,
      todasLasCarpetas.length + ' carpetas encontradas');
    rep.config.carpetasEnRaiz = todasLasCarpetas;
  } catch(e) {
    paso(3, 'Listar carpetas en raíz ACTAS', false, e.message);
  }

  // ── PASO 4: Normalización del nombre ─────────────────────────────────
  function norm(s) {
    return (s||'').toUpperCase().trim()
      .replace(/[ÁÀÂÄ]/g,'A').replace(/[ÉÈÊË]/g,'E')
      .replace(/[ÍÌÎÏ]/g,'I').replace(/[ÓÒÔÖ]/g,'O')
      .replace(/[ÚÙÛÜ]/g,'U').replace(/[Ñ]/g,'N')
      .replace(/\s+/g,' ');
  }
  var busq   = norm(nombreUsuario);
  var tokens = busq.split(/[,\s]+/).filter(function(t){ return t.length >= 3; });
  paso(4, 'Normalización del nombre', true,
    '"' + nombreUsuario + '" → "' + busq + '" · tokens: [' + tokens.join(', ') + ']');
  rep.config.nombreNorm  = busq;
  rep.config.tokensUsados = tokens;

  // ── PASO 5: Buscar la carpeta del usuario con scoring ─────────────────
  Logger.log('');
  Logger.log('── Resultado de la búsqueda por scoring ──');
  var carpetaEncontrada = null, carpetaNombre = '';
  try {
    // a) Exacta normalizada
    var itE = raizActas.getFoldersByName(busq);
    if (itE.hasNext()) {
      carpetaEncontrada = itE.next();
      carpetaNombre = carpetaEncontrada.getName();
      paso(5, 'Búsqueda exacta (normalizada)', true, '"' + carpetaNombre + '"');
    } else {
      // b) Exacta sin normalizar
      var itE2 = raizActas.getFoldersByName(nombreUsuario.toUpperCase().trim());
      if (itE2.hasNext()) {
        carpetaEncontrada = itE2.next();
        carpetaNombre = carpetaEncontrada.getName();
        paso(5, 'Búsqueda exacta (sin normalizar)', true, '"' + carpetaNombre + '"');
      } else {
        // c) Scoring por tokens
        var scoringLog = [];
        var itAll = raizActas.getFolders();
        var best = 0, bestF = null;
        while (itAll.hasNext()) {
          var ff = itAll.next();
          var fn = norm(ff.getName());
          var score = tokens.filter(function(t){ return fn.indexOf(t) >= 0; }).length;
          if (dniUsuario && fn.indexOf(dniUsuario) >= 0) score += 3;
          scoringLog.push(ff.getName() + ' (score=' + score + ')');
          if (score > best) { best = score; bestF = ff; }
        }
        Logger.log('  Scoring: ' + scoringLog.slice(0,10).join(' | '));
        var umbral = tokens.length === 1 ? 1 : 2;
        if (best >= umbral && bestF) {
          carpetaEncontrada = bestF;
          carpetaNombre = bestF.getName();
          paso(5, 'Búsqueda por scoring', true,
            '"' + carpetaNombre + '" (score=' + best + ', umbral=' + umbral + ')');
        } else {
          paso(5, 'Búsqueda de carpeta', false,
            'No se encontró carpeta. Mayor score=' + best + ' (umbral=' + umbral + '). ' +
            'Carpetas disponibles: ' + todasLasCarpetas.slice(0,5).join(' | ') +
            (todasLasCarpetas.length > 5 ? ' ...y ' + (todasLasCarpetas.length - 5) + ' más' : ''));
          rep.error = 'Carpeta del usuario no encontrada. Nombre buscado: "' + busq + '". ' +
            'Verifica que el nombre en Drive coincida con el de la planilla RH.';
          rep.sugerencia = 'Carpetas en Drive: ' + todasLasCarpetas.slice(0,8).join(' | ');
          return rep;
        }
      }
    }
  } catch(e) {
    paso(5, 'Búsqueda de carpeta', false, e.message);
    rep.error = e.message;
    return rep;
  }

  rep.config.carpetaUrl = carpetaEncontrada.getUrl();

  // ── PASO 6: Listar archivos en la carpeta del usuario ─────────────────
  Logger.log('');
  Logger.log('── Archivos en "' + carpetaNombre + '" (raíz) ──');
  var todosLosArchivos = [], pdfsAsignacion = [], pdfsDevolucion = [], otrosArchivos = [];
  try {
    var itF = carpetaEncontrada.getFiles();
    while (itF.hasNext()) {
      var f = itF.next();
      var n = f.getName();
      var nl = n.toLowerCase();
      todosLosArchivos.push(n);
      if (nl.endsWith('.pdf')) {
        if (nl.indexOf('devolucion') >= 0 || nl.indexOf('devolución') >= 0 || nl.indexOf('dev_') >= 0) {
          pdfsDevolucion.push(n);
          Logger.log('  📋 [DEVOLUCIÓN - excluido] ' + n);
        } else {
          pdfsAsignacion.push({ nombre: n, url: f.getUrl(),
            fecha: Utilities.formatDate(f.getDateCreated(), CFG.TZ, 'dd/MM/yyyy HH:mm'),
            ts: f.getDateCreated().getTime() });
          Logger.log('  📄 [ASIGNACIÓN ✅] ' + n + ' — ' + Utilities.formatDate(f.getDateCreated(), CFG.TZ, 'dd/MM/yyyy HH:mm'));
        }
      } else {
        otrosArchivos.push(n);
        Logger.log('  📎 [OTRO - no PDF] ' + n);
      }
    }
    paso(6, 'Listar archivos en carpeta raíz del usuario', true,
      todosLosArchivos.length + ' total · ' + pdfsAsignacion.length + ' PDFs asignación · ' +
      pdfsDevolucion.length + ' PDFs devolución · ' + otrosArchivos.length + ' otros');
  } catch(e) {
    paso(6, 'Listar archivos en carpeta', false, e.message);
  }

  // ── PASO 7: Buscar en subcarpeta Asignaciones ─────────────────────────
  Logger.log('');
  Logger.log('── Subcarpetas de "' + carpetaNombre + '" ──');
  var NOMBRES_ASIG = ['asignaciones','asignacion','asignación','asignaciones ti','asig'];
  var NOMBRES_DEV  = ['devoluciones','devolucion','devolución','dev'];
  try {
    var itSub = carpetaEncontrada.getFolders();
    var subcarpetasEncontradas = [];
    while (itSub.hasNext()) {
      var sub = itSub.next();
      var subNorm = sub.getName().toLowerCase().trim();
      subcarpetasEncontradas.push(sub.getName());
      Logger.log('  📁 Subcarpeta: "' + sub.getName() + '"');
      if (NOMBRES_DEV.indexOf(subNorm) >= 0) {
        Logger.log('    → IGNORADA (es de devoluciones)');
        continue;
      }
      if (NOMBRES_ASIG.indexOf(subNorm) >= 0) {
        Logger.log('    → ENTRANDO (es de asignaciones)');
        var itSubF = sub.getFiles();
        while (itSubF.hasNext()) {
          var sf = itSubF.next();
          var sn = sf.getName();
          var snl = sn.toLowerCase();
          if (snl.endsWith('.pdf') && snl.indexOf('devolucion') < 0 && snl.indexOf('dev_') < 0) {
            pdfsAsignacion.push({ nombre: sn, url: sf.getUrl(),
              fecha: Utilities.formatDate(sf.getDateCreated(), CFG.TZ, 'dd/MM/yyyy HH:mm'),
              ts: sf.getDateCreated().getTime() });
            Logger.log('    📄 [ASIGNACIÓN en subcarpeta ✅] ' + sn);
          }
        }
      } else {
        Logger.log('    → OMITIDA (no es de asignaciones ni devoluciones)');
      }
    }
    paso(7, 'Buscar en subcarpetas', true,
      subcarpetasEncontradas.length > 0
        ? 'Subcarpetas: ' + subcarpetasEncontradas.join(', ')
        : 'Sin subcarpetas');
  } catch(e) {
    paso(7, 'Buscar en subcarpetas', false, e.message);
  }

  // ── PASO 8: Resultado final ───────────────────────────────────────────
  Logger.log('');
  Logger.log('── RESULTADO FINAL ──');
  pdfsAsignacion.sort(function(a,b){ return b.ts - a.ts; });

  if (pdfsAsignacion.length > 0) {
    paso(8, 'PDFs de asignación encontrados', true,
      pdfsAsignacion.length + ' PDF(s). Más reciente: "' + pdfsAsignacion[0].nombre + '" del ' + pdfsAsignacion[0].fecha);
    rep.ok    = true;
    rep.actas = pdfsAsignacion;
    rep.config.carpetaNombre = carpetaNombre;
    Logger.log('');
    Logger.log('✅ ÉXITO — ' + pdfsAsignacion.length + ' acta(s) encontrada(s)');
    pdfsAsignacion.forEach(function(a, i) {
      Logger.log('  ' + (i+1) + '. ' + a.nombre + ' — ' + a.fecha);
    });
  } else {
    paso(8, 'PDFs de asignación encontrados', false,
      'Cero PDFs de asignación. Archivos en carpeta: ' + todosLosArchivos.join(' | '));
    rep.error = 'La carpeta "' + carpetaNombre + '" no tiene PDFs de asignación. ' +
      'Archivos encontrados: ' + (todosLosArchivos.length > 0 ? todosLosArchivos.join(', ') : 'ninguno');
    rep.sugerencia = pdfsDevolucion.length > 0
      ? 'Hay ' + pdfsDevolucion.length + ' PDF(s) de devolución (excluidos). Los PDFs de asignación deben no contener "devolucion" en el nombre.'
      : 'La carpeta está vacía o solo tiene archivos no-PDF.';
    Logger.log('');
    Logger.log('❌ Sin actas — ' + rep.error);
    Logger.log('💡 Sugerencia: ' + rep.sugerencia);
  }

  Logger.log('════════════════════════════════════════════════');
  return rep;
}

// Alias para ejecutar desde el editor rápidamente
function testBusquedaActa() {
  // ← CAMBIAR ESTOS VALORES ANTES DE EJECUTAR
  var NOMBRE = 'APELLIDO, NOMBRE';  // nombre exacto del colaborador como aparece en RH
  var DNI    = '';                   // DNI opcional, ayuda si el nombre no coincide
  return diagnosticarBusquedaActa({ nombre: NOMBRE, dni: DNI });
}

function testGenerarPDF() {
  Logger.log('=== TEST GENERAR PDF ===');
  try {
    // 1. Verificar acceso a Drive
    var raiz = DriveApp.getFolderById(CFG.DRIVE_ACTAS_ROOT_ID);
    Logger.log('✅ Carpeta raíz OK: ' + raiz.getName());

    // 2. Verificar acceso a carpeta de firmas
    var firmas = DriveApp.getFolderById(CFG.DRIVE_FIRMAS_TEC_ID);
    Logger.log('✅ Carpeta firmas OK: ' + firmas.getName());

    // 3. Buscar logo
    var logoOk = false;
    ['LOGO.png','Vive_Sharf.png','logo.jpg'].forEach(function(n){
      var it = firmas.getFilesByName(n);
      if (it.hasNext()) { Logger.log('✅ Logo encontrado: ' + n); logoOk = true; }
    });
    if (!logoOk) Logger.log('⚠️  No se encontró logo (LOGO.png o Vive_Sharf.png)');

    // 4. Generar HTML mínimo y convertir a PDF
    var html = '<!DOCTYPE html><html><body>' +
      '<h1 style="color:#68002B;font-family:Arial">SHARF - Test PDF</h1>' +
      '<p style="font-family:Arial">PDF generado: ' + new Date().toISOString() + '</p>' +
      '</body></html>';
    var htmlBlob = Utilities.newBlob(html, 'text/html', 'test.html');
    Logger.log('Convirtiendo HTML a PDF...');
    var driveTemp = DriveApp.createFile(htmlBlob);
    var pdfBlob   = driveTemp.getAs('application/pdf');
    Logger.log('✅ PDF generado: ' + pdfBlob.getBytes().length + ' bytes');

    // 5. Guardar en carpeta raíz (test)
    pdfBlob.setName('TEST_PDF_' + Utilities.formatDate(new Date(), CFG.TZ, 'yyyyMMdd_HHmm') + '.pdf');
    var pdfFile = raiz.createFile(pdfBlob);
    Logger.log('✅ PDF guardado en Drive: ' + pdfFile.getUrl());

    // 6. Limpiar temp
    try { driveTemp.setTrashed(true); } catch(e) {}
    // Opcional: borrar el test también
    // try { pdfFile.setTrashed(true); } catch(e) {}

    Logger.log('=== TEST PDF EXITOSO ✅ ===');
    Logger.log('El PDF funciona correctamente.');
    Logger.log('Si el PDF no se genera en la app, verificar:');
    Logger.log('  1. Que ST.activoSel tenga datos válidos');
    Logger.log('  2. Que ST.empleado.nombre no esté vacío');
    Logger.log('  3. Revisar el log de procesarDevolucion() para el error exacto');

  } catch(e) {
    Logger.log('❌ ERROR: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    Logger.log('');
    Logger.log('POSIBLES CAUSAS:');
    Logger.log('  • El script no tiene permiso de Drive — re-autoriza el script');
    Logger.log('  • DRIVE_ACTAS_ROOT_ID incorrecto: ' + CFG.DRIVE_ACTAS_ROOT_ID);
    Logger.log('  • La carpeta no existe o fue eliminada');
  }
}

function _clasificarTipoActivo(texto) {
  var t = (texto || '').toUpperCase();
  // Categorías reales del catálogo Snipe-IT SHARF:
  // [6] LAPTOP, [4] DESKTOP, [26] MiniDesk, [21] MONITOR, [5] TABLET, [32] CELULAR
  if (t.indexOf('LAPTOP') >= 0 || t.indexOf('NOTEBOOK') >= 0)             return 'Laptop';
  if (t.indexOf('ALL IN ONE') >= 0 || t.indexOf('ALL-IN-ONE') >= 0)        return 'All-in-One';
  if (t.indexOf('MINIDESKTOP') >= 0 || t.indexOf('MINI DESKTOP') >= 0 ||
      t.indexOf('MINIDESKT') >= 0 || t.indexOf('MINIDESK') >= 0)           return 'Minidesktop';
  if (t.indexOf('DESKTOP') >= 0 || t.indexOf('TOWER') >= 0)                return 'PC';
  if (t.indexOf('MONITOR') >= 0)                                            return 'Monitor';
  if (t.indexOf('TABLET') >= 0)                                             return 'Tablet';
  if (t.indexOf('CELULAR') >= 0 || t.indexOf('SMARTPHONE') >= 0)           return 'Celular';
  if (t.indexOf('IMPRESORA') >= 0 || t.indexOf('PRINTER') >= 0)            return 'Impresora';
  if (t.indexOf('ESCANER') >= 0 || t.indexOf('SCANNER') >= 0)              return 'Escáner';
  return 'Laptop'; // default para activos TI sin categoría identificada
}

// ════════════════════════════════════════════════════════════════════════
// 16. HELPERS DRIVE
// ════════════════════════════════════════════════════════════════════════
function _carpetaDevolucionesUsuario(empleado) {
  var nomCarpeta = (empleado && empleado.nombre) ? empleado.nombre.toUpperCase().trim() : 'SIN_NOMBRE';
  var raiz = DriveApp.getFolderById(CFG.DRIVE_ACTAS_ROOT_ID);

  // Buscar carpeta del usuario (ya existe, creada por SharfTI asignaciones)
  var carpetaUsuario = null;
  var iter = raiz.getFoldersByName(nomCarpeta);
  if (iter.hasNext()) {
    carpetaUsuario = iter.next();
  } else {
    // Si no existe crearla
    carpetaUsuario = raiz.createFolder(nomCarpeta);
  }

  // Subcarpeta Devoluciones
  var iterDev = carpetaUsuario.getFoldersByName('Devoluciones');
  if (iterDev.hasNext()) return iterDev.next();
  return carpetaUsuario.createFolder('Devoluciones');
}

function _driveSetPublicReader(fileId) {
  try {
    var token = ScriptApp.getOAuthToken();
    UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + fileId + '/permissions?supportsAllDrives=true',
      {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify({ role: 'reader', type: 'anyone' }),
        muteHttpExceptions: true
      }
    );
    return true;
  } catch(e) { Logger.log('_driveSetPublicReader: ' + e.message); return false; }
}

function _imgDrive(nombreArchivo) {
  try {
    var folder = DriveApp.getFolderById(CFG.DRIVE_FIRMAS_TEC_ID);
    // Intentar con el nombre solicitado primero, luego alternativas del logo
    var nombres = [nombreArchivo];
    if (nombreArchivo.toUpperCase() === 'LOGO.PNG')
      nombres = ['LOGO.png', 'Vive_Sharf.png', 'logo.jpg', 'LOGO.jpg'];
    for (var i = 0; i < nombres.length; i++) {
      var fs = folder.getFilesByName(nombres[i]);
      if (fs.hasNext()) {
        var b = fs.next().getBlob();
        return 'data:' + (b.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(b.getBytes());
      }
    }
    return null;
  } catch(e) { Logger.log('_imgDrive(' + nombreArchivo + '): ' + e.message); return null; }
}

function _imgDriveTec(firmaKey) {
  if (!firmaKey) return null;
  try {
    var files = DriveApp.getFolderById(CFG.DRIVE_FIRMAS_TEC_ID).getFiles();
    while (files.hasNext()) {
      var f = files.next();
      if (f.getName().toUpperCase().replace(/\.[^.]+$/, '') === firmaKey.toUpperCase()) {
        var b = f.getBlob();
        return 'data:' + (b.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(b.getBytes());
      }
    }
    return null;
  } catch(e) { Logger.log('_imgDriveTec: ' + e.message); return null; }
}

function _obtenerFirmaTecBlob(firmaKey) {
  if (!firmaKey) return null;
  try {
    var files = DriveApp.getFolderById(CFG.DRIVE_FIRMAS_TEC_ID).getFiles();
    while (files.hasNext()) {
      var f = files.next();
      if (f.getName().toUpperCase().replace(/\.[^.]+$/, '') === firmaKey.toUpperCase()) return f.getBlob();
    }
    return null;
  } catch(e) { Logger.log('_obtenerFirmaTecBlob: ' + e.message); return null; }
}


function _slug(texto) {
  return (texto || '').replace(/[,\s]+/g,'_').replace(/[^a-zA-Z0-9_]/g,'').slice(0,40);
}

// ════════════════════════════════════════════════════════════════════════
// MODO PRUEBA — gestión de correos override para pruebas en caliente
//   Almacena la configuración en PropertiesService del script
//   Los correos originales NO se modifican — solo se redirigen durante el envío
// ════════════════════════════════════════════════════════════════════════
function _getModoPrueba_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(CFG.MODO_PRUEBA_PROP);
    if (!raw) return { activo: false };
    return JSON.parse(raw);
  } catch(e) { return { activo: false }; }
}

function getModoPrueba() {
  var mp = _getModoPrueba_();
  return {
    ok:         true,
    activo:     mp.activo     || false,
    emailPara:  mp.emailPara  || '',
    emailCC:    mp.emailCC    || '',
    emailBCC:   mp.emailBCC   || '',
    // Siempre devolver los originales para que el panel los muestre
    originalPara: '(correo personal — cese / correo corp — otros tipos)',
    originalCC:   'jefe + Sheily+Melanie (cese) / Gabriel+Anais (renting, no usará, equipo)',
    originalBCC:  '(ninguno)'
  };
}

function setModoPrueba(params) {
  /*
    params = {
      activo:    bool,
      emailPara: string,   correo override del Para (colaborador)
      emailCC:   string,   correo override del CC
      emailBCC:  string    correo override del BCC
    }
  */
  try {
    var mp = {
      activo:    !!params.activo,
      emailPara: (params.emailPara || '').trim(),
      emailCC:   (params.emailCC   || '').trim(),
      emailBCC:  (params.emailBCC  || '').trim(),
      setBy:     Session.getActiveUser().getEmail(),
      setAt:     Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy HH:mm:ss')
    };
    PropertiesService.getScriptProperties().setProperty(
      CFG.MODO_PRUEBA_PROP, JSON.stringify(mp)
    );
    Logger.log((mp.activo ? '⚠️ MODO PRUEBA ACTIVADO' : '✅ MODO PRUEBA DESACTIVADO') +
               ' por ' + mp.setBy + ' a las ' + mp.setAt);
    return { ok: true, modoPrueba: mp };
  } catch(e) {
    Logger.log('setModoPrueba ERROR: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════════════════════
// CATÁLOGOS SNIPE PARA FORMULARIO MANUAL
//   Retorna ubicaciones y departamentos para dropdowns en modo manual
// ════════════════════════════════════════════════════════════════════════
function getSnipeCatalogos() {
  var resultado = { ok: true, ubicaciones: [], departamentos: [] };
  try {
    var locs = snipeGET('/locations?limit=100&sort=name&order=asc');
    if (locs && locs.rows) {
      resultado.ubicaciones = locs.rows.map(function(l) {
        return { id: l.id, nombre: l.name };
      });
    }
  } catch(e) { Logger.log('getSnipeCatalogos ubicaciones: ' + e.message); }
  try {
    var deps = snipeGET('/departments?limit=100&sort=name&order=asc');
    if (deps && deps.rows) {
      resultado.departamentos = deps.rows.map(function(d) {
        return { id: d.id, nombre: d.name };
      });
    }
  } catch(e) { Logger.log('getSnipeCatalogos departamentos: ' + e.message); }
  return resultado;
}

// ════════════════════════════════════════════════════════════════════════
// getCecosSnipe — extrae los valores únicos del campo "Centro de Costos"
//   de los activos en Snipe-IT, ordenados alfabéticamente.
//   Los CECOs viven en custom_fields["Centro de Costos"].value de cada activo.
//   Paginamos hasta 500 activos para cubrir todo el inventario.
// ════════════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════════════
// AUDITORÍA: URLs del Sheet y exportación XLSX
// ════════════════════════════════════════════════════════════════════════

// Retorna la URL del Sheet de log y de la carpeta Drive donde vive
// ════════════════════════════════════════════════════════════════════════
// AUDITORÍA DETALLE — devuelve el listado de devoluciones de un segmento
//   params.tipo:   'tecnico' | 'mes' | 'tipoDev' | 'empresa' | 'area'
//   params.valor:  el label del segmento clickeado
//   params.meses:  período a analizar
// ════════════════════════════════════════════════════════════════════════
function getAuditoriaDetalle(params) {
  try {
    var tipo   = params.tipo  || '';
    var valor  = params.valor || '';
    var meses  = parseInt(params.meses) || 99;

    var hoja  = _sheetLog();
    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return { ok: true, filas: [] };

    var hdrs = datos[0];

    function col(nombre, fallback) {
      var n = nombre.toLowerCase().trim();
      for (var i = 0; i < hdrs.length; i++) {
        if (String(hdrs[i]).toLowerCase().trim() === n) return i;
      }
      for (var j = 0; j < hdrs.length; j++) {
        if (String(hdrs[j]).toLowerCase().indexOf(n) === 0) return j;
      }
      return fallback !== undefined ? fallback : -1;
    }

    var iDate  = col('Fecha/Hora', 1);
    var iTipo  = col('Tipo Devolución', 2);
    var iColab = col('Colaborador', 4);
    var iEmp   = col('Empresa', 6);
    var iArea  = col('Área', 7);
    var iSerial= col('Serial', 12);
    var iActivo= col('Tipo Activo', 10);
    var iEst   = col('Estado Equipo', 20);
    var iTec   = -1;
    for (var hi = 0; hi < hdrs.length; hi++) {
      var hn = String(hdrs[hi]).toLowerCase().trim();
      if (hn === 'técnico' || hn === 'tecnico') { iTec = hi; break; }
    }
    if (iTec < 0) iTec = 29;

    var ahora  = new Date();
    var limite = new Date(ahora.getFullYear(), ahora.getMonth() - meses, 1);

    var MAPA = { tecnico: iTec, mes: iDate, tipoDev: iTipo,
                 empresa: iEmp, area: iArea };
    var iCol = MAPA[tipo];
    if (iCol === undefined) return { ok: false, error: 'Tipo no reconocido: ' + tipo };

    var filas = [];
    datos.slice(1).forEach(function(f) {
      if (!f[iDate]) return;
      var d = new Date(f[iDate]);
      if (d < limite) return;

      var clave = '';
      if (tipo === 'mes') {
        clave = Utilities.formatDate(d, CFG.TZ, 'yyyy-MM');
      } else {
        clave = String(f[iCol] || '').trim();
      }

      if (clave !== valor) return;

      var ts = f[iDate] ? Utilities.formatDate(new Date(f[iDate]), CFG.TZ, 'dd/MM/yyyy HH:mm') : '';
      filas.push({
        fecha:    ts,
        colab:    String(f[iColab]  || '').trim(),
        tipo:     String(f[iTipo]   || '').trim(),
        serial:   String(f[iSerial] || '').trim(),
        activo:   String(f[iActivo] || '').trim(),
        empresa:  String(f[iEmp]    || '').trim(),
        area:     String(f[iArea]   || '').trim(),
        tecnico:  String(f[iTec]    || '').trim(),
        estado:   String(f[iEst]    || '').trim()
      });
    });

    // Ordenar por fecha descendente
    filas.sort(function(a, b) { return b.fecha.localeCompare(a.fecha); });

    return { ok: true, tipo: tipo, valor: valor, total: filas.length, filas: filas };

  } catch(e) {
    Logger.log('getAuditoriaDetalle ERROR: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════════════════════
// LIMPIAR DRIVE — Elimina carpetas sin nombre de técnico válido
//   Busca en la raíz de actas carpetas cuyo nombre sea:
//   'SIN_NOMBRE', vacías sin PDFs, o que no correspondan a ningún colaborador.
//   Solo elimina carpetas vacías — nunca borra archivos.
// ════════════════════════════════════════════════════════════════════════
function limpiarDriveHuerfanos() {
  try {
    var raiz     = DriveApp.getFolderById(CFG.DRIVE_ACTAS_ROOT_ID);
    var iter     = raiz.getFolders();
    var eliminadas = [];
    var revisadas  = [];

    while (iter.hasNext()) {
      var carpeta   = iter.next();
      var nombre    = carpeta.getName();
      var esHuerfana = false;
      var razon      = '';

      // Carpeta SIN_NOMBRE
      if (nombre === 'SIN_NOMBRE' || nombre === '' || nombre === 'undefined') {
        esHuerfana = true;
        razon = 'nombre inválido';
      }

      // Carpeta completamente vacía (sin subcarpetas ni archivos)
      if (!esHuerfana) {
        var tieneArchivos = carpeta.getFiles().hasNext();
        var tieneSubcarpetas = carpeta.getFolders().hasNext();
        if (!tieneArchivos && !tieneSubcarpetas) {
          esHuerfana = true;
          razon = 'carpeta vacía';
        }
      }

      if (esHuerfana) {
        try {
          carpeta.setTrashed(true);
          eliminadas.push(nombre + ' (' + razon + ')');
          Logger.log('Carpeta huérfana eliminada: "' + nombre + '" — ' + razon);
        } catch(eDel) {
          Logger.log('No se pudo eliminar "' + nombre + '": ' + eDel.message);
        }
      } else {
        revisadas.push(nombre);
      }
    }

    Logger.log('limpiarDriveHuerfanos: ' + eliminadas.length + ' eliminadas, ' +
               revisadas.length + ' conservadas');
    return {
      ok:         true,
      eliminadas: eliminadas,
      conservadas: revisadas.length,
      mensaje:    eliminadas.length + ' carpeta(s) huérfana(s) eliminada(s)'
    };

  } catch(e) {
    Logger.log('limpiarDriveHuerfanos ERROR: ' + e.message);
    return { ok: false, error: e.message };
  }
}

function getAuditoriaUrls() {
  try {
    var ssId = PropertiesService.getScriptProperties().getProperty('SHEET_LOG_ID');
    var sheetUrl   = '';
    var folderUrl  = '';
    var sheetNombre = '';

    if (ssId) {
      try {
        var ss = SpreadsheetApp.openById(ssId);
        sheetUrl    = ss.getUrl();
        sheetNombre = ss.getName();
      } catch(e) { Logger.log('getAuditoriaUrls sheet: ' + e.message); }
    }

    try {
      var folder  = DriveApp.getFolderById(CFG.DRIVE_SEGUIMIENTO_FOLDER);
      folderUrl   = 'https://drive.google.com/drive/folders/' + CFG.DRIVE_SEGUIMIENTO_FOLDER;
    } catch(e) { Logger.log('getAuditoriaUrls folder: ' + e.message); }

    return {
      ok:          true,
      sheetUrl:    sheetUrl,
      sheetNombre: sheetNombre,
      folderUrl:   folderUrl,
      ssId:        ssId || ''
    };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Exporta el Sheet de log completo como .xlsx y retorna base64
// El rango exportado puede filtrarse por período si params.meses está dado
function exportarAuditoriaXLSX(params) {
  try {
    var ssId = PropertiesService.getScriptProperties().getProperty('SHEET_LOG_ID');
    if (!ssId) return { ok: false, error: 'Sheet de log no inicializado. Ejecuta inicializarSheetLog() primero.' };

    var ss   = SpreadsheetApp.openById(ssId);
    var hoja = ss.getSheetByName(CFG.SHEET_LOG_NAME);
    if (!hoja) return { ok: false, error: 'Hoja Devoluciones_TI no encontrada.' };

    var gid  = hoja.getSheetId();
    var meses = (params && params.meses) ? parseInt(params.meses) : 99;

    // Construir URL de exportación de Google Sheets a XLSX
    // Incluye solo la hoja del log, con todos los datos
    var exportUrl = 'https://docs.google.com/spreadsheets/d/' + ssId +
      '/export?format=xlsx' +
      '&gid=' + gid +
      '&portrait=false' +
      '&fitw=true' +
      '&sheetnames=true' +
      '&gridlines=true';

    var token = ScriptApp.getOAuthToken();
    var resp  = UrlFetchApp.fetch(exportUrl, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      return { ok: false, error: 'Error exportando XLSX: HTTP ' + resp.getResponseCode() };
    }

    var bytes  = resp.getContent();
    var b64    = Utilities.base64Encode(bytes);
    var fecha  = Utilities.formatDate(new Date(), CFG.TZ, 'yyyyMMdd');

    Logger.log('exportarAuditoriaXLSX: OK — ' + bytes.length + ' bytes');
    return {
      ok:       true,
      base64:   b64,
      nombre:   'SHARF_Auditoria_Devoluciones_TI_' + fecha + '.xlsx',
      tamanio:  bytes.length
    };

  } catch(e) {
    Logger.log('exportarAuditoriaXLSX ERROR: ' + e.message);
    return { ok: false, error: e.message };
  }
}

function getCecosSnipe() {
  try {
    var cecos = {};
    var offset = 0;
    var limit  = 500;
    var maxPags = 4;  // máximo 2000 activos para no exceder tiempo de ejecución

    for (var p = 0; p < maxPags; p++) {
      var resp = snipeGET('/hardware?limit=' + limit + '&offset=' + offset +
                          '&sort=asset_tag&order=asc');
      if (!resp || !resp.rows || resp.rows.length === 0) break;

      resp.rows.forEach(function(a) {
        var cf = a.custom_fields || {};
        // El campo puede llamarse "Centro de Costos", "CECO", etc.
        var ceco = '';
        var posiblesNombres = ['Centro de Costos','Centro de costos','CECO','ceco',
                               'Cost Center','CostCenter'];
        for (var ni = 0; ni < posiblesNombres.length; ni++) {
          if (cf[posiblesNombres[ni]] && cf[posiblesNombres[ni]].value) {
            ceco = String(cf[posiblesNombres[ni]].value).trim();
            break;
          }
        }
        if (ceco && ceco !== '' && ceco.toUpperCase() !== 'N/A' && ceco !== '0') {
          cecos[ceco] = true;
        }
      });

      if (resp.rows.length < limit) break;  // no hay más páginas
      offset += limit;
    }

    var lista = Object.keys(cecos).sort(function(a, b) {
      return a.localeCompare(b, 'es', { numeric: true });
    });

    Logger.log('getCecosSnipe: ' + lista.length + ' CECOs únicos encontrados');
    return { ok: true, cecos: lista };

  } catch(e) {
    Logger.log('getCecosSnipe ERROR: ' + e.message);
    return { ok: false, error: e.message, cecos: [] };
  }
}

// ════════════════════════════════════════════════════════════════════════
// BUSCAR ACTAS DE ASIGNACIÓN EN DRIVE
//   Busca SOLO archivos .pdf (no .docx, no de devolución) en la carpeta
//   del usuario, dentro de DRIVE_ASIGNACIONES_ROOT_ID si está configurado,
//   o en la raíz del usuario en DRIVE_ACTAS_ROOT_ID como fallback.
//   NO busca en subcarpetas (solo la raíz de la carpeta del usuario).
//   Devuelve la lista ordenada por fecha descendente (más reciente primero).
// ════════════════════════════════════════════════════════════════════════
function buscarActaAsignacion(params) {
  var nombreUsuario = (typeof params === 'string') ? params : (params && params.nombreUsuario || '');
  var dniUsuario    = (typeof params === 'object' && params) ? (params.dni || '') : '';
  if (!nombreUsuario && !dniUsuario) return { ok: false, error: 'Nombre o DNI requerido' };

  try {
    function norm(s) {
      return (s || '').toUpperCase().trim()
        .replace(/[ÁÀÂÄ]/g,'A').replace(/[ÉÈÊË]/g,'E')
        .replace(/[ÍÌÎÏ]/g,'I').replace(/[ÓÒÔÖ]/g,'O')
        .replace(/[ÚÙÛÜ]/g,'U').replace(/[Ñ]/g,'N')
        .replace(/\s+/g,' ');
    }

    var busq   = norm(nombreUsuario);
    var tokens = busq.split(/[,\s]+/).filter(function(t){ return t.length >= 3; });

    // ── Buscar carpeta del usuario en una raíz dada ──────────────────────
    function buscarCarpetaEnRaiz(raizFolder) {
      // a) Exacta normalizada
      var it = raizFolder.getFoldersByName(busq);
      if (it.hasNext()) { var c = it.next(); return { c: c, n: c.getName() }; }
      // b) Exacta sin normalizar
      it = raizFolder.getFoldersByName(nombreUsuario.toUpperCase().trim());
      if (it.hasNext()) { var c2 = it.next(); return { c: c2, n: c2.getName() }; }
      // c) Por tokens con scoring
      if (tokens.length > 0) {
        var allF = raizFolder.getFolders();
        var best = 0, bestF = null;
        while (allF.hasNext()) {
          var f = allF.next();
          var fn = norm(f.getName());
          var score = tokens.filter(function(t){ return fn.indexOf(t) >= 0; }).length;
          if (dniUsuario && fn.indexOf(dniUsuario) >= 0) score += 3;
          if (score > best) { best = score; bestF = f; }
        }
        var umbral = tokens.length === 1 ? 1 : 2;
        if (best >= umbral && bestF) return { c: bestF, n: bestF.getName() };
      }
      return { c: null, n: '' };
    }

    // ── Buscar carpeta: primero en ASIGNACIONES, luego en ACTAS ──────────
    var carpeta = null, carpetaNombre = '';

    if (CFG.DRIVE_ASIGNACIONES_ROOT_ID) {
      try {
        var res1 = buscarCarpetaEnRaiz(DriveApp.getFolderById(CFG.DRIVE_ASIGNACIONES_ROOT_ID));
        if (res1.c) { carpeta = res1.c; carpetaNombre = res1.n; }
      } catch(e) { Logger.log('buscarActa ASIGNACIONES error: ' + e.message); }
    }

    if (!carpeta) {
      var res2 = buscarCarpetaEnRaiz(DriveApp.getFolderById(CFG.DRIVE_ACTAS_ROOT_ID));
      if (res2.c) { carpeta = res2.c; carpetaNombre = res2.n; }
    }

    if (!carpeta) {
      return {
        ok: false, noEncontrado: true,
        error: 'No se encontró carpeta para "' + (nombreUsuario || dniUsuario) + '" en Drive.'
      };
    }

    var carpetaUrl = carpeta.getUrl();

    // ── Recopilar PDFs de asignación ────────────────────────────────────
    // Reglas:
    //   · Solo archivos .pdf
    //   · Excluir los que tengan "devolucion" / "dev_" en el nombre
    //   · Buscar en: (1) raíz de la carpeta del usuario
    //                (2) subcarpeta llamada "Asignaciones" si existe
    //   · IGNORAR la subcarpeta "Devoluciones" completamente

    function esAsignacion(nombre) {
      var n = nombre.toLowerCase();
      return n.endsWith('.pdf') &&
             n.indexOf('devolucion') < 0 &&
             n.indexOf('devolución') < 0 &&
             n.indexOf('dev_') < 0;
    }

    var pdfs = [];
    var vistosIds = {};

    function recopilarPDFsDeFolder(folder) {
      var fi = folder.getFiles();
      while (fi.hasNext()) {
        var f = fi.next();
        if (!esAsignacion(f.getName())) continue;
        if (vistosIds[f.getId()]) continue;
        vistosIds[f.getId()] = true;
        var created = f.getDateCreated();
        pdfs.push({
          nombre: f.getName(),
          url:    f.getUrl(),
          id:     f.getId(),
          fecha:  Utilities.formatDate(created, CFG.TZ, 'dd/MM/yyyy HH:mm'),
          ts:     created.getTime()
        });
      }
    }

    // 1. Raíz de la carpeta del usuario
    recopilarPDFsDeFolder(carpeta);

    // 2. Subcarpetas:
    //    · IGNORAR cualquier cosa que empiece con "devoluc"
    //    · ENTRAR en todas las demás (Asignacion-DD-MM-YYYY, Asignaciones, etc.)
    var itSub = carpeta.getFolders();
    while (itSub.hasNext()) {
      var sub     = itSub.next();
      var subNorm = sub.getName().toLowerCase().trim();
      // Saltar devoluciones en cualquier formato
      if (subNorm.indexOf('devoluc') === 0 || subNorm === 'dev' || subNorm.indexOf('dev_') === 0) continue;
      // Entrar en todas las demás subcarpetas (asignacion-*, asignaciones, otros)
      recopilarPDFsDeFolder(sub);
    }

    if (!pdfs.length) {
      return {
        ok: false, noEncontrado: true,
        carpetaNombre: carpetaNombre,
        carpetaUrl:    carpetaUrl,
        error: 'No se encontraron PDFs de asignación en la carpeta "' + carpetaNombre + '".'
      };
    }

    // Ordenar por fecha descendente — el más reciente primero
    pdfs.sort(function(a, b){ return b.ts - a.ts; });

    Logger.log('buscarActa: ' + pdfs.length + ' PDF(s) de asignación para "' + carpetaNombre + '"');

    return {
      ok:            true,
      actas:         pdfs,
      carpetaNombre: carpetaNombre,
      carpetaUrl:    carpetaUrl,
      total:         pdfs.length
    };

  } catch(e) {
    Logger.log('buscarActaAsignacion ERROR: ' + e.message);
    return { ok: false, error: 'Error al buscar en Drive: ' + e.message };
  }
}

// ════════════════════════════════════════════════════════════════════════
// NUEVA: REENVIAR CORREO (cuando el primero falló)
//   Permite corregir el correo y reintentar sin volver a generar todo
// ════════════════════════════════════════════════════════════════════════
function reenviarCorreo(params) {
  /*
    params = {
      datos:         <objeto empleado/activo/accesorios completo>,
      ts:            timestamp original,
      nombre:        string,
      serial:        string,
      devId:         string,
      tipoActivo:    string,
      pdfDriveId:    string,
      emailCorregido: string  (opcional — correo personal corregido)
      emailJefeCorregido: string (opcional)
    }
  */
  try {
    var datos = params.datos;
    if (!datos || !datos.empleado) return { ok: false, error: 'Datos incompletos para reenvío' };

    // Aplicar correos corregidos si se proporcionaron
    if (params.emailCorregido)    datos.empleado.emailPersonal = params.emailCorregido;
    if (params.emailJefeCorregido) datos.empleado.emailJefe    = params.emailJefeCorregido;

    var ts         = params.ts     || Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy HH:mm:ss');
    var nombre     = params.nombre || datos.empleado.nombre || '';
    var serial     = params.serial || (datos.activo ? datos.activo.serial : '') || '';
    var devId      = params.devId  || 'DEV-REENVIO';
    var tipoActivo = params.tipoActivo || _clasificarTipoActivo(datos.activo ? datos.activo.nombre : '');

    var logoData     = _imgDrive('LOGO.png');
    var firmaTecData = null;
    try { firmaTecData = _imgDriveTec(datos.tecnico ? datos.tecnico.firmaKey : ''); } catch(ef) {}

    // Recuperar PDF desde Drive si existe
    var pdfBlob = null;
    if (params.pdfDriveId) {
      try {
        var pdfFetch = UrlFetchApp.fetch(
          'https://www.googleapis.com/drive/v3/files/' + params.pdfDriveId + '?alt=media&supportsAllDrives=true',
          { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
        );
        if (pdfFetch.getResponseCode() === 200)
          pdfBlob = pdfFetch.getBlob().setName('Devolucion_' + _slug(nombre) + '_' + serial + '.pdf');
      } catch(ef2) { Logger.log('reenviarCorreo PDF fetch: ' + ef2.message); }
    }

    var emailRes = _enviarCorreoDevolucion(datos, ts, nombre, serial, devId, tipoActivo, logoData, firmaTecData, pdfBlob);

    return {
      ok: emailRes.ok,
      msg: emailRes.ok
        ? 'Correo reenviado a: ' + (emailRes.para || '')
        : (emailRes.error || 'Error al reenviar'),
      emailInfo: emailRes
    };
  } catch(e) {
    Logger.log('reenviarCorreo: ' + e.message);
    var esInvalido = e.message.indexOf('Invalid email') >= 0 || e.message.indexOf('recipient') >= 0;
    return { ok: false, error: e.message, correoInvalido: esInvalido };
  }
}


// ════════════════════════════════════════════════════════════════════════
// SETUP Y DIAGNÓSTICO
// ════════════════════════════════════════════════════════════════════════
function setupApiKey() {
  setKeyManual();
}

/**
 * Guarda la API Key de Snipe-IT en las propiedades del script.
 * Ejecutar UNA sola vez desde el editor, luego borrar la línea del token.
 *
 * Uso:
 *   1. Reemplaza TU_TOKEN_AQUI con el Bearer Token real
 *   2. Ejecuta esta función desde el editor (▶ Run)
 *   3. Borra el token del código (seguridad)
 */
/**
 * testSnipeUsuario — muestra TODOS los campos de un usuario en Snipe-IT.
 * Ejecutar para confirmar en qué campo está guardado el DNI.
 * Ver resultado en: Ejecutar → Registros de ejecución
 */
function testSnipeUsuario() {
  var DNI_PRUEBA = '70406659';   // DNI de prueba

  Logger.log('═══ TEST USUARIO SNIPE: ' + DNI_PRUEBA + ' ═══');

  // Buscar por texto libre — devuelve cualquier campo que coincida
  var r = snipeGET('/users?search=' + encodeURIComponent(DNI_PRUEBA) + '&limit=10');

  if (!r || !r.rows || !r.rows.length) {
    Logger.log('❌ Ningún usuario encontrado buscando: ' + DNI_PRUEBA);
    Logger.log('   Verifica que el usuario existe en Snipe-IT con ese DNI.');
    return { encontrado: false };
  }

  Logger.log('Usuarios encontrados: ' + r.rows.length);

  r.rows.forEach(function(u, i) {
    Logger.log('');
    Logger.log('── Usuario #' + (i+1) + ' ──────────────────');
    Logger.log('  id:            ' + u.id);
    Logger.log('  name:          ' + u.name);
    Logger.log('  username:      ' + (u.username      || '(vacío)'));
    Logger.log('  employee_num:  ' + (u.employee_num  || '(vacío)'));
    Logger.log('  email:         ' + (u.email         || '(vacío)'));
    Logger.log('  department:    ' + ((u.department && u.department.name) || '(vacío)'));
    Logger.log('  location:      ' + ((u.location   && u.location.name)   || '(vacío)'));
    Logger.log('  company:       ' + ((u.company    && u.company.name)    || '(vacío)'));
    // Mostrar custom_fields si existen
    if (u.custom_fields) {
      Logger.log('  custom_fields:');
      Object.keys(u.custom_fields).forEach(function(k) {
        Logger.log('    "' + k + '": ' + JSON.stringify(u.custom_fields[k]));
      });
    }
    Logger.log('  ── ¿Dónde está el DNI? ──');
    Logger.log('  username = DNI?      ' + (u.username     === DNI_PRUEBA ? '✅ SÍ' : '❌ NO (' + u.username + ')'));
    Logger.log('  employee_num = DNI?  ' + (u.employee_num === DNI_PRUEBA ? '✅ SÍ' : '❌ NO (' + (u.employee_num||'vacío') + ')'));
  });

  return r.rows;
}

/**
 * testConexionSnipeCompleto
 * Ejecutar desde el editor → Ver → Registros de ejecución
 * Muestra el JSON RAW del usuario y sus activos para diagnóstico definitivo
 */
function testConexionSnipeCompleto() {
  var DNI = '70406659';
  Logger.log('═══ DIAGNÓSTICO SNIPE COMPLETO ═══');

  // 1. Buscar usuario por employee_num
  Logger.log('\n[1] GET /users?employee_num=' + DNI);
  try {
    var r1 = snipeGET('/users?employee_num=' + encodeURIComponent(DNI) + '&limit=5');
    Logger.log('Total: ' + (r1.total || 0) + '  Rows: ' + (r1.rows ? r1.rows.length : 0));
    if (r1.rows && r1.rows.length) {
      r1.rows.forEach(function(u) { Logger.log('  → ' + JSON.stringify({id:u.id,name:u.name,username:u.username,employee_num:u.employee_num,email:u.email})); });
    }
  } catch(e) { Logger.log('ERROR: ' + e.message); }

  // 2. Buscar usuario por search=DNI
  Logger.log('\n[2] GET /users?search=' + DNI);
  try {
    var r2 = snipeGET('/users?search=' + encodeURIComponent(DNI) + '&limit=10');
    Logger.log('Total: ' + (r2.total || 0) + '  Rows: ' + (r2.rows ? r2.rows.length : 0));
    if (r2.rows && r2.rows.length) {
      r2.rows.forEach(function(u) { Logger.log('  → ' + JSON.stringify({id:u.id,name:u.name,username:u.username,employee_num:u.employee_num,email:u.email})); });
    }
  } catch(e) { Logger.log('ERROR: ' + e.message); }

  // 3. Primeros 10 usuarios — ver estructura real
  Logger.log('\n[3] GET /users?limit=10 (primeros 10 usuarios del sistema)');
  try {
    var r3 = snipeGET('/users?limit=10&sort=id&order=asc');
    Logger.log('Total usuarios en Snipe: ' + (r3.total || 0));
    if (r3.rows && r3.rows.length) {
      r3.rows.forEach(function(u) { Logger.log('  → ' + JSON.stringify({id:u.id,name:u.name,username:u.username,employee_num:u.employee_num,email:u.email})); });
    }
  } catch(e) { Logger.log('ERROR: ' + e.message); }

  // 4. Si encontramos el usuario en [1] o [2], buscar sus activos
  Logger.log('\n[4] Buscar activos asignados al usuario (si se encontró)');
  try {
    var rU = snipeGET('/users?employee_num=' + encodeURIComponent(DNI) + '&limit=1');
    var userId = (rU.rows && rU.rows.length) ? rU.rows[0].id : null;
    if (!userId) {
      var rU2 = snipeGET('/users?search=' + encodeURIComponent(DNI) + '&limit=1');
      userId = (rU2.rows && rU2.rows.length) ? rU2.rows[0].id : null;
    }
    if (userId) {
      Logger.log('Usuario ID: ' + userId + ' → buscando activos...');
      var rA = snipeGET('/hardware?assigned_to=' + userId + '&assigned_type=user&limit=50');
      Logger.log('Activos asignados: ' + (rA.total || 0));
      if (rA.rows && rA.rows.length) {
        rA.rows.forEach(function(a) {
          Logger.log('  → ' + JSON.stringify({id:a.id,name:a.name,serial:a.serial,asset_tag:a.asset_tag,
            status:(a.status_label&&a.status_label.name)||'',category:(a.category&&a.category.name)||'',
            assigned_to:(a.assigned_to&&a.assigned_to.name)||''}));
        });
      } else {
        // También intentar sin assigned_type por si está configurado diferente
        Logger.log('Sin activos con assigned_type=user. Intentando sin ese parámetro...');
        var rA2 = snipeGET('/hardware?assigned_to=' + userId + '&limit=50');
        Logger.log('Activos (sin assigned_type): ' + (rA2.total || 0));
        if (rA2.rows) rA2.rows.forEach(function(a) { Logger.log('  → ' + a.name + ' S/N:' + a.serial); });
      }
    } else {
      Logger.log('No se encontró el usuario con DNI ' + DNI + ' — no se pueden buscar activos');
      Logger.log('ACCIÓN REQUERIDA: Confirma en Snipe-IT en qué campo está el DNI de este usuario');
    }
  } catch(e) { Logger.log('ERROR activos: ' + e.message); }

  Logger.log('\n═══ FIN DIAGNÓSTICO ═══');
}

/**
 * testSnipeActivo
 * Inspecciona la estructura completa de un activo real en Snipe-IT
 * y prueba si el token tiene permisos de escritura (PATCH).
 *
 * Ejecutar desde el editor → Ver → Registros de ejecución
 * Reemplaza SERIAL_PRUEBA con un serial real antes de ejecutar.
 */
function testSnipeActivo() {
  var SERIAL_PRUEBA = '4CE321BHXK';  // ← serial del activo de prueba

  Logger.log('═══ TEST ACTIVO SNIPE: ' + SERIAL_PRUEBA + ' ═══');

  // 1. Buscar el activo por serial
  var activo = null;
  try {
    var r = snipeGET('/hardware?search=' + encodeURIComponent(SERIAL_PRUEBA) + '&limit=5');
    Logger.log('Búsqueda serial → rows: ' + (r.rows ? r.rows.length : 0));
    if (r.rows && r.rows.length) {
      activo = r.rows.filter(function(a) {
        return (a.serial || '').toUpperCase() === SERIAL_PRUEBA.toUpperCase();
      })[0] || r.rows[0];
    }
  } catch(e) { Logger.log('ERROR búsqueda: ' + e.message); return; }

  if (!activo) { Logger.log('Activo no encontrado'); return; }

  Logger.log('\n── Campos estándar ──');
  Logger.log('id:           ' + activo.id);
  Logger.log('name:         ' + activo.name);
  Logger.log('serial:       ' + activo.serial);
  Logger.log('asset_tag:    ' + activo.asset_tag);
  Logger.log('status:       ' + JSON.stringify(activo.status_label));
  Logger.log('category:     ' + JSON.stringify(activo.category));
  Logger.log('location:     ' + JSON.stringify(activo.location));
  Logger.log('rtd_location: ' + JSON.stringify(activo.rtd_location));
  Logger.log('company:      ' + JSON.stringify(activo.company));
  Logger.log('assigned_to:  ' + JSON.stringify(activo.assigned_to));

  Logger.log('\n── Custom Fields ──');
  if (activo.custom_fields) {
    Object.keys(activo.custom_fields).forEach(function(k) {
      Logger.log('  "' + k + '": ' + JSON.stringify(activo.custom_fields[k]));
    });
  } else {
    Logger.log('  (sin custom_fields en la respuesta)');
  }

  // 2. Obtener el activo por ID para ver todos los campos
  Logger.log('\n── GET /hardware/' + activo.id + ' (detalle completo) ──');
  try {
    var det = snipeGET('/hardware/' + activo.id);
    Logger.log('custom_fields detalle:');
    if (det.custom_fields) {
      Object.keys(det.custom_fields).forEach(function(k) {
        var cf = det.custom_fields[k];
        Logger.log('  field: "' + k + '"  field_id:' + (cf.field || '?') + '  value:"' + (cf.value || '') + '"');
      });
    }
  } catch(e) { Logger.log('GET detalle ERROR: ' + e.message); }

  // 3. Probar PATCH — intentar actualizar la location del activo (sin cambiarla)
  Logger.log('\n── TEST PATCH /hardware/' + activo.id + ' ──');
  try {
    var patchPayload = {};
    // Solo enviamos el mismo location_id que ya tiene (no cambia nada)
    if (activo.location && activo.location.id) {
      patchPayload.location_id = activo.location.id;
    } else {
      patchPayload.notes = activo.notes || '';  // campo seguro para probar
    }
    var patchResp = UrlFetchApp.fetch(CFG.SNIPE_BASE + '/hardware/' + activo.id, {
      method: 'PATCH',
      headers: {
        'Authorization': 'Bearer ' + _snipeKey(),
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(patchPayload),
      muteHttpExceptions: true
    });
    var patchCode = patchResp.getResponseCode();
    var patchBody = JSON.parse(patchResp.getContentText());
    Logger.log('PATCH HTTP: ' + patchCode);
    Logger.log('PATCH status: ' + (patchBody.status || '?'));
    Logger.log('PATCH messages: ' + JSON.stringify(patchBody.messages || patchBody.error || ''));
    if (patchCode === 200 && patchBody.status === 'success') {
      Logger.log('✅ Token tiene permisos de ESCRITURA sobre activos');
    } else {
      Logger.log('❌ Sin permisos de escritura o error: HTTP ' + patchCode);
    }
  } catch(e) { Logger.log('PATCH ERROR: ' + e.message); }

  Logger.log('\n═══ FIN TEST ═══');
}

/**
 * testSnipeUploadPDF
 * ─────────────────────────────────────────────────────────────────────────
 * Prueba ESPECÍFICA de subida de archivo al activo 4CE321BHXK (id=2659)
 * Crea un PDF de prueba mínimo y lo sube al endpoint /hardware/{id}/files
 * Ejecutar desde el editor → Ver → Registros de ejecución
 * ─────────────────────────────────────────────────────────────────────────
 */
function testSnipeUploadPDF() {
  var ASSET_ID   = 2659;    // id del activo 4CE321BHXK confirmado en el log
  var ASSET_SN   = '4CE321BHXK';
  var snipeKey   = _snipeKey();

  Logger.log('═══ TEST UPLOAD PDF → Snipe activo ' + ASSET_ID + ' (' + ASSET_SN + ') ═══');

  // ── 1. Crear un PDF de prueba mínimo usando HTML→Drive ─────────────
  var htmlPrueba =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>' +
    '<h2 style="font-family:Arial">PRUEBA DE SUBIDA DE ARCHIVO</h2>' +
    '<p style="font-family:Arial">Activo: ' + ASSET_SN + ' | Fecha: ' +
    Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy HH:mm:ss') + '</p>' +
    '<p style="font-family:Arial">Este archivo es generado automáticamente por el sistema SHARF Devoluciones TI.</p>' +
    '</body></html>';

  var pdfBlob = null;
  var driveTemp = null;
  try {
    var htmlBlob = Utilities.newBlob(htmlPrueba, 'text/html', 'test.html');
    driveTemp    = DriveApp.createFile(htmlBlob);
    pdfBlob      = driveTemp.getAs('application/pdf');
    pdfBlob.setName('TEST_Upload_' + ASSET_SN + '.pdf');
    Logger.log('PDF de prueba creado: ' + pdfBlob.getBytes().length + ' bytes');
  } catch(e) {
    Logger.log('ERROR creando PDF de prueba: ' + e.message);
    return;
  } finally {
    if (driveTemp) try { driveTemp.setTrashed(true); } catch(e2) {}
  }

  // ── 2. Intentar subida con campo "file" ────────────────────────────
  Logger.log('\n── Intento 1: field name = "file" ──');
  var resultado1 = _intentarUpload(ASSET_ID, pdfBlob, 'file', snipeKey);
  Logger.log('HTTP: ' + resultado1.code + ' | Body: ' + resultado1.body.slice(0, 300));
  if (resultado1.code === 200) {
    try {
      var b1 = JSON.parse(resultado1.body);
      if (b1.status === 'success') {
        Logger.log('✅ ÉXITO con field "file" — el PDF fue subido al activo en Snipe-IT');
        Logger.log('═══ FIN TEST UPLOAD ═══');
        return;
      }
      Logger.log('HTTP 200 pero status: ' + (b1.status || '?') + ' — ' + JSON.stringify(b1.messages || b1.error || ''));
    } catch(ep) { Logger.log('HTTP 200 pero body no es JSON válido'); }
  }

  // ── 3. Retry con campo "files[]" ───────────────────────────────────
  Logger.log('\n── Intento 2: field name = "files[]" ──');
  var resultado2 = _intentarUpload(ASSET_ID, pdfBlob, 'files[]', snipeKey);
  Logger.log('HTTP: ' + resultado2.code + ' | Body: ' + resultado2.body.slice(0, 300));
  if (resultado2.code === 200) {
    try {
      var b2 = JSON.parse(resultado2.body);
      if (b2.status === 'success') {
        Logger.log('✅ ÉXITO con field "files[]" — el PDF fue subido al activo en Snipe-IT');
        Logger.log('ACCION: cambiar fieldName en el código principal a "files[]"');
        Logger.log('═══ FIN TEST UPLOAD ═══');
        return;
      }
    } catch(ep2) {}
  }

  // ── 4. Retry con campo "file_upload" ──────────────────────────────
  Logger.log('\n── Intento 3: field name = "file_upload" ──');
  var resultado3 = _intentarUpload(ASSET_ID, pdfBlob, 'file_upload', snipeKey);
  Logger.log('HTTP: ' + resultado3.code + ' | Body: ' + resultado3.body.slice(0, 300));

  // ── 5. Diagnóstico final ───────────────────────────────────────────
  Logger.log('\n── DIAGNÓSTICO ──');
  if (resultado1.code === 403 || resultado2.code === 403) {
    Logger.log('❌ HTTP 403 — El token NO tiene permisos para subir archivos.');
    Logger.log('   SOLUCION: Ir a scharff.snipe-it.io → Admin → Usuarios → usuario del API token');
    Logger.log('   → verificar que el rol sea "Superadmin" o "Admin"');
  } else if (resultado1.code === 401) {
    Logger.log('❌ HTTP 401 — Token inválido o expirado. Ejecutar setKeyManual() con un token nuevo.');
  } else if (resultado1.code === 404) {
    Logger.log('❌ HTTP 404 — El endpoint /hardware/' + ASSET_ID + '/files no existe en esta versión de Snipe-IT.');
  } else {
    Logger.log('Ningún intento fue exitoso. Revisa los body arriba para más detalles.');
  }
  Logger.log('═══ FIN TEST UPLOAD ═══');
}

function _intentarUpload(assetId, pdfBlob, fieldName, snipeKey) {
  try {
    // En Apps Script, la única forma confiable de hacer multipart es
    // pasar payload como objeto {fieldName: blob} SIN especificar contentType.
    // Apps Script genera el boundary automáticamente solo así.
    var payload = {};
    // Asegurarse que el blob tiene el tipo MIME correcto
    var uploadBlob = pdfBlob;
    try { uploadBlob = pdfBlob.getAs('application/pdf'); } catch(e) {}
    uploadBlob.setName(pdfBlob.getName() || 'devolucion.pdf');
    payload[fieldName] = uploadBlob;

    var resp = UrlFetchApp.fetch(
      CFG.SNIPE_BASE + '/hardware/' + assetId + '/files',
      {
        method:             'POST',
        headers:            { 'Authorization': 'Bearer ' + snipeKey, 'Accept': 'application/json' },
        payload:            payload
        // NO contentType — Apps Script lo maneja solo con el boundary correcto
      }
    );
    return { code: resp.getResponseCode(), body: resp.getContentText() };
  } catch(e) {
    return { code: 0, body: 'ERROR: ' + e.message };
  }
}

/**
 * testSnipeUploadPDF_Debug
 * Versión de diagnóstico profundo — muestra exactamente qué envía Apps Script
 * Ejecutar desde el editor para ver headers y payload enviados
 */
function testSnipeUploadPDF_Debug() {
  var ASSET_ID = 2659;
  var snipeKey = _snipeKey();
  Logger.log('═══ DEBUG UPLOAD ═══');

  // Crear PDF mínimo
  var htmlB = Utilities.newBlob(
    '<!DOCTYPE html><html><body><p>TEST SHARF ' +
    Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy HH:mm:ss') +
    '</p></body></html>', 'text/html', 'test.html');
  var driveF = DriveApp.createFile(htmlB);
  var pdfB   = driveF.getAs('application/pdf');
  pdfB.setName('TEST_DEBUG_' + ASSET_ID + '.pdf');
  try { driveF.setTrashed(true); } catch(e) {}
  Logger.log('PDF: ' + pdfB.getBytes().length + ' bytes, type: ' + pdfB.getContentType());

  // Mostrar primeros bytes del payload multipart para diagnóstico
  // Probar con requestbin o httpbin para ver exactamente qué llega
  var testUrl = 'https://httpbin.org/post';
  try {
    var debugResp = UrlFetchApp.fetch(testUrl, {
      method:  'POST',
      headers: { 'Authorization': 'Bearer TEST', 'Accept': 'application/json' },
      payload: { 'file[]': pdfB },
      muteHttpExceptions: true
    });
    var debugBody = JSON.parse(debugResp.getContentText());
    Logger.log('HTTPBin files recibidos: ' + JSON.stringify(debugBody.files || {}));
    Logger.log('HTTPBin form recibido: ' + JSON.stringify(debugBody.form || {}));
    Logger.log('HTTPBin headers: Content-Type = ' + (debugBody.headers && debugBody.headers['Content-Type'] ? debugBody.headers['Content-Type'] : 'no encontrado'));
  } catch(eDebug) {
    Logger.log('httpbin test error: ' + eDebug.message);
  }

  // Ahora el intento real a Snipe
  Logger.log('\n── Intento real a Snipe ──');
  var r = UrlFetchApp.fetch(
    CFG.SNIPE_BASE + '/hardware/' + ASSET_ID + '/files',
    {
      method:             'POST',
      headers:            { 'Authorization': 'Bearer ' + snipeKey, 'Accept': 'application/json' },
      payload:            { 'file[]': pdfB },
      muteHttpExceptions: true
    }
  );
  Logger.log('HTTP: ' + r.getResponseCode());
  Logger.log('Body: ' + r.getContentText().slice(0, 400));
  Logger.log('═══ FIN DEBUG ═══');
}
function setKeyManual() {
  var token = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiNjZlZmMyOTllOTMyMzVmMzJmMTg5NDg3MGMyYzNhYzQ0ODhmNTZkNTJlNjEwNjQ2NWM1N2FmNmJkYjdjNDE4YjNlNmI4NmQ2Y2I4ZWQ0OTUiLCJpYXQiOjE3NzI5MzM2NjguMzM2MzA0LCJuYmYiOjE3NzI5MzM2NjguMzM2MzA5LCJleHAiOjI0MDQwODU2NjguMzE2Mjg2LCJzdWIiOiI3Iiwic2NvcGVzIjpbXX0.KhW_WJ6Y4hTtlL6bNC9wXlLDiiHdsptWSTfa2NXn9iLiKAgLnTgJBzbGGhi83hjz7y3adyc5slRS-YEvDqdlVY-VERHtlU6JAIRuYcrzFZlbaEBD-iuuaalA6OnfphG4C-5OLZd1Y0BZWXna37f6yNns0EPwRnhdINL2QqV7Y-pehzAYraZfmZg6aOL5HENH-WiWRt_kP305u3mECOZhUMjsiNZIygqG1MipcXB6NMGtfvaD6JmkhHUNFpwnLZdgfHjWsIekAxVCZjeRvf2oRRg6z9T3U3QGks3lHF46yhb1qo9vLrA9wMngjiEkWi80pbi3005hPCvePm01b78lZNJ1Ckxlxq45A0vtg7OPBfQWzRbFG_NZp6lOLDfEotxEddbfUy0EiWQUPAs7oNK8tKh5jPYZvvwhSWGC2EpD3BTMpfxlmU_kCRJ3PobOGk9o_4GQuYkD3n52vkXoAzrJcYikJCcRdNN2isn7JPWH386VeVE8mQxNieS0H2g6bzOuFKMghKcp_uHf1eaV6hDq98Tv_6kmAK1PXISY0lzD7l_7S849iKpoYQscy2xHtwRXYIgCWSGY339aUG3iuqbg4MDynisnfurABSj9zEUgYy2DKKLgDnugq2IXtJP3dlV0aB6qUQ89d6mzAUfRtPd39SlvVvtN8piVSFt2YAgMS78';
  PropertiesService.getScriptProperties().setProperty(CFG.SNIPE_KEY_PROP, token);
  Logger.log('✅ SNIPE_API_KEY guardada correctamente.');
}

/**
 * inicializarSheetLog
 *
 * Crea el Google Sheet "SHARF_Log_Devoluciones_TI" dentro de la carpeta
 * Drive de helpdesk (DRIVE_SEGUIMIENTO_FOLDER = 0APknu_tBOg5SUk9PVA).
 *
 * Si el Sheet ya existe en esa carpeta lo reutiliza.
 * Guarda el ID del Sheet en las PropertiesService para uso posterior.
 */
// ════════════════════════════════════════════════════════════════════════
// MÓDULO DE AUDITORÍA
//   Genera un resumen gerencial desde el Sheet de log.
//   Acceso restringido: Anais Chero, Eddie Fernandez, Ismael Gomez.
//   params.meses: número de meses a analizar (default 3)
// ════════════════════════════════════════════════════════════════════════
function getAuditoriaData(params) {
  try {
    var meses = (params && params.meses) ? parseInt(params.meses) : 3;
    var hoja  = _sheetLog();
    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return { ok: true, vacio: true, filas: 0 };

    var hdrs = datos[0];

    // ── Búsqueda de columna robusta ───────────────────────────────────────
    // Primero exacta (toLowerCase), luego substring como fallback.
    // Esto evita que "Técnico" matchee "Email Técnico" antes que "Técnico".
    function col(nombre, fallbackPos) {
      var n = nombre.toLowerCase().trim();
      // 1. Coincidencia exacta
      for (var i = 0; i < hdrs.length; i++) {
        if (String(hdrs[i]).toLowerCase().trim() === n) return i;
      }
      // 2. Empieza con el nombre (ej: "Técnico" no matchea "Email Técnico")
      for (var j = 0; j < hdrs.length; j++) {
        if (String(hdrs[j]).toLowerCase().trim().indexOf(n) === 0) return j;
      }
      // 3. Substring (cualquier posición)
      for (var k = 0; k < hdrs.length; k++) {
        if (String(hdrs[k]).toLowerCase().indexOf(n) >= 0) return k;
      }
      // 4. Posición fija (0-based) — para Sheets con encabezados viejos
      return (fallbackPos !== undefined) ? fallbackPos : -1;
    }

    // Columnas del Sheet (encabezados nuevos — posiciones 0-based como fallback)
    // ID(0) Fecha(1) TipoDev(2) ModoIngreso(3) Colaborador(4) DNI(5) Empresa(6)
    // Area(7) Cargo(8) Sede(9) TipoActivo(10) NombreActivo(11) Serial(12)
    // AssetTag(13) SnipeID(14) Cargador(15) Mochila(16) Mouse(17) Docking(18) Teclado(19)
    // EstadoEquipo(20) ObsEquipo(21) CalidadEquipo(22) EstadoAccesorios(23)
    // AccNoDevueltos(24) CantNoDevueltos(25) AccCotizar(26) CantCotizaciones(27)
    // ObsAccesorio(28) Tecnico(29) EmailTecnico(30) SedeTecnico(31)
    // Snipe(32) PDF(33) Drive(34) Email(35) UrlDrive(36) PdfId(37) SnipeDesact(38)
    var iDate    = col('Fecha/Hora',                 1);
    var iTipo    = col('Tipo Devolución',             2);
    var iColab   = col('Colaborador',                 4);
    var iEmp     = col('Empresa',                     6);
    var iArea    = col('Área',                        7);
    var iActivo  = col('Tipo Activo',                10);
    var iEstado  = col('Estado Equipo',              20);
    var iNoDev   = col('Accesorios NO Devueltos',    24);
    var iCantND  = col('Cant. No Devueltos',         25);
    var iCotiz   = col('Accesorios a Cotizar',       26);
    var iCantCot = col('Cant. Cotizaciones',         27);
    var iSnipe   = col('Snipe',                      32);  // resultado Snipe (no Snipe ID)
    var iEmail   = col('Email',                      35);  // resultado Email (no Email Técnico)

    // Para "Técnico" buscamos exacto primero para no confundir con "Email Técnico"
    // Posición 29 en el nuevo formato
    var iTec = -1;
    for (var hi = 0; hi < hdrs.length; hi++) {
      if (String(hdrs[hi]).toLowerCase().trim() === 'técnico' ||
          String(hdrs[hi]).toLowerCase().trim() === 'tecnico') {
        iTec = hi; break;
      }
    }
    // Fallback: buscar "Técnico" que NO sea "Email Técnico" ni "Sede Técnico"
    if (iTec < 0) {
      for (var hj = 0; hj < hdrs.length; hj++) {
        var hn = String(hdrs[hj]).toLowerCase().trim();
        if (hn === 'técnico' || hn === 'tecnico' ||
            (hn.indexOf('técnico') >= 0 && hn.indexOf('email') < 0 && hn.indexOf('sede') < 0)) {
          iTec = hj; break;
        }
      }
    }
    if (iTec < 0) iTec = 29;  // posición fija definitiva

    // Log diagnóstico de columnas detectadas
    Logger.log('getAuditoriaData — columnas detectadas:');
    Logger.log('  iTec=' + iTec + ' ("' + (hdrs[iTec]||'?') + '")');
    Logger.log('  iArea=' + iArea + ' ("' + (hdrs[iArea]||'?') + '")');
    Logger.log('  iTipo=' + iTipo + ' ("' + (hdrs[iTipo]||'?') + '")');
    Logger.log('  iEmp=' + iEmp + ' ("' + (hdrs[iEmp]||'?') + '")');
    Logger.log('  iSnipe=' + iSnipe + ' ("' + (hdrs[iSnipe]||'?') + '")');
    Logger.log('  iEmail=' + iEmail + ' ("' + (hdrs[iEmail]||'?') + '")');
    Logger.log('  Total hdrs=' + hdrs.length + ' | Encabezados: ' +
               hdrs.slice(0,10).map(function(h,i){return i+'='+h;}).join(', '));

    var ahora  = new Date();
    var limite = new Date(ahora.getFullYear(), ahora.getMonth() - meses, 1);
    var filas  = datos.slice(1).filter(function(f) {
      if (!f[iDate]) return false;
      var d = new Date(f[iDate]);
      return d >= limite;
    });

    var total = filas.length;
    if (total === 0) return { ok: true, vacio: true, filas: 0, meses: meses };

    var porTipo = {}, porTec = {}, porEmp = {}, porArea = {}, porActivo = {};
    var conObs = 0, sinObs = 0;
    var totalNoDev = 0, totalCotiz = 0;
    var errSnipe = 0, errEmail = 0;
    var porMes = {};
    var accNoDev = { cargador:0, mouse:0, mochila:0, docking:0, teclado:0 };
    var accCotiz = { cargador:0, mouse:0, mochila:0, docking:0, teclado:0 };

    filas.forEach(function(f) {
      var tipo   = String(f[iTipo]  || '').trim();
      var tec    = String(f[iTec]   || '').trim();
      var emp    = String(f[iEmp]   || '').trim();
      var area   = String(f[iArea]  || '').trim();
      var activo = String(f[iActivo]|| '').trim();
      var estado = String(f[iEstado]|| '').toUpperCase();
      var cantND = parseInt(f[iCantND]) || 0;
      var cantCot= parseInt(f[iCantCot])|| 0;
      var noDev  = String(f[iNoDev] || '').toLowerCase();
      var cotiz  = String(f[iCotiz] || '').toLowerCase();

      // Filtrar valores vacíos o inválidos (que se colaron por columna errónea)
      if (!tec || tec === '—') tec = '(Sin nombre)';
      // Si el técnico parece ser una observación de equipo, hay desalineamiento de columna
      var INVALIDOS_TEC = ['bueno','malo','con observ','sin observ','equipo','golpe','raya','daño'];
      var tecLow = tec.toLowerCase();
      var esTecInvalido = INVALIDOS_TEC.some(function(k){ return tecLow.indexOf(k) >= 0; });
      if (esTecInvalido) {
        tec = '(columna desalineada — ver log)';
        Logger.log('WARN: técnico parece observación: "' + tec + '" en fila iTec=' + iTec);
      }

      if (tipo) porTipo[tipo]   = (porTipo[tipo]   || 0) + 1;
      if (tec)  porTec[tec]     = (porTec[tec]     || 0) + 1;
      if (emp)  porEmp[emp]     = (porEmp[emp]      || 0) + 1;
      if (area) porArea[area]   = (porArea[area]    || 0) + 1;
      if (activo) porActivo[activo] = (porActivo[activo] || 0) + 1;

      if (estado.indexOf('CON') >= 0) conObs++; else sinObs++;
      totalNoDev += cantND;
      totalCotiz += cantCot;

      ['cargador','mouse','mochila','docking','teclado'].forEach(function(k) {
        if (noDev.indexOf(k) >= 0) accNoDev[k]++;
        if (cotiz.indexOf(k) >= 0) accCotiz[k]++;
      });

      if (String(f[iSnipe]||'').indexOf('ERROR') >= 0) errSnipe++;
      if (String(f[iEmail]||'').indexOf('ERROR') >= 0) errEmail++;

      var d = new Date(f[iDate]);
      var mes = Utilities.formatDate(d, CFG.TZ, 'yyyy-MM');
      porMes[mes] = (porMes[mes] || 0) + 1;
    });

    var mesesOrdenados = Object.keys(porMes).sort().map(function(m) {
      return { mes: m, total: porMes[m] };
    });

    function top5(obj) {
      return Object.keys(obj)
        .filter(function(k){ return k && k !== '(Sin nombre)'; })
        .map(function(k){ return { label:k, val:obj[k] }; })
        .sort(function(a,b){ return b.val - a.val; })
        .slice(0,8);
    }

    return {
      ok:      true,
      filas:   total,
      meses:   meses,
      periodo: Utilities.formatDate(limite, CFG.TZ, 'dd/MM/yyyy') + ' — ' +
               Utilities.formatDate(ahora,  CFG.TZ, 'dd/MM/yyyy'),
      colsDetectadas: {
        iTec: iTec, nomTec: String(hdrs[iTec]||'?'),
        iArea: iArea, nomArea: String(hdrs[iArea]||'?')
      },
      resumen: {
        total:      total,
        conObs:     conObs,
        sinObs:     sinObs,
        pctObs:     Math.round(conObs / total * 100),
        totalNoDev: totalNoDev,
        totalCotiz: totalCotiz,
        errSnipe:   errSnipe,
        errEmail:   errEmail
      },
      porTipo:    top5(porTipo),
      porTecnico: top5(porTec),
      porEmpresa: top5(porEmp),
      porArea:    top5(porArea),
      porActivo:  top5(porActivo),
      porMes:     mesesOrdenados,
      accNoDev:   accNoDev,
      accCotiz:   accCotiz
    };

  } catch(e) {
    Logger.log('getAuditoriaData ERROR: ' + e.message);
    return { ok: false, error: e.message };
  }
}

function inicializarSheetLog() {
  var NOMBRE_FILE  = 'SHARF_Log_Devoluciones_TI';
  var FOLDER_ID    = CFG.DRIVE_SEGUIMIENTO_FOLDER;

  Logger.log('Inicializando Sheet de log en carpeta: ' + FOLDER_ID);

  // ── 1. Obtener la carpeta destino ────────────────────────────────────
  var folder;
  try {
    folder = DriveApp.getFolderById(FOLDER_ID);
    Logger.log('Carpeta encontrada: ' + folder.getName());
  } catch(e) {
    Logger.log('❌ No se pudo acceder a la carpeta ' + FOLDER_ID + ': ' + e.message);
    Logger.log('   Asegúrate de que helpdesk@holasharf.com compartió la carpeta con tu cuenta.');
    throw new Error('Carpeta Drive no accesible: ' + e.message);
  }

  // ── 2. Buscar si ya existe el Sheet en esa carpeta ───────────────────
  var ssId    = '';
  var existentes = folder.getFilesByName(NOMBRE_FILE);
  if (existentes.hasNext()) {
    var archivo = existentes.next();
    ssId = archivo.getId();
    Logger.log('Sheet ya existe: ' + ssId);
  } else {
    // ── 3. Crear el Sheet dentro de la carpeta usando Drive API v3 ─────
    //    SpreadsheetApp.create() siempre crea en "Mi unidad".
    //    Para crearlo directamente en una Shared/carpeta específica
    //    usamos la Drive REST API.
    var token = ScriptApp.getOAuthToken();
    var metaResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true',
      {
        method:  'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type':  'application/json'
        },
        payload: JSON.stringify({
          name:     NOMBRE_FILE,
          mimeType: 'application/vnd.google-apps.spreadsheet',
          parents:  [FOLDER_ID]
        }),
        muteHttpExceptions: true
      }
    );

    if (metaResp.getResponseCode() !== 200) {
      Logger.log('❌ Error creando Sheet: HTTP ' + metaResp.getResponseCode());
      Logger.log(metaResp.getContentText());
      throw new Error('No se pudo crear el Sheet en Drive: ' + metaResp.getContentText().slice(0, 200));
    }

    ssId = JSON.parse(metaResp.getContentText()).id;
    Logger.log('✅ Sheet creado con ID: ' + ssId);
  }

  // ── 4. Guardar el ID en PropertiesService para uso del sistema ───────
  PropertiesService.getScriptProperties().setProperty('SHEET_LOG_ID', ssId);
  Logger.log('ID guardado en PropertiesService → SHEET_LOG_ID: ' + ssId);

  // ── 5. Crear/verificar la hoja "Devoluciones_TI" dentro del Sheet ────
  var ss = SpreadsheetApp.openById(ssId);
  var h  = ss.getSheetByName(CFG.SHEET_LOG_NAME);

  if (!h) {
    h = ss.insertSheet(CFG.SHEET_LOG_NAME);
    Logger.log('Hoja "' + CFG.SHEET_LOG_NAME + '" creada.');
  } else {
    Logger.log('Hoja "' + CFG.SHEET_LOG_NAME + '" ya existe — verificando headers.');
    if (h.getLastRow() > 0) {
      Logger.log('✅ Sheet listo (ya tiene datos). No se modificó.');
      return;
    }
  }

  // ── 6. Insertar fila de encabezados ──────────────────────────────────
  var hdrs = [
    'ID Devolución', 'Fecha/Hora', 'Tipo Devolución', 'Modo Ingreso',
    'Colaborador', 'DNI', 'Empresa', 'Área', 'Cargo', 'Sede',
    'Tipo Activo', 'Nombre Activo', 'Serial', 'Asset Tag', 'Snipe ID',
    'Cargador Dev.', 'Mochila Dev.', 'Mouse Dev.', 'Docking Dev.', 'Teclado Dev.',
    'Estado Equipo', 'Observaciones Equipo', 'Calidad Equipo',
    'Estado Accesorios',
    'Accesorios NO Devueltos', 'Cant. No Devueltos',
    'Accesorios a Cotizar', 'Cant. Cotizaciones',
    'Obs. por Accesorio',
    'Técnico', 'Email Técnico', 'Sede Técnico',
    'Snipe', 'PDF', 'Drive', 'Email',
    'URL Drive', 'PDF ID', 'Snipe Desactualizado'
  ];
  h.appendRow(hdrs);
  h.getRange(1, 1, 1, hdrs.length)
    .setBackground(CFG.BRAND_COLOR)
    .setFontColor('#000000')
    .setFontWeight('bold')
    .setFontSize(9);
  h.setFrozenRows(1);
  h.setColumnWidth(1, 160);   // ID Devolución
  h.setColumnWidth(2, 140);   // Fecha/Hora
  h.setColumnWidth(5, 200);   // Colaborador
  h.setColumnWidth(28, 300);  // URL Drive

  Logger.log('✅ Sheet de log inicializado correctamente.');
  Logger.log('   Archivo: ' + NOMBRE_FILE);
  Logger.log('   ID:      ' + ssId);
  Logger.log('   Carpeta: ' + folder.getName() + ' (' + FOLDER_ID + ')');
  Logger.log('   URL:     https://docs.google.com/spreadsheets/d/' + ssId);
}

/**
 * _sheetLog
 * Obtiene la hoja "Devoluciones_TI" del Sheet de log.
 * Busca el ID en PropertiesService (guardado por inicializarSheetLog).
 * Si no existe el ID guardado, intenta buscarlo por nombre en la carpeta.
 */
function _sheetLog() {
  // Intentar obtener el ID guardado
  var ssId = PropertiesService.getScriptProperties().getProperty('SHEET_LOG_ID');

  if (!ssId) {
    // Fallback: buscar por nombre en la carpeta de seguimiento
    Logger.log('SHEET_LOG_ID no encontrado en PropertiesService. Buscando por nombre...');
    try {
      var folder = DriveApp.getFolderById(CFG.DRIVE_SEGUIMIENTO_FOLDER);
      var files  = folder.getFilesByName('SHARF_Log_Devoluciones_TI');
      if (files.hasNext()) {
        ssId = files.next().getId();
        PropertiesService.getScriptProperties().setProperty('SHEET_LOG_ID', ssId);
        Logger.log('Sheet encontrado y guardado: ' + ssId);
      }
    } catch(e) {
      Logger.log('_sheetLog fallback error: ' + e.message);
    }
  }

  if (!ssId) {
    // Si definitivamente no existe, crear
    Logger.log('Sheet de log no existe. Ejecutando inicializarSheetLog...');
    inicializarSheetLog();
    ssId = PropertiesService.getScriptProperties().getProperty('SHEET_LOG_ID');
  }

  var ss = SpreadsheetApp.openById(ssId);
  var h  = ss.getSheetByName(CFG.SHEET_LOG_NAME);
  if (!h) {
    inicializarSheetLog();
    ss = SpreadsheetApp.openById(ssId);
    h  = ss.getSheetByName(CFG.SHEET_LOG_NAME);
  }
  return h;
}

// ════════════════════════════════════════════════════════════════════════
// CHEQUEO COMPLETO DE ACCESOS + CATÁLOGOS SNIPE-IT
// Ejecutar desde el editor → Ver → Registros de ejecución
// ════════════════════════════════════════════════════════════════════════
function chequeoCompleto() {
  var R = {};   // objeto resultado — visible en panel Ejecución
  var L = [];   // líneas de log

  function sep(titulo) { L.push(''); L.push('══ ' + titulo + ' ══'); }
  function ok(msg)     { L.push('  ✅  ' + msg); }
  function err(msg)    { L.push('  ❌  ' + msg); }
  function warn(msg)   { L.push('  ⚠️   ' + msg); }
  function info(msg)   { L.push('       ' + msg); }

  L.push('SHARF Devoluciones TI · Chequeo Completo');
  L.push(Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy HH:mm:ss'));

  // ────────────────────────────────────────────────────────────────────
  // 1. SESIÓN
  // ────────────────────────────────────────────────────────────────────
  sep('SESIÓN');
  try {
    var emailActivo = Session.getActiveUser().getEmail();
    var emailEfec   = Session.getEffectiveUser().getEmail();
    ok('getActiveUser():    ' + (emailActivo || '(vacío)'));
    ok('getEffectiveUser(): ' + (emailEfec   || '(vacío)'));
    R.sesion = { activo: emailActivo, efectivo: emailEfec };
  } catch(e) { err('Sesión: ' + e.message); R.sesion = { error: e.message }; }

  // ────────────────────────────────────────────────────────────────────
  // 2. SNIPE-IT API KEY
  // ────────────────────────────────────────────────────────────────────
  sep('SNIPE-IT · API KEY');
  var snipeKey = PropertiesService.getScriptProperties().getProperty(CFG.SNIPE_KEY_PROP);
  if (!snipeKey) {
    err('SNIPE_API_KEY NO configurada — ejecuta setKeyManual()');
    R.snipeKey = false;
  } else {
    ok('SNIPE_API_KEY configurada (' + snipeKey.slice(0,6) + '...)');
    R.snipeKey = true;
  }

  // ────────────────────────────────────────────────────────────────────
  // 3. SNIPE-IT CONEXIÓN + CATÁLOGOS
  // ────────────────────────────────────────────────────────────────────
  sep('SNIPE-IT · CONEXIÓN Y CATÁLOGOS');
  R.snipe = {};

  if (snipeKey) {
    // 3a. Conexión general — total de activos
    try {
      var hw0 = snipeGET('/hardware?limit=1');
      ok('Conexión OK · Total activos en Snipe-IT: ' + (hw0.total || 0));
      R.snipe.totalActivos = hw0.total || 0;
      R.snipe.instancia    = CFG.SNIPE_BASE;
    } catch(e) {
      err('Conexión fallida: ' + e.message);
      R.snipe.error = e.message;
    }

    // 3c. Categorías
    try {
      var cats = snipeGET('/categories?limit=50');
      R.snipe.categorias = [];
      if (cats && cats.rows && cats.rows.length) {
        ok('Categorías (' + cats.rows.length + '):');
        cats.rows.forEach(function(c) {
          info('→ [' + c.id + '] ' + c.name + ' (' + (c.category_type||'') + ') — ' + (c.assets_count||0) + ' activos');
          R.snipe.categorias.push({ id: c.id, nombre: c.name, tipo: c.category_type, activos: c.assets_count });
        });
      } else {
        warn('Sin categorías en Snipe-IT');
      }
    } catch(e) { err('Categorías: ' + e.message); }

    // 3d. Estados / Status Labels
    try {
      var sts = snipeGET('/statuslabels?limit=50');
      R.snipe.estados = [];
      if (sts && sts.rows && sts.rows.length) {
        ok('Estados / Status Labels (' + sts.rows.length + '):');
        sts.rows.forEach(function(s) {
          info('→ [' + s.id + '] ' + s.name + ' (tipo: ' + (s.type||'') + ')');
          R.snipe.estados.push({ id: s.id, nombre: s.name, tipo: s.type });
        });
      } else {
        warn('Sin estados en Snipe-IT');
      }
    } catch(e) { err('Estados: ' + e.message); }

    // 3e. Fabricantes
    try {
      var mfg = snipeGET('/manufacturers?limit=30');
      R.snipe.fabricantes = [];
      if (mfg && mfg.rows && mfg.rows.length) {
        ok('Fabricantes (' + mfg.rows.length + '):');
        mfg.rows.forEach(function(m) {
          info('→ [' + m.id + '] ' + m.name);
          R.snipe.fabricantes.push({ id: m.id, nombre: m.name });
        });
      }
    } catch(e) { err('Fabricantes: ' + e.message); }

    // 3f. Empresas / Companies
    try {
      var comp = snipeGET('/companies?limit=30');
      R.snipe.empresas = [];
      if (comp && comp.rows && comp.rows.length) {
        ok('Empresas (' + comp.rows.length + '):');
        comp.rows.forEach(function(c) {
          info('→ [' + c.id + '] ' + c.name + ' — ' + (c.assets_count||0) + ' activos');
          R.snipe.empresas.push({ id: c.id, nombre: c.name, activos: c.assets_count });
        });
      }
    } catch(e) { err('Empresas: ' + e.message); }

    // 3g. Ubicaciones
    try {
      var locs = snipeGET('/locations?limit=50');
      R.snipe.ubicaciones = [];
      if (locs && locs.rows && locs.rows.length) {
        ok('Ubicaciones (' + locs.rows.length + '):');
        locs.rows.forEach(function(l) {
          info('→ [' + l.id + '] ' + l.name);
          R.snipe.ubicaciones.push({ id: l.id, nombre: l.name });
        });
      }
    } catch(e) { err('Ubicaciones: ' + e.message); }

  } else {
    warn('Snipe-IT omitido — sin API Key');
  }

  // ────────────────────────────────────────────────────────────────────
  // 4. SHEET RH — 3 PESTAÑAS
  // ────────────────────────────────────────────────────────────────────
  sep('SHEET RH · ' + CFG.SHEET_RH_ID);
  R.sheetRH = { accesible: false, pestanas: [] };
  try {
    var ss = SpreadsheetApp.openById(CFG.SHEET_RH_ID);
    ok('Accesible: ' + ss.getName());
    R.sheetRH.accesible = true;
    R.sheetRH.nombre    = ss.getName();

    var todasHojas = ss.getSheets();
    info('Total pestañas en el Sheet: ' + todasHojas.length);

    // Mostrar TODAS las pestañas del Sheet
    info('Listado completo:');
    todasHojas.forEach(function(h) {
      info('  · "' + h.getName() + '"  GID:' + h.getSheetId() + '  Filas:' + h.getLastRow());
    });

    // Verificar cada pestaña configurada
    L.push('');
    info('Verificando pestañas configuradas:');
    CFG.SHEET_RH_TABS.forEach(function(tabCfg) {
      var hoja = ss.getSheetByName(tabCfg.nombre);
      if (!hoja) {
        // Intentar por GID
        todasHojas.forEach(function(h) {
          if (String(h.getSheetId()) === String(tabCfg.gid)) hoja = h;
        });
      }
      var datPest = { nombre: tabCfg.nombre, gid: tabCfg.gid };
      if (hoja) {
        var lastRow  = hoja.getLastRow();
        var lastCol  = hoja.getLastColumn();
        var numDatos = Math.max(0, lastRow - CFG.SHEET_RH_HDR_ROW);
        ok('"' + hoja.getName() + '"  GID:' + hoja.getSheetId() +
           '  Filas datos: ' + numDatos + '  Columnas: ' + lastCol);

        // Leer encabezados (fila 4)
        if (lastRow >= CFG.SHEET_RH_HDR_ROW) {
          var hdrsArr = hoja.getRange(CFG.SHEET_RH_HDR_ROW, 1, 1, lastCol).getValues()[0];
          var hdrsLimpios = hdrsArr.map(function(h) { return String(h||'').trim(); })
                                   .filter(function(h) { return h !== ''; });
          info('  Encabezados (fila ' + CFG.SHEET_RH_HDR_ROW + '): ' + hdrsLimpios.join(' | '));

          // Verificar columnas clave
          var hdrsNorm = _hdrs(hdrsArr);
          var colsClave = ['dni','nombre','empresa','area','emailpers','emailcorp','cargo','centroops','ceco'];
          var okCols = [], faltanCols = [];
          colsClave.forEach(function(c) {
            if (hdrsNorm[c] !== undefined) okCols.push(c + '→col' + (hdrsNorm[c]+1));
            else faltanCols.push(c);
          });
          ok('  Columnas mapeadas: ' + okCols.join(', '));
          if (faltanCols.length > 0) {
            warn('  Sin mapear: ' + faltanCols.join(', ') + ' (revisa nombres de columna en el Sheet)');
            info('  Todas las claves normalizadas: ' + Object.keys(hdrsNorm).sort().join(', '));
          }
          datPest.columnas = hdrsLimpios;
          datPest.filasDatos = numDatos;
        }
      } else {
        err('"' + tabCfg.nombre + '" (GID:' + tabCfg.gid + ') — NO ENCONTRADA');
        datPest.error = 'No encontrada';
      }
      R.sheetRH.pestanas.push(datPest);
    });
  } catch(e) {
    err('No se pudo abrir el Sheet RH: ' + e.message);
    R.sheetRH.error = e.message;
  }

  // ────────────────────────────────────────────────────────────────────
  // 5. DRIVE — carpetas configuradas
  // ────────────────────────────────────────────────────────────────────
  sep('DRIVE · CARPETAS');
  R.drive = {};
  var carpetas = [
    { clave: 'DRIVE_ACTAS_ROOT_ID',      id: CFG.DRIVE_ACTAS_ROOT_ID,      label: 'Carpeta raíz de actas' },
    { clave: 'DRIVE_FIRMAS_TEC_ID',      id: CFG.DRIVE_FIRMAS_TEC_ID,      label: 'Carpeta firmas y logos' },
    { clave: 'DRIVE_SEGUIMIENTO_FOLDER', id: CFG.DRIVE_SEGUIMIENTO_FOLDER, label: 'Carpeta log seguimiento' }
  ];
  carpetas.forEach(function(c) {
    try {
      var f = DriveApp.getFolderById(c.id);
      ok(c.label + ': "' + f.getName() + '"  ID:' + c.id);
      R.drive[c.clave] = { ok: true, nombre: f.getName() };

      // Para carpeta de firmas, listar archivos
      if (c.clave === 'DRIVE_FIRMAS_TEC_ID') {
        var files = f.getFiles(), archivos = [];
        while (files.hasNext()) {
          var fi = files.next();
          archivos.push(fi.getName());
        }
        info('  Archivos (' + archivos.length + '): ' + archivos.join(', '));
        // Verificar qué firmas faltan
        var faltanFirmas = CFG.TECNICOS.filter(function(t) {
          return !archivos.some(function(a) {
            return a.toUpperCase().replace(/\.[^.]+$/,'') === t.firmaKey;
          });
        }).map(function(t) { return t.firmaKey + '.png'; });
        var tienelogo = archivos.some(function(a) { return a.toUpperCase() === 'LOGO.PNG'; });
        tienelogo ? ok('  LOGO.png: ✅') : err('  LOGO.png: NO encontrado');
        faltanFirmas.length === 0
          ? ok('  Firmas técnicos: todas presentes')
          : warn('  Firmas faltantes: ' + faltanFirmas.join(', '));
        R.drive[c.clave].archivos = archivos;
      }
    } catch(e) {
      err(c.label + ' (ID:' + c.id + '): ' + e.message);
      R.drive[c.clave] = { ok: false, error: e.message };
    }
  });

  // Sheet log (buscar por PropertiesService o por nombre en carpeta)
  var ssLogId = PropertiesService.getScriptProperties().getProperty('SHEET_LOG_ID');
  if (ssLogId) {
    try {
      var ssLog = SpreadsheetApp.openById(ssLogId);
      ok('Sheet Log: "' + ssLog.getName() + '"  ID:' + ssLogId);
      var hLog = ssLog.getSheetByName(CFG.SHEET_LOG_NAME);
      hLog
        ? ok('  Hoja "' + CFG.SHEET_LOG_NAME + '": ' + Math.max(0, hLog.getLastRow()-1) + ' registros')
        : warn('  Hoja "' + CFG.SHEET_LOG_NAME + '" no encontrada — ejecuta inicializarSheetLog()');
      R.drive.sheetLog = { ok: true, id: ssLogId, registros: hLog ? Math.max(0,hLog.getLastRow()-1) : 0 };
    } catch(e) {
      err('Sheet Log (ID guardado ' + ssLogId + '): ' + e.message);
      R.drive.sheetLog = { ok: false, error: e.message };
    }
  } else {
    warn('Sheet Log: ID no guardado — ejecuta inicializarSheetLog() primero');
    R.drive.sheetLog = { ok: false, error: 'SHEET_LOG_ID no en PropertiesService' };
  }

  // ────────────────────────────────────────────────────────────────────
  // 6. RESUMEN FINAL
  // ────────────────────────────────────────────────────────────────────
  sep('RESUMEN');
  var checks = [
    { label: 'API Key Snipe-IT',  ok: R.snipeKey },
    { label: 'Snipe-IT conexión', ok: !!(R.snipe && !R.snipe.error) },
    { label: 'Sheet RH',          ok: !!(R.sheetRH && R.sheetRH.accesible) },
    { label: 'Drive Actas',       ok: !!(R.drive && R.drive.DRIVE_ACTAS_ROOT_ID && R.drive.DRIVE_ACTAS_ROOT_ID.ok) },
    { label: 'Drive Firmas',      ok: !!(R.drive && R.drive.DRIVE_FIRMAS_TEC_ID && R.drive.DRIVE_FIRMAS_TEC_ID.ok) },
    { label: 'Drive Log',         ok: !!(R.drive && R.drive.DRIVE_SEGUIMIENTO_FOLDER && R.drive.DRIVE_SEGUIMIENTO_FOLDER.ok) },
    { label: 'Sheet Log',         ok: !!(R.drive && R.drive.sheetLog && R.drive.sheetLog.ok) }
  ];
  checks.forEach(function(c) {
    c.ok ? ok(c.label) : err(c.label);
  });
  var totalOk  = checks.filter(function(c){return c.ok;}).length;
  L.push('');
  L.push('  ' + totalOk + '/' + checks.length + ' checks pasados');
  if (totalOk === checks.length) {
    L.push('  🚀 TODO OK — el sistema está listo para usar');
  } else {
    L.push('  ⚠️  Revisa los items con ❌ antes de usar el aplicativo');
  }

  var logFinal = L.join('\n');
  Logger.log(logFinal);
  R._log = logFinal;
  return R;
}

function diagnosticarConexiones() {
  var log = [];
  log.push('═══════════════════════════════════════');
  log.push('SHARF Devoluciones TI · Diagnóstico');
  log.push(Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy HH:mm:ss'));
  log.push('═══════════════════════════════════════');

  // API Key Snipe-IT
  var key = PropertiesService.getScriptProperties().getProperty(CFG.SNIPE_KEY_PROP);
  log.push(key ? '✅ SNIPE_API_KEY: configurada' : '❌ SNIPE_API_KEY: NO configurada — ejecuta setKeyManual()');

  // Snipe-IT conexión
  if (key) {
    try {
      var r = snipeGET('/hardware?limit=1');
      log.push('✅ Snipe-IT API: OK · ' + (r.total || 0) + ' activos totales');
    } catch(e) { log.push('❌ Snipe-IT API: ' + e.message); }
  }

  // Sheet RH — verificar las 3 pestañas configuradas
  try {
    var ss = SpreadsheetApp.openById(CFG.SHEET_RH_ID);
    log.push('✅ Sheet RH: accesible · "' + ss.getName() + '"');
    var todasHojas = ss.getSheets();
    var nombresHojas = todasHojas.map(function(h) { return h.getName(); });
    log.push('   Pestañas: ' + nombresHojas.join(' | '));

    CFG.SHEET_RH_TABS.forEach(function(tabCfg) {
      // Buscar por nombre, si no por GID
      var hoja = ss.getSheetByName(tabCfg.nombre);
      if (!hoja) {
        todasHojas.forEach(function(h) {
          if (String(h.getSheetId()) === String(tabCfg.gid)) hoja = h;
        });
      }
      if (hoja) {
        var filasDatos = Math.max(0, hoja.getLastRow() - CFG.SHEET_RH_HDR_ROW);
        log.push('   ✅ "' + hoja.getName() + '" (GID:' + hoja.getSheetId() + ') · ' + filasDatos + ' registros');
      } else {
        log.push('   ❌ "' + tabCfg.nombre + '" (GID:' + tabCfg.gid + ') NO encontrada');
      }
    });
  } catch(e) { log.push('❌ Sheet RH: ' + e.message); }

  // Carpeta Drive actas
  try {
    var fa = DriveApp.getFolderById(CFG.DRIVE_ACTAS_ROOT_ID);
    log.push('✅ Drive Actas raíz: "' + fa.getName() + '"');
  } catch(e) { log.push('❌ Drive Actas raíz: ' + e.message); }

  // Carpeta Drive firmas/logo
  try {
    var ff = DriveApp.getFolderById(CFG.DRIVE_FIRMAS_TEC_ID);
    log.push('✅ Drive Firmas/Logo: "' + ff.getName() + '"');
    // Verificar logo
    var logoIter = ff.getFilesByName('LOGO.png');
    log.push(logoIter.hasNext() ? '✅ LOGO.png: encontrado' : '⚠️  LOGO.png: NO encontrado en carpeta firmas');
    // Verificar firmas técnicos
    var faltantes = [];
    CFG.TECNICOS.forEach(function(t) {
      if (!t.firmaKey) return;
      var iter2 = ff.getFiles();
      var found = false;
      while (iter2.hasNext()) {
        if (iter2.next().getName().toUpperCase().replace(/\.[^.]+$/, '') === t.firmaKey) { found = true; break; }
      }
      if (!found) faltantes.push(t.firmaKey + '.png');
    });
    log.push(faltantes.length === 0
      ? '✅ Firmas técnicos: todas presentes (' + CFG.TECNICOS.length + ')'
      : '⚠️  Firmas faltantes: ' + faltantes.join(', '));
  } catch(e) { log.push('❌ Drive Firmas: ' + e.message); }

  // Carpeta Drive seguimiento (log)
  try {
    var fs = DriveApp.getFolderById(CFG.DRIVE_SEGUIMIENTO_FOLDER);
    log.push('✅ Drive Seguimiento folder: "' + fs.getName() + '"');
    // Verificar si el Sheet de log ya existe
    var ssId = PropertiesService.getScriptProperties().getProperty('SHEET_LOG_ID');
    if (ssId) {
      try {
        var ssLog = SpreadsheetApp.openById(ssId);
        log.push('✅ Sheet Log: "' + ssLog.getName() + '" (ID: ' + ssId + ')');
      } catch(eLog) {
        log.push('⚠️  Sheet Log ID guardado pero no accesible — ejecuta inicializarSheetLog()');
      }
    } else {
      var logFiles = fs.getFilesByName('SHARF_Log_Devoluciones_TI');
      if (logFiles.hasNext()) {
        log.push('✅ Sheet Log: encontrado en carpeta (ejecuta inicializarSheetLog() para registrar el ID)');
      } else {
        log.push('⚠️  Sheet Log: no existe aún — ejecuta inicializarSheetLog() para crearlo');
      }
    }
  } catch(e) { log.push('❌ Drive Seguimiento folder: ' + e.message + '\n   → Verifica que helpdesk@holasharf.com compartió la carpeta con tu cuenta'); }

  log.push('═══════════════════════════════════════');
  var msg = log.join('\n');
  Logger.log(msg);
}

// ════════════════════════════════════════════════════════════════════════
// SISTEMA DE BACKUP SEMANAL — SHARF Devoluciones TI
//
//   Se dispara automáticamente cada lunes a las 8:00 AM (Lima, GMT-5).
//   Genera un .zip con todo el proyecto y lo guarda en una carpeta
//   compartida con Eddie, Anais e Ismael.
//   También envía un correo de confirmación con el resumen.
//
//   INSTALACIÓN:
//     1. Ejecuta instalarTriggerBackup() UNA SOLA VEZ desde el editor
//        para registrar el trigger semanal.
//     2. Crea la carpeta de backups en Drive y compártela con los 3 correos,
//        luego ejecuta setBackupFolderId('ID_DE_LA_CARPETA') para guardar el ID.
//     3. Ejecuta ejecutarBackupManual() para probar sin esperar el lunes.
//
// ════════════════════════════════════════════════════════════════════════

var BACKUP_CFG = {
  EMAILS_BACKUP:  ['ismael.helpdesk@holasharf.com',
                   'anais.chero@holasharf.com',
                   'eddie.fernandez@holasharf.com'],
  FOLDER_PROP:    'SHARF_BACKUP_FOLDER_ID',   // ID guardado en PropertiesService
  MAX_BACKUPS:    8,                           // conservar solo los últimos 8 (2 meses)
  TZ:             'America/Lima'
};

// ── Instalar/desinstalar el trigger ──────────────────────────────────────
function instalarTriggerBackup() {
  // Eliminar triggers previos del mismo tipo para no duplicar
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'ejecutarBackupSemanal') {
      ScriptApp.deleteTrigger(t);
      Logger.log('Trigger anterior eliminado');
    }
  });

  // Crear trigger: cada lunes a las 8:00-9:00 AM Lima (UTC-5 → 13:00 UTC)
  ScriptApp.newTrigger('ejecutarBackupSemanal')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)        // 8 AM en la zona del servidor GAS (aproximado)
    .create();

  Logger.log('✅ Trigger de backup instalado: cada lunes ~8 AM');
  Logger.log('   Handler: ejecutarBackupSemanal');
  Logger.log('   Para probar ahora: ejecuta ejecutarBackupManual()');
}

function desinstalarTriggerBackup() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'ejecutarBackupSemanal') {
      ScriptApp.deleteTrigger(t);
      n++;
    }
  });
  Logger.log(n > 0 ? '✅ Trigger de backup eliminado' : '⚠️  No se encontró trigger de backup');
}

// Configurar la carpeta donde se guardan los backups
function setBackupFolderId(folderId) {
  PropertiesService.getScriptProperties().setProperty(BACKUP_CFG.FOLDER_PROP, folderId);
  Logger.log('✅ Backup folder ID guardado: ' + folderId);
  try {
    var f = DriveApp.getFolderById(folderId);
    Logger.log('   Carpeta: "' + f.getName() + '"');
  } catch(e) { Logger.log('⚠️  No se pudo acceder a la carpeta: ' + e.message); }
}

// ── Handler del trigger semanal ──────────────────────────────────────────
function ejecutarBackupSemanal() {
  Logger.log('═══ BACKUP SEMANAL SHARF — ' +
    Utilities.formatDate(new Date(), BACKUP_CFG.TZ, 'dd/MM/yyyy HH:mm') + ' ═══');
  _ejecutarBackup(false);
}

// Ejecutar manualmente para pruebas
function ejecutarBackupManual() {
  Logger.log('═══ BACKUP MANUAL SHARF — ' +
    Utilities.formatDate(new Date(), BACKUP_CFG.TZ, 'dd/MM/yyyy HH:mm') + ' ═══');
  _ejecutarBackup(true);
}

// ── Lógica principal del backup ──────────────────────────────────────────
function _ejecutarBackup(esManual) {
  var ts        = Utilities.formatDate(new Date(), BACKUP_CFG.TZ, 'yyyyMMdd_HHmm');
  var tsLegible = Utilities.formatDate(new Date(), BACKUP_CFG.TZ, 'dd/MM/yyyy HH:mm');
  var log       = [];
  var errores   = [];
  var archivos  = [];  // { nombre, blob }

  function ok(m)  { log.push('✅ ' + m); Logger.log('✅ ' + m); }
  function err(m) { log.push('❌ ' + m); errores.push(m); Logger.log('❌ ' + m); }
  function info(m){ log.push('   ' + m); Logger.log('   ' + m); }

  Logger.log('Iniciando backup' + (esManual ? ' MANUAL' : ' semanal') + ' — ' + tsLegible);

  // ── 1. Exportar archivos HTML desde Drive ──────────────────────────────
  // La Apps Script API requiere scope adicional (script.projects.readonly)
  // que no siempre está disponible sin habilitarlo manualmente.
  // Alternativa robusta: leer los archivos HTML desde la carpeta del proyecto
  // en Drive, donde GAS los almacena como archivos de texto accesibles.
  //
  // NOTA IMPORTANTE sobre el .gs:
  // El código backend (Código.gs) NO puede exportarse automáticamente sin
  // habilitar la Apps Script API en Google Cloud Console.
  // Instrucción para habilitarla: ve a script.google.com → Tu proyecto →
  // Configuración del proyecto → Google Cloud Platform → habilita
  // "Apps Script API" en la biblioteca de APIs.
  // Mientras tanto, el backup incluye todos los HTML y el Sheet de log.

  var scriptId = ScriptApp.getScriptId();
  var intentoApi = false;

  // Intentar con la API primero (funciona si el scope está habilitado)
  try {
    var token   = ScriptApp.getOAuthToken();
    var apiUrl  = 'https://script.googleapis.com/v1/projects/' + scriptId + '/content';
    var apiResp = UrlFetchApp.fetch(apiUrl, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (apiResp.getResponseCode() === 200) {
      var content = JSON.parse(apiResp.getContentText());
      var files   = content.files || [];
      files.forEach(function(f) {
        var ext  = (f.type === 'HTML') ? '.html' : '.gs';
        var nom  = (f.name || 'archivo') + ext;
        var blob = Utilities.newBlob(f.source || '', 'text/plain', nom);
        archivos.push({ nombre: nom, blob: blob });
        ok('Código (API): ' + nom + ' (' + (f.source||'').length + ' chars)');
      });
      intentoApi = true;

    } else if (apiResp.getResponseCode() === 403) {
      // 403 = scope insuficiente → fallback a Drive
      info('Apps Script API no disponible (scope 403) → usando Drive como fuente');
    } else {
      info('Apps Script API: HTTP ' + apiResp.getResponseCode() + ' → usando Drive');
    }
  } catch(eApi) {
    info('Apps Script API no accesible: ' + eApi.message + ' → usando Drive');
  }

  // Fallback: buscar los archivos HTML en Drive (carpeta del script)
  if (!intentoApi) {
    try {
      // Los archivos HTML del proyecto están en el mismo Drive del propietario
      // con extensión .html y accesibles via DriveApp
      var nombresBuscar = ['Index', 'Mobile', 'Index.html', 'Mobile.html'];
      var carpetaScript = null;

      // Intentar encontrar los archivos por nombre exacto
      var htmlEncontrados = 0;
      ['Index', 'Mobile'].forEach(function(nom) {
        var iter = DriveApp.getFilesByName(nom + '.html');
        if (!iter.hasNext()) iter = DriveApp.getFilesByName(nom);
        if (iter.hasNext()) {
          var f = iter.next();
          try {
            var blob = f.getBlob().setName(nom + '.html');
            archivos.push({ nombre: nom + '.html', blob: blob });
            ok('HTML Drive: ' + nom + '.html (' + f.getSize() + ' bytes)');
            htmlEncontrados++;
          } catch(ef) { info('No se pudo leer ' + nom + ': ' + ef.message); }
        }
      });

      if (htmlEncontrados === 0) {
        info('No se encontraron archivos HTML en Drive con nombre Index/Mobile');
        info('Para incluir el código: habilita Apps Script API en Google Cloud Console');
      }

      // Incluir un aviso sobre cómo obtener el .gs manualmente
      var avisoGs = [
        'AVISO — CÓDIGO BACKEND (.gs)',
        '══════════════════════════════════════════════',
        '',
        'El archivo Código.gs no pudo exportarse automáticamente.',
        'Razón: Se requiere habilitar la Apps Script API.',
        '',
        'PARA HABILITARLA (una sola vez):',
        '  1. Ve a script.google.com',
        '  2. Abre el proyecto SHARF Devoluciones TI',
        '  3. Configuración del proyecto (⚙️)',
        '  4. En "Proyecto de Google Cloud" → Ir a Google Cloud Console',
        '  5. Busca "Apps Script API" en la biblioteca',
        '  6. Haz clic en Habilitar',
        '  7. Vuelve a ejecutar ejecutarBackupManual()',
        '',
        'ALTERNATIVA INMEDIATA — Copia manual del código:',
        '  1. En el editor GAS, abre cada archivo',
        '  2. Ctrl+A → Ctrl+C',
        '  3. Pega en un archivo .txt y guárdalo',
        '',
        'Script ID: ' + scriptId,
        'URL directa: https://script.google.com/d/' + scriptId + '/edit',
        '',
        'Este aviso se eliminará automáticamente una vez que',
        'la Apps Script API esté habilitada.',
      ].join('\n');

      var avisoBlob = Utilities.newBlob(avisoGs, 'text/plain', 'AVISO_CodigoGs_LeerPrimero.txt');
      archivos.push({ nombre: 'AVISO_CodigoGs_LeerPrimero.txt', blob: avisoBlob });
      info('Aviso sobre .gs incluido en el ZIP');

    } catch(eDrive) {
      err('Fallback Drive: ' + eDrive.message);
    }
  }

  // ── 2. Exportar Sheet de log (XLSX) ───────────────────────────────────
  var sheetLogNombre = 'NO_DISPONIBLE';
  try {
    var ssLogId = PropertiesService.getScriptProperties().getProperty('SHEET_LOG_ID');
    if (ssLogId) {
      var ssLog    = SpreadsheetApp.openById(ssLogId);
      sheetLogNombre = ssLog.getName();
      var gidLog   = ssLog.getSheetByName(CFG.SHEET_LOG_NAME);
      var gid      = gidLog ? gidLog.getSheetId() : 0;
      var xlsxUrl  = 'https://docs.google.com/spreadsheets/d/' + ssLogId +
                     '/export?format=xlsx&gid=' + gid;
      var xlsxResp = UrlFetchApp.fetch(xlsxUrl, {
        headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      if (xlsxResp.getResponseCode() === 200) {
        var bLog = xlsxResp.getBlob().setName('SHARF_Log_Devoluciones_TI.xlsx');
        archivos.push({ nombre: 'SHARF_Log_Devoluciones_TI.xlsx', blob: bLog });
        ok('Sheet Log exportado: ' + sheetLogNombre + ' (' + xlsxResp.getBlob().getBytes().length + ' bytes)');
      } else {
        err('Sheet Log export XLSX: HTTP ' + xlsxResp.getResponseCode());
      }
    } else {
      err('Sheet Log ID no configurado — ejecuta inicializarSheetLog()');
    }
  } catch(e) { err('Sheet Log: ' + e.message); }

  // ── 3. Crear README de configuración y restauración ───────────────────
  var readme = [
    'SHARF DEVOLUCIONES TI — BACKUP DE RESTAURACIÓN',
    'Generado: ' + tsLegible + (esManual ? ' (MANUAL)' : ' (AUTOMÁTICO — lunes 8 AM)'),
    '═══════════════════════════════════════════════════════════',
    '',
    'CONTENIDO DE ESTE BACKUP:',
    '  • Código.gs         — Backend Google Apps Script',
    '  • Index.html        — Frontend versión PC',
    '  • Mobile.html       — Frontend versión Móvil',
    '  • SHARF_Log_Devoluciones_TI.xlsx — Historial completo de devoluciones',
    '  • README_Restauracion.txt — Este archivo',
    '',
    'CONFIGURACIÓN ACTUAL DEL SISTEMA:',
    '  Script ID:           ' + ScriptApp.getScriptId(),
    '  Sheet RH ID:         ' + CFG.SHEET_RH_ID,
    '  Drive Actas Root:    ' + CFG.DRIVE_ACTAS_ROOT_ID,
    '  Drive Firmas Tec:    ' + CFG.DRIVE_FIRMAS_TEC_ID,
    '  Drive Seguimiento:   ' + CFG.DRIVE_SEGUIMIENTO_FOLDER,
    '  Sheet Log ID:        ' + (PropertiesService.getScriptProperties().getProperty('SHEET_LOG_ID') || 'NO CONFIGURADO'),
    '  Snipe-IT Base:       ' + CFG.SNIPE_BASE,
    '',
    'PASOS DE RESTAURACIÓN:',
    '  1. Crear nuevo proyecto en script.google.com',
    '  2. Copiar el contenido de Código.gs al archivo Código.gs del proyecto',
    '  3. Crear archivo HTML llamado "Index" y pegar el contenido de Index.html',
    '  4. Crear archivo HTML llamado "Mobile" y pegar el contenido de Mobile.html',
    '  5. En el editor, ejecutar: setKeyManual() para reconfigurar la API key de Snipe-IT',
    '  6. Ejecutar: inicializarSheetLog() para recrear el Sheet de log',
    '  7. Si el Sheet de log se puede restaurar, importar SHARF_Log_Devoluciones_TI.xlsx',
    '  8. Ejecutar: instalarTriggerBackup() para reactivar el backup semanal',
    '  9. Reimplementar la Web App: Implementar > Nueva implementación',
    ' 10. Actualizar las URLs del aplicativo en los dispositivos de los técnicos',
    '',
    'PROPERTIES SERVICE (requieren reconfiguración manual):',
    '  SNIPE_API_KEY        — API key de Snipe-IT (clave secreta, no incluida en backup)',
    '  SHEET_LOG_ID         — Se recrea con inicializarSheetLog()',
    '  SHARF_BACKUP_FOLDER_ID — ID de esta carpeta de backups',
    '',
    'USUARIOS DEL SISTEMA:',
    '  Técnicos: gabriel.helpdesk, michael.helpdesk, misael.helpdesk,',
    '            ismael.helpdesk, perla.helpdesk, alfredo.helpdesk,',
    '            doris.quispe, jesus.alvarez',
    '  Auditoría: ismael.helpdesk, anais.chero, eddie.fernandez',
    '  Dominio: @holasharf.com',
    '',
    '═══════════════════════════════════════════════════════════',
    'SHARF · Sistema de Gestión de Activos TI · Mesa de Servicio'
  ].join('\n');

  var readmeBlob = Utilities.newBlob(readme, 'text/plain', 'README_Restauracion.txt');
  archivos.push({ nombre: 'README_Restauracion.txt', blob: readmeBlob });
  ok('README de restauración generado');

  // ── 4. Empaquetar todo en un ZIP ──────────────────────────────────────
  var zipNombre = 'SHARF_Backup_' + ts + (esManual ? '_MANUAL' : '') + '.zip';
  var zipBlob   = null;
  try {
    var blobs = archivos.map(function(a) { return a.blob; });
    zipBlob   = Utilities.zip(blobs, zipNombre);
    ok('ZIP generado: ' + zipNombre + ' (' + zipBlob.getBytes().length + ' bytes — ' +
       archivos.length + ' archivos)');
  } catch(e) { err('Crear ZIP: ' + e.message); }

  if (!zipBlob) {
    err('No se pudo generar el ZIP — backup abortado');
    _enviarNotificacionBackup(false, tsLegible, log, errores, zipNombre, esManual);
    return;
  }

  // ── 5. Guardar ZIP en la carpeta de backups en Drive ──────────────────
  var driveUrl = '';
  try {
    var folderId = PropertiesService.getScriptProperties().getProperty(BACKUP_CFG.FOLDER_PROP);
    if (!folderId) {
      // Intentar buscar/crear automáticamente en la carpeta de seguimiento
      try {
        var parentFolder = DriveApp.getFolderById(CFG.DRIVE_SEGUIMIENTO_FOLDER);
        var backupFolderIter = parentFolder.getFoldersByName('SHARF_Backups');
        var backupFolder;
        if (backupFolderIter.hasNext()) {
          backupFolder = backupFolderIter.next();
          info('Carpeta backups encontrada automáticamente');
        } else {
          backupFolder = parentFolder.createFolder('SHARF_Backups');
          info('Carpeta SHARF_Backups creada en Drive seguimiento');
          // Compartir con los 3 usuarios
          BACKUP_CFG.EMAILS_BACKUP.forEach(function(email) {
            try {
              backupFolder.addViewer(email);
              info('Compartida con: ' + email);
            } catch(es) { info('No se pudo compartir con ' + email + ': ' + es.message); }
          });
        }
        folderId = backupFolder.getId();
        PropertiesService.getScriptProperties().setProperty(BACKUP_CFG.FOLDER_PROP, folderId);
        info('Backup folder ID guardado automáticamente: ' + folderId);
      } catch(eAuto) {
        err('No se pudo crear carpeta de backups automáticamente: ' + eAuto.message);
        err('Ejecuta: setBackupFolderId("ID_CARPETA") para configurar manualmente');
      }
    }

    if (folderId) {
      var folder   = DriveApp.getFolderById(folderId);
      var savedZip = folder.createFile(zipBlob);
      driveUrl     = savedZip.getUrl();
      ok('ZIP guardado en Drive: ' + folder.getName());
      info('URL: ' + driveUrl);

      // Limpiar backups antiguos (conservar solo los últimos MAX_BACKUPS)
      var todosBackups = [];
      var iter = folder.getFilesByType('application/zip');
      while (iter.hasNext()) {
        var f2 = iter.next();
        todosBackups.push({ id: f2.getId(), nombre: f2.getName(), fecha: f2.getDateCreated() });
      }
      todosBackups.sort(function(a,b){ return b.fecha - a.fecha; });
      if (todosBackups.length > BACKUP_CFG.MAX_BACKUPS) {
        var aEliminar = todosBackups.slice(BACKUP_CFG.MAX_BACKUPS);
        aEliminar.forEach(function(bk) {
          try {
            DriveApp.getFileById(bk.id).setTrashed(true);
            info('Backup antiguo eliminado: ' + bk.nombre);
          } catch(eDel) { info('No se pudo eliminar ' + bk.nombre + ': ' + eDel.message); }
        });
      }
    }
  } catch(e) { err('Guardar ZIP en Drive: ' + e.message); }

  // ── 6. Notificar por correo ───────────────────────────────────────────
  _enviarNotificacionBackup(errores.length === 0, tsLegible, log, errores, zipNombre, esManual, driveUrl, zipBlob);
}

// ── Correo de notificación ────────────────────────────────────────────────
function _enviarNotificacionBackup(exito, tsLegible, logLines, errores, zipNombre, esManual, driveUrl, zipBlob) {
  var FONT   = "'Segoe UI',Arial,sans-serif";
  var C_N    = '#68002B';
  var C_S    = '#FF6568';
  var C_P    = '#FFC1C2';
  var estado = exito ? '✅ EXITOSO' : '⚠️ CON ERRORES';
  var asunto = 'SHARF · Backup ' + (esManual ? 'Manual' : 'Semanal') + ' ' +
               estado + ' — ' + tsLegible;

  var logHtml = logLines.map(function(l) {
    var color = l.indexOf('✅') >= 0 ? '#16a34a' :
                l.indexOf('❌') >= 0 ? '#dc2626' : '#5c3040';
    return '<div style="font-size:11px;font-family:monospace;color:'+color+';' +
           'padding:2px 0;line-height:1.5">' + l.replace(/</g,'&lt;') + '</div>';
  }).join('');

  var htmlBody =
    '<div style="font-family:'+FONT+';max-width:600px;margin:0 auto;background:#f8f0f1;padding:12px">' +

    // Header
    '<div style="background:'+C_N+';border-radius:10px 10px 0 0;padding:16px 24px;' +
      'border-bottom:3px solid '+C_S+'">' +
      '<div style="color:#fff;font-size:20px;font-weight:900;font-style:italic;' +
        'letter-spacing:-1px">sharf <span style="color:'+C_P+';font-size:12px;' +
        'font-style:normal;font-weight:600;letter-spacing:1px">BACKUP TI</span></div>' +
      '<div style="color:'+C_P+';font-size:11px;margin-top:2px">Devoluciones TI · Mesa de Servicio</div>' +
    '</div>' +

    '<div style="background:#fff;border:1px solid #f0d6d8;border-top:none;' +
      'border-radius:0 0 10px 10px;padding:24px">' +

      // Estado
      '<div style="display:flex;align-items:center;gap:12px;margin-bottom:20px;' +
        'background:'+(exito?'#f0fdf4':'#fff7ed')+';border:1.5px solid '+
        (exito?'#86efac':'#fed7aa')+';border-radius:10px;padding:14px 16px">' +
        '<div style="font-size:28px">'+(exito?'✅':'⚠️')+'</div>' +
        '<div>' +
          '<div style="font-size:15px;font-weight:800;color:'+(exito?'#15803d':'#92400e')+'">'+
            'Backup ' + (esManual ? 'manual' : 'semanal') + ' — ' + estado + '</div>' +
          '<div style="font-size:12px;color:'+(exito?'#166534':'#78350f')+';margin-top:2px">'+
            tsLegible + (esManual ? ' · Ejecutado manualmente' : ' · Ejecución automática (lunes 8 AM)') +
          '</div>' +
        '</div>' +
      '</div>' +

      // Archivo generado
      '<h4 style="font-size:12px;color:'+C_N+';border-bottom:2px solid '+C_S+';' +
        'padding-bottom:4px;margin:0 0 10px;font-weight:800">Archivo generado</h4>' +
      '<div style="background:#f8f0f1;border-radius:8px;padding:10px 14px;' +
        'font-family:monospace;font-size:12px;color:#1a0a0e;margin-bottom:16px">' +
        '📦 ' + zipNombre +
      '</div>' +

      // Enlace a Drive
      (driveUrl ?
        '<div style="margin-bottom:16px">' +
          '<a href="'+driveUrl+'" target="_blank" ' +
            'style="display:inline-flex;align-items:center;gap:8px;background:'+C_N+';' +
            'color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;' +
            'font-size:13px;font-weight:700">📁 Ver backup en Drive</a>' +
        '</div>' : '') +

      // Log detallado
      '<h4 style="font-size:12px;color:'+C_N+';border-bottom:2px solid '+C_S+';' +
        'padding-bottom:4px;margin:0 0 10px;font-weight:800">Log del proceso</h4>' +
      '<div style="background:#f8f0f1;border-radius:8px;padding:12px 14px;' +
        'max-height:200px;overflow-y:auto">' + logHtml + '</div>' +

      // Errores si los hay
      (errores.length > 0 ?
        '<div style="margin-top:14px;background:#fee2e2;border:1px solid #fecaca;' +
          'border-radius:8px;padding:12px 14px">' +
          '<div style="font-size:12px;font-weight:700;color:#991b1b;margin-bottom:8px">' +
            '❌ ' + errores.length + ' error(es) — requiere atención:</div>' +
          errores.map(function(e){
            return '<div style="font-size:11px;color:#7f1d1d;font-family:monospace;' +
                   'padding:2px 0">• '+e.replace(/</g,'&lt;')+'</div>';
          }).join('') +
        '</div>' : '') +

      // Footer
      '<div style="margin-top:20px;padding-top:12px;border-top:1px solid #f0d6d8;' +
        'font-size:10px;color:#9c7080;text-align:center">' +
        'SHARF · Backup automático semanal — cada lunes 8 AM · ' +
        'Conserva los últimos ' + BACKUP_CFG.MAX_BACKUPS + ' backups' +
      '</div>' +
    '</div></div>';

  var opciones = {
    htmlBody: htmlBody,
    name:     'SHARF - Backup TI'
  };
  // Adjuntar el ZIP si no pesa más de 20MB (límite Gmail)
  if (zipBlob && zipBlob.getBytes().length < 20 * 1024 * 1024) {
    opciones.attachments = [zipBlob];
  } else if (zipBlob) {
    // Si es muy grande, solo incluir el enlace al Drive
    opciones.htmlBody = opciones.htmlBody.replace(
      'Log del proceso',
      'Nota: El ZIP supera 20 MB y no se adjuntó. Descárgalo desde el enlace de Drive.<br><br>Log del proceso'
    );
  }

  BACKUP_CFG.EMAILS_BACKUP.forEach(function(email) {
    try {
      GmailApp.sendEmail(email, asunto, '', opciones);
      Logger.log('Notificación enviada a: ' + email);
    } catch(e) {
      Logger.log('Error enviando notificación a ' + email + ': ' + e.message);
    }
  });
}


// ════════════════════════════════════════════════════════════════════════
// CONFIGURAR PERMISOS — Ejecutar UNA SOLA VEZ como propietario
// ════════════════════════════════════════════════════════════════════════
// Da acceso Editor a todas las carpetas Drive y Sheets del sistema
// a todos los técnicos + Anais + Eddie + Ismael con @holasharf.com
// ════════════════════════════════════════════════════════════════════════
function configurarPermisosSistema() {
  var log = [];
  function ok(m)  { log.push('✅ ' + m); Logger.log('✅ ' + m); }
  function err(m) { log.push('❌ ' + m); Logger.log('❌ ' + m); }
  function info(m){ log.push('   ' + m); Logger.log('   ' + m); }

  // ── Lista completa de usuarios con acceso al sistema ─────────────────
  var USUARIOS = [
    'gabriel.helpdesk@holasharf.com',
    'michael.helpdesk@holasharf.com',
    'misael.helpdesk@holasharf.com',
    'ismael.helpdesk@holasharf.com',
    'perla.helpdesk@holasharf.com',
    'alfredo.helpdesk@holasharf.com',
    'rocio.helpdesk@holasharf.com',
    'jesus.helpdesk@holasharf.com',
    'anais.chero@holasharf.com',
    'eddie.fernandez@holasharf.com'
  ];

  Logger.log('═══ CONFIGURAR PERMISOS SHARF ═══');
  Logger.log('Usuarios: ' + USUARIOS.join(', '));

  // ── 1. Carpetas Drive ─────────────────────────────────────────────────
  var carpetas = [
    { id: CFG.DRIVE_ACTAS_ROOT_ID,      nombre: 'Actas Raíz (Devoluciones)' },
    { id: CFG.DRIVE_FIRMAS_TEC_ID,      nombre: 'Firmas Técnicos + Logo' },
    { id: CFG.DRIVE_SEGUIMIENTO_FOLDER, nombre: 'Seguimiento / Log' }
  ];

  // Agregar carpeta de backups si existe
  var backupFolderId = PropertiesService.getScriptProperties().getProperty('SHARF_BACKUP_FOLDER_ID');
  if (backupFolderId) {
    carpetas.push({ id: backupFolderId, nombre: 'Backups semanales' });
  }

  carpetas.forEach(function(c) {
    if (!c.id) { err('Carpeta sin ID: ' + c.nombre); return; }
    try {
      var folder = DriveApp.getFolderById(c.id);
      USUARIOS.forEach(function(email) {
        try {
          folder.addEditor(email);
          ok('Carpeta "' + c.nombre + '" → Editor: ' + email);
        } catch(eU) {
          // addEditor falla si ya tiene acceso → ignorar
          if (eU.message && eU.message.indexOf('already') >= 0) {
            info('Ya tenía acceso: ' + email + ' en ' + c.nombre);
          } else {
            err('Carpeta "' + c.nombre + '" → ' + email + ': ' + eU.message);
          }
        }
      });
    } catch(eF) {
      err('No se pudo abrir carpeta "' + c.nombre + '" (' + c.id + '): ' + eF.message);
    }
  });

  // ── 2. Sheet de Log ───────────────────────────────────────────────────
  var ssLogId = PropertiesService.getScriptProperties().getProperty('SHEET_LOG_ID');
  if (ssLogId) {
    try {
      var ssLog = SpreadsheetApp.openById(ssLogId);
      USUARIOS.forEach(function(email) {
        try {
          ssLog.addEditor(email);
          ok('Sheet Log → Editor: ' + email);
        } catch(eU2) {
          if (eU2.message && eU2.message.indexOf('already') >= 0) {
            info('Ya tenía acceso: ' + email + ' en Sheet Log');
          } else {
            err('Sheet Log → ' + email + ': ' + eU2.message);
          }
        }
      });
    } catch(eSS) {
      err('No se pudo abrir Sheet Log (' + ssLogId + '): ' + eSS.message);
    }
  } else {
    err('Sheet Log no inicializado — ejecuta inicializarSheetLog() primero');
  }

  // ── 3. Sheet RH (solo lectura para los técnicos) ──────────────────────
  try {
    var ssRH = SpreadsheetApp.openById(CFG.SHEET_RH_ID);
    USUARIOS.forEach(function(email) {
      try {
        ssRH.addViewer(email);
        ok('Sheet RH → Viewer: ' + email);
      } catch(eU3) {
        if (eU3.message && eU3.message.indexOf('already') >= 0) {
          info('Ya tenía acceso: ' + email + ' en Sheet RH');
        } else {
          err('Sheet RH → ' + email + ': ' + eU3.message);
        }
      }
    });
  } catch(eSS2) {
    err('No se pudo abrir Sheet RH (' + CFG.SHEET_RH_ID + '): ' + eSS2.message);
  }

  // ── 4. Verificar configuración de la Web App ──────────────────────────
  info('');
  info('══════════════════════════════════════════════════');
  info('VERIFICAR en el editor GAS:');
  info('  Implementar > Administrar implementaciones');
  info('  Ejecutar como: MÍ (helpdesk@holasharf.com)');
  info('  Quién puede acceder: Cualquier usuario de dominio holasharf.com');
  info('  → Haz una nueva implementación después de estos cambios');
  info('══════════════════════════════════════════════════');

  // ── Resumen final ─────────────────────────────────────────────────────
  var errores   = log.filter(function(l){ return l.indexOf('❌') >= 0; }).length;
  var correctos = log.filter(function(l){ return l.indexOf('✅') >= 0; }).length;
  Logger.log('');
  Logger.log('══ RESULTADO: ' + correctos + ' OK · ' + errores + ' errores ══');
  Logger.log(log.join('\n'));

  // Correo resumen a Ismael + Anais + Eddie
  try {
    var html = '<h3>SHARF — Configuración de permisos</h3>' +
      '<p><b>' + correctos + ' accesos configurados · ' + errores + ' errores</b></p>' +
      '<pre style="font-size:11px;background:#f5f5f5;padding:10px">' +
      log.join('\n').replace(/</g,'&lt;') + '</pre>';
    GmailApp.sendEmail(
      'ismael.helpdesk@holasharf.com',
      'SHARF · Permisos configurados — ' + correctos + ' OK · ' + errores + ' errores',
      log.join('\n'),
      { htmlBody: html, cc: 'anais.chero@holasharf.com,eddie.fernandez@holasharf.com',
        name: 'SHARF Sistema' }
    );
    Logger.log('Correo de resumen enviado');
  } catch(eMail) {
    Logger.log('No se pudo enviar correo resumen: ' + eMail.message);
  }
}


// ════════════════════════════════════════════════════════════════════════
// PASO 1 — Ejecutar PRIMERO: verificar que el script corre como propietario
// ════════════════════════════════════════════════════════════════════════
function paso1_verificarPropietario() {
  var activo   = Session.getActiveUser().getEmail();
  var efectivo = Session.getEffectiveUser().getEmail();
  Logger.log('══════════════════════════════════════════');
  Logger.log('PASO 1 — VERIFICAR PROPIETARIO DEL SCRIPT');
  Logger.log('══════════════════════════════════════════');
  Logger.log('Usuario activo    (quien ejecuta ahora): ' + activo);
  Logger.log('Usuario efectivo  (como corre el script): ' + efectivo);
  Logger.log('');
  if (efectivo === activo) {
    Logger.log('⚠️  PROBLEMA DETECTADO: effectiveUser === activeUser');
    Logger.log('   El script está configurado como "Ejecutar como: Usuario que accede".');
    Logger.log('   SOLUCIÓN: ve a Implementar > Administrar implementaciones > Editar');
    Logger.log('   Cambia "Ejecutar como" a: Yo (' + activo + ')');
    Logger.log('   Luego haz clic en Implementar (nueva versión).');
  } else {
    Logger.log('✅ CORRECTO: el script corre como: ' + efectivo);
    Logger.log('   Los técnicos usan los permisos de esta cuenta.');
  }
  Logger.log('');
  Logger.log('Si effectiveUser aparece vacío también indica problema.');
}

// ════════════════════════════════════════════════════════════════════════
// PASO 2 — Ejecutar SEGUNDO: dar permisos de Drive/Sheets a todos los usuarios
//          (ya existe como configurarPermisosSistema, este es un alias claro)
// ════════════════════════════════════════════════════════════════════════
function paso2_configurarPermisos() {
  configurarPermisosSistema();
}

// ════════════════════════════════════════════════════════════════════════
// PASO 3 — Ejecutar TERCERO: verificar que todo está listo
// ════════════════════════════════════════════════════════════════════════
function paso3_verificarTodo() {
  var log = [];
  function ok(m)  { log.push('✅ ' + m); Logger.log('✅ ' + m); }
  function err(m) { log.push('❌ ' + m); Logger.log('❌ ' + m); }

  Logger.log('══════════════════════════════════════');
  Logger.log('PASO 3 — VERIFICACIÓN COMPLETA SHARF');
  Logger.log('══════════════════════════════════════');

  // 1. Propietario
  var efectivo = Session.getEffectiveUser().getEmail();
  var activo   = Session.getActiveUser().getEmail();
  if (efectivo && efectivo !== activo) {
    ok('Script corre como propietario: ' + efectivo);
  } else {
    err('Script NO corre como propietario. Reconfigura la implementación.');
  }

  // 2. Snipe-IT
  try {
    var snipeKey = PropertiesService.getScriptProperties().getProperty(CFG.SNIPE_KEY_PROP);
    if (snipeKey && snipeKey.length > 10) {
      ok('Snipe-IT API Key configurada (' + snipeKey.length + ' chars)');
      var testSnipe = snipeGET('/hardware?limit=1');
      ok('Snipe-IT responde OK — total activos: ' + (testSnipe.total || '?'));
    } else {
      err('Snipe-IT API Key no configurada. Ejecuta setKeyManual()');
    }
  } catch(eS) { err('Snipe-IT: ' + eS.message); }

  // 3. Drive — carpetas
  var carpetasTest = [
    { id: CFG.DRIVE_ACTAS_ROOT_ID,    nombre: 'Actas' },
    { id: CFG.DRIVE_FIRMAS_TEC_ID,    nombre: 'Firmas' },
    { id: CFG.DRIVE_SEGUIMIENTO_FOLDER, nombre: 'Seguimiento' }
  ];
  carpetasTest.forEach(function(c) {
    try {
      var f = DriveApp.getFolderById(c.id);
      ok('Drive carpeta "' + c.nombre + '": ' + f.getName());
    } catch(eF) { err('Drive carpeta "' + c.nombre + '": ' + eF.message); }
  });

  // 4. Sheet Log
  var ssLogId = PropertiesService.getScriptProperties().getProperty('SHEET_LOG_ID');
  if (ssLogId) {
    try {
      var ss = SpreadsheetApp.openById(ssLogId);
      ok('Sheet Log accesible: ' + ss.getName());
    } catch(eSS) { err('Sheet Log: ' + eSS.message); }
  } else {
    err('Sheet Log no inicializado. Ejecuta inicializarSheetLog()');
  }

  // 5. Sheet RH
  try {
    var ssRH = SpreadsheetApp.openById(CFG.SHEET_RH_ID);
    ok('Sheet RH accesible: ' + ssRH.getName());
  } catch(eRH) { err('Sheet RH: ' + eRH.message); }

  // 6. Gmail
  try {
    var draft = GmailApp.createDraft(
      'helpdesk@holasharf.com',
      '[TEST SHARF] Verificación de permisos Gmail',
      'Este borrador se creó automáticamente para verificar permisos. Puedes eliminarlo.'
    );
    ok('Gmail: puede crear correos (borrador de prueba creado)');
    draft.deleteDraft();
  } catch(eG) { err('Gmail: ' + eG.message); }

  // Resumen
  Logger.log('');
  Logger.log('══ RESUMEN ══');
  var errores = log.filter(function(l){ return l.indexOf('❌')>=0; }).length;
  var oks     = log.filter(function(l){ return l.indexOf('✅')>=0; }).length;
  Logger.log(oks + ' OK · ' + errores + ' errores');
  if (errores === 0) {
    Logger.log('');
    Logger.log('🎉 TODO LISTO — El sistema está configurado correctamente.');
    Logger.log('   Los técnicos pueden realizar devoluciones.');
  } else {
    Logger.log('');
    Logger.log('⚠️  Hay ' + errores + ' problema(s) que resolver antes de usar el sistema.');
  }
}
