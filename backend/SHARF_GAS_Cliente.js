/**
 * SHARF Devoluciones TI — Google Apps Script (CLIENTE LIGERO)
 * ============================================================
 * ANTES: GAS contenía toda la lógica (riesgo alto ⚠️)
 * AHORA: GAS solo es un puente que llama a la API en Railway 🔐
 *
 * Toda la lógica sensible está en Railway:
 *   https://sharf-ti.railway.app
 *
 * Este archivo contiene:
 *   ✅ doGet()       → sirve el frontend HTML
 *   ✅ handleRequest() → proxy hacia Railway API
 *   ✅ Sin tokens, sin credenciales, sin lógica de negocio
 */

// ══ CONFIGURACIÓN ════════════════════════════════════════════════════════
var RAILWAY_API_URL = PropertiesService.getScriptProperties()
                        .getProperty('RAILWAY_API_URL')
                      || 'https://sharf-ti.railway.app';

// Token JWT del técnico actual (se obtiene al hacer login)
// Se guarda en PropertiesService para persistir entre llamadas GAS
var _jwt_cache = null;

// ══ ENTRY POINT ══════════════════════════════════════════════════════════

function doGet(e) {
  var params = e ? e.parameter : {};

  // Servir la versión móvil
  if (params.v === 'mobile') {
    return HtmlService
      .createTemplateFromFile('SHARF_Mobile_v2')
      .evaluate()
      .setTitle('SHARF · Devoluciones TI')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Servir la versión PC
  return HtmlService
    .createTemplateFromFile('SHARF_Index_v5b')
    .evaluate()
    .setTitle('SHARF · Devoluciones TI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ══ PROXY HACIA RAILWAY ══════════════════════════════════════════════════
/**
 * Punto de entrada para llamadas del frontend.
 * Antes: ejecutaba la lógica aquí.
 * Ahora: reenvía a Railway y devuelve la respuesta.
 */
function handleRequest(accion, params) {
  try {
    // Obtener JWT (del cache o hacer login)
    var jwt = _obtenerJWT();

    // Mapear accion → endpoint Railway
    var endpoint = _mapearAccion(accion, params);
    if (!endpoint) {
      return { ok: false, error: 'Acción desconocida: ' + accion };
    }

    // Llamar a Railway
    var response = _llamarRailway(endpoint.method, endpoint.path,
                                   endpoint.body, jwt);
    return response;

  } catch (e) {
    Logger.log('handleRequest error (' + accion + '): ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ══ OBTENER JWT ══════════════════════════════════════════════════════════
function _obtenerJWT() {
  // Usar cache si existe y no expiró
  if (_jwt_cache && _jwt_cache.expires > Date.now()) {
    return _jwt_cache.token;
  }

  // Obtener email del técnico actual
  var email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error('Sin sesión Google activa');
  }

  // Login en Railway con el email
  var resp = _llamarRailway('POST', '/api/sesion', { email: email }, null);
  if (!resp || !resp.access_token) {
    throw new Error(resp ? (resp.error || 'Login fallido') : 'Sin respuesta de Railway');
  }

  // Guardar en cache (expira 5 min antes del JWT)
  _jwt_cache = {
    token:   resp.access_token,
    expires: Date.now() + (resp.expires_in - 300) * 1000,
  };

  return resp.access_token;
}

// ══ MAPEO DE ACCIONES ════════════════════════════════════════════════════
function _mapearAccion(accion, params) {
  var p = params || {};
  switch (accion) {
    case 'getSession':
      return { method:'POST', path:'/api/sesion', body:{ email: Session.getActiveUser().getEmail() } };
    case 'getFirmaBase64':
      return { method:'GET', path:'/api/firma/' + encodeURIComponent(p.firmaKey || ''), body:null };
    case 'getLogoBase64':
      return { method:'GET', path:'/api/logo', body:null };
    case 'buscarPorDNI':
      return { method:'GET', path:'/api/activos/' + encodeURIComponent(p.dni || ''), body:null };
    case 'parsearQR':
      return { method:'POST', path:'/api/qr/texto', body:{ texto: p.qr || '' } };
    case 'buscarActivoPorSerial':
      return { method:'GET', path:'/api/activo/serial/' + encodeURIComponent(p.serial || ''), body:null };
    case 'procesarDevolucion':
      return { method:'POST', path:'/api/devolucion', body:p };
    case 'validarConexiones':
      return { method:'GET', path:'/api/validar', body:null };
    case 'getAuditoriaData':
      return { method:'GET', path:'/api/auditoria?meses=' + (p.meses || 3), body:null };
    case 'getModoPrueba':
      return { method:'GET', path:'/api/modo-prueba', body:null };
    case 'setModoPrueba':
      return { method:'POST', path:'/api/modo-prueba', body:p };
    case 'reenviarCorreo':
      return { method:'POST', path:'/api/reenviar-correo', body:p };
    case 'getCatalogos':
    case 'getCatalogoSnipe':
      return { method:'GET', path:'/api/catalogos', body:null };
    case 'diagnosticarSheetRH':
      return { method:'POST', path:'/api/diagnostico', body:p };
    case 'buscarActaAsignacion':
      return { method:'GET', path:'/api/acta-asignacion/' +
               encodeURIComponent(p.serial || '') + '?nombre=' +
               encodeURIComponent(p.nombreUsuario || '') + '&dni=' +
               encodeURIComponent(p.dni || ''), body:null };
    default:
      Logger.log('Acción no mapeada: ' + accion);
      return null;
  }
}

// ══ HTTP CLIENT HACIA RAILWAY ════════════════════════════════════════════
function _llamarRailway(method, path, body, jwt) {
  var url     = RAILWAY_API_URL + path;
  var options = {
    method:             method.toLowerCase(),
    contentType:        'application/json',
    muteHttpExceptions: true,
    followRedirects:    true,
    headers:            {},
  };

  if (jwt) {
    options.headers['Authorization'] = 'Bearer ' + jwt;
  }

  if (body && method !== 'GET') {
    options.payload = JSON.stringify(body);
  }

  try {
    var resp     = UrlFetchApp.fetch(url, options);
    var code     = resp.getResponseCode();
    var text     = resp.getContentText();
    var parsed   = JSON.parse(text);

    Logger.log('[Railway] ' + method + ' ' + path + ' → ' + code);

    if (code === 401) {
      // Token expirado — limpiar cache y reintentar una vez
      _jwt_cache = null;
      Logger.log('JWT expirado — reintentando con nuevo token');
      var newJwt = _obtenerJWT();
      options.headers['Authorization'] = 'Bearer ' + newJwt;
      resp   = UrlFetchApp.fetch(url, options);
      text   = resp.getContentText();
      parsed = JSON.parse(text);
    }

    return parsed;
  } catch (e) {
    Logger.log('[Railway] ERROR: ' + e.message);
    return { ok: false, error: 'Error de red: ' + e.message };
  }
}

// ══ CONFIGURACIÓN INICIAL (ejecutar UNA VEZ como admin) ════════════════
/**
 * Ejecuta esta función desde el editor GAS para configurar
 * la URL de Railway. Solo necesitas hacerlo una vez.
 */
function configurarRailwayURL() {
  var url = 'https://TU-APP.railway.app';  // ← cambia esto
  PropertiesService.getScriptProperties().setProperty('RAILWAY_API_URL', url);
  Logger.log('✅ Railway URL configurada: ' + url);
}

/**
 * Verifica la conexión con Railway.
 * Ejecuta desde el editor para testear.
 */
function testConexion() {
  var resp = _llamarRailway('GET', '/api/health', null, null);
  Logger.log('Health check: ' + JSON.stringify(resp));
}
