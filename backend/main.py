"""
SHARF Devoluciones TI — Backend API v2 (Arquitectura PRO)
Arquitectura: Frontend → GAS cliente → Railway API → Snipe/Drive/Sheets/Gmail
Seguridad: JWT + Google OAuth + Rate limiting + Logs + CORS estricto
"""
from fastapi import FastAPI,HTTPException,UploadFile,File,Depends,Request
from fastapi.responses import JSONResponse,FileResponse,HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from slowapi import Limiter,_rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded
import uvicorn,os,logging
from contextlib import asynccontextmanager
from typing import Optional,Dict
from datetime import datetime

from core.config     import CFG
from core.snipe      import SnipeClient
from core.sheets     import SheetsClient
from core.drive      import DriveClient
from core.mail       import MailClient
from core.pdf        import generar_acta_pdf
from core.qr         import decodificar_qr,_parsear_qr
from core.auth       import autenticar_tecnico,USUARIOS_AUTORIZADOS
from core.models     import PayloadDevolucion,SesionRequest
from core.security   import (crear_token,verificar_token,revocar_token,
                              get_current_user,get_supervisor,verificar_google_token)
from core.access_log import AccessLogMiddleware,get_recent_logs,get_stats

logging.basicConfig(level=getattr(logging,CFG.LOG_LEVEL,logging.INFO),
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
log=logging.getLogger("sharf")
limiter=Limiter(key_func=get_remote_address,default_limits=["60/minute"])

snipe:SnipeClient=None; sheets:SheetsClient=None
drive:DriveClient=None; mail=None
_MP:Dict={"activo":False,"emailPara":"","emailCC":"","emailBCC":""}

@asynccontextmanager
async def lifespan(app):
    global snipe,sheets,drive,mail
    log.info(f"▶ SHARF API v2 | ENV={CFG.ENV} | mail={CFG.mail_mode}")
    snipe =SnipeClient(CFG.SNIPE_BASE,CFG.SNIPE_TOKEN)
    sheets=SheetsClient(CFG.SHEET_LOG_ID,CFG.GOOGLE_CREDS_JSON)
    drive =DriveClient(CFG.DRIVE_ACTAS_ROOT,CFG.GOOGLE_CREDS_JSON)
    m=CFG.mail_mode
    if m=="smtp":
        from core.mail_hybrid import MailSMTPClient
        mail=MailSMTPClient(CFG.GMAIL_USER,CFG.GMAIL_APP_PASSWORD)
    elif m=="hybrid":
        from core.mail_hybrid import MailHybridClient
        mail=MailHybridClient(CFG.GAS_WEBHOOK_URL,CFG.GAS_WEBHOOK_TOKEN)
    elif m=="service_account":
        mail=MailClient(CFG.GMAIL_USER,CFG.GOOGLE_CREDS_JSON)
    else:
        log.warning("⚠️  Sin cliente de correo")
        mail=None
    log.info(f"✅ OK — correo:{m}")
    yield
    log.info("■ Apagando")

app=FastAPI(title="SHARF TI API v2",version="2.0.0",lifespan=lifespan,
            docs_url=None if CFG.is_production else "/docs",redoc_url=None)
app.state.limiter=limiter
app.add_exception_handler(RateLimitExceeded,_rate_limit_exceeded_handler)
app.add_middleware(CORSMiddleware,
    allow_origins=CFG.ALLOWED_ORIGINS if CFG.is_production else ["*"],
    allow_methods=["GET","POST","OPTIONS"],
    allow_headers=["Authorization","Content-Type","X-SHARF-Token","X-Admin-Token"],
    allow_credentials=True)
app.add_middleware(AccessLogMiddleware)

static_dir=os.path.join(os.path.dirname(__file__),"static")
if os.path.isdir(static_dir):
    app.mount("/static",StaticFiles(directory=static_dir),name="static")

def _check_mail():
    if not mail: raise HTTPException(503,detail="Correo no configurado")
def _verify_admin(request:Request):
    t=request.headers.get("X-Admin-Token","")
    if not CFG.ADMIN_TOKEN or t!=CFG.ADMIN_TOKEN:
        raise HTTPException(403,detail="Admin token requerido")

# ── 0. RAÍZ ──────────────────────────────────────────────────
@app.get("/",response_class=HTMLResponse,include_in_schema=False)
async def raiz():
    idx=os.path.join(static_dir,"index.html")
    if os.path.exists(idx): return FileResponse(idx)
    return HTMLResponse('<h1 style="font-family:Arial;color:#9B1035;padding:40px">SHARF API v2 ✅</h1>'
                        '<p><a href="/api/health">Health</a> · <a href="/docs">Docs</a></p>')

# ── 1. HEALTH ─────────────────────────────────────────────────
@app.get("/api/health",tags=["Sistema"])
async def health():
    return {"status":"ok","version":"2.0.0","env":CFG.ENV,
            "timestamp":datetime.now().isoformat(),"mail_mode":CFG.mail_mode}

# ── 2. AUTH ───────────────────────────────────────────────────
@app.post("/api/auth/google",tags=["Auth"])
@limiter.limit("10/minute")
async def login_google(request:Request,body:dict):
    """Recibe id_token de Google → verifica → devuelve JWT propio."""
    id_token=body.get("id_token","")
    if not id_token: raise HTTPException(400,detail="id_token requerido")
    gdata=await verificar_google_token(id_token)
    email=gdata["email"]
    tec=autenticar_tecnico(email)
    if not tec:
        log.warning(f"⛔ Denegado: {email}")
        raise HTTPException(403,detail=f"Acceso denegado: {email}")
    token_data=crear_token(email,tec["nombre"],tec.get("rol","tecnico"))
    token_data["tecnico"]=tec
    log.info(f"✅ Login: {email}")
    return token_data

@app.post("/api/auth/refresh",tags=["Auth"])
async def refresh_token(user:Dict=Depends(get_current_user)):
    tec=autenticar_tecnico(user["sub"])
    if not tec: raise HTTPException(403,detail="No autorizado")
    return crear_token(user["sub"],user["nombre"],user["rol"])

@app.post("/api/auth/logout",tags=["Auth"])
async def logout(request:Request,user:Dict=Depends(get_current_user)):
    auth=request.headers.get("authorization","")
    if auth.startswith("Bearer "): revocar_token(auth[7:])
    log.info(f"Logout: {user.get('sub','?')}")
    return {"ok":True,"message":"Sesión cerrada"}

@app.post("/api/sesion",tags=["Auth"])
@limiter.limit("20/minute")
async def sesion_simple(request:Request,req:SesionRequest):
    """Compatibilidad con GAS cliente — usa /api/auth/google en prod."""
    tec=autenticar_tecnico(req.email)
    if not tec: raise HTTPException(403,detail=f"Denegado: {req.email}")
    token_data=crear_token(req.email,tec["nombre"],tec.get("rol","tecnico"))
    return {"ok":True,"tecnico":tec,**token_data}

# ── 3. ACTIVOS ────────────────────────────────────────────────
@app.get("/api/activos/{dni}",tags=["Activos"])
async def buscar_activos(dni:str,user:Dict=Depends(get_current_user)):
    try:
        r=await snipe.buscar_por_custom_field("DNI",dni)
        return {"ok":True,"activos":r,"total":len(r)}
    except Exception as e: raise HTTPException(500,detail=str(e))

@app.get("/api/activo/{asset_id}",tags=["Activos"])
async def detalle_activo(asset_id:int,user:Dict=Depends(get_current_user)):
    try: return {"ok":True,"activo":await snipe.get_activo(asset_id)}
    except Exception as e: raise HTTPException(500,detail=str(e))

@app.get("/api/activo/serial/{serial}",tags=["Activos"])
async def activo_serial(serial:str,user:Dict=Depends(get_current_user)):
    try:
        d=await snipe.buscar_por_serial(serial)
        if not d: return {"ok":False,"error":f"Serial {serial} no encontrado"}
        return {"ok":True,"activo":d}
    except Exception as e: raise HTTPException(500,detail=str(e))

@app.get("/api/catalogos",tags=["Activos"])
async def catalogos(user:Dict=Depends(get_current_user)):
    try: return {"ok":True,"areas":await snipe.get_categorias(),"sedes":await snipe.get_locations()}
    except Exception as e: return {"ok":False,"areas":[],"sedes":[],"error":str(e)}

# ── 4. QR ─────────────────────────────────────────────────────
@app.post("/api/qr",tags=["QR"])
async def leer_qr(imagen:UploadFile=File(...),user:Dict=Depends(get_current_user)):
    try:
        r=decodificar_qr(await imagen.read())
        if not r: return {"ok":False,"error":"QR no detectado"}
        return {"ok":True,**r}
    except Exception as e: raise HTTPException(500,detail=str(e))

@app.post("/api/qr/texto",tags=["QR"])
async def parsear_qr(body:dict,user:Dict=Depends(get_current_user)):
    t=body.get("texto","")
    if not t: return {"ok":False,"error":"Texto vacío"}
    return {"ok":True,**_parsear_qr(t)}

# ── 5. DEVOLUCIÓN ─────────────────────────────────────────────
@app.post("/api/devolucion",tags=["Devolución"])
@limiter.limit("30/minute")
async def procesar_devolucion(request:Request,payload:PayloadDevolucion,
                               user:Dict=Depends(get_current_user)):
    dev_id=payload.devId or f"DEV-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    log.info(f"▶ {dev_id} | {user.get('sub','?')} | {payload.serial} | {payload.tipoDev}")
    pasos={};errors=[]

    try:
        await snipe.checkin(payload.snipeId,{"note":f"Dev {payload.tipoDev} {dev_id}","status_id":4,
              "location_id":payload.empleado.get("locationId")})
        pasos["snipe"]={"ok":True,"msg":"Activo → Disponible"}
    except Exception as e: pasos["snipe"]={"ok":False,"msg":str(e)};errors.append(f"Snipe:{e}")

    pdf_bytes=None
    try:
        pdf_bytes=generar_acta_pdf(payload,dev_id)
        pasos["pdf"]={"ok":True,"msg":f"PDF {len(pdf_bytes)//1024}KB"}
    except Exception as e: pasos["pdf"]={"ok":False,"msg":str(e)};errors.append(f"PDF:{e}")

    drive_url=None
    try:
        nombre=f"Acta_{dev_id}_{payload.empleado.get('nombre','X').split(',')[0]}.pdf"
        drive_url=await drive.subir_pdf(pdf_bytes,nombre,
            empleado=payload.empleado.get("nombre",""),tecnico=payload.tecnico.get("nombre",""))
        pasos["drive"]={"ok":True,"msg":"PDF en Drive","url":drive_url}
    except Exception as e: pasos["drive"]={"ok":False,"msg":str(e)};errors.append(f"Drive:{e}")

    try:
        await sheets.registrar_devolucion(payload,dev_id,drive_url,
            bool(pasos.get("snipe",{}).get("ok")),bool(pasos.get("pdf",{}).get("ok")))
        pasos["log"]={"ok":True,"msg":"Registrado"}
    except Exception as e: pasos["log"]={"ok":False,"msg":str(e)};errors.append(f"Sheets:{e}")

    if mail:
        try:
            if _MP.get("activo") and _MP.get("emailPara"):
                payload.modoPrueba=True; payload.emailPrueba=_MP["emailPara"]
            await mail.enviar_devolucion(payload,dev_id,pdf_bytes,drive_url)
            pasos["email"]={"ok":True,"msg":"Correo enviado"}
        except Exception as e: pasos["email"]={"ok":False,"msg":str(e)};errors.append(f"Mail:{e}")
    else: pasos["email"]={"ok":False,"msg":"Correo no configurado"}

    ok=all(p.get("ok") for p in pasos.values())
    log.info(f"■ {dev_id}: {'✅' if ok else '⚠️'} {errors or ''}")
    return {"ok":ok,"devId":dev_id,"driveUrl":drive_url,"pasos":pasos,"errors":errors,
            "snipeOk":pasos.get("snipe",{}).get("ok",False),
            "pdfOk":pasos.get("pdf",{}).get("ok",False),
            "emailOk":pasos.get("email",{}).get("ok",False),
            "logOk":pasos.get("log",{}).get("ok",False)}

# ── 6. VALIDAR CONEXIONES ─────────────────────────────────────
@app.get("/api/validar",tags=["Sistema"])
async def validar(user:Dict=Depends(get_current_user)):
    r={"timestamp":datetime.now().strftime("%d/%m/%Y %H:%M:%S"),"ok":True,"pasos":[],"resumen":[]}
    def ok(m): r["resumen"].append(f"✅ {m}");r["pasos"].append(f"✅ {m}")
    def err(m): r["resumen"].append(f"❌ {m}");r["pasos"].append(f"❌ {m}");r.__setitem__("ok",False)
    def warn(m): r["resumen"].append(f"⚠️ {m}");r["pasos"].append(f"⚠️ {m}")
    try: i=await snipe.status(); ok(f"Snipe-IT {i.get('total_assets',0)} activos")
    except Exception as e: err(f"Snipe-IT: {e}")
    try: f=await sheets.contar_filas(); ok(f"Sheets {f} registros")
    except Exception as e: err(f"Sheets: {e}")
    try: n=await drive.verificar_carpeta(); ok(f"Drive: {n}")
    except Exception as e: err(f"Drive: {e}")
    if mail:
        try: await mail.verificar(); ok("Gmail OK")
        except Exception as e: warn(f"Gmail: {e}")
    else: warn("Gmail no configurado")
    return r

# ── 7. AUDITORÍA ──────────────────────────────────────────────
@app.get("/api/auditoria",tags=["Auditoría"])
async def auditoria(meses:int=3,sup:Dict=Depends(get_supervisor)):
    try: return {"ok":True,**(await sheets.get_auditoria(meses))}
    except Exception as e: raise HTTPException(500,detail=str(e))

# ── 8. MODO PRUEBA ────────────────────────────────────────────
@app.get("/api/modo-prueba",tags=["Config"])
async def get_mp(user:Dict=Depends(get_current_user)):
    return {"ok":True,**_MP}

@app.post("/api/modo-prueba",tags=["Config"])
async def set_mp(body:dict,sup:Dict=Depends(get_supervisor)):
    global _MP
    _MP={"activo":bool(body.get("activo",False)),"emailPara":body.get("emailPara",""),
         "emailCC":body.get("emailCC",""),"emailBCC":body.get("emailBCC","")}
    log.info(f"Modo prueba: {'ON' if _MP['activo'] else 'OFF'} — {sup.get('sub','?')}")
    return {"ok":True,**_MP}

# ── 9. IMÁGENES ───────────────────────────────────────────────
@app.get("/api/firma/{firma_key}",tags=["Recursos"])
async def get_firma(firma_key:str,user:Dict=Depends(get_current_user)):
    try: return {"ok":True,"data":await drive.obtener_imagen_b64(firma_key)}
    except: return {"ok":False,"data":""}

@app.get("/api/logo",tags=["Recursos"])
async def get_logo(user:Dict=Depends(get_current_user)):
    try: return {"ok":True,"data":await drive.obtener_imagen_b64("LOGO.png")}
    except: return {"ok":False,"data":""}

# ── 10. REENVIAR CORREO ───────────────────────────────────────
@app.post("/api/reenviar-correo",tags=["Correo"])
async def reenviar(body:dict,user:Dict=Depends(get_current_user)):
    _check_mail()
    ep=body.get("emailPersonal","").strip(); ej=body.get("emailJefe","").strip()
    if not ep and not ej: raise HTTPException(400,detail="Sin correo")
    pd=body.get("payload")
    if not pd: raise HTTPException(400,detail="Sin payload")
    try:
        p=PayloadDevolucion(**pd); p.emailPara=ep or ej; p.emailCC=ej if ep else ""
        await mail.enviar_devolucion(p,pd.get("devId","REENV"),None,pd.get("driveUrl",""))
        return {"ok":True,"msg":f"Reenviado a {p.emailPara}"}
    except Exception as e: raise HTTPException(500,detail=str(e))

# ── 11. DIAGNÓSTICO ───────────────────────────────────────────
@app.post("/api/diagnostico",tags=["Sistema"])
async def diag(body:dict,user:Dict=Depends(get_current_user)):
    dni=body.get("dni","00000000"); pasos=[]
    async def paso(n,d,fn):
        try: det=await fn(); pasos.append({"num":n,"desc":d,"ok":True,"detalle":str(det or "OK")})
        except Exception as e: pasos.append({"num":n,"desc":d,"ok":False,"detalle":str(e)})
    await paso(1,"Snipe-IT",lambda:snipe.status())
    await paso(2,f"DNI {dni}",lambda:snipe.buscar_por_custom_field("DNI",dni))
    await paso(3,"Sheets",lambda:sheets.contar_filas())
    await paso(4,"Drive",lambda:drive.verificar_carpeta())
    await paso(5,"Gmail",lambda:mail.verificar() if mail else (_ for _ in()).throw(Exception("No conf.")))
    return {"ok":all(p["ok"] for p in pasos),"pasos":pasos}

# ── 12. ACTA ASIGNACIÓN ───────────────────────────────────────
@app.get("/api/acta-asignacion/{serial}",tags=["Recursos"])
async def acta(serial:str,nombre:str="",dni:str="",user:Dict=Depends(get_current_user)):
    try: return {"ok":True,**(await drive.buscar_acta_asignacion(nombre,dni,serial))}
    except Exception as e: return {"ok":False,"error":str(e),"archivos":[]}

# ── 13. ADMIN ─────────────────────────────────────────────────
@app.get("/api/admin/logs",tags=["Admin"])
async def admin_logs(request:Request,n:int=100,user:str="",path:str=""):
    _verify_admin(request)
    return {"ok":True,"logs":get_recent_logs(n,user,path)}

@app.get("/api/admin/stats",tags=["Admin"])
async def admin_stats(request:Request):
    _verify_admin(request); return {"ok":True,"stats":get_stats()}

@app.get("/api/admin/usuarios",tags=["Admin"])
async def admin_usuarios(request:Request):
    _verify_admin(request)
    return {"ok":True,"usuarios":[
        {"email":k,"nombre":v["nombre"],"rol":v["rol"],"sede":v["sede"]}
        for k,v in USUARIOS_AUTORIZADOS.items()]}

# ── PWA manifest ──────────────────────────────────────────────
@app.get("/manifest.json",include_in_schema=False)
async def manifest():
    return {"name":"SHARF Devoluciones TI","short_name":"SHARF TI",
            "start_url":"/","scope":"/","display":"standalone","orientation":"portrait",
            "background_color":"#9B1035","theme_color":"#9B1035","lang":"es",
            "icons":[{"src":"/static/icon-192.png","sizes":"192x192","type":"image/png","purpose":"any maskable"},
                     {"src":"/static/icon-512.png","sizes":"512x512","type":"image/png","purpose":"any maskable"}]}

if __name__=="__main__":
    uvicorn.run("main:app",host="0.0.0.0",port=CFG.PORT,
                reload=not CFG.is_production)
