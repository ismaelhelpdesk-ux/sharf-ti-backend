"""
core/mail_hybrid.py — Solución híbrida para Gmail

PROBLEMA: Enviar correos con cuenta de servicio de Google requiere
"Delegación de Dominio" en Google Workspace — configuración de admin.

SOLUCIÓN HÍBRIDA: Python hace todo (Snipe, PDF, Drive, Sheets)
y delega el envío de correo al GAS original via webhook.

Ventaja: No necesitas permisos de admin de Workspace.
El GAS ya tiene acceso a GmailApp y funciona hoy.
"""

import aiohttp, logging, json
from typing import Optional

log = logging.getLogger("sharf.mail")

class MailHybridClient:
    """
    Envía correos delegando al GAS original via HTTP POST.
    El GAS expone un endpoint /doPost que recibe el payload
    y usa GmailApp.sendEmail() internamente.
    """
    def __init__(self, gas_url: str, gas_token: str = ""):
        """
        gas_url   → URL de deployment del GAS (la misma Web App URL)
        gas_token → token secreto compartido para autenticar la petición
        """
        self.gas_url   = gas_url.rstrip("/")
        self.gas_token = gas_token

    async def verificar(self):
        if not self.gas_url:
            raise RuntimeError("GAS_URL no configurado")
        log.info("Mail híbrido: GAS URL configurado")

    async def enviar_devolucion(self, p, dev_id: str,
                                pdf_bytes: Optional[bytes],
                                drive_url: str):
        """
        Llama al endpoint GAS para que envíe el correo.
        El PDF ya está en Drive — se adjunta como enlace.
        """
        if not self.gas_url:
            log.warning("GAS_URL no configurado — omitiendo envío de correo")
            return

        emp = p.empleado or {}
        tec = p.tecnico  or {}

        payload = {
            "action":     "enviarCorreoDesdeAPI",
            "token":      self.gas_token,
            "devId":      dev_id,
            "tipoDev":    p.tipoDev,
            "empleado":   emp,
            "tecnico":    tec,
            "serial":     p.serial,
            "tipoActivo": p.tipoActivo,
            "driveUrl":   drive_url or "",
            "emailPara":  p.emailPara or emp.get("emailPers",""),
            "emailCC":    p.emailCC  or "helpdesk@holasharf.com",
            "emailBCC":   p.emailBCC or "",
            "modoPrueba": p.modoPrueba,
            "emailPrueba": p.emailPrueba or "",
        }

        try:
            async with aiohttp.ClientSession() as s:
                async with s.post(
                    self.gas_url,
                    json=payload,
                    timeout=aiohttp.ClientTimeout(total=30)
                ) as r:
                    resp = await r.json(content_type=None)
                    if resp.get("ok"):
                        log.info(f"Mail híbrido: enviado via GAS → {payload['emailPara']}")
                    else:
                        raise RuntimeError(resp.get("error", "GAS respondió con error"))
        except Exception as e:
            log.error(f"Mail híbrido: {e}")
            raise

class MailSMTPClient:
    """
    Alternativa SMTP directa — no requiere Google API.
    Usa el servidor SMTP de Gmail con usuario y contraseña de aplicación.
    
    Configurar en Google Account → Seguridad → Contraseñas de aplicación
    """
    def __init__(self, usuario: str, password_app: str):
        self.usuario = usuario
        self.password = password_app

    async def verificar(self):
        import smtplib
        import asyncio
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(None, self._test_smtp)

    def _test_smtp(self):
        import smtplib
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(self.usuario, self.password)
        log.info("SMTP: conexión OK")

    async def enviar_devolucion(self, p, dev_id: str,
                                pdf_bytes: Optional[bytes],
                                drive_url: str):
        import asyncio
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(None, self._enviar_sync, p, dev_id, pdf_bytes, drive_url)

    def _enviar_sync(self, p, dev_id, pdf_bytes, drive_url):
        import smtplib, base64
        from email.mime.multipart import MIMEMultipart
        from email.mime.text      import MIMEText
        from email.mime.base      import MIMEBase
        from email                import encoders
        from datetime             import datetime

        emp = p.empleado or {}
        tec = p.tecnico  or {}
        dest = p.emailPara or emp.get("emailPers","")
        if not dest:
            raise RuntimeError("Sin correo de destino")

        asunto = f"Acta de Devolución — {emp.get('nombre','')} [{dev_id}]"
        if p.modoPrueba:
            asunto += " [PRUEBA]"
            dest = p.emailPrueba or dest

        cuerpo = f"""
<html><body style="font-family:Arial,sans-serif">
<div style="max-width:580px;margin:0 auto">
  <div style="background:#9B1035;padding:20px;text-align:center">
    <h1 style="color:#fff;margin:0;font-style:italic;font-size:28px">sharf</h1>
    <p style="color:#FFC1C2;margin:4px 0 0;font-size:12px">Mesa de Servicio TI</p>
  </div>
  <div style="padding:24px;background:#fdf5f6;border:1px solid #f0d6d8">
    <h2 style="color:#9B1035;margin-top:0">Devolución registrada</h2>
    <p>Estimado/a <b>{emp.get('nombre','colaborador')}</b>,</p>
    <p>Se registró la devolución de sus equipos tecnológicos.</p>
    <table style="width:100%;border-collapse:collapse;margin:16px 0">
      <tr style="background:#FF6568;color:#fff">
        <th style="padding:8px;text-align:left">Campo</th>
        <th style="padding:8px;text-align:left">Detalle</th>
      </tr>
      <tr><td style="padding:8px;border:1px solid #f0d6d8"><b>ID</b></td>
          <td style="padding:8px;border:1px solid #f0d6d8">{dev_id}</td></tr>
      <tr style="background:#fff8f8">
          <td style="padding:8px;border:1px solid #f0d6d8"><b>Tipo</b></td>
          <td style="padding:8px;border:1px solid #f0d6d8">{p.tipoDev}</td></tr>
      <tr><td style="padding:8px;border:1px solid #f0d6d8"><b>Equipo</b></td>
          <td style="padding:8px;border:1px solid #f0d6d8">{p.tipoActivo} · S/N: {p.serial}</td></tr>
      <tr style="background:#fff8f8">
          <td style="padding:8px;border:1px solid #f0d6d8"><b>Técnico</b></td>
          <td style="padding:8px;border:1px solid #f0d6d8">{tec.get('nombre','')}</td></tr>
      <tr><td style="padding:8px;border:1px solid #f0d6d8"><b>Fecha</b></td>
          <td style="padding:8px;border:1px solid #f0d6d8">{datetime.now().strftime('%d/%m/%Y %H:%M')}</td></tr>
    </table>
    {'<p><a href="'+drive_url+'" style="background:#9B1035;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;font-weight:bold">📄 Ver Acta en Drive</a></p>' if drive_url else ''}
  </div>
  <div style="background:#f0d6d8;padding:12px;text-align:center;font-size:11px;color:#5c3040">
    SHARF Mesa de Servicio TI · helpdesk@holasharf.com
  </div>
</div></body></html>"""

        msg = MIMEMultipart("mixed")
        msg["From"]    = self.usuario
        msg["To"]      = dest
        if p.emailCC:  msg["Cc"]  = p.emailCC
        if p.emailBCC: msg["Bcc"] = p.emailBCC
        msg["Subject"] = asunto
        msg.attach(MIMEText(cuerpo, "html"))

        if pdf_bytes:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(pdf_bytes)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f'attachment; filename="Acta_{dev_id}.pdf"')
            msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(self.usuario, self.password)
            all_dest = [d.strip() for d in
                        f"{dest},{p.emailCC},{p.emailBCC}".split(",") if d.strip()]
            s.sendmail(self.usuario, all_dest, msg.as_bytes())
        log.info(f"SMTP: correo enviado a {dest}")
