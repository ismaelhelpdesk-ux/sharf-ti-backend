"""
core/models.py — Modelos de datos (Pydantic v2)
Validan y documentan los payloads del API.
"""

from pydantic import BaseModel, Field, field_validator
from typing   import Dict, List, Optional, Any
import re

class SesionRequest(BaseModel):
    email: str = Field(..., description="Email @holasharf.com del técnico")

    @field_validator("email")
    @classmethod
    def email_valido(cls, v: str) -> str:
        v = v.strip().lower()
        if not re.match(r'^[\w.+\-]+@[\w\-]+\.[\w.\-]+$', v):
            raise ValueError("Email inválido")
        return v

class Empleado(BaseModel):
    nombre:     str = ""
    dni:        str = ""
    area:       str = ""
    cargo:      str = ""
    sede:       str = ""
    ceco:       str = ""
    empresa:    str = ""
    emailCorp:  str = ""
    emailPers:  str = ""
    emailJefe:  str = ""
    locationId: Optional[int] = None

class Tecnico(BaseModel):
    nombre:   str = ""
    email:    str = ""
    sede:     str = ""
    firmaKey: str = ""

class PayloadDevolucion(BaseModel):
    devId:            Optional[str] = None
    tipoDev:          str  = Field(..., description="Por cese | Por renting | Por cambio | Por préstamo")
    modoIngreso:      str  = "manual"
    empleado:         Dict[str, Any] = {}
    activo:           Dict[str, Any] = {}
    snipeId:          Optional[int]  = None
    serial:           str  = ""
    tipoActivo:       str  = "Laptop"
    accesorios:       Dict[str, Any] = {}   # {cargadorSerial, mouseDesc, …}
    accesoriosDevueltos: Dict[str, bool] = {}
    accesoriosEstado: Dict[str, str]  = {}  # bueno | malo
    accesoriosCotizar: Dict[str, bool] = {}
    accesoriosCosto:  Dict[str, float] = {}
    accesoriosObs:    Dict[str, str]  = {}
    accesoriosNuevos: Dict[str, bool] = {}
    compromisosDiferidos: List[Dict] = []
    equipoBueno:      bool = True
    hayObservaciones: bool = False
    observacionDesc:  str  = ""
    fotosBase64:      List[str] = []        # fotos en base64
    tecnico:          Dict[str, Any] = {}
    firmaBase64:      str  = ""
    logoBase64:       str  = ""
    emailPara:        str  = ""
    emailCC:          str  = ""
    emailBCC:         str  = ""
    snipeDesact:      bool = False
    modoPrueba:       bool = False
    emailPrueba:      str  = ""

class BusquedaDNI(BaseModel):
    dni: str = Field(..., min_length=7, max_length=12)

class ValidarResp(BaseModel):
    ok:        bool
    timestamp: str
    pasos:     List[str]
    resumen:   List[str]
