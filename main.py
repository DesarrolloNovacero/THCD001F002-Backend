import os
import json
import io
import unicodedata
from datetime import datetime, timedelta
from typing import Optional
from pydantic import BaseModel
from fastapi import FastAPI, Depends, HTTPException, status, UploadFile, File, Form, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from sqlalchemy import create_engine, Column, String, Boolean, DateTime, Text, Numeric, ForeignKey, or_, func, case
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
import sqlalchemy
from jose import JWTError, jwt
from passlib.context import CryptContext
import pandas as pd

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./trainform.db")

if DATABASE_URL.startswith("sqlite"):
    engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
else:
    engine = create_engine(DATABASE_URL)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

SECRET_KEY = os.getenv("SECRET_KEY", "CLAVE_ULTRA_SECRETA_TRAINFORM")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 480 

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token")

class Usuario(Base):
    __tablename__ = "usuarios"
    id = Column(UUID(as_uuid=True), primary_key=True, server_default=sqlalchemy.text("gen_random_uuid()"))
    email = Column(String(100), unique=True, nullable=False)
    password_hash = Column(String(255), nullable=False)
    nombre_completo = Column(String(100), nullable=False)
    rol = Column(String(50), nullable=False)
    localidad = Column(String(50), nullable=False, default="Quito")
    activo = Column(Boolean, nullable=False, default=True)
    fecha_creacion = Column(DateTime, server_default=sqlalchemy.text("CURRENT_TIMESTAMP"))

class EmpresaCapacitadora(Base):
    __tablename__ = "empresas_capacitadoras"
    id = Column(UUID(as_uuid=True), primary_key=True, server_default=sqlalchemy.text("gen_random_uuid()"))
    nombre = Column(String(150), unique=True, nullable=False)
    fecha_creacion = Column(DateTime, server_default=sqlalchemy.text("CURRENT_TIMESTAMP"))

class Colaborador(Base):
    __tablename__ = "colaboradores"
    cedula = Column(String(20), primary_key=True)
    apellidos = Column(String(100))
    nombres = Column(String(100))
    cargo = Column(String(100))
    genero = Column(String(20))
    unidad = Column(String(100))
    area = Column(String(100))
    seccion = Column(String(100))
    centro_costo = Column(String(50))
    grupo_personal = Column(String(50))
    area_personal = Column(String(50))
    jefe_area = Column(String(100))
    gerente_area = Column(String(100))
    localidad = Column(String(100))
    origen = Column(String(20))
    estado_laboral = Column(String(20), default="ACTIVO")
    ultima_actualizacion = Column(DateTime, server_default=sqlalchemy.text("CURRENT_TIMESTAMP"))

class Evento(Base):
    __tablename__ = "eventos"
    id = Column(UUID(as_uuid=True), primary_key=True, server_default=sqlalchemy.text("gen_random_uuid()"))
    codigo_curso = Column(String(50), unique=True)
    creado_por_usuario_id = Column(UUID(as_uuid=True), ForeignKey("usuarios.id"), nullable=False)
    nombre_curso = Column(String(200), nullable=False)
    objetivo = Column(String(200))
    empresa = Column(String(100))
    facilitador = Column(String(100))
    dimension_evento = Column(String(100))
    lugar = Column(String(100))
    modalidad = Column(String(50))
    fecha_hora_inicio = Column(DateTime)
    fecha_hora_cierre = Column(DateTime)
    total_horas = Column(Numeric(5, 2))
    tipo_evento = Column(String(50))
    mes_anio = Column(String(20))
    estado = Column(String(50), default="PENDIENTE")
    localidad = Column(String(50))
    fecha_creacion = Column(DateTime, server_default=sqlalchemy.text("CURRENT_TIMESTAMP"))

class Asistencia(Base):
    __tablename__ = "asistencias"
    id = Column(UUID(as_uuid=True), primary_key=True, server_default=sqlalchemy.text("gen_random_uuid()"))
    evento_id = Column(UUID(as_uuid=True), ForeignKey("eventos.id", ondelete="CASCADE"), nullable=False)
    colaborador_cedula = Column(String(20), ForeignKey("colaboradores.cedula"), nullable=False)
    estado_validacion = Column(String(50), nullable=False)
    fecha_registro = Column(DateTime, server_default=sqlalchemy.text("CURRENT_TIMESTAMP"))

class HistorialEvento(Base):
    __tablename__ = "historial_eventos"
    id = Column(UUID(as_uuid=True), primary_key=True, server_default=sqlalchemy.text("gen_random_uuid()"))
    evento_id = Column(UUID(as_uuid=True), ForeignKey("eventos.id", ondelete="CASCADE"), nullable=False)
    usuario_id = Column(UUID(as_uuid=True), ForeignKey("usuarios.id"), nullable=False)
    accion = Column(String(50), nullable=False)
    comentario = Column(Text)
    fecha_registro = Column(DateTime, server_default=sqlalchemy.text("CURRENT_TIMESTAMP"))

class SesionWeb(Base):
    __tablename__ = "sesiones_web"
    id = Column(UUID(as_uuid=True), primary_key=True, server_default=sqlalchemy.text("gen_random_uuid()"))
    usuario_id = Column(UUID(as_uuid=True), ForeignKey("usuarios.id"), nullable=False)
    token_sesion = Column(String(255), nullable=False, unique=True)
    estado_borrador_json = Column(Text)
    ultima_modificacion = Column(DateTime, server_default=sqlalchemy.text("CURRENT_TIMESTAMP"))
    expira_en = Column(DateTime, nullable=False)

class MetricaMensual(Base):
    __tablename__ = "metricas_mensuales"
    mes_anio = Column(String(20), primary_key=True)
    total_activos = Column(sqlalchemy.Integer, nullable=False)
    ultima_actualizacion = Column(DateTime, server_default=sqlalchemy.text("CURRENT_TIMESTAMP"))

class NuevoUsuario(BaseModel):
    email: str
    password: str
    nombre_completo: str
    rol: str
    localidad: str

class NuevaEmpresa(BaseModel):
    nombre: str

class UpdatePasswordModel(BaseModel):
    password: str

class AuditoriaAccion(BaseModel):
    comentario: str

Base.metadata.create_all(bind=engine)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def parse_iso_date(date_str):
    if not date_str: return None
    try:
        dt = pd.to_datetime(date_str)
        if pd.isna(dt): return None
        return dt.replace(tzinfo=None).to_pydatetime()
    except:
        return None

def verify_password(plain_password, hashed_password):
    return pwd_context.verify(plain_password, hashed_password)

def create_access_token(data: dict, expires_delta: Optional[timedelta] = None):
    to_encode = data.copy()
    expire = datetime.utcnow() + (expires_delta if expires_delta else timedelta(minutes=15))
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)

def get_current_user(token: str = Depends(oauth2_scheme), db: Session = Depends(get_db)):
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Credenciales invalidas",
        headers={"WWW-Authenticate": "Bearer"},
    )
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        email: str = payload.get("sub")
        if email is None: raise credentials_exception
    except JWTError:
        raise credentials_exception
        
    user = db.query(Usuario).filter(Usuario.email == email).first()
    if user is None or not user.activo:
        raise credentials_exception
    return user

def get_current_admin(current_user: Usuario = Depends(get_current_user)):
    if current_user.rol != "ADMIN":
        raise HTTPException(status_code=403, detail="No tienes permisos")
    return current_user

@app.post("/token")
def login_for_access_token(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user = db.query(Usuario).filter(Usuario.email == form_data.username).first()
    if not user or not verify_password(form_data.password, user.password_hash):
        raise HTTPException(status_code=401, detail="Email o contraseña incorrectos")
    
    access_token = create_access_token(
        data={"sub": user.email, "rol": user.rol}, 
        expires_delta=timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    )
    return {
        "access_token": access_token, 
        "token_type": "bearer",
        "user_name": user.nombre_completo,
        "user_role": user.rol,
        "user_location": user.localidad
    }

@app.post("/crear-usuario")
def crear_usuario(data: NuevoUsuario, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    usuario_existente = db.query(Usuario).filter(Usuario.email == data.email).first()
    if usuario_existente:
        raise HTTPException(status_code=400, detail="Este correo ya está registrado")
        
    hash_generado = pwd_context.hash(data.password)
    nuevo_usuario = Usuario(email=data.email, password_hash=hash_generado, nombre_completo=data.nombre_completo, rol=data.rol, localidad=data.localidad)
    db.add(nuevo_usuario)
    db.commit()
    return {"message": "Usuario creado"}

@app.get("/usuarios")
def listar_usuarios(current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    usuarios = db.query(Usuario).order_by(Usuario.fecha_creacion.desc()).all()
    return [{"id": str(u.id), "email": u.email, "nombre_completo": u.nombre_completo, "rol": u.rol, "activo": u.activo, "localidad": u.localidad} for u in usuarios]

@app.put("/usuarios/{user_id}/toggle")
def toggle_usuario(user_id: str, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    if str(current_admin.id) == user_id: raise HTTPException(status_code=400)
    usuario = db.query(Usuario).filter(Usuario.id == user_id).first()
    if not usuario: raise HTTPException(status_code=404)
    usuario.activo = not usuario.activo
    db.commit()
    return {"message": "Estado actualizado", "activo": usuario.activo}

@app.put("/usuarios/{user_id}/password")
def cambiar_password(user_id: str, data: UpdatePasswordModel, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    usuario = db.query(Usuario).filter(Usuario.id == user_id).first()
    if not usuario: raise HTTPException(status_code=404)
    usuario.password_hash = pwd_context.hash(data.password)
    db.commit()
    return {"message": "Contraseña actualizada"}

@app.delete("/usuarios/{user_id}")
def eliminar_usuario(user_id: str, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    if str(current_admin.id) == user_id: raise HTTPException(status_code=400)
    usuario = db.query(Usuario).filter(Usuario.id == user_id).first()
    if not usuario: raise HTTPException(status_code=404)
    if db.query(Evento).filter(Evento.creado_por_usuario_id == user_id).first(): raise HTTPException(status_code=400, detail="Desactívalo, tiene eventos.")
    db.query(SesionWeb).filter(SesionWeb.usuario_id == user_id).delete()
    db.delete(usuario)
    db.commit()
    return {"message": "Usuario eliminado"}

@app.get("/empresas")
def listar_empresas(current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    empresas = db.query(EmpresaCapacitadora).order_by(EmpresaCapacitadora.nombre).all()
    return [{"id": str(e.id), "nombre": e.nombre} for e in empresas]

@app.post("/empresas")
def crear_empresa(data: NuevaEmpresa, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    existe = db.query(EmpresaCapacitadora).filter(EmpresaCapacitadora.nombre.ilike(data.nombre)).first()
    if existe: raise HTTPException(status_code=400, detail="La empresa ya existe")
    db.add(EmpresaCapacitadora(nombre=data.nombre.upper().strip()))
    db.commit()
    return {"message": "Empresa añadida"}

@app.delete("/empresas/{empresa_id}")
def eliminar_empresa(empresa_id: str, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    emp = db.query(EmpresaCapacitadora).filter(EmpresaCapacitadora.id == empresa_id).first()
    if not emp: raise HTTPException(status_code=404)
    db.delete(emp)
    db.commit()
    return {"message": "Empresa eliminada"}

@app.get("/check-db-status")
def check_db_status(db: Session = Depends(get_db)):
    count = db.query(Colaborador).count()
    return {"ready": count > 0, "count": count}

@app.get("/load-state")
def load_state(current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    sesion = db.query(SesionWeb).filter(SesionWeb.usuario_id == current_user.id).first()
    if sesion and sesion.estado_borrador_json: return json.loads(sesion.estado_borrador_json)
    return None

@app.post("/save-state")
def save_state(payload: dict = Body(...), current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    sesion = db.query(SesionWeb).filter(SesionWeb.usuario_id == current_user.id).first()
    if not sesion:
        sesion = SesionWeb(usuario_id=current_user.id, token_sesion=f"draft_{current_user.id}", estado_borrador_json=json.dumps(payload), expira_en=datetime.utcnow() + timedelta(days=7))
        db.add(sesion)
    else:
        sesion.estado_borrador_json = json.dumps(payload)
        sesion.ultima_modificacion = datetime.utcnow()
    db.commit()
    return {"status": "ok"}

@app.post("/upload-masters")
async def upload_masters(file: UploadFile = File(...), mes_corte: str = Form(...), current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    contents = await file.read()
    df = pd.read_excel(io.BytesIO(contents), dtype=str)
    df = df.fillna("")
    df.columns = [str(c).strip().upper() for c in df.columns]

    def find_col(keywords):
        for col in df.columns:
            for kw in keywords:
                if kw in col: return col
        return None

    col_cedula = find_col(["CÉDULA DE IDENTIFICACIÓN", "CEDULA", "IDENTIFICACIÓN NACIONAL"])
    col_apellidos = find_col(["APELLIDOS"])
    col_nombres = find_col(["NOMBRES"])
    col_cargo = find_col(["CARGO NOMBRE DEL PUESTO", "CARGO"])
    col_genero = find_col(["SEXO", "GENERO"])
    col_unidad = find_col(["UNIDAD DE NEGOCIO", "UNIDAD"])
    col_area = find_col(["ÁREA NOMBRE", "AREA NOMBRE"])
    col_seccion = find_col(["SECCIÓN NOMBRE", "SECCION NOMBRE"])
    col_cc = find_col(["CENTRO DE COSTO"])
    col_gp = find_col(["GRUPO DE PERSONAL"])
    col_ap = find_col(["ÁREA DE PERSONAL", "AREA DE PERSONAL"])
    col_ja = find_col(["JEFE INMEDIATO", "JEFE"])
    col_ga = find_col(["GERENTE DE AREA", "GERENTE"])
    col_loc = find_col(["LOCACIÓN", "LOCACION", "LOCALIDAD"])
    col_desv = find_col(["FECHA DE DESVINCULACIÓN", "FECHA DE DESVINCULACION", "DESVINCULACION"])

    conteo_activos = 0

    for index, row in df.iterrows():
        if not col_cedula: continue
        cedula = str(row[col_cedula]).strip()
        if cedula.endswith('.0'): cedula = cedula[:-2]
        if not cedula or cedula.lower() == "nan": continue
        if cedula.isdigit() and len(cedula) < 10: cedula = cedula.zfill(10)
            
        colaborador = db.query(Colaborador).filter(Colaborador.cedula == cedula).first()
        if not colaborador:
            colaborador = Colaborador(cedula=cedula)
            db.add(colaborador)

        val_desv = str(row[col_desv]).strip().upper() if col_desv else ""
        es_cesado = bool(val_desv and val_desv != "NAN" and val_desv != "NAT" and val_desv != "NONE")
        estado_final = "CESADO" if es_cesado else "ACTIVO"
        if estado_final == "ACTIVO": conteo_activos += 1
            
        colaborador.apellidos = str(row[col_apellidos]).strip() if col_apellidos else ""
        colaborador.nombres = str(row[col_nombres]).strip() if col_nombres else ""
        colaborador.cargo = str(row[col_cargo]).strip() if col_cargo else ""
        colaborador.genero = str(row[col_genero]).strip() if col_genero else ""
        colaborador.unidad = str(row[col_unidad]).strip() if col_unidad else ""
        colaborador.area = str(row[col_area]).strip() if col_area else ""
        colaborador.seccion = str(row[col_seccion]).strip() if col_seccion else ""
        colaborador.centro_costo = str(row[col_cc]).strip() if col_cc else ""
        colaborador.grupo_personal = str(row[col_gp]).strip() if col_gp else ""
        colaborador.area_personal = str(row[col_ap]).strip() if col_ap else ""
        colaborador.jefe_area = str(row[col_ja]).strip() if col_ja else ""
        colaborador.gerente_area = str(row[col_ga]).strip() if col_ga else ""
        colaborador.localidad = str(row[col_loc]).strip() if col_loc else ""
        colaborador.origen = "maestro"
        colaborador.estado_laboral = estado_final
        
    metrica = db.query(MetricaMensual).filter(MetricaMensual.mes_anio == mes_corte).first()
    if metrica:
        metrica.total_activos = conteo_activos
        metrica.ultima_actualizacion = datetime.utcnow()
    else:
        db.add(MetricaMensual(mes_anio=mes_corte, total_activos=conteo_activos))
        
    db.commit()
    return {"message": "Datos maestros unificados"}

@app.post("/validate-cedula")
def validate_cedula(cedulas_json: str = Form(...), current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    cedulas = json.loads(cedulas_json)
    resultados = []
    for cedula in cedulas:
        cedula_str = str(cedula).strip()
        colaborador = db.query(Colaborador).filter(Colaborador.cedula == cedula_str).first()
        if colaborador:
            data = {
                "apellidos": colaborador.apellidos, "nombres": colaborador.nombres, "cargo": colaborador.cargo,
                "genero": colaborador.genero, "unidad": colaborador.unidad, "area": colaborador.area, "seccion": colaborador.seccion,
                "centro_costo": colaborador.centro_costo, "grupo_personal": colaborador.grupo_personal, "area_personal": colaborador.area_personal,
                "jefe_area": colaborador.jefe_area, "gerente_area": colaborador.gerente_area, "localidad": colaborador.localidad
            }
            fake_source = "cesantes" if colaborador.estado_laboral == "CESADO" else "headcount"
            resultados.append({"cedula": cedula_str, "found": True, "source": fake_source, "data": data})
        else:
            resultados.append({"cedula": cedula_str, "found": False, "source": None, "data": {}})
    return resultados

@app.post("/suggest-cedulas")
def suggest_cedulas(search_term: str = Form(...), current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    search_term = search_term.strip()
    if len(search_term) < 2: return []
    search_pattern = f"%{search_term}%"
    resultados = db.query(Colaborador).filter(
        or_(Colaborador.cedula.ilike(search_pattern), Colaborador.nombres.ilike(search_pattern), Colaborador.apellidos.ilike(search_pattern))
    ).limit(10).all()
    sugs = []
    for r in resultados:
        fake_source = "cesantes" if r.estado_laboral == "CESADO" else "headcount"
        sugs.append({"cedula": r.cedula, "nombre": f"{r.apellidos} {r.nombres}".strip(), "source": fake_source})
    return sugs

@app.post("/enviar-revision")
def enviar_revision(payload: dict = Body(...), current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    event_data = payload.get("eventData", {})
    registros = payload.get("registros", [])
    evento_id = payload.get("eventoId")

    if not registros: raise HTTPException(status_code=400, detail="No hay registros")

    try: horas = float(event_data.get("totalHoras", 0)) if event_data.get("totalHoras") else 0.0
    except ValueError: horas = 0.0

    inicio_dt = parse_iso_date(event_data.get("fechaHoraInicio"))
    cierre_dt = parse_iso_date(event_data.get("fechaHoraCierre"))
    loc_curso = event_data.get("localidadCurso")
    if not loc_curso: loc_curso = current_user.localidad

    if evento_id:
        evento = db.query(Evento).filter(Evento.id == evento_id).first()
        if not evento: raise HTTPException(status_code=404)
        evento.nombre_curso = event_data.get("nombreCurso", "Sin Nombre")
        evento.objetivo = event_data.get("objetivo")
        evento.empresa = event_data.get("empresa")
        evento.facilitador = event_data.get("facilitador")
        evento.dimension_evento = event_data.get("dimensionEvento")
        evento.lugar = event_data.get("lugar")
        evento.modalidad = event_data.get("modalidad")
        evento.fecha_hora_inicio = inicio_dt
        evento.fecha_hora_cierre = cierre_dt
        evento.total_horas = horas
        evento.tipo_evento = event_data.get("tipoEvento")
        evento.mes_anio = event_data.get("mesAnio")
        evento.localidad = loc_curso
        evento.estado = "PENDIENTE"
        db.query(Asistencia).filter(Asistencia.evento_id == evento_id).delete()
    else:
        count = db.query(Evento).count() + 1
        codigo = f"NOV-{datetime.utcnow().year}-{str(count).zfill(4)}"
        evento = Evento(
            codigo_curso=codigo, creado_por_usuario_id=current_user.id, nombre_curso=event_data.get("nombreCurso", "Sin Nombre"),
            objetivo=event_data.get("objetivo"), empresa=event_data.get("empresa"), facilitador=event_data.get("facilitador"),
            dimension_evento=event_data.get("dimensionEvento"), lugar=event_data.get("lugar"), modalidad=event_data.get("modalidad"),
            fecha_hora_inicio=inicio_dt, fecha_hora_cierre=cierre_dt, total_horas=horas, tipo_evento=event_data.get("tipoEvento"),
            mes_anio=event_data.get("mesAnio"), localidad=loc_curso, estado="PENDIENTE"
        )
        db.add(evento)
        db.flush()

    for r in registros:
        cedula_colaborador = str(r.get("CÉDULA", "")).strip()
        if cedula_colaborador: db.add(Asistencia(evento_id=evento.id, colaborador_cedula=cedula_colaborador, estado_validacion="VALIDADO"))

    db.add(HistorialEvento(evento_id=evento.id, usuario_id=current_user.id, accion="ENVIADO A REVISION", comentario=""))
    db.commit()
    return {"message": "Enviado a revisión", "evento_id": str(evento.id)}

@app.get("/mis-eventos")
def mis_eventos(current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    eventos = db.query(Evento).filter(Evento.creado_por_usuario_id == current_user.id).order_by(Evento.fecha_creacion.desc()).all()
    res = []
    for e in eventos:
        historial = db.query(HistorialEvento).filter(HistorialEvento.evento_id == e.id).order_by(HistorialEvento.fecha_registro.desc()).first()
        res.append({
            "id": str(e.id), "codigo": e.codigo_curso, "nombre": e.nombre_curso, "estado": e.estado,
            "fecha": e.fecha_creacion, "comentario": historial.comentario if historial else ""
        })
    return res

@app.get("/admin/eventos")
def admin_eventos(current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    eventos = db.query(Evento).order_by(Evento.fecha_creacion.desc()).all()
    res = []
    for e in eventos:
        usuario = db.query(Usuario).filter(Usuario.id == e.creado_por_usuario_id).first()
        res.append({
            "id": str(e.id), "codigo": e.codigo_curso, "nombre": e.nombre_curso, "estado": e.estado,
            "creador": usuario.nombre_completo if usuario else "Desconocido", "fecha": e.fecha_creacion
        })
    return res

@app.put("/admin/eventos/{evento_id}/aprobar")
def aprobar_evento(evento_id: str, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    evento = db.query(Evento).filter(Evento.id == evento_id).first()
    if not evento: raise HTTPException(status_code=404)
    evento.estado = "APROBADO"
    db.add(HistorialEvento(evento_id=evento.id, usuario_id=current_admin.id, accion="APROBADO", comentario=""))
    db.commit()
    return {"message": "Aprobado"}

@app.put("/admin/eventos/{evento_id}/rechazar")
def rechazar_evento(evento_id: str, accion: AuditoriaAccion, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    evento = db.query(Evento).filter(Evento.id == evento_id).first()
    if not evento: raise HTTPException(status_code=404)
    evento.estado = "RECHAZADO"
    db.add(HistorialEvento(evento_id=evento.id, usuario_id=current_admin.id, accion="RECHAZADO", comentario=accion.comentario))
    db.commit()
    return {"message": "Rechazado"}

@app.get("/admin/eventos/{evento_id}/exportar")
def exportar_evento(evento_id: str, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    evento = db.query(Evento).filter(Evento.id == evento_id).first()
    if not evento: raise HTTPException(status_code=404)
    asistencias = db.query(Asistencia).filter(Asistencia.evento_id == evento.id).all()
    procesados = []
    for a in asistencias:
        colab = db.query(Colaborador).filter(Colaborador.cedula == a.colaborador_cedula).first()
        if not colab: continue
        procesados.append({
            "NOMBRE DEL CURSO": evento.nombre_curso,
            "OBJETIVO": evento.objetivo,
            "EMPRESA CAPACITADORA": evento.empresa,
            "FACILITADOR": evento.facilitador,
            "DIMENSIÓN DE EVENTO": evento.dimension_evento,
            "LUGAR DONDE SE DIO LA CAPACITACION": evento.lugar,
            "MODALIDAD": evento.modalidad,
            "FECHA INICIO": evento.fecha_hora_inicio.strftime('%d/%m/%Y') if evento.fecha_hora_inicio else '',
            "FECHA CIERRE": evento.fecha_hora_cierre.strftime('%d/%m/%Y') if evento.fecha_hora_cierre else '',
            "DURACION DE LA CAPACITACION (HORAS)": str(evento.total_horas),
            "TIPO EVENTO": evento.tipo_evento,
            "MES-AÑO": evento.mes_anio,
            "CÉDULA": colab.cedula,
            "APELLIDOS Y NOMBRE DEL COLABORADOR": f"{colab.apellidos} {colab.nombres}".strip(),
            "GÉNERO": colab.genero,
            "CARGO": colab.cargo,
            "UNIDAD": colab.unidad,
            "ÁREA": colab.area,
            "SECCIÓN": colab.seccion,
            "CENTRO DE COSTO": colab.centro_costo,
            "GRUPO DE PERSONAL": colab.grupo_personal,
            "ÁREA DE PERSONAL": colab.area_personal,
            "JEFE DE ÁREA": colab.jefe_area,
            "GERENTE DE AREA": colab.gerente_area,
            "LOCALIDAD": evento.localidad,
        })
    df = pd.DataFrame(procesados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='Auditoria')
    output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=auditoria_{evento.codigo_curso}.xlsx"})

#---

def get_kpis_for_month(mes_str: str, estado: str, db: Session):
    mes_busqueda = mes_str
    if len(mes_str) == 7 and mes_str[5] == '0':
        mes_alternativo = mes_str[:5] + mes_str[6:]
    else:
        mes_alternativo = mes_str

    query = db.query(Asistencia.id, Asistencia.colaborador_cedula, Evento.total_horas)\
              .join(Evento, Asistencia.evento_id == Evento.id)\
              .outerjoin(Colaborador, Asistencia.colaborador_cedula == Colaborador.cedula)\
              .filter(or_(Evento.mes_anio.ilike(f"%{mes_busqueda}%"), Evento.mes_anio.ilike(f"%{mes_alternativo}%")))
    
    query = query.filter(Evento.estado != "RECHAZADO")

    if estado != "TODOS":
        query = query.filter(Colaborador.estado_laboral == estado)

    asistencias = query.all()
    if not asistencias:
        return 0, 0

    metrica_mes = db.query(MetricaMensual).filter(MetricaMensual.mes_anio == mes_str).first()
    total_empresa = metrica_mes.total_activos if metrica_mes and metrica_mes.total_activos > 0 else 1

    cedulas_unicas = set()
    total_horas = 0
    for a in asistencias:
        cedulas_unicas.add(a.colaborador_cedula)
        total_horas += float(a.total_horas or 0)

    porcentaje_capacitado = (len(cedulas_unicas) / total_empresa) * 100
    return total_horas, porcentaje_capacitado



#---
@app.get("/dashboard/metricas")
def obtener_metricas(mes: str, estado: str = "TODOS", current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    try:
        y, m = map(int, mes.split('-'))
        prev_mes = f"{y-1}-12" if m == 1 else f"{y}-{m-1:02d}"
    except:
        prev_mes = ""

    prev_horas, prev_pct = get_kpis_for_month(prev_mes, estado, db) if prev_mes else (0, 0)
    
    mes_busqueda = mes
    mes_alternativo = mes
    if len(mes) == 7 and mes[5] == '0':
        mes_alternativo = mes[:5] + mes[6:]

    query_base = db.query(
        Asistencia.id,
        Asistencia.colaborador_cedula, 
        Evento.total_horas, 
        Colaborador.genero, 
        Colaborador.unidad, 
        Evento.localidad, 
        Colaborador.grupo_personal, 
        Evento.modalidad, 
        Evento.dimension_evento,
        Evento.nombre_curso
    ).join(Evento, Asistencia.evento_id == Evento.id)\
     .outerjoin(Colaborador, Asistencia.colaborador_cedula == Colaborador.cedula)\
     .filter(Evento.estado != "RECHAZADO")\
     .filter(or_(Evento.mes_anio.ilike(f"%{mes_busqueda}%"), Evento.mes_anio.ilike(f"%{mes_alternativo}%")))

    if estado != "TODOS":
        query_base = query_base.filter(Colaborador.estado_laboral == estado)

    asistencias = query_base.all()

    if not asistencias:
        return {
            "kpis": {"total_colaboradores": 0, "total_horas": 0, "horas_promedio": 0, "total_cursos": 0, "personal_capacitado_pct": 0},
            "tendencias": {"diferencia_horas": round(0 - prev_horas, 1), "diferencia_pct": round(0 - prev_pct, 2)},
            "graficos": {"modalidad": [], "genero": [], "dimension_grupo": [], "unidad_negocio": [], "localidad": []}
        }
        
    metrica_mes = db.query(MetricaMensual).filter(MetricaMensual.mes_anio == mes).first()
    total_empresa = metrica_mes.total_activos if metrica_mes and metrica_mes.total_activos > 0 else 1

    nombres_cursos_unicos = set()
    cedulas_unicas = set()
    total_horas = 0
    mod_dic, gen_dic, uni_dic, loc_dic, dim_grp_dic = {}, {}, {}, {}, {}

    for a in asistencias:
        nombres_cursos_unicos.add(a.nombre_curso)
        cedulas_unicas.add(a.colaborador_cedula)
        
        hrs = float(a.total_horas or 0)
        total_horas += hrs
        
        g = a.genero or "NO DEFINIDO"
        m_val = a.modalidad or "NO DEFINIDA"
        u = a.unidad or "NO DEFINIDA"
        l = a.localidad or "NO DEFINIDA"
        gp = a.grupo_personal or "NO DEFINIDO"
        d = a.dimension_evento or "NO DEFINIDA"

        mod_dic[m_val] = mod_dic.get(m_val, 0) + hrs
        gen_dic[g] = gen_dic.get(g, 0) + hrs
        uni_dic[u] = uni_dic.get(u, 0) + hrs
        loc_dic[l] = loc_dic.get(l, 0) + hrs
        
        if d not in dim_grp_dic: dim_grp_dic[d] = {}
        dim_grp_dic[d][gp] = dim_grp_dic[d].get(gp, 0) + hrs

    total_colaboradores = len(cedulas_unicas)
    porcentaje_capacitado = (total_colaboradores / total_empresa) * 100
    
    horas_promedio = total_horas / len(asistencias) if len(asistencias) > 0 else 0

    formatear_dic = lambda d: [{"name": k, "value": round(v, 1)} for k, v in d.items()]
    
    dim_grp_lista = []
    for d_name, grp_data in dim_grp_dic.items():
        obj = {"dimension": d_name}
        for g_name, hrs_val in grp_data.items():
            obj[g_name] = round(hrs_val, 1)
        dim_grp_lista.append(obj)

    return {
        "kpis": {
            "total_colaboradores": total_colaboradores, 
            "total_horas": round(total_horas, 1), 
            "horas_promedio": round(horas_promedio, 1), 
            "total_cursos": len(nombres_cursos_unicos), 
            "personal_capacitado_pct": round(porcentaje_capacitado, 2)
        },
        "tendencias": {
            "diferencia_horas": round(total_horas - prev_horas, 1), 
            "diferencia_pct": round(porcentaje_capacitado - prev_pct, 2)
        },
        "graficos": {
            "modalidad": formatear_dic(mod_dic), 
            "genero": formatear_dic(gen_dic), 
            "unidad_negocio": formatear_dic(uni_dic), 
            "localidad": formatear_dic(loc_dic), 
            "dimension_grupo": dim_grp_lista
        }
    }
    # Buscamos coincidencias flexibles para el mes (por si acaso hay 2026-03 o 2026-3)
    mes_busqueda = mes_str
    if len(mes_str) == 7 and mes_str[5] == '0':
        mes_alternativo = mes_str[:5] + mes_str[6:] # "2026-03" -> "2026-3"
    else:
        mes_alternativo = mes_str

    query = db.query(Asistencia.colaborador_cedula, Evento.total_horas)\
              .join(Evento, Asistencia.evento_id == Evento.id)\
              .outerjoin(Colaborador, Asistencia.colaborador_cedula == Colaborador.cedula)\
              .filter(or_(Evento.mes_anio == mes_busqueda, Evento.mes_anio == mes_alternativo))
    
    # IMPORTANTE: En el total de horas general, Power BI no suele filtrar por estado "Aprobado"
    # pero para mantener integridad con tu sistema, buscaremos los que no estén RECHAZADOS.
    query = query.filter(Evento.estado != "RECHAZADO")

    if estado != "TODOS":
        query = query.filter(Colaborador.estado_laboral == estado)

    asistencias = query.all()
    if not asistencias:
        return 0, 0

    metrica_mes = db.query(MetricaMensual).filter(MetricaMensual.mes_anio == mes_str).first()
    total_empresa = metrica_mes.total_activos if metrica_mes and metrica_mes.total_activos > 0 else 1

    cedulas_unicas = set()
    total_horas = 0
    for a in asistencias:
        cedulas_unicas.add(a.colaborador_cedula)
        total_horas += float(a.total_horas or 0)

    porcentaje_capacitado = (len(cedulas_unicas) / total_empresa) * 100
    return total_horas, porcentaje_capacitado

@app.get("/dashboard/metricas")
def obtener_metricas(mes: str, estado: str = "TODOS", current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    try:
        y, m = map(int, mes.split('-'))
        prev_mes = f"{y-1}-12" if m == 1 else f"{y}-{m-1:02d}"
    except:
        prev_mes = ""

    prev_horas, prev_pct = get_kpis_for_month(prev_mes, estado, db) if prev_mes else (0, 0)
    
    # Ajuste de mes para búsqueda flexible
    mes_busqueda = mes
    mes_alternativo = mes
    if len(mes) == 7 and mes[5] == '0':
        mes_alternativo = mes[:5] + mes[6:]

    query_base = db.query(
        Asistencia.colaborador_cedula, 
        Evento.total_horas, 
        Colaborador.genero, 
        Colaborador.unidad, 
        Evento.localidad, 
        Colaborador.grupo_personal, 
        Evento.modalidad, 
        Evento.dimension_evento,
        Evento.nombre_curso
    ).join(Evento, Asistencia.evento_id == Evento.id)\
     .outerjoin(Colaborador, Asistencia.colaborador_cedula == Colaborador.cedula)\
     .filter(Evento.estado != "RECHAZADO")\
     .filter(or_(Evento.mes_anio == mes_busqueda, Evento.mes_anio == mes_alternativo))

    if estado != "TODOS":
        query_base = query_base.filter(Colaborador.estado_laboral == estado)

    asistencias = query_base.all()

    if not asistencias:
        return {
            "kpis": {"total_colaboradores": 0, "total_horas": 0, "horas_promedio": 0, "total_cursos": 0, "personal_capacitado_pct": 0},
            "tendencias": {"diferencia_horas": round(0 - prev_horas, 1), "diferencia_pct": round(0 - prev_pct, 2)},
            "graficos": {"modalidad": [], "genero": [], "dimension_grupo": [], "unidad_negocio": [], "localidad": []}
        }
        
    metrica_mes = db.query(MetricaMensual).filter(MetricaMensual.mes_anio == mes).first()
    total_empresa = metrica_mes.total_activos if metrica_mes and metrica_mes.total_activos > 0 else 1

    nombres_cursos_unicos = set()
    cedulas_unicas = set()
    total_horas = 0
    mod_dic, gen_dic, uni_dic, loc_dic, dim_grp_dic = {}, {}, {}, {}, {}

    for a in asistencias:
        nombres_cursos_unicos.add(a.nombre_curso)
        cedulas_unicas.add(a.colaborador_cedula)
        
        hrs = float(a.total_horas or 0)
        total_horas += hrs
        
        # Agrupaciones para gráficos con nombres amigables para "vacíos"
        g = a.genero or "NO DEFINIDO"
        m = a.modalidad or "NO DEFINIDA"
        u = a.unidad or "NO DEFINIDA"
        l = a.localidad or "NO DEFINIDA"
        gp = a.grupo_personal or "NO DEFINIDO"
        d = a.dimension_evento or "NO DEFINIDA"

        mod_dic[m] = mod_dic.get(m, 0) + hrs
        gen_dic[g] = gen_dic.get(g, 0) + hrs
        uni_dic[u] = uni_dic.get(u, 0) + hrs
        loc_dic[l] = loc_dic.get(l, 0) + hrs
        
        if d not in dim_grp_dic: dim_grp_dic[d] = {}
        dim_grp_dic[d][gp] = dim_grp_dic[d].get(gp, 0) + hrs

    total_colaboradores = len(cedulas_unicas)
    porcentaje_capacitado = (total_colaboradores / total_empresa) * 100
    
    # Power BI suele promediar Horas Totales / Conteo de registros en la tabla plana
    horas_promedio = total_horas / len(asistencias) if len(asistencias) > 0 else 0

    formatear_dic = lambda d: [{"name": k, "value": round(v, 1)} for k, v in d.items()]
    
    dim_grp_lista = []
    for d_name, grp_data in dim_grp_dic.items():
        obj = {"dimension": d_name}
        for g_name, hrs_val in grp_data.items():
            obj[g_name] = round(hrs_val, 1)
        dim_grp_lista.append(obj)

    return {
        "kpis": {
            "total_colaboradores": total_colaboradores, 
            "total_horas": round(total_horas, 1), 
            "horas_promedio": round(horas_promedio, 1), 
            "total_cursos": len(nombres_cursos_unicos), 
            "personal_capacitado_pct": round(porcentaje_capacitado, 2)
        },
        "tendencias": {
            "diferencia_horas": round(total_horas - prev_horas, 1), 
            "diferencia_pct": round(porcentaje_capacitado - prev_pct, 2)
        },
        "graficos": {
            "modalidad": formatear_dic(mod_dic), 
            "genero": formatear_dic(gen_dic), 
            "unidad_negocio": formatear_dic(uni_dic), 
            "localidad": formatear_dic(loc_dic), 
            "dimension_grupo": dim_grp_lista
        }
    }



@app.get("/dashboard/metricas")
def obtener_metricas(mes: str, estado: str = "TODOS", current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    try:
        y, m = map(int, mes.split('-'))
        prev_mes = f"{y-1}-12" if m == 1 else f"{y}-{m-1:02d}"
    except:
        prev_mes = ""

    prev_horas, prev_pct = get_kpis_for_month(prev_mes, estado, db) if prev_mes else (0, 0)
    
    query_base = db.query(
        Asistencia.colaborador_cedula, 
        Evento.total_horas, 
        Colaborador.genero, 
        Colaborador.unidad, 
        Evento.localidad, 
        Colaborador.grupo_personal, 
        Evento.modalidad, 
        Evento.dimension_evento,
        Evento.nombre_curso
    ).join(Evento, Asistencia.evento_id == Evento.id).outerjoin(Colaborador, Asistencia.colaborador_cedula == Colaborador.cedula).filter(Evento.estado == "APROBADO", Evento.mes_anio == mes)

    if estado != "TODOS":
        query_base = query_base.filter(Colaborador.estado_laboral == estado)

    asistencias = query_base.all()

    if not asistencias:
        return {
            "kpis": {"total_colaboradores": 0, "total_horas": 0, "horas_promedio": 0, "total_cursos": 0, "personal_capacitado_pct": 0},
            "tendencias": {"diferencia_horas": 0 - prev_horas, "diferencia_pct": 0 - prev_pct},
            "graficos": {"modalidad": [], "genero": [], "dimension_grupo": [], "unidad_negocio": [], "localidad": []}
        }
        
    metrica_mes = db.query(MetricaMensual).filter(MetricaMensual.mes_anio == mes).first()
    total_empresa = metrica_mes.total_activos if metrica_mes and metrica_mes.total_activos > 0 else 1

    nombres_cursos_unicos = set()
    cedulas_unicas = set()
    total_horas = 0
    mod_dic, gen_dic, uni_dic, loc_dic, dim_grp_dic = {}, {}, {}, {}, {}

    for a in asistencias:
        nombres_cursos_unicos.add(a.nombre_curso)
        cedula = a.colaborador_cedula
        horas = float(a.total_horas or 0)
        gen = a.genero or "NO DEFINIDO"
        mod = a.modalidad or "NO DEFINIDA"
        uni = a.unidad or "NO DEFINIDA"
        loc = a.localidad or "NO DEFINIDA"
        grp = a.grupo_personal or "NO DEFINIDO"
        dim = a.dimension_evento or "NO DEFINIDA"

        cedulas_unicas.add(cedula)
        total_horas += horas
        
        mod_dic[mod] = mod_dic.get(mod, 0) + horas
        gen_dic[gen] = gen_dic.get(gen, 0) + horas
        uni_dic[uni] = uni_dic.get(uni, 0) + horas
        loc_dic[loc] = loc_dic.get(loc, 0) + horas
        
        if dim not in dim_grp_dic: dim_grp_dic[dim] = {}
        dim_grp_dic[dim][grp] = dim_grp_dic[dim].get(grp, 0) + horas

    total_cursos = len(nombres_cursos_unicos)
    total_colaboradores = len(cedulas_unicas)
    porcentaje_capacitado = (total_colaboradores / total_empresa) * 100
    horas_promedio = total_horas / len(asistencias) if len(asistencias) > 0 else 0

    formatear_dic = lambda d: [{"name": k, "value": round(v, 1)} for k, v in d.items()]
    
    dim_grp_lista = []
    for d_name, grp_data in dim_grp_dic.items():
        obj = {"dimension": d_name}
        for g_name, hrs in grp_data.items(): obj[g_name] = round(hrs, 1)
        dim_grp_lista.append(obj)

    return {
        "kpis": {
            "total_colaboradores": total_colaboradores, 
            "total_horas": round(total_horas, 1), 
            "horas_promedio": round(horas_promedio, 1), 
            "total_cursos": total_cursos, 
            "personal_capacitado_pct": round(porcentaje_capacitado, 2)
        },
        "tendencias": {
            "diferencia_horas": round(total_horas - prev_horas, 1), 
            "diferencia_pct": round(porcentaje_capacitado - prev_pct, 2)
        },
        "graficos": {
            "modalidad": formatear_dic(mod_dic), 
            "genero": formatear_dic(gen_dic), 
            "unidad_negocio": formatear_dic(uni_dic), 
            "localidad": formatear_dic(loc_dic), 
            "dimension_grupo": dim_grp_lista
        }
    }


@app.get("/dashboard/metricas")
def obtener_metricas(mes: str, estado: str = "TODOS", current_admin: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    
    try:
        y, m = map(int, mes.split('-'))
        prev_mes = f"{y-1}-12" if m == 1 else f"{y}-{m-1:02d}"
    except:
        prev_mes = ""

    prev_horas, prev_pct = get_kpis_for_month(prev_mes, estado, db) if prev_mes else (0, 0)
    
    eventos_aprobados = db.query(Evento).filter(Evento.estado == "APROBADO", Evento.mes_anio == mes).all()
    ids_eventos = [e.id for e in eventos_aprobados]
    
    if not ids_eventos:
        return {
            "kpis": {"total_colaboradores": 0, "total_horas": 0, "horas_promedio": 0, "total_cursos": 0, "personal_capacitado_pct": 0},
            "tendencias": {"diferencia_horas": 0 - prev_horas, "diferencia_pct": 0 - prev_pct},
            "graficos": {"modalidad": [], "genero": [], "dimension_grupo": [], "unidad_negocio": [], "localidad": []}
        }
        
    metrica_mes = db.query(MetricaMensual).filter(MetricaMensual.mes_anio == mes).first()
    total_empresa = metrica_mes.total_activos if metrica_mes and metrica_mes.total_activos > 0 else 1
    total_cursos = len(eventos_aprobados)

    query = db.query(
        Asistencia.colaborador_cedula, 
        Evento.total_horas, 
        Colaborador.genero, 
        Colaborador.unidad, 
        Evento.localidad, 
        Colaborador.grupo_personal, 
        Evento.modalidad, 
        Evento.dimension_evento
    ).join(Evento, Asistencia.evento_id == Evento.id).outerjoin(Colaborador, Asistencia.colaborador_cedula == Colaborador.cedula).filter(Evento.id.in_(ids_eventos))

    if estado != "TODOS":
        query = query.filter(Colaborador.estado_laboral == estado)

    asistencias = query.all()

    cedulas_unicas = set()
    total_horas = 0
    mod_dic, gen_dic, uni_dic, loc_dic, dim_grp_dic = {}, {}, {}, {}, {}

    for a in asistencias:
        cedula = a.colaborador_cedula
        horas = float(a.total_horas or 0)
        gen = a.genero or "NO DEFINIDO"
        mod = a.modalidad or "NO DEFINIDA"
        uni = a.unidad or "NO DEFINIDA"
        loc = a.localidad or "NO DEFINIDA"
        grp = a.grupo_personal or "NO DEFINIDO"
        dim = a.dimension_evento or "NO DEFINIDA"

        cedulas_unicas.add(cedula)
        total_horas += horas
        
        mod_dic[mod] = mod_dic.get(mod, 0) + horas
        gen_dic[gen] = gen_dic.get(gen, 0) + horas
        uni_dic[uni] = uni_dic.get(uni, 0) + horas
        loc_dic[loc] = loc_dic.get(loc, 0) + horas
        
        if dim not in dim_grp_dic: dim_grp_dic[dim] = {}
        dim_grp_dic[dim][grp] = dim_grp_dic[dim].get(grp, 0) + horas

    total_colaboradores = len(cedulas_unicas)
    porcentaje_capacitado = (total_colaboradores / total_empresa) * 100
    horas_promedio = total_horas / len(asistencias) if len(asistencias) > 0 else 0

    formatear_dic = lambda d: [{"name": k, "value": round(v, 1)} for k, v in d.items()]
    
    dim_grp_lista = []
    for d_name, grp_data in dim_grp_dic.items():
        obj = {"dimension": d_name}
        for g_name, hrs in grp_data.items(): obj[g_name] = round(hrs, 1)
        dim_grp_lista.append(obj)

    return {
        "kpis": {
            "total_colaboradores": total_colaboradores, 
            "total_horas": round(total_horas, 1), 
            "horas_promedio": round(horas_promedio, 1), 
            "total_cursos": total_cursos, 
            "personal_capacitado_pct": round(porcentaje_capacitado, 2)
        },
        "tendencias": {
            "diferencia_horas": round(total_horas - prev_horas, 1), 
            "diferencia_pct": round(porcentaje_capacitado - prev_pct, 2)
        },
        "graficos": {
            "modalidad": formatear_dic(mod_dic), 
            "genero": formatear_dic(gen_dic), 
            "unidad_negocio": formatear_dic(uni_dic), 
            "localidad": formatear_dic(loc_dic), 
            "dimension_grupo": dim_grp_lista
        }
    }

@app.get("/dashboard/metricas")
def obtener_metricas(mes: str, estado: str = "TODOS", current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    
    try:
        y, m = map(int, mes.split('-'))
        prev_mes = f"{y-1}-12" if m == 1 else f"{y}-{m-1:02d}"
    except:
        prev_mes = ""

    prev_horas, prev_pct = get_kpis_for_month(prev_mes, estado, db) if prev_mes else (0, 0)
    
    eventos_aprobados = db.query(Evento).filter(Evento.estado == "APROBADO", Evento.mes_anio == mes).all()
    ids_eventos = [e.id for e in eventos_aprobados]
    
    if not ids_eventos:
        return {
            "kpis": {"total_colaboradores": 0, "total_horas": 0, "horas_promedio": 0, "total_cursos": 0, "personal_capacitado_pct": 0},
            "tendencias": {"diferencia_horas": 0 - prev_horas, "diferencia_pct": 0 - prev_pct},
            "graficos": {"modalidad": [], "genero": [], "dimension_grupo": [], "unidad_negocio": [], "localidad": []}
        }
        
    metrica_mes = db.query(MetricaMensual).filter(MetricaMensual.mes_anio == mes).first()
    total_empresa = metrica_mes.total_activos if metrica_mes and metrica_mes.total_activos > 0 else 1
    total_cursos = len(eventos_aprobados)
    horas_promedio = sum(float(e.total_horas or 0) for e in eventos_aprobados) / total_cursos if total_cursos > 0 else 0

    query = db.query(Asistencia.colaborador_cedula, Evento.total_horas, Colaborador.genero, Colaborador.unidad, Evento.localidad, Colaborador.grupo_personal, Evento.modalidad, Evento.dimension_evento).join(Evento, Asistencia.evento_id == Evento.id).join(Colaborador, Asistencia.colaborador_cedula == Colaborador.cedula).filter(Evento.id.in_(ids_eventos))

    if estado != "TODOS":
        query = query.filter(Colaborador.estado_laboral == estado)

    asistencias = query.all()

    cedulas_unicas = set()
    total_horas = 0
    mod_dic, gen_dic, uni_dic, loc_dic, dim_grp_dic = {}, {}, {}, {}, {}

    for a in asistencias:
        cedula = a.colaborador_cedula
        horas = float(a.total_horas or 0)
        gen = a.genero or "NO DEFINIDO"
        mod = a.modalidad or "NO DEFINIDA"
        uni = a.unidad or "NO DEFINIDA"
        loc = a.localidad or "NO DEFINIDA"
        grp = a.grupo_personal or "NO DEFINIDO"
        dim = a.dimension_evento or "NO DEFINIDA"

        cedulas_unicas.add(cedula)
        total_horas += horas
        mod_dic[mod] = mod_dic.get(mod, 0) + horas
        gen_dic[gen] = gen_dic.get(gen, 0) + horas
        uni_dic[uni] = uni_dic.get(uni, 0) + horas
        loc_dic[loc] = loc_dic.get(loc, 0) + horas
        if dim not in dim_grp_dic: dim_grp_dic[dim] = {}
        dim_grp_dic[dim][grp] = dim_grp_dic[dim].get(grp, 0) + horas

    total_colaboradores = len(cedulas_unicas)
    porcentaje_capacitado = (total_colaboradores / total_empresa) * 100
    formatear_dic = lambda d: [{"name": k, "value": round(v, 1)} for k, v in d.items()]
    
    dim_grp_lista = []
    for d_name, grp_data in dim_grp_dic.items():
        obj = {"dimension": d_name}
        for g_name, hrs in grp_data.items(): obj[g_name] = round(hrs, 1)
        dim_grp_lista.append(obj)

    return {
        "kpis": {"total_colaboradores": total_colaboradores, "total_horas": round(total_horas, 1), "horas_promedio": round(horas_promedio, 1), "total_cursos": total_cursos, "personal_capacitado_pct": round(porcentaje_capacitado, 2)},
        "tendencias": {"diferencia_horas": round(total_horas - prev_horas, 1), "diferencia_pct": round(porcentaje_capacitado - prev_pct, 2)},
        "graficos": {"modalidad": formatear_dic(mod_dic), "genero": formatear_dic(gen_dic), "unidad_negocio": formatear_dic(uni_dic), "localidad": formatear_dic(loc_dic), "dimension_grupo": dim_grp_lista}
    }