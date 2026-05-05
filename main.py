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
from sqlalchemy.orm import sessionmaker, Session, declarative_base
import sqlalchemy
from jose import JWTError, jwt
from passlib.context import CryptContext
import pandas as pd

DATABASE_URL = os.getenv("DATABASE_URL", "postgresql://neondb_owner:npg_cz5oNRL7Snwj@ep-wild-tooth-an30z9iu.c-6.us-east-1.aws.neon.tech/neondb?sslmode=require")

engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

SECRET_KEY = os.getenv("SECRET_KEY", "CLAVE_ULTRA_SECRETA_TRAINFORM")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 480 

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token")

# --- MODELOS ---

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
    allow_origins=["https://desarrollonovacero.github.io", "http://localhost:5173", "http://127.0.0.1:5173"],
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
        return dt.replace(tzinfo=None).to_pydatetime() if not pd.isna(dt) else None
    except: return None

def verify_password(plain_password, hashed_password):
    return pwd_context.verify(plain_password, hashed_password)

def create_access_token(data: dict, expires_delta: Optional[timedelta] = None):
    to_encode = data.copy()
    expire = datetime.utcnow() + (expires_delta if expires_delta else timedelta(minutes=15))
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)

def get_current_user(token: str = Depends(oauth2_scheme), db: Session = Depends(get_db)):
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        email = payload.get("sub")
        if email is None: raise HTTPException(status_code=401)
    except JWTError: raise HTTPException(status_code=401)
    user = db.query(Usuario).filter(Usuario.email == email).first()
    if not user or not user.activo: raise HTTPException(status_code=401)
    return user

def get_current_admin(current_user: Usuario = Depends(get_current_user)):
    if current_user.rol != "ADMIN": raise HTTPException(status_code=403)
    return current_user

@app.post("/token")
def login_for_access_token(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user = db.query(Usuario).filter(Usuario.email == form_data.username).first()
    if not user or not verify_password(form_data.password, user.password_hash):
        raise HTTPException(status_code=401, detail="Credenciales incorrectas")
    token = create_access_token(data={"sub": user.email}, expires_delta=timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES))
    return {"access_token": token, "token_type": "bearer", "user_name": user.nombre_completo, "user_role": user.rol, "user_location": user.localidad}

@app.post("/crear-usuario")
def crear_usuario(data: NuevoUsuario, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    if db.query(Usuario).filter(Usuario.email == data.email).first(): raise HTTPException(status_code=400, detail="Ya existe")
    db.add(Usuario(email=data.email, password_hash=pwd_context.hash(data.password), nombre_completo=data.nombre_completo, rol=data.rol, localidad=data.localidad))
    db.commit()
    return {"status": "ok"}

@app.get("/usuarios")
def listar_usuarios(db: Session = Depends(get_db), current_admin: Usuario = Depends(get_current_admin)):
    users = db.query(Usuario).all()
    return [{"id": str(u.id), "email": u.email, "nombre_completo": u.nombre_completo, "rol": u.rol, "activo": u.activo, "localidad": u.localidad} for u in users]

@app.put("/usuarios/{user_id}/toggle")
def toggle_usuario(user_id: str, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
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
    usuario = db.query(Usuario).filter(Usuario.id == user_id).first()
    if not usuario: raise HTTPException(status_code=404)
    db.delete(usuario)
    db.commit()
    return {"message": "Usuario eliminado"}

@app.get("/empresas")
def listar_empresas(db: Session = Depends(get_db), current_user: Usuario = Depends(get_current_user)):
    return [{"id": str(e.id), "nombre": e.nombre} for e in db.query(EmpresaCapacitadora).order_by(EmpresaCapacitadora.nombre).all()]

@app.post("/empresas")
def crear_empresa(data: NuevaEmpresa, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
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
    c = db.query(Colaborador).count()
    return {"ready": c > 0, "count": c}

@app.get("/load-state")
def load_state(current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    sesion = db.query(SesionWeb).filter(SesionWeb.usuario_id == current_user.id).first()
    return json.loads(sesion.estado_borrador_json) if sesion and sesion.estado_borrador_json else None

@app.post("/save-state")
def save_state(payload: dict = Body(...), current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    sesion = db.query(SesionWeb).filter(SesionWeb.usuario_id == current_user.id).first()
    if not sesion:
        sesion = SesionWeb(usuario_id=current_user.id, token_sesion=f"draft_{current_user.id}", estado_borrador_json=json.dumps(payload), expira_en=datetime.utcnow()+timedelta(days=7))
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

    col_cedula = find_col(["CEDULA", "IDENTIFICACIÓN"])
    col_apellidos = find_col(["APELLIDOS"])
    col_nombres = find_col(["NOMBRES"])
    col_cargo = find_col(["CARGO"])
    col_genero = find_col(["GENERO", "SEXO"])
    col_unidad = find_col(["UNIDAD"])
    col_area = find_col(["AREA"])
    col_seccion = find_col(["SECCION"])
    col_cc = find_col(["CENTRO DE COSTO"])
    col_gp = find_col(["GRUPO DE PERSONAL"])
    col_ap = find_col(["AREA DE PERSONAL"])
    col_ja = find_col(["JEFE INMEDIATO", "JEFE"])
    col_ga = find_col(["GERENTE"])
    col_loc = find_col(["LOCALIDAD", "LOCACION"])
    col_desv = find_col(["DESVINCULACION", "EGRESO"])

    conteo_activos = 0
    for index, row in df.iterrows():
        cedula = str(row[col_cedula]).strip() if col_cedula else ""
        if not cedula or cedula.lower() == "nan": continue
        
        colaborador = db.query(Colaborador).filter(Colaborador.cedula == cedula).first()
        if not colaborador:
            colaborador = Colaborador(cedula=cedula)
            db.add(colaborador)

        val_desv = str(row[col_desv]).strip().upper() if col_desv else ""
        estado_final = "CESADO" if val_desv and val_desv != "NAN" else "ACTIVO"
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
    if metrica: metrica.total_activos = conteo_activos
    else: db.add(MetricaMensual(mes_anio=mes_corte, total_activos=conteo_activos))
    db.commit()
    return {"message": "ok"}

@app.post("/suggest-cedulas")
def suggest_cedulas(search_term: str = Form(...), db: Session = Depends(get_db), current_user: Usuario = Depends(get_current_user)):
    search = f"%{search_term.strip()}%"
    res = db.query(Colaborador).filter(or_(Colaborador.cedula.ilike(search), Colaborador.nombres.ilike(search), Colaborador.apellidos.ilike(search))).limit(10).all()
    return [{"cedula": r.cedula, "nombre": f"{r.apellidos} {r.nombres}".strip(), "source": "headcount" if r.estado_laboral=="ACTIVO" else "cesantes"} for r in res]

@app.post("/validate-cedula")
def validate_cedula(cedulas_json: str = Form(...), db: Session = Depends(get_db), current_user: Usuario = Depends(get_current_user)):
    cedulas = json.loads(cedulas_json)
    resultados = []
    for c in cedulas:
        colab = db.query(Colaborador).filter(Colaborador.cedula == str(c).strip()).first()
        if colab:
            resultados.append({"cedula": colab.cedula, "found": True, "source": "headcount" if colab.estado_laboral=="ACTIVO" else "cesantes", "data": {"nombres": colab.nombres, "apellidos": colab.apellidos, "cargo": colab.cargo, "unidad": colab.unidad, "area": colab.area, "localidad": colab.localidad, "genero": colab.genero, "centro_costo": colab.centro_costo, "grupo_personal": colab.grupo_personal, "area_personal": colab.area_personal, "jefe_area": colab.jefe_area, "gerente_area": colab.gerente_area}})
        else: resultados.append({"cedula": str(c), "found": False})
    return resultados

@app.post("/enviar-revision")
def enviar_revision(payload: dict = Body(...), current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    try:
        event_data = payload.get("eventData", {})
        registros = payload.get("registros", [])
        evento_id = payload.get("eventoId")
        if not registros: raise HTTPException(status_code=400, detail="No hay asistentes")
        try: horas = float(event_data.get("totalHoras", 0))
        except: horas = 0.0
        inicio_dt = parse_iso_date(event_data.get("fechaHoraInicio"))
        cierre_dt = parse_iso_date(event_data.get("fechaHoraCierre"))

        if evento_id:
            evento = db.query(Evento).filter(Evento.id == evento_id).first()
            if not evento: raise HTTPException(status_code=404)
            db.query(Asistencia).filter(Asistencia.evento_id == evento_id).delete()
        else:
            anio_act = datetime.utcnow().year
            prefijo = f"NOV-{anio_act}-"
            ultimo = db.query(Evento).filter(Evento.codigo_curso.like(f"{prefijo}%")).order_by(Evento.codigo_curso.desc()).first()
            nuevo_num = (int(ultimo.codigo_curso.split('-')[-1]) + 1) if ultimo else 1
            evento = Evento(codigo_curso=f"{prefijo}{str(nuevo_num).zfill(4)}", creado_por_usuario_id=current_user.id)
            db.add(evento)

        evento.nombre_curso, evento.objetivo, evento.empresa, evento.facilitador = event_data.get("nombreCurso"), event_data.get("objetivo"), event_data.get("empresa"), event_data.get("facilitador")
        evento.dimension_evento, evento.lugar, evento.modalidad, evento.fecha_hora_inicio = event_data.get("dimensionEvento"), event_data.get("lugar"), event_data.get("modalidad"), inicio_dt
        evento.fecha_hora_cierre, evento.total_horas, evento.tipo_evento, evento.mes_anio = cierre_dt, horas, event_data.get("tipoEvento"), event_data.get("mesAnio")
        evento.estado = "PENDIENTE"
        db.flush()

        for r in registros:
            cedula_raw = str(r.get("CÉDULA") or r.get("cedula") or "").strip()
            if not cedula_raw: continue
            colab = db.query(Colaborador).filter(Colaborador.cedula == cedula_raw).first()
            if not colab:
                colab = Colaborador(cedula=cedula_raw, nombres=str(r.get("APELLIDOS Y NOMBRE DEL COLABORADOR", "REGISTRO MANUAL")), origen="auto", estado_laboral="ACTIVO")
                db.add(colab)
                db.flush()
            db.add(Asistencia(evento_id=evento.id, colaborador_cedula=colab.cedula, estado_validacion="VALIDADO"))

        db.add(HistorialEvento(evento_id=evento.id, usuario_id=current_user.id, accion="ENVIADO A REVISION"))
        db.commit()
        return {"message": "ok", "evento_id": str(evento.id)}
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/mis-eventos")
def mis_eventos(current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    eventos = db.query(Evento).filter(Evento.creado_por_usuario_id == current_user.id).order_by(Evento.fecha_creacion.desc()).all()
    res = []
    for e in eventos:
        hist = db.query(HistorialEvento).filter(HistorialEvento.evento_id == e.id).order_by(HistorialEvento.fecha_registro.desc()).first()
        res.append({"id": str(e.id), "codigo": e.codigo_curso, "nombre": e.nombre_curso, "estado": e.estado, "fecha": e.fecha_creacion, "comentario": hist.comentario if hist else ""})
    return res

@app.get("/admin/eventos")
def admin_eventos(db: Session = Depends(get_db), current_admin: Usuario = Depends(get_current_admin)):
    evs = db.query(Evento).order_by(Evento.fecha_creacion.desc()).all()
    res = []
    for e in evs:
        u = db.query(Usuario).filter(Usuario.id == e.creado_por_usuario_id).first()
        res.append({"id": str(e.id), "codigo": e.codigo_curso, "nombre": e.nombre_curso, "estado": e.estado, "creador": u.nombre_completo if u else "N/A", "fecha": e.fecha_creacion})
    return res

@app.put("/admin/eventos/{id}/aprobar")
def aprobar_evento(id: str, db: Session = Depends(get_db), current_admin: Usuario = Depends(get_current_admin)):
    ev = db.query(Evento).filter(Evento.id == id).first()
    ev.estado = "APROBADO"
    db.add(HistorialEvento(evento_id=ev.id, usuario_id=current_admin.id, accion="APROBADO"))
    db.commit()
    return {"status": "ok"}

@app.put("/admin/eventos/{id}/rechazar")
def rechazar_evento(id: str, accion: AuditoriaAccion, db: Session = Depends(get_db), current_admin: Usuario = Depends(get_current_admin)):
    ev = db.query(Evento).filter(Evento.id == id).first()
    ev.estado = "RECHAZADO"
    db.add(HistorialEvento(evento_id=ev.id, usuario_id=current_admin.id, accion="RECHAZADO", comentario=accion.comentario))
    db.commit()
    return {"status": "ok"}

@app.get("/admin/eventos/{id}/exportar")
def exportar_evento(id: str, db: Session = Depends(get_db), current_admin: Usuario = Depends(get_current_admin)):
    try:
        query = db.query(
            Evento.nombre_curso.label("NOMBRE DEL CURSO"), Evento.objetivo.label("OBJETIVO"), Evento.empresa.label("EMPRESA CAPACITADORA"),
            Evento.facilitador.label("FACILITADOR"), Evento.dimension_evento.label("DIMENSIÓN DE EVENTO"), Evento.lugar.label("LUGAR DONDE SE DIO LA CAPACITACION"),
            Evento.modalidad.label("MODALIDAD"), Evento.fecha_hora_inicio.label("FECHA INICIO"), Evento.fecha_hora_cierre.label("FECHA CIERRE"),
            Evento.total_horas.label("DURACION DE LA CAPACITACION (HORAS)"), Evento.tipo_evento.label("TIPO EVENTO"), Evento.mes_anio.label("MES-AÑO"),
            Colaborador.cedula.label("CÉDULA"), (Colaborador.apellidos + " " + Colaborador.nombres).label("APELLIDOS Y NOMBRE DEL COLABORADOR"),
            Colaborador.genero.label("GÉNERO"), Colaborador.cargo.label("CARGO"), Colaborador.unidad.label("UNIDAD"),
            Colaborador.area.label("ÁREA"), Colaborador.seccion.label("SECCIÓN"), Colaborador.centro_costo.label("CENTRO DE COSTO"),
            Colaborador.grupo_personal.label("GRUPO DE PERSONAL"), Colaborador.area_personal.label("ÁREA DE PERSONAL"),
            Colaborador.jefe_area.label("JEFE DE ÁREA"), Colaborador.gerente_area.label("GERENTE DE AREA"), Colaborador.localidad.label("LOCALIDAD")
        ).join(Asistencia, Asistencia.evento_id == Evento.id).join(Colaborador, Colaborador.cedula == Asistencia.colaborador_cedula).filter(Evento.id == id)

        df = pd.read_sql(query.statement, engine)
        for col in ["FECHA INICIO", "FECHA CIERRE"]: df[col] = pd.to_datetime(df[col]).dt.strftime('%d/%m/%Y')
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='Reporte')
        output.seek(0)
        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=auditoria_evento.xlsx"})
    except Exception as e: raise HTTPException(status_code=500, detail=str(e))

@app.get("/dashboard/metricas")
def obtener_metricas(mes: str, vista: str = "MENSUAL", estado: str = "TODOS", db: Session = Depends(get_db), current_admin: Usuario = Depends(get_current_admin)):
    try:
        y = int(mes.split('-')[0])
        query = db.query(Evento.total_horas, Evento.nombre_curso, Evento.modalidad, Evento.dimension_evento, Colaborador.genero, Colaborador.unidad, Colaborador.localidad, Colaborador.grupo_personal, Colaborador.cedula).join(Asistencia, Asistencia.evento_id == Evento.id).join(Colaborador, Colaborador.cedula == Asistencia.colaborador_cedula).filter(Evento.estado == "APROBADO")
        if vista == "ANUAL": query = query.filter(Evento.mes_anio.like(f"{y}%"))
        else: query = query.filter(Evento.mes_anio == mes)
        if estado != "TODOS": query = query.filter(Colaborador.estado_laboral == estado)
        data = query.all()
        total_activos = db.query(func.avg(MetricaMensual.total_activos)).filter(MetricaMensual.mes_anio == mes).scalar() or 1
        nombres_unicos, ceds_unicas, total_h = set(), set(), 0.0
        mod_dic, gen_dic, uni_dic, loc_dic, dim_grp_dic = {}, {}, {}, {}, {}
        for r in data:
            hrs = float(r.total_horas or 0)
            total_h += hrs
            nombres_unicos.add(r.nombre_curso)
            ceds_unicas.add(r.cedula)
            mod_dic[r.modalidad or "N/A"] = mod_dic.get(r.modalidad or "N/A", 0) + hrs
            gen_dic[r.genero or "N/A"] = gen_dic.get(r.genero or "N/A", 0) + hrs
            uni_dic[r.unidad or "N/A"] = uni_dic.get(r.unidad or "N/A", 0) + hrs
            loc_dic[r.localidad or "N/A"] = loc_dic.get(r.localidad or "N/A", 0) + hrs
            d, gp = (r.dimension_evento or "Otros"), (r.grupo_personal or "N/A")
            if d not in dim_grp_dic: dim_grp_dic[d] = {}
            dim_grp_dic[d][gp] = dim_grp_dic[d].get(gp, 0) + hrs
        return {"kpis": {"total_colaboradores": len(ceds_unicas), "total_horas": round(total_h, 1), "horas_promedio": round(total_h/len(data), 1) if data else 0, "total_cursos": len(nombres_unicos), "personal_capacitado_pct": round((len(ceds_unicas)/total_activos)*100, 1)}, "tendencias": {"diferencia_horas": 0, "diferencia_pct": 0}, "graficos": {"modalidad": [{"name": k, "value": v} for k, v in mod_dic.items()], "genero": [{"name": k, "value": v} for k, v in gen_dic.items()], "unidad_negocio": [{"name": k, "value": v} for k, v in uni_dic.items()], "localidad": [{"name": k, "value": v} for k, v in loc_dic.items()], "dimension_grupo": [{"dimension": d, **grps} for d, grps in dim_grp_dic.items()]}}
    except Exception as e: raise HTTPException(status_code=400, detail=str(e))

@app.get("/dashboard/exportar")
def exportar_dashboard(mes: str, vista: str = "MENSUAL", estado: str = "TODOS", db: Session = Depends(get_db), current_user: Usuario = Depends(get_current_user)):
    try:
        query = db.query(
            Evento.nombre_curso.label("NOMBRE DEL CURSO"), Evento.objetivo.label("OBJETIVO"), Evento.empresa.label("EMPRESA CAPACITADORA"),
            Evento.facilitador.label("FACILITADOR"), Evento.dimension_evento.label("DIMENSIÓN DE EVENTO"), Evento.lugar.label("LUGAR DONDE SE DIO LA CAPACITACION"),
            Evento.modalidad.label("MODALIDAD"), Evento.fecha_hora_inicio.label("FECHA INICIO"), Evento.fecha_hora_cierre.label("FECHA CIERRE"),
            Evento.total_horas.label("DURACION DE LA CAPACITACION (HORAS)"), Evento.tipo_evento.label("TIPO EVENTO"), Evento.mes_anio.label("MES-AÑO"),
            Colaborador.cedula.label("CÉDULA"), (Colaborador.apellidos + " " + Colaborador.nombres).label("APELLIDOS Y NOMBRE DEL COLABORADOR"),
            Colaborador.genero.label("GÉNERO"), Colaborador.cargo.label("CARGO"), Colaborador.unidad.label("UNIDAD"),
            Colaborador.area.label("ÁREA"), Colaborador.seccion.label("SECCIÓN"), Colaborador.centro_costo.label("CENTRO DE COSTO"),
            Colaborador.grupo_personal.label("GRUPO DE PERSONAL"), Colaborador.area_personal.label("ÁREA DE PERSONAL"),
            Colaborador.jefe_area.label("JEFE DE ÁREA"), Colaborador.gerente_area.label("GERENTE DE AREA"), Colaborador.localidad.label("LOCALIDAD")
        ).join(Asistencia, Asistencia.evento_id == Evento.id).join(Colaborador, Colaborador.cedula == Asistencia.colaborador_cedula).filter(Evento.estado == "APROBADO")

        if vista == "ANUAL": query = query.filter(Evento.mes_anio.like(f"{mes.split('-')[0]}%"))
        else: query = query.filter(Evento.mes_anio == mes)
        if estado != "TODOS": query = query.filter(Colaborador.estado_laboral == estado)
        
        df = pd.read_sql(query.statement, engine)
        for col in ["FECHA INICIO", "FECHA CIERRE"]: df[col] = pd.to_datetime(df[col]).dt.strftime('%d/%m/%Y')
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='Dashboard_Report')
        output.seek(0)
        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=dashboard_completo.xlsx"})
    except Exception as e: raise HTTPException(status_code=500, detail=str(e))