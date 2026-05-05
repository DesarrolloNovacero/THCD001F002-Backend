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
    localidad = Column(String(50)) # Se mantiene por compatibilidad, pero no se usa en métricas
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

# --- SCHEMAS Pydantic ---
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

# --- ENDPOINTS ---

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

@app.get("/empresas")
def listar_empresas(db: Session = Depends(get_db), current_user: Usuario = Depends(get_current_user)):
    return [{"id": str(e.id), "nombre": e.nombre} for e in db.query(EmpresaCapacitadora).order_by(EmpresaCapacitadora.nombre).all()]

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
            # Lógica inteligente para el código del curso
            ultimo = db.query(Evento).filter(Evento.codigo_curso.like(f"{prefijo}%")).order_by(Evento.codigo_curso.desc()).first()
            nuevo_num = (int(ultimo.codigo_curso.split('-')[-1]) + 1) if ultimo else 1
            evento = Evento(codigo_curso=f"{prefijo}{str(nuevo_num).zfill(4)}", creado_por_usuario_id=current_user.id)
            db.add(evento)

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
        evento.estado = "PENDIENTE"
        
        db.flush()

        for r in registros:
            cedula_raw = str(r.get("CÉDULA") or r.get("cedula") or "").strip()
            if not cedula_raw: continue
            
            # Verificamos si existe el colaborador para no romper la DB
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

@app.get("/dashboard/metricas")
def obtener_metricas(mes: str, vista: str = "MENSUAL", estado: str = "TODOS", db: Session = Depends(get_db), current_admin: Usuario = Depends(get_current_admin)):
    try:
        y = int(mes.split('-')[0])
        # Consulta base uniendo Asistencia con Colaborador para sacar la localidad automática
        query = db.query(
            Evento.total_horas, Evento.nombre_curso, Evento.modalidad, Evento.dimension_evento,
            Colaborador.genero, Colaborador.unidad, Colaborador.localidad, Colaborador.grupo_personal, Colaborador.cedula
        ).join(Asistencia, Asistencia.evento_id == Evento.id).join(Colaborador, Colaborador.cedula == Asistencia.colaborador_cedula).filter(Evento.estado == "APROBADO")

        if vista == "ANUAL": query = query.filter(Evento.mes_anio.like(f"{y}%"))
        else: query = query.filter(Evento.mes_anio == mes)
        
        if estado != "TODOS": query = query.filter(Colaborador.estado_laboral == estado)

        data = query.all()
        
        # Snapshot de activos para el KPI de porcentaje
        total_activos = db.query(func.avg(MetricaMensual.total_activos)).filter(MetricaMensual.mes_anio == mes).scalar() or 1
        
        nombres_unicos = set()
        ceds_unicas = set()
        total_h = 0.0
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

        return {
            "kpis": {"total_colaboradores": len(ceds_unicas), "total_horas": round(total_h, 1), "horas_promedio": round(total_h/len(data), 1) if data else 0, "total_cursos": len(nombres_unicos), "personal_capacitado_pct": round((len(ceds_unicas)/total_activos)*100, 1)},
            "tendencias": {"diferencia_horas": 0, "diferencia_pct": 0}, # Simplificado
            "graficos": {
                "modalidad": [{"name": k, "value": v} for k, v in mod_dic.items()],
                "genero": [{"name": k, "value": v} for k, v in gen_dic.items()],
                "unidad_negocio": [{"name": k, "value": v} for k, v in uni_dic.items()],
                "localidad": [{"name": k, "value": v} for k, v in loc_dic.items()],
                "dimension_grupo": [{"dimension": d, **grps} for d, grps in dim_grp_dic.items()]
            }
        }
    except Exception as e: raise HTTPException(status_code=400, detail=str(e))

@app.get("/dashboard/exportar")
def exportar_dashboard(mes: str, vista: str = "MENSUAL", estado: str = "TODOS", db: Session = Depends(get_db), current_user: Usuario = Depends(get_current_user)):
    try:
        query = db.query(
            Evento.nombre_curso.label("CURSO"), Evento.total_horas.label("HORAS"), 
            Colaborador.cedula.label("CÉDULA"), (Colaborador.apellidos + " " + Colaborador.nombres).label("COLABORADOR"),
            Colaborador.localidad.label("LOCALIDAD"), Colaborador.unidad.label("UNIDAD"), 
            Colaborador.cargo.label("CARGO"), Evento.mes_anio.label("MES")
        ).join(Asistencia, Asistencia.evento_id == Evento.id).join(Colaborador, Colaborador.cedula == Asistencia.colaborador_cedula).filter(Evento.estado == "APROBADO")

        if vista == "ANUAL": query = query.filter(Evento.mes_anio.like(f"{mes.split('-')[0]}%"))
        else: query = query.filter(Evento.mes_anio == mes)
        
        df = pd.read_sql(query.statement, engine)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False)
        output.seek(0)
        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=reporte_{mes}.xlsx"})
    except Exception as e: raise HTTPException(status_code=500, detail=str(e))