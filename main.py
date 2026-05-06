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

class NombreCurso(Base):
    __tablename__ = "nombres_cursos"
    id = Column(UUID(as_uuid=True), primary_key=True, server_default=sqlalchemy.text("gen_random_uuid()"))
    nombre = Column(String(200), unique=True, nullable=False)
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
    jefe_inmediato = Column(String(100))
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

class NuevoNombreCurso(BaseModel):
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
    try: yield db
    finally: db.close()

def parse_iso_date(date_str):
    if not date_str: return None
    try:
        dt = pd.to_datetime(date_str)
        return dt.replace(tzinfo=None).to_pydatetime() if not pd.isna(dt) else None
    except: return None

def create_access_token(data: dict, expires_delta: Optional[timedelta] = None):
    to_encode = data.copy()
    expire = datetime.utcnow() + (expires_delta if expires_delta else timedelta(minutes=15))
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)

@app.post("/token")
def login_for_access_token(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user = db.query(Usuario).filter(Usuario.email == form_data.username).first()
    if not user or not pwd_context.verify(form_data.password, user.password_hash):
        raise HTTPException(status_code=401, detail="Credenciales incorrectas")
    token = create_access_token(data={"sub": user.email}, expires_delta=timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES))
    return {"access_token": token, "token_type": "bearer", "user_name": user.nombre_completo, "user_role": user.rol, "user_location": user.localidad}

@app.get("/nombres-cursos")
def listar_nombres_cursos(db: Session = Depends(get_db), current_user: Usuario = Depends(get_current_user)):
    return [{"id": str(n.id), "nombre": n.nombre} for n in db.query(NombreCurso).order_by(NombreCurso.nombre).all()]

@app.post("/nombres-cursos")
def crear_nombre_curso(data: NuevoNombreCurso, db: Session = Depends(get_db), current_admin: Usuario = Depends(get_current_admin)):
    existe = db.query(NombreCurso).filter(NombreCurso.nombre.ilike(data.nombre)).first()
    if existe: raise HTTPException(status_code=400, detail="Ya existe")
    db.add(NombreCurso(nombre=data.nombre.upper().strip()))
    db.commit()
    return {"message": "ok"}

@app.delete("/nombres-cursos/{id}")
def eliminar_nombre_curso(id: str, db: Session = Depends(get_db), current_admin: Usuario = Depends(get_current_admin)):
    obj = db.query(NombreCurso).filter(NombreCurso.id == id).first()
    if obj: db.delete(obj); db.commit()
    return {"message": "ok"}

@app.get("/empresas")
def listar_empresas(db: Session = Depends(get_db), current_user: Usuario = Depends(get_current_user)):
    return [{"id": str(e.id), "nombre": e.nombre} for e in db.query(EmpresaCapacitadora).order_by(EmpresaCapacitadora.nombre).all()]

@app.post("/empresas")
def crear_empresa(data: NuevaEmpresa, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    db.add(EmpresaCapacitadora(nombre=data.nombre.upper().strip()))
    db.commit()
    return {"status": "ok"}

@app.delete("/empresas/{empresa_id}")
def eliminar_empresa(empresa_id: str, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    emp = db.query(EmpresaCapacitadora).filter(EmpresaCapacitadora.id == empresa_id).first()
    if emp: db.delete(emp); db.commit()
    return {"message": "ok"}

@app.get("/dashboard/metricas")
def obtener_metricas(mes: str, vista: str = "MENSUAL", estado: str = "TODOS", db: Session = Depends(get_db), current_admin: Usuario = Depends(get_current_admin)):
    try:
        def get_period_data(p_mes, p_vista):
            try: y = int(p_mes.split('-')[0])
            except: y = datetime.utcnow().year
            q = db.query(Evento.total_horas, Evento.nombre_curso, Evento.modalidad, Evento.dimension_evento, Colaborador.genero, Colaborador.unidad, Colaborador.localidad, Colaborador.grupo_personal, Colaborador.cedula).join(Asistencia, Asistencia.evento_id == Evento.id).join(Colaborador, Colaborador.cedula == Asistencia.colaborador_cedula).filter(Evento.estado == "APROBADO")
            if p_vista == "ANUAL": q = q.filter(Evento.mes_anio.like(f"{y}%"))
            else: q = q.filter(Evento.mes_anio == p_mes)
            if estado != "TODOS": q = q.filter(Colaborador.estado_laboral == estado)
            rows = q.all()
            act = db.query(func.avg(MetricaMensual.total_activos)).filter(MetricaMensual.mes_anio == p_mes).scalar() or 1
            h = sum(float(r.total_horas or 0) for r in rows)
            c = len(set(r.cedula for r in rows))
            p = round((c/act)*100, 1) if act > 0 else 0
            return rows, h, c, p

        cur_rows, cur_h, cur_c, cur_p = get_period_data(mes, vista)
        try:
            parts = mes.split('-')
            y = int(parts[0])
            m = int(parts[1]) if len(parts) > 1 else 1
            if vista == "ANUAL": prev_mes = f"{y-1}-01"
            else: prev_mes = f"{y-1}-12" if m == 1 else f"{y}-{m-1:02d}"
        except: prev_mes = mes

        _, pre_h, _, pre_p = get_period_data(prev_mes, vista)

        mod_dic, gen_dic, uni_dic, loc_dic, d_g, n_u = {}, {}, {}, {}, {}, set()
        for r in cur_rows:
            hrs = float(r.total_horas or 0); n_u.add(r.nombre_curso)
            mod_dic[r.modalidad or "N/A"] = mod_dic.get(r.modalidad or "N/A", 0) + hrs
            gen_dic[r.genero or "N/A"] = gen_dic.get(r.genero or "N/A", 0) + hrs
            uni_dic[r.unidad or "N/A"] = uni_dic.get(r.unidad or "N/A", 0) + hrs
            loc_dic[r.localidad or "N/A"] = loc_dic.get(r.localidad or "N/A", 0) + hrs
            d, gp = (r.dimension_evento or "Otros"), (r.grupo_personal or "N/A")
            if d not in d_g: d_g[d] = {}
            d_g[d][gp] = d_g[d].get(gp, 0) + hrs

        return {
            "kpis": {"total_colaboradores": cur_c, "total_horas": round(cur_h, 1), "horas_promedio": round(cur_h/len(cur_rows), 1) if cur_rows else 0, "total_cursos": len(n_u), "personal_capacitado_pct": cur_p},
            "tendencias": {"diferencia_horas": round(cur_h - pre_h, 1), "diferencia_pct": round(cur_p - pre_p, 1)},
            "graficos": {"modalidad": [{"name": k, "value": v} for k, v in mod_dic.items()], "genero": [{"name": k, "value": v} for k, v in gen_dic.items()], "unidad_negocio": [{"name": k, "value": v} for k, v in uni_dic.items()], "localidad": [{"name": k, "value": v} for k, v in loc_dic.items()], "dimension_grupo": [{"dimension": d, **grps} for d, grps in d_g.items()]}
        }
    except Exception as e: raise HTTPException(status_code=400, detail=str(e))

@app.get("/admin/eventos/{id}/exportar")
def exportar_evento_individual(id: str, db: Session = Depends(get_db), current_user: Usuario = Depends(get_current_user)):
    try:
        evento = db.query(Evento).filter(Evento.id == id).first()
        if not evento: raise HTTPException(status_code=404)
        query = db.query(
            Evento.codigo_curso.label("CÓDIGO"), Evento.nombre_curso.label("NOMBRE DEL CURSO"), Evento.objetivo.label("OBJETIVO"),
            Evento.empresa.label("EMPRESA CAPACITADORA"), Evento.facilitador.label("FACILITADOR"), Evento.dimension_evento.label("DIMENSIÓN DE EVENTO"),
            Evento.lugar.label("LUGAR DONDE SE DIO LA CAPACITACION"), Evento.modalidad.label("MODALIDAD"),
            Evento.fecha_hora_inicio.label("FECHA INICIO"), Evento.fecha_hora_cierre.label("FECHA CIERRE"),
            Evento.total_horas.label("DURACION DE LA CAPACITACION (HORAS)"), Evento.tipo_evento.label("TIPO EVENTO"),
            Evento.mes_anio.label("MES-AÑO"), Colaborador.cedula.label("CÉDULA"),
            (Colaborador.apellidos + " " + Colaborador.nombres).label("APELLIDOS Y NOMBRE DEL COLABORADOR"),
            Colaborador.genero.label("GÉNERO"), Colaborador.cargo.label("CARGO"), Colaborador.unidad.label("UNIDAD"),
            Colaborador.area.label("ÁREA"), Colaborador.seccion.label("SECCIÓN"), Colaborador.centro_costo.label("CENTRO DE COSTO"),
            Colaborador.grupo_personal.label("GRUPO DE PERSONAL"), Colaborador.area_personal.label("ÁREA DE PERSONAL"),
            Colaborador.jefe_inmediato.label("JEFE DE ÁREA"), Colaborador.gerente_area.label("GERENTE DE AREA"), Colaborador.localidad.label("LOCALIDAD")
        ).join(Asistencia, Asistencia.evento_id == Evento.id).join(Colaborador, Colaborador.cedula == Asistencia.colaborador_cedula).filter(Evento.id == id)
        df = pd.read_sql(query.statement, engine)
        for col in ["FECHA INICIO", "FECHA CIERRE"]: df[col] = pd.to_datetime(df[col]).dt.strftime('%d/%m/%Y').fillna('')
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='Reporte')
        output.seek(0)
        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=auditoria_{evento.codigo_curso}.xlsx"})
    except Exception as e: raise HTTPException(status_code=500, detail=str(e))

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
            Colaborador.jefe_inmediato.label("JEFE DE ÁREA"), Colaborador.gerente_area.label("GERENTE DE AREA"), Colaborador.localidad.label("LOCALIDAD")
        ).join(Asistencia, Asistencia.evento_id == Evento.id).join(Colaborador, Colaborador.cedula == Asistencia.colaborador_cedula).filter(Evento.estado == "APROBADO")
        if vista == "ANUAL": query = query.filter(Evento.mes_anio.like(f"{mes.split('-')[0]}%"))
        else: query = query.filter(Evento.mes_anio == mes)
        if estado != "TODOS": query = query.filter(Colaborador.estado_laboral == estado)
        df = pd.read_sql(query.statement, engine)
        for col in ["FECHA INICIO", "FECHA CIERRE"]: df[col] = pd.to_datetime(df[col]).dt.strftime('%d/%m/%Y')
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='Reporte')
        output.seek(0)
        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=reporte_dashboard.xlsx"})
    except Exception as e: raise HTTPException(status_code=500, detail=str(e))

@app.get("/mis-eventos")
def mis_eventos(current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    evs = db.query(Evento).filter(Evento.creado_por_usuario_id == current_user.id).order_by(Evento.fecha_creacion.desc()).all()
    res = []
    for e in evs:
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

@app.put("/admin/eventos/{id}/revertir")
def revertir_aprobacion(id: str, db: Session = Depends(get_db), current_admin: Usuario = Depends(get_current_admin)):
    evento = db.query(Evento).filter(Evento.id == id).first()
    if not evento: raise HTTPException(status_code=404)
    evento.estado = "PENDIENTE"
    db.add(HistorialEvento(evento_id=evento.id, usuario_id=current_admin.id, accion="REVERTIDO A PENDIENTE"))
    db.commit()
    return {"status": "ok"}

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
