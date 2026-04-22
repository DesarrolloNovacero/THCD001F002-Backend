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
from sqlalchemy import create_engine, Column, String, Boolean, DateTime, Text, Numeric, ForeignKey, or_
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
    activo = Column(Boolean, nullable=False, default=True)
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
    ultima_actualizacion = Column(DateTime, server_default=sqlalchemy.text("CURRENT_TIMESTAMP"))

class Evento(Base):
    __tablename__ = "eventos"
    id = Column(UUID(as_uuid=True), primary_key=True, server_default=sqlalchemy.text("gen_random_uuid()"))
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

class SesionWeb(Base):
    __tablename__ = "sesiones_web"
    id = Column(UUID(as_uuid=True), primary_key=True, server_default=sqlalchemy.text("gen_random_uuid()"))
    usuario_id = Column(UUID(as_uuid=True), ForeignKey("usuarios.id"), nullable=False)
    token_sesion = Column(String(255), nullable=False, unique=True)
    estado_borrador_json = Column(Text)
    ultima_modificacion = Column(DateTime, server_default=sqlalchemy.text("CURRENT_TIMESTAMP"))
    expira_en = Column(DateTime, nullable=False)

class Asistencia(Base):
    __tablename__ = "asistencias"
    id = Column(UUID(as_uuid=True), primary_key=True, server_default=sqlalchemy.text("gen_random_uuid()"))
    evento_id = Column(UUID(as_uuid=True), ForeignKey("eventos.id"), nullable=False)
    colaborador_cedula = Column(String(20), ForeignKey("colaboradores.cedula"), nullable=False)
    estado_validacion = Column(String(50), nullable=False)
    fecha_registro = Column(DateTime, server_default=sqlalchemy.text("CURRENT_TIMESTAMP"))

class NuevoUsuario(BaseModel):
    email: str
    password: str
    nombre_completo: str
    rol: str

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

def remove_accents(text):
    if text is None: return ""
    if not isinstance(text, str): text = str(text)
    text = text.strip().upper()
    text = unicodedata.normalize('NFD', text)
    return "".join(c for c in text if unicodedata.category(c) != 'Mn')

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
        raise HTTPException(status_code=403, detail="No tienes permisos de administrador")
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
        "user_role": user.rol
    }

@app.post("/crear-usuario")
def crear_usuario(data: NuevoUsuario, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    usuario_existente = db.query(Usuario).filter(Usuario.email == data.email).first()
    if usuario_existente:
        raise HTTPException(status_code=400, detail="Este correo ya está registrado")
        
    hash_generado = pwd_context.hash(data.password)
    
    nuevo_usuario = Usuario(
        email=data.email,
        password_hash=hash_generado,
        nombre_completo=data.nombre_completo,
        rol=data.rol
    )
    db.add(nuevo_usuario)
    db.commit()
    
    return {"message": f"Usuario {data.nombre_completo} creado exitosamente"}

@app.get("/")
def read_root():
    return {"status": "API en linea, Segura y Conectada a Neon"}

@app.get("/check-db-status")
def check_db_status(db: Session = Depends(get_db)):
    count = db.query(Colaborador).count()
    return {"ready": count > 0, "count": count}

@app.get("/load-state")
def load_state(current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    sesion = db.query(SesionWeb).filter(SesionWeb.usuario_id == current_user.id).first()
    if sesion and sesion.estado_borrador_json:
        return json.loads(sesion.estado_borrador_json)
    return None

@app.post("/save-state")
def save_state(payload: dict = Body(...), current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    sesion = db.query(SesionWeb).filter(SesionWeb.usuario_id == current_user.id).first()
    if not sesion:
        sesion = SesionWeb(
            usuario_id=current_user.id,
            token_sesion=f"draft_{current_user.id}",
            estado_borrador_json=json.dumps(payload),
            expira_en=datetime.utcnow() + timedelta(days=7)
        )
        db.add(sesion)
    else:
        sesion.estado_borrador_json = json.dumps(payload)
        sesion.ultima_modificacion = datetime.utcnow()
    db.commit()
    return {"status": "ok"}

@app.post("/upload-masters")
async def upload_masters(file: UploadFile = File(...), source: str = Form(...), current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
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
        colaborador.origen = source
        
    db.commit()
    return {"message": "Datos maestros procesados y actualizados correctamente"}

@app.post("/validate-cedula")
def validate_cedula(cedulas_json: str = Form(...), current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    cedulas = json.loads(cedulas_json)
    resultados = []
    
    for cedula in cedulas:
        cedula_str = str(cedula).strip()
        colaborador = db.query(Colaborador).filter(Colaborador.cedula == cedula_str).first()
        
        if colaborador:
            data = {
                "apellidos": colaborador.apellidos,
                "nombres": colaborador.nombres,
                "cargo": colaborador.cargo,
                "genero": colaborador.genero,
                "unidad": colaborador.unidad,
                "area": colaborador.area,
                "seccion": colaborador.seccion,
                "centro_costo": colaborador.centro_costo,
                "grupo_personal": colaborador.grupo_personal,
                "area_personal": colaborador.area_personal,
                "jefe_area": colaborador.jefe_area,
                "gerente_area": colaborador.gerente_area,
                "localidad": colaborador.localidad
            }
            resultados.append({"cedula": cedula_str, "found": True, "source": colaborador.origen, "data": data})
        else:
            resultados.append({"cedula": cedula_str, "found": False, "source": None, "data": {}})
            
    return resultados

@app.post("/suggest-cedulas")
def suggest_cedulas(search_term: str = Form(...), current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    search_term = search_term.strip()
    if len(search_term) < 2: return []
        
    search_pattern = f"%{search_term}%"
    resultados = db.query(Colaborador).filter(
        or_(
            Colaborador.cedula.ilike(search_pattern),
            Colaborador.nombres.ilike(search_pattern),
            Colaborador.apellidos.ilike(search_pattern)
        )
    ).limit(10).all()
    
    suggestions = [{"cedula": r.cedula, "nombre": f"{r.apellidos} {r.nombres}".strip(), "source": r.origen} for r in resultados]
    return suggestions

@app.post("/export-excel")
async def export_excel(payload: dict = Body(...), current_user: Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    event_data = payload.get("eventData", {})
    registros = payload.get("registros", [])
    
    if not registros:
        raise HTTPException(status_code=400, detail="No hay registros para exportar")

    try:
        horas = float(event_data.get("totalHoras", 0)) if event_data.get("totalHoras") else 0.0
    except ValueError:
        horas = 0.0
        
    nuevo_evento = Evento(
        creado_por_usuario_id=current_user.id,
        nombre_curso=event_data.get("nombreCurso", "Sin Nombre"),
        objetivo=event_data.get("objetivo"),
        empresa=event_data.get("empresa"),
        facilitador=event_data.get("facilitador"),
        dimension_evento=event_data.get("dimensionEvento"),
        lugar=event_data.get("lugar"),
        modalidad=event_data.get("modalidad"),
        total_horas=horas,
        tipo_evento=event_data.get("tipoEvento"),
        mes_anio=event_data.get("mesAnio")
    )
    db.add(nuevo_evento)
    db.flush() 
    
    procesados = []
    for r in registros:
        cedula_colaborador = str(r.get("CÉDULA", "")).strip()
        if cedula_colaborador:
            nueva_asistencia = Asistencia(
                evento_id=nuevo_evento.id,
                colaborador_cedula=cedula_colaborador,
                estado_validacion="VALIDADO"
            )
            db.add(nueva_asistencia)
            
        fila = {}
        for key, value in r.items():
            fila[key] = remove_accents(value)
        procesados.append(fila)

    db.commit()

    df = pd.DataFrame(procesados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Capacitacion')
    output.seek(0)
    
    return StreamingResponse(
        output, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=capacitacion_{nuevo_evento.id}.xlsx"}
    )