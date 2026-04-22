import os
import json
import io
import unicodedata
from datetime import datetime, timedelta
from fastapi import FastAPI, Depends, HTTPException, status, UploadFile, File, Form, Bodyfrom fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from sqlalchemy import create_engine, Column, String, Boolean, DateTime, Text, or_
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
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

class User(Base):
    __tablename__ = "usuarios"
    id = Column(String(36), primary_key=True)
    email = Column(String(100), unique=True, nullable=False)
    password_hash = Column(String(255), nullable=False)
    nombre_completo = Column(String(100))
    rol = Column(String(50))
    activo = Column(Boolean, default=True)
    fecha_creacion = Column(DateTime, default=datetime.utcnow)

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
    ultima_actualizacion = Column(DateTime, default=datetime.utcnow)

class AppState(Base):
    __tablename__ = "app_state"
    id = Column(String(36), primary_key=True)
    estado_json = Column(Text)
    ultima_modificacion = Column(DateTime, default=datetime.utcnow)

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
    if text is None:
        return ""
    if not isinstance(text, str):
        text = str(text)
    text = text.strip().upper()
    text = unicodedata.normalize('NFD', text)
    return "".join(c for c in text if unicodedata.category(c) != 'Mn')

@app.get("/")
def read_root():
    return {"status": "API en linea"}

@app.get("/check-db-status")
def check_db_status(db: Session = Depends(get_db)):
    count = db.query(Colaborador).count()
    return {"ready": count > 0, "count": count}

@app.get("/load-state")
def load_state(db: Session = Depends(get_db)):
    state = db.query(AppState).first()
    if state and state.estado_json:
        return json.loads(state.estado_json)
    return None

@app.post("/save-state")
def save_state(payload: dict, db: Session = Depends(get_db)):
    state = db.query(AppState).first()
    if not state:
        state = AppState(id="1", estado_json=json.dumps(payload))
        db.add(state)
    else:
        state.estado_json = json.dumps(payload)
        state.ultima_modificacion = datetime.utcnow()
    db.commit()
    return {"status": "ok"}

@app.post("/upload-masters")
async def upload_masters(file: UploadFile = File(...), source: str = Form(...), db: Session = Depends(get_db)):
    contents = await file.read()
    df = pd.read_excel(io.BytesIO(contents), dtype=str)
    df = df.fillna("")
    df.columns = [str(c).strip().upper() for c in df.columns]

    def find_col(keywords):
        for col in df.columns:
            for kw in keywords:
                if kw in col:
                    return col
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
        if not col_cedula:
            continue
            
        cedula = str(row[col_cedula]).strip()
        
        if cedula.endswith('.0'):
            cedula = cedula[:-2]
            
        if not cedula or cedula.lower() == "nan":
            continue
            
        if cedula.isdigit() and len(cedula) < 10:
            cedula = cedula.zfill(10)
            
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
        colaborador.ultima_actualizacion = datetime.utcnow()
        
    db.commit()
    return {"message": "Datos procesados correctamente"}

@app.post("/validate-cedula")
def validate_cedula(cedulas_json: str = Form(...), db: Session = Depends(get_db)):
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
            resultados.append({
                "cedula": cedula_str,
                "found": True,
                "source": colaborador.origen,
                "data": data
            })
        else:
            resultados.append({
                "cedula": cedula_str,
                "found": False,
                "source": None,
                "data": {}
            })
            
    return resultados

@app.post("/suggest-cedulas")
def suggest_cedulas(search_term: str = Form(...), db: Session = Depends(get_db)):
    search_term = search_term.strip()
    if len(search_term) < 2:
        return []
        
    search_pattern = f"%{search_term}%"
    
    resultados = db.query(Colaborador).filter(
        or_(
            Colaborador.cedula.ilike(search_pattern),
            Colaborador.nombres.ilike(search_pattern),
            Colaborador.apellidos.ilike(search_pattern)
        )
    ).limit(10).all()
    
    suggestions = []
    for r in resultados:
        suggestions.append({
            "cedula": r.cedula,
            "nombre": f"{r.apellidos} {r.nombres}".strip(),
            "source": r.origen
        })
        
    return suggestions

@app.post("/export-excel")
async def export_excel(registros: list = Body(...)):
    procesados = []
    for r in registros:
        fila = {}
        for key, value in r.items():
            fila[key] = remove_accents(value)
        procesados.append(fila)

    df = pd.DataFrame(procesados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Capacitacion')
    output.seek(0)
    
    return StreamingResponse(
        output, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=capacitacion.xlsx"}
    )
    #verga para emelec