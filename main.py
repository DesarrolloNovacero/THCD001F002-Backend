import sqlite3
import json
import io
import os
import sys
import unicodedata
from datetime import datetime
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

if getattr(sys, "frozen", False):
    app_data = os.environ.get("APPDATA") or os.path.expanduser("~")
    DB_PATH = os.path.join(app_data, "trainform_data", "trainform.db")
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
else:
    DB_PATH = "trainform.db"

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute('''
    CREATE TABLE IF NOT EXISTS colaboradores (
        cedula TEXT PRIMARY KEY,
        apellidos TEXT,
        nombres TEXT,
        cargo TEXT,
        genero TEXT,
        unidad TEXT,
        area TEXT,
        seccion TEXT,
        centro_costo TEXT,
        grupo_personal TEXT,
        area_personal TEXT,
        jefe_area TEXT,
        gerente_area TEXT,
        localidad TEXT,
        origen TEXT,
        ultima_actualizacion DATETIME
    )
    ''')
    c.execute('''
    CREATE TABLE IF NOT EXISTS app_state (
        id INTEGER PRIMARY KEY,
        state_json TEXT,
        updated_at DATETIME
    )
    ''')
    conn.commit()
    conn.close()

init_db()

def remove_accents(text):
    if not text: return ""
    text = str(text).upper().strip()
    text = unicodedata.normalize("NFD", text)
    return "".join(c for c in text if unicodedata.category(c) != "Mn")

@app.get("/check-db-status")
async def check_db_status():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT count(*) FROM colaboradores")
    count = cursor.fetchone()[0]
    conn.close()
    return {"ready": count > 0, "count": count}

@app.post("/upload-masters")
async def upload_masters(source: str = Form(...), file: UploadFile = File(...)):
    try:
        content = await file.read()
        file_bytes = io.BytesIO(content)
        
        if file.filename.lower().endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_bytes, dtype=str)
            df = df.replace(r'^\s*$', float('nan'), regex=True)
            df = df.ffill()
        else:
            try:
                df = pd.read_csv(file_bytes, dtype=str, sep=';', encoding='utf-8')
            except:
                file_bytes.seek(0)
                df = pd.read_csv(file_bytes, dtype=str, sep=',', encoding='latin-1')
        
        df.fillna("", inplace=True)
        conn = get_db()
        cursor = conn.cursor()
        timestamp = datetime.now().isoformat()

        for _, row in df.iterrows():
            if source == "headcount":
                cedula = str(row.iloc[0]).strip()
                if not cedula: continue
                cedula = cedula.zfill(10)
                
                data = {
                    "cedula": cedula,
                    "apellidos": row.iloc[2],
                    "nombres": row.iloc[3],
                    "cargo": row.iloc[5],
                    "genero": row.iloc[8],
                    
                    "unidad": row.iloc[10],         
                    "area": row.iloc[12],           
                    "seccion": row.iloc[13],        
                    "centro_costo": row.iloc[20],   
                    "grupo_personal": row.iloc[21], 
                    "area_personal": row.iloc[14],  
                    "jefe_area": row.iloc[15],      
                    "gerente_area": row.iloc[15],   
                    
                    "localidad": row.iloc[4],
                    "origen": "headcount"
                }
            else: 
                cedula = str(row.iloc[31]).strip()
                if not cedula: continue
                cedula = cedula.zfill(10)

                data = {
                    "cedula": cedula, "apellidos": row.iloc[1], "nombres": row.iloc[2],
                    "cargo": row.iloc[4], "genero": row.iloc[7], "unidad": row.iloc[9],
                    "area": row.iloc[11], "seccion": row.iloc[12], "centro_costo": row.iloc[13],
                    "grupo_personal": row.iloc[24], "area_personal": row.iloc[25],
                    "jefe_area": row.iloc[29], "gerente_area": row.iloc[26],
                    "localidad": row.iloc[3], "origen": "cesantes"
                }
            
            cursor.execute('''
                INSERT INTO colaboradores (
                    cedula, apellidos, nombres, cargo, genero, unidad, area, seccion,
                    centro_costo, grupo_personal, area_personal, jefe_area, gerente_area,
                    localidad, origen, ultima_actualizacion
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(cedula) DO UPDATE SET
                    apellidos=excluded.apellidos, nombres=excluded.nombres, cargo=excluded.cargo,
                    genero=excluded.genero, unidad=excluded.unidad, area=excluded.area,
                    seccion=excluded.seccion, centro_costo=excluded.centro_costo,
                    grupo_personal=excluded.grupo_personal, area_personal=excluded.area_personal,
                    jefe_area=excluded.jefe_area, gerente_area=excluded.gerente_area,
                    localidad=excluded.localidad, origen=excluded.origen,
                    ultima_actualizacion=excluded.ultima_actualizacion
            ''', tuple(data.values()) + (timestamp,))

        conn.commit()
        conn.close()
        return {"status": "success"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/validate-cedula")
async def validate_cedula(cedulas_json: str = Form(...)):
    lista_cedulas = json.loads(cedulas_json)
    conn = get_db()
    cursor = conn.cursor()
    resultados = []
    
    for cedula in lista_cedulas:
        cedula_formateada = str(cedula).strip().zfill(10)
        cursor.execute("SELECT * FROM colaboradores WHERE cedula = ?", (cedula_formateada,))
        row = cursor.fetchone()
        item = {"cedula": cedula_formateada, "found": False, "source": None, "data": {}}
        if row:
            item["found"] = True
            item["source"] = row["origen"]
            item["data"] = {k: row[k] for k in row.keys() if k not in ['cedula', 'origen', 'ultima_actualizacion']}
        resultados.append(item)
        
    conn.close()
    return resultados

@app.post("/suggest-cedulas")
async def suggest_cedulas(search_term: str = Form(...)):
    term = search_term.strip()
    if len(term) < 2: return []
    
    conn = get_db()
    cursor = conn.cursor()
    
    query = f"%{term}%"
    cursor.execute("""
        SELECT cedula, apellidos, nombres, origen 
        FROM colaboradores 
        WHERE cedula LIKE ? OR apellidos LIKE ? OR nombres LIKE ?
        LIMIT 10
    """, (query, query, query))
    
    rows = cursor.fetchall()
    conn.close()
    
    sugerencias = []
    for row in rows:
        sugerencias.append({
            "cedula": row["cedula"],
            "nombre": f"{row['apellidos']} {row['nombres']}",
            "source": row["origen"]
        })
    return sugerencias

@app.post("/save-state")
async def save_state(state: dict = Body(...)):
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO app_state (id, state_json, updated_at) VALUES (1, ?, ?)", 
                       (json.dumps(state), datetime.now().isoformat()))
        conn.commit()
        conn.close()
        return {"status": "saved"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/load-state")
async def load_state():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT state_json FROM app_state WHERE id = 1")
    row = cursor.fetchone()
    conn.close()
    return json.loads(row[0]) if row and row[0] else None

@app.post("/export-excel")
async def export_excel(registros: list = Body(...)):
    # Mapeo que traduce el nombre de la variable al encabezado de Excel
    column_mapping = {
        "nombreCurso": "NOMBRE DEL CURSO", "objetivo": "OBJETIVO", "empresa": "EMPRESA CAPACITADORA",
        "facilitador": "FACILITADOR", "dimensionEvento": "DIMENSIÓN DE EVENTO", "lugar": "LUGAR DONDE SE DIO LA CAPACITACION",
        "modalidad": "MODALIDAD", "fechaInicio": "FECHA INICIO", "fechaCierre": "FECHA CIERRE",
        "totalHoras": "DURACION DE LA CAPACITACION (HORAS)", "tipoEvento": "TIPO EVENTO", "mesAnio": "MES-AÑO",
        "cedula": "CÉDULA", "apellidosNombre": "APELLIDOS Y NOMBRE DEL COLABORADOR", "genero": "GÉNERO",
        "cargo": "CARGO", "unidad": "UNIDAD", "area": "ÁREA", "seccion": "SECCIÓN", "centroCosto": "CENTRO DE COSTO",
        "grupoPersonal": "GRUPO DE PERSONAL", "areaPersonal": "ÁREA DE PERSONAL", "jefeArea": "JEFE DE ÁREA",
        "gerenteArea": "GERENTE DE AREA", "localidad": "LOCALIDAD"
    }
    
    # Lista de campos que suelen venir en snake_case del backend
    snake_to_camel = {
        "centro_costo": "centroCosto",
        "grupo_personal": "grupoPersonal",
        "area_personal": "areaPersonal",
        "jefe_area": "jefeArea",
        "gerente_area": "gerenteArea"
    }

    processed = []
    for r in registros:
        # Normalizar: Si viene centro_costo, moverlo a centroCosto para que coincida con column_mapping
        for snake, camel in snake_to_camel.items():
            if snake in r and (camel not in r or not r[camel]):
                r[camel] = r[snake]

        row_data = {header: remove_accents(str(r.get(k, ""))) for k, header in column_mapping.items()}
        
        if r.get("cedula"):
            row_data["CÉDULA"] = str(r["cedula"]).strip().zfill(10)
        
        processed.append(row_data)

    df = pd.DataFrame(processed)
    df = df.reindex(columns=list(column_mapping.values()), fill_value="")
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Capacitaciones')
    output.seek(0)
    
    return StreamingResponse(output, headers={'Content-Disposition': 'attachment; filename="capacitacion.xlsx"'}, 
                             media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                             