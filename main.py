@app.get("/usuarios")
def listar_usuarios(current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    usuarios = db.query(Usuario).order_by(Usuario.fecha_creacion.desc()).all()
    return [{"id": str(u.id), "email": u.email, "nombre_completo": u.nombre_completo, "rol": u.rol, "activo": u.activo} for u in usuarios]

@app.put("/usuarios/{user_id}/toggle")
def toggle_usuario(user_id: str, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    if str(current_admin.id) == user_id:
        raise HTTPException(status_code=400, detail="No puedes desactivar tu propia cuenta")
    usuario = db.query(Usuario).filter(Usuario.id == user_id).first()
    if not usuario:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    usuario.activo = not usuario.activo
    db.commit()
    return {"message": "Estado actualizado", "activo": usuario.activo}

@app.delete("/usuarios/{user_id}")
def eliminar_usuario(user_id: str, current_admin: Usuario = Depends(get_current_admin), db: Session = Depends(get_db)):
    if str(current_admin.id) == user_id:
        raise HTTPException(status_code=400, detail="No puedes eliminar tu propia cuenta")
    usuario = db.query(Usuario).filter(Usuario.id == user_id).first()
    if not usuario:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    
    tiene_eventos = db.query(Evento).filter(Evento.creado_por_usuario_id == user_id).first()
    if tiene_eventos:
        raise HTTPException(status_code=400, detail="Este usuario ya ha exportado cursos. Por seguridad de trazabilidad, debes 'Desactivarlo' en lugar de eliminarlo.")
        
    db.query(SesionWeb).filter(SesionWeb.usuario_id == user_id).delete()
    db.delete(usuario)
    db.commit()
    return {"message": "Usuario eliminado permanentemente"}