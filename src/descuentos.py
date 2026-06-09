from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy import text
from sqlalchemy.orm import Session
from typing import Optional, List
from pydantic import BaseModel
from database import get_db

router = APIRouter(prefix="/descuentos", tags=["Descuentos"])


class DetalleDescuentoIn(BaseModel):
    idgrupodescuento: int
    descuento: Optional[float] = 0


class EstrategiaIn(BaseModel):
    nombre: str
    descripcion: Optional[str] = None
    detalle: List[DetalleDescuentoIn] = []


@router.get("/")
def listar(db: Session = Depends(get_db)):
    """Lista todas las estrategias de descuento."""
    sql = "SELECT id, nombre, descripcion, fecha FROM descuentos ORDER BY nombre"
    return [dict(r._mapping) for r in db.execute(text(sql)).fetchall()]


@router.get("/{id}")
def obtener(id: int, db: Session = Depends(get_db)):
    """Cabecera de la estrategia + su detalle (con datos del grupo descuento)."""
    cab = db.execute(
        text("SELECT id, nombre, descripcion, fecha FROM descuentos WHERE id = :id"),
        {"id": id}
    ).fetchone()
    if not cab:
        raise HTTPException(status_code=404, detail="Estrategia no encontrada")

    det_sql = """
        SELECT dd.id, dd.iddescuento, dd.idgrupodescuento, dd.descuento,
               gd.grupodescuentospain AS grupodescuento,
               gd.descripcion         AS descripciongrupo
        FROM detalledescuentos dd
        LEFT JOIN gruposdescuento gd ON dd.idgrupodescuento = gd.id
        WHERE dd.iddescuento = :id
        ORDER BY gd.grupodescuentospain
    """
    detalle = db.execute(text(det_sql), {"id": id}).fetchall()
    return {
        "cabecera": dict(cab._mapping),
        "detalle": [dict(r._mapping) for r in detalle],
    }


@router.post("/")
def crear(estrategia: EstrategiaIn, db: Session = Depends(get_db)):
    """Crea una estrategia y su detalle."""
    r = db.execute(text("""
        INSERT INTO descuentos (nombre, descripcion, fecha)
        VALUES (:nombre, :descripcion, NOW())
        RETURNING id
    """), {"nombre": estrategia.nombre, "descripcion": estrategia.descripcion}).fetchone()
    id_estrategia = r[0]
    for d in estrategia.detalle:
        db.execute(text("""
            INSERT INTO detalledescuentos (iddescuento, idgrupodescuento, descuento)
            VALUES (:idd, :idg, :dto)
        """), {"idd": id_estrategia, "idg": d.idgrupodescuento, "dto": d.descuento or 0})
    db.commit()
    return {"id": id_estrategia, "lineas": len(estrategia.detalle)}


@router.put("/{id}")
def actualizar(id: int, estrategia: EstrategiaIn, db: Session = Depends(get_db)):
    """Actualiza la cabecera y reemplaza el detalle completo."""
    exist = db.execute(text("SELECT id FROM descuentos WHERE id = :id"), {"id": id}).fetchone()
    if not exist:
        raise HTTPException(status_code=404, detail="Estrategia no encontrada")
    db.execute(text("""
        UPDATE descuentos SET nombre = :nombre, descripcion = :descripcion, fecha = NOW()
        WHERE id = :id
    """), {"id": id, "nombre": estrategia.nombre, "descripcion": estrategia.descripcion})
    db.execute(text("DELETE FROM detalledescuentos WHERE iddescuento = :id"), {"id": id})
    for d in estrategia.detalle:
        db.execute(text("""
            INSERT INTO detalledescuentos (iddescuento, idgrupodescuento, descuento)
            VALUES (:idd, :idg, :dto)
        """), {"idd": id, "idg": d.idgrupodescuento, "dto": d.descuento or 0})
    db.commit()
    return {"id": id, "lineas": len(estrategia.detalle)}


@router.delete("/{id}")
def borrar(id: int, db: Session = Depends(get_db)):
    """Borra una estrategia y todo su detalle."""
    exist = db.execute(text("SELECT id FROM descuentos WHERE id = :id"), {"id": id}).fetchone()
    if not exist:
        raise HTTPException(status_code=404, detail="Estrategia no encontrada")
    db.execute(text("DELETE FROM detalledescuentos WHERE iddescuento = :id"), {"id": id})
    db.execute(text("DELETE FROM descuentos WHERE id = :id"), {"id": id})
    db.commit()
    return {"id": id, "borrado": True}
