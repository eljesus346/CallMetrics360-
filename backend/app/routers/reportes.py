from fastapi import APIRouter, Query
from fastapi.responses import StreamingResponse
from datetime import date, datetime, timedelta

from app.services.agentes_service import obtener_reporte_agentes
from app.services.callcenter_service import (
    obtener_reporte_callcenter,
    obtener_reporte_dashboard,
    obtener_dashboard_por_cola
)

from app.services.excel_agentes_service import generar_excel_agentes
from app.excel.excel_callcenter import generar_excel_callcenter

router = APIRouter(
    prefix="/reportes",
    tags=["Reportes"]
)


@router.get("/agentes")
def reporte_agentes(
    fecha_inicio: str = Query(...),
    fecha_fin: str = Query(...)
):
    inicio = f"{fecha_inicio} 00:00:00"
    fin = f"{fecha_fin} 23:59:59"
    return obtener_reporte_agentes(inicio, fin)


@router.get("/excel/agentes")
def descargar_excel_agentes(
    fecha_inicio: str = Query(...),
    fecha_fin: str = Query(...)
):
    inicio = f"{fecha_inicio} 00:00:00"
    fin = f"{fecha_fin} 23:59:59"
    data = obtener_reporte_agentes(inicio, fin)
    output = generar_excel_agentes(data)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=reporte_agentes.xlsx"}
    )


@router.get("/callcenter/semana")
def reporte_callcenter(
    queue: str = Query(...),
    fecha_inicio: str = Query(...),
    fecha_fin: str = Query(...)
):
    try:
        dt_inicio = datetime.strptime(fecha_inicio, "%Y-%m-%d")
        dt_fin = datetime.strptime(fecha_fin, "%Y-%m-%d") + timedelta(days=1)

        inicio = dt_inicio.strftime("%Y-%m-%d 00:00:00")
        fin = dt_fin.strftime("%Y-%m-%d 00:00:00")
    except Exception:
        inicio = f"{fecha_inicio} 00:00:00"
        fin = f"{fecha_fin} 23:59:59"

    return obtener_reporte_callcenter(queue, inicio, fin)


@router.get("/callcenter/semana/excel")
def descargar_excel_callcenter(
    fecha_inicio: str = Query(...),
    fecha_fin: str = Query(...)
):
    try:
        dt_inicio = datetime.strptime(fecha_inicio, "%Y-%m-%d")
        dt_fin = datetime.strptime(fecha_fin, "%Y-%m-%d") + timedelta(days=1)

        inicio = dt_inicio.strftime("%Y-%m-%d 00:00:00")
        fin = dt_fin.strftime("%Y-%m-%d 00:00:00")
    except Exception:
        inicio = f"{fecha_inicio} 00:00:00"
        fin = f"{fecha_fin} 23:59:59"

    # Datos para hoja Resumen Colas (por fecha y cola)
    data_por_cola = obtener_dashboard_por_cola(inicio, fin)

    # Obtener las colas únicas presentes en el período
    colas = list({d["cola"] for d in data_por_cola})

    # Llamar obtener_reporte_callcenter para cada cola → datos completos con horarios, espera, duración
    data_detalle_por_cola = {}
    for cola in colas:
        data_detalle_por_cola[cola] = obtener_reporte_callcenter(cola, inicio, fin)

    output = generar_excel_callcenter(data_por_cola, data_detalle_por_cola)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=reporte_callcenter.xlsx"}
    )


@router.get("/dashboard")
def dashboard_general():
    hoy = date.today()
    inicio = f"{hoy} 00:00:00"
    fin   = f"{hoy} 23:59:59"

    datos_colas = obtener_dashboard_por_cola(inicio, fin)

    if not datos_colas:
        return {
            "fecha": str(hoy),
            "colas": [],
            "totales": {
                "llamadas_totales": 0,
                "respondidas": 0,
                "abandonadas": 0,
                "pct_abandonadas": 0.0
            }
        }

    total_llamadas    = sum(r["llamadas_totales"] for r in datos_colas)
    total_respondidas = sum(r["respondidas"] for r in datos_colas)
    total_abandonadas = sum(r["abandonadas"] for r in datos_colas)

    return {
        "fecha": str(hoy),
        "colas": datos_colas,
        "totales": {
            "llamadas_totales": total_llamadas,
            "respondidas": total_respondidas,
            "abandonadas": total_abandonadas,
            "pct_abandonadas": round((total_abandonadas / total_llamadas * 100), 1) if total_llamadas > 0 else 0.0
        }
    }