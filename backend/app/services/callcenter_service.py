from app.database import get_connection


def obtener_reporte_callcenter(queue: str, fecha_inicio: str, fecha_fin: str):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    query_general = """
    SELECT
        DATE(c.datetime_entry_queue) AS fecha,
        q.queue AS cola,
        COUNT(*) AS llamadas_totales,
        SUM(c.status = 'terminada') AS respondidas,
        ROUND(SUM(c.status = 'terminada') / COUNT(*) * 100, 2) AS pct_respondidas,
        SUM(c.status = 'abandonada') AS abandonadas,
        ROUND(SUM(c.status = 'abandonada') / COUNT(*) * 100, 2) AS pct_abandonadas,
        ROUND(AVG(c.duration_wait), 0) AS promedio_espera,
        MAX(c.duration_wait) AS espera_mas_larga,
        ROUND(AVG(c.duration), 0) AS promedio_duracion_llamada,
        MAX(c.duration) AS duracion_mas_larga
    FROM call_entry c
    JOIN queue_call_entry q ON c.id_queue_call_entry = q.id
    WHERE q.queue = %s
      AND c.datetime_entry_queue >= %s
      AND c.datetime_entry_queue < %s
      AND c.status IN ('terminada', 'abandonada')
    GROUP BY DATE(c.datetime_entry_queue), q.queue
    ORDER BY fecha;
    """

    cursor.execute(query_general, (queue, fecha_inicio, fecha_fin))
    data = cursor.fetchall()

    query_horas = """
    SELECT
        DATE(c.datetime_entry_queue) AS fecha,
        HOUR(IF(c.status = 'abandonada', c.datetime_entry_queue, c.datetime_init)) AS hora,
        COUNT(*) AS cantidad
    FROM call_entry c
    JOIN queue_call_entry q ON c.id_queue_call_entry = q.id
    WHERE q.queue = %s
      AND c.datetime_entry_queue >= %s
      AND c.datetime_entry_queue < %s
      AND c.status IN ('terminada', 'abandonada')
      AND HOUR(IF(c.status = 'abandonada', c.datetime_entry_queue, c.datetime_init)) BETWEEN 6 AND 19
    GROUP BY DATE(c.datetime_entry_queue), HOUR(IF(c.status = 'abandonada', c.datetime_entry_queue, c.datetime_init))
    ORDER BY fecha, hora;
    """

    cursor.execute(query_horas, (queue, fecha_inicio, fecha_fin))
    horas_data = cursor.fetchall()

    cursor.close()
    conn.close()

    horas_por_fecha = {}
    for row in horas_data:
        fecha = str(row['fecha'])
        if fecha not in horas_por_fecha:
            horas_por_fecha[fecha] = {}
        hora = int(row['hora'])
        horas_por_fecha[fecha][hora] = int(row['cantidad'])

    for fecha, horas in horas_por_fecha.items():
        suma_horas = sum(horas.values())
        registro_dia = next((r for r in data if str(r['fecha']) == fecha), None)
        if registro_dia and suma_horas != registro_dia['llamadas_totales']:
            print(f"ALERTA - Fecha {fecha}: suma horas {suma_horas} != total llamadas {registro_dia['llamadas_totales']}")

    for registro in data:
        fecha = str(registro['fecha'])

        if fecha not in horas_por_fecha or not horas_por_fecha[fecha]:
            registro['hora_pico'] = "-"
            registro['cantidad_hora_pico'] = 0
            registro['hora_menos_pico'] = "-"
            registro['cantidad_hora_menos_pico'] = 0
            continue

        horas = horas_por_fecha[fecha]

        pares = {}
        for hora in range(6, 19):
            actual    = horas.get(hora, 0)
            siguiente = horas.get(hora + 1, 0)
            pares[hora] = actual + siguiente

        if pares:
            hora_max = max(pares, key=pares.get)
            registro['hora_pico']           = f"{hora_max:02d}h+{hora_max + 1:02d}h"
            registro['cantidad_hora_pico']  = pares[hora_max]

            hora_min = min(pares, key=pares.get)
            registro['hora_menos_pico']          = f"{hora_min:02d}h+{hora_min + 1:02d}h"
            registro['cantidad_hora_menos_pico'] = pares[hora_min]
        else:
            registro['hora_pico']                = "-"
            registro['cantidad_hora_pico']       = 0
            registro['hora_menos_pico']          = "-"
            registro['cantidad_hora_menos_pico'] = 0

    return data


def obtener_reporte_dashboard(fecha_inicio: str, fecha_fin: str):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    query = """
    SELECT
        DATE(c.datetime_entry_queue) AS fecha,
        COUNT(*) AS llamadas_totales,
        SUM(CASE WHEN c.status = 'terminada' THEN 1 ELSE 0 END) AS respondidas,
        ROUND(SUM(CASE WHEN c.status = 'terminada' THEN 1 ELSE 0 END) / COUNT(*) * 100, 2) AS pct_respondidas,
        SUM(CASE WHEN c.status = 'abandonada' THEN 1 ELSE 0 END) AS abandonadas,
        ROUND(SUM(CASE WHEN c.status = 'abandonada' THEN 1 ELSE 0 END) / COUNT(*) * 100, 2) AS pct_abandonadas,
        ROUND(AVG(c.duration_wait), 0) AS promedio_espera,
        MAX(c.duration_wait) AS espera_mas_larga,
        ROUND(AVG(c.duration), 0) AS promedio_duracion_llamada,
        MAX(c.duration) AS duracion_mas_larga
    FROM call_entry c
    WHERE c.datetime_entry_queue >= %s
      AND c.datetime_entry_queue < %s
      AND c.status IN ('terminada', 'abandonada')
    GROUP BY DATE(c.datetime_entry_queue)
    ORDER BY fecha;
    """

    cursor.execute(query, (fecha_inicio, fecha_fin))
    data = cursor.fetchall()

    cursor.close()
    conn.close()

    return data


def obtener_dashboard_por_cola(fecha_inicio: str, fecha_fin: str):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    query = """
    SELECT
        DATE(c.datetime_entry_queue) AS fecha,
        q.queue AS cola,
        COUNT(*) AS llamadas_totales,
        SUM(CASE WHEN c.status = 'terminada' THEN 1 ELSE 0 END) AS respondidas,
        SUM(CASE WHEN c.status = 'abandonada' THEN 1 ELSE 0 END) AS abandonadas,
        ROUND(SUM(CASE WHEN c.status = 'terminada' THEN 1 ELSE 0 END) / COUNT(*) * 100, 1) AS pct_exito,
        ROUND(SUM(CASE WHEN c.status = 'abandonada' THEN 1 ELSE 0 END) / COUNT(*) * 100, 1) AS pct_abandonadas
    FROM call_entry c
    JOIN queue_call_entry q ON c.id_queue_call_entry = q.id
    WHERE c.datetime_entry_queue >= %s
      AND c.datetime_entry_queue < %s
      AND c.status IN ('terminada', 'abandonada')
    GROUP BY DATE(c.datetime_entry_queue), q.queue
    HAVING llamadas_totales > 0
    ORDER BY cola, fecha
    """

    cursor.execute(query, (fecha_inicio, fecha_fin))
    data = cursor.fetchall()

    cursor.close()
    conn.close()

    return data