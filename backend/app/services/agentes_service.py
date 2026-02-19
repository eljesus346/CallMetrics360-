from app.database import get_connection

def obtener_reporte_agentes(fecha_inicio: str, fecha_fin: str):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    query = """
    SELECT
        ag.name AS agente,

        SUM(TIME_TO_SEC(a.duration)) AS tiempo_logueado,

        SUM(
            CASE
                WHEN a.id_break IS NULL THEN TIME_TO_SEC(a.duration)
                ELSE 0
            END
        ) AS tiempo_activo,

        COALESCE(b.name, 'SIN PAUSA') AS tipo_pausa,

        SUM(
            CASE
                WHEN a.id_break IS NOT NULL THEN TIME_TO_SEC(a.duration)
                ELSE 0
            END
        ) AS tiempo_pausa,

        CASE
            WHEN LOWER(b.name) LIKE '%almuerzo%'
                AND SUM(TIME_TO_SEC(a.duration)) > 3600 THEN 'INCUMPLE'

            WHEN LOWER(b.name) LIKE '%descanso%'
                AND SUM(TIME_TO_SEC(a.duration)) > 600 THEN 'INCUMPLE'

            -- BAÑO = 15 MIN (900 SEGUNDOS)
            WHEN LOWER(b.name) LIKE '%ba%'
                AND SUM(TIME_TO_SEC(a.duration)) > 900 THEN 'INCUMPLE'

            WHEN b.name IS NULL THEN 'TRABAJO'

            ELSE 'CUMPLE'
        END AS estado

    FROM audit a
    JOIN agent ag ON a.id_agent = ag.id
    LEFT JOIN break b ON a.id_break = b.id

    WHERE a.datetime_init >= %s
      AND a.datetime_end <= %s

    GROUP BY ag.name, b.name
    ORDER BY ag.name;
    """

    cursor.execute(query, (fecha_inicio, fecha_fin))
    rows = cursor.fetchall()

    cursor.close()
    conn.close()

    return rows
