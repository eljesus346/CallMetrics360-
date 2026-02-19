const API_BASE = "http://127.0.0.1:8000/api";

/* =========================
   MENSAJES UI
========================= */
function mostrarMensaje(texto, tipo = "info") {
    const msg = document.getElementById("mensaje");
    if (!msg) return;

    msg.className = `mensaje mensaje-${tipo}`;
    msg.innerText = texto;
    msg.classList.remove("hidden");

    setTimeout(() => {
        msg.classList.add("hidden");
    }, 4000);
}

/* =========================
   UTILIDAD TIEMPO
========================= */
function segundosATiempo(segundos) {
    if (segundos === null || segundos === undefined) return "00:00";

    const h = Math.floor(segundos / 3600);
    const m = Math.floor((segundos % 3600) / 60);
    const s = segundos % 60;

    return h > 0
        ? `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`
        : `${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`;
}

/* =========================
   CARGAR REPORTE
========================= */
async function cargarReporte() {
    const inicio = document.getElementById("fecha_inicio").value;
    const fin = document.getElementById("fecha_fin").value;
    const loader = document.getElementById("loader");
    const tbody = document.getElementById("tabla_agentes");
    const contador = document.getElementById("contador");
    const resumenDiv = document.getElementById("resumen");

    if (!inicio || !fin) {
        mostrarMensaje("⚠️ Selecciona ambas fechas", "error");
        return;
    }

    loader.classList.remove("hidden");
    tbody.innerHTML = "";
    contador.innerText = "";
    resumenDiv.innerHTML = "";

    try {
        const res = await fetch(
            `${API_BASE}/reportes/agentes?fecha_inicio=${inicio}&fecha_fin=${fin}`
        );

        if (!res.ok) throw new Error("Error API");

        const data = await res.json();
        loader.classList.add("hidden");

        if (!data.length) {
            mostrarMensaje("📭 No se encontraron registros", "info");
            return;
        }

        mostrarMensaje("✅ Reporte de agentes cargado", "exito");
        contador.innerText = `📌 Registros encontrados: ${data.length}`;

        data.forEach(row => {
            let estadoClass = "";
            switch (row.estado) {
                case "TRABAJO": estadoClass = "estado-trabajo"; break;
                case "CUMPLE": estadoClass = "estado-cumple"; break;
                case "INCUMPLE": estadoClass = "estado-incumple"; break;
                case "MONITOREAR": estadoClass = "estado-monitorear"; break;
            }

            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td>${row.agente}</td>
                <td>${segundosATiempo(row.tiempo_logueado)}</td>
                <td>${segundosATiempo(row.tiempo_activo)}</td>
                <td>${row.tipo_pausa}</td>
                <td>${segundosATiempo(row.tiempo_pausa)}</td>
                <td class="${estadoClass}"><strong>${row.estado}</strong></td>
            `;
            tbody.appendChild(tr);
        });

    } catch (error) {
        loader.classList.add("hidden");
        mostrarMensaje("❌ Error conectando con el servidor", "error");
        console.error(error);
    }
}

/* =========================
   DESCARGAR EXCEL
========================= */
function descargarExcel() {
    const inicio = document.getElementById("fecha_inicio").value;
    const fin = document.getElementById("fecha_fin").value;

    if (!inicio || !fin) {
        mostrarMensaje("⚠️ Selecciona las fechas antes de descargar", "error");
        return;
    }

    window.location.href =
        `${API_BASE}/reportes/excel/agentes?fecha_inicio=${inicio}&fecha_fin=${fin}`;
}
