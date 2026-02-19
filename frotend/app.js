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
   UTILIDADES
========================= */
function segundosAMinutos(segundos) {
    if (!segundos || segundos === "-" || segundos === 0) return "0 min";
    return `${Math.round(segundos / 60)} min`;
}

function segundosAMMS(segundos) {
    if (!segundos || segundos === "-" || segundos === 0) return "0:00:00";
    const mins = Math.floor(segundos / 60);
    const segs = segundos % 60;
    return `0:${mins.toString().padStart(2, "0")}:${segs
        .toString()
        .padStart(2, "0")}`;
}

/* =========================
   ELEMENTOS DOM
========================= */
const btnBuscar = document.getElementById("btnBuscar");
const loader = document.getElementById("loader");
const thead = document.getElementById("thead");
const tbody = document.getElementById("tbody");

/* =========================
   GENERAR SEMANAS CORRECTAMENTE
========================= */
function generarSemanasPersonalizadas(fechaInicio, fechaFin) {

    const semanas = [];

    let inicioFiltro = new Date(fechaInicio);
    let finFiltro = new Date(fechaFin);

    inicioFiltro.setHours(6, 0, 0, 0);
    finFiltro.setHours(19, 0, 0, 0);

    let cursor = new Date(inicioFiltro);
    let esPrimeraSemana = true;

    while (cursor <= finFiltro) {

        let semanaInicio = new Date(cursor);
        semanaInicio.setHours(6, 0, 0, 0);

        let semanaFin = new Date(semanaInicio);

        const dia = semanaInicio.getDay();

        if (esPrimeraSemana) {
            const diasHastaSabado = (6 - dia);
            semanaFin.setDate(semanaInicio.getDate() + diasHastaSabado);
            esPrimeraSemana = false;
        } else {
            if (dia !== 1) {
                const diasHastaLunes = (8 - dia) % 7;
                semanaInicio.setDate(semanaInicio.getDate() + diasHastaLunes);
                semanaInicio.setHours(6, 0, 0, 0);
            }

            semanaFin = new Date(semanaInicio);
            semanaFin.setDate(semanaInicio.getDate() + 5);
        }

        semanaFin.setHours(19, 0, 0, 0);

        if (semanaFin > finFiltro) {
            semanaFin = new Date(finFiltro);
        }

        semanas.push({
            inicio: semanaInicio.toISOString().split("T")[0],
            fin: semanaFin.toISOString().split("T")[0]
        });

        cursor = new Date(semanaFin);
        cursor.setDate(semanaFin.getDate() + 1);
        cursor.setHours(6, 0, 0, 0);
    }

    return semanas;
}

/* =========================
   EVENTO BUSCAR
========================= */
btnBuscar.addEventListener("click", async () => {

    const queue = document.getElementById("queue").value;
    const inicio = document.getElementById("fechaInicio").value;
    const fin = document.getElementById("fechaFin").value;

    if (!queue || !inicio || !fin) {
        mostrarMensaje("⚠️ Completa todos los campos", "error");
        return;
    }

    loader.classList.remove("hidden");
    thead.innerHTML = "";
    tbody.innerHTML = "";

    try {

        const diferenciaDias =
            (new Date(fin) - new Date(inicio)) / (1000 * 60 * 60 * 24);

        if (diferenciaDias > 6) {

            const semanas = generarSemanasPersonalizadas(inicio, fin);
            await renderSemanas(semanas, queue);

            loader.classList.add("hidden");
            mostrarMensaje("✅ Reporte dividido correctamente por semanas", "exito");
            return;
        }

        const res = await fetch(
            `${API_BASE}/reportes/callcenter/semana?queue=${queue}&fecha_inicio=${inicio}&fecha_fin=${fin}`
        );

        if (!res.ok) throw new Error("Error API");

        const data = await res.json();
        loader.classList.add("hidden");

        if (!data.length) {
            mostrarMensaje("📭 No hay datos para el rango seleccionado", "info");
            return;
        }

        mostrarMensaje("✅ Reporte cargado correctamente", "exito");
        renderTablaExcelStyle(data);

    } catch (error) {
        loader.classList.add("hidden");
        mostrarMensaje("❌ Error conectando con el backend", "error");
        console.error(error);
    }
});

/* =========================
   RENDER MULTIPLES SEMANAS
========================= */
async function renderSemanas(semanas, queue) {

    // El contenedor principal queda transparente para mostrar el fondo azul de la página
    tbody.style.cssText = "background: transparent; border: none; padding: 0; box-shadow: none;";

    for (let i = 0; i < semanas.length; i++) {

        const semana = semanas[i];

        const res = await fetch(
            `${API_BASE}/reportes/callcenter/semana?queue=${queue}&fecha_inicio=${semana.inicio}&fecha_fin=${semana.fin}`
        );

        if (!res.ok) continue;

        const data = await res.json();
        if (!data.length) continue;

        // Cada semana es su propio bloque flotante con fondo blanco
        const bloque = document.createElement("div");
        bloque.style.cssText = `
            background: #ffffff;
            border-radius: 12px;
            box-shadow: 0 4px 16px rgba(0,0,0,0.18);
            overflow: hidden;
            margin-bottom: 32px;
        `;

        // Encabezado de la semana pegado arriba del bloque
        const header = document.createElement("div");
        header.style.cssText = `
            background: linear-gradient(135deg, #4f46e5, #7c3aed);
            color: #ffffff;
            font-size: 13px;
            font-weight: 700;
            padding: 10px 20px;
            letter-spacing: 0.4px;
        `;
        header.innerText = `📅  Semana ${i + 1}   ·   ${semana.inicio}  →  ${semana.fin}`;

        const tabla = document.createElement("table");
        tabla.classList.add("tabla-excel");
        tabla.style.cssText = "margin: 0; border-radius: 0;";

        const theadLocal = document.createElement("thead");
        const tbodyLocal = document.createElement("tbody");

        tabla.appendChild(theadLocal);
        tabla.appendChild(tbodyLocal);

        bloque.appendChild(header);
        bloque.appendChild(tabla);
        tbody.appendChild(bloque);

        renderTablaEnElemento(data, theadLocal, tbodyLocal);
    }
}

/* =========================
   TABLA NORMAL
========================= */
function renderTablaExcelStyle(data) {
    renderTablaEnElemento(data, thead, tbody);
}

/* =========================
   TABLA GENERICA
========================= */
function renderTablaEnElemento(data, theadRef, tbodyRef) {

    let header = `<tr><th rowspan="2">MÉTRICA</th>`;
    data.forEach(d => {
        header += `<th colspan="2">${d.fecha}</th>`;
    });
    header += `</tr>`;
    theadRef.innerHTML = header;

    const filas = [
        { label: "Llamadas Total", render: d => [d.llamadas_totales, null] },
        { label: "Respondidas", render: d => [d.respondidas, `${d.pct_respondidas}%`] },
        { label: "Abandonadas", render: d => [d.abandonadas, `${d.pct_abandonadas}%`] },
        {
            label: "Horario + Tráfico",
            render: d => [
                d.hora_pico || "-",
                d.cantidad_hora_pico ? `${d.cantidad_hora_pico} Llam` : "-"
            ]
        },
        {
            label: "Horario - Tráfico",
            render: d => [
                d.hora_menos_pico || "-",
                d.cantidad_hora_menos_pico ? `${d.cantidad_hora_menos_pico} Llam` : "-"
            ]
        },
        {
            label: "Promedio Espera",
            render: d => [
                d.promedio_espera ? `${d.promedio_espera} s` : "0 s",
                segundosAMinutos(d.promedio_espera)
            ]
        },
        {
            label: "Espera + Larga",
            render: d => [
                d.espera_mas_larga ? `${d.espera_mas_larga} s` : "0 s",
                segundosAMinutos(d.espera_mas_larga)
            ]
        },
        {
            label: "Prom. Dur. Llam.",
            render: d => [null, segundosAMMS(d.promedio_duracion_llamada)]
        },
        {
            label: "Dur. + Larga Llam.",
            render: d => [null, segundosAMMS(d.duracion_mas_larga)]
        }
    ];

    let bodyHTML = "";

    filas.forEach(fila => {
        bodyHTML += `<tr><td><strong>${fila.label}</strong></td>`;

        data.forEach(d => {
            const [v1, v2] = fila.render(d);
            if (v1 === null) {
                bodyHTML += `<td colspan="2">${v2}</td>`;
            } else if (v2 === null) {
                bodyHTML += `<td colspan="2">${v1}</td>`;
            } else {
                bodyHTML += `<td>${v1}</td><td>${v2}</td>`;
            }
        });

        bodyHTML += `</tr>`;
    });

    tbodyRef.innerHTML = bodyHTML;
}

/* =========================
   DESCARGA EXCEL
========================= */
function descargarExcel() {

    const queue = document.getElementById("queue").value;
    const inicio = document.getElementById("fechaInicio").value;
    const fin = document.getElementById("fechaFin").value;

    if (!queue || !inicio || !fin) {
        mostrarMensaje("⚠️ Completa los filtros antes de descargar", "error");
        return;
    }

    window.location.href =
        `${API_BASE}/reportes/callcenter/semana/excel?queue=${queue}&fecha_inicio=${inicio}&fecha_fin=${fin}`;
}