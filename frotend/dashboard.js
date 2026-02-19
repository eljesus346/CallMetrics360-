const API_BASE = "http://127.0.0.1:8000/api";

let chartDistribucion;

function inicializarGrafico() {
    const ctx = document.getElementById('chartDistribucion')?.getContext('2d');
    if (!ctx) {
        console.warn("No se encontró el canvas chartDistribucion");
        return;
    }

    chartDistribucion = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: [],
            datasets: [{
                label: 'Llamadas Totales por Cola',
                data: [],
                backgroundColor: [
                    '#6366f1', '#8b5cf6', '#a78bfa', '#c084fc', '#d946ef',
                    '#ec4899', '#f472b6', '#fb7185', '#f43f5e', '#e11d48',
                    '#3b82f6', '#60a5fa', '#93c5fd', '#10b981', '#34d399'
                ],
                borderColor: '#0f172a',
                borderWidth: 2,
                hoverOffset: 20
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '45%',
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        color: '#cbd5e1',
                        font: { size: 13 },
                        padding: 20,
                        boxWidth: 15
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(15, 23, 42, 0.95)',
                    titleFont: { size: 15 },
                    bodyFont: { size: 13 },
                    padding: 14,
                    callbacks: {
                        title: (items) => items[0].label,
                        label: (context) => {
                            const d = context.raw;
                            const pctResp = d.llamadas_totales > 0 
                                ? ((d.respondidas / d.llamadas_totales) * 100).toFixed(1) 
                                : '0.0';
                            return [
                                `Totales:     ${d.llamadas_totales.toLocaleString()}`,
                                `Respondidas: ${d.respondidas.toLocaleString()} (${pctResp}%)`,
                                `Abandonadas: ${d.abandonadas.toLocaleString()} (${d.pct_abandonadas}%)`
                            ];
                        }
                    }
                }
            }
        }
    });
}

async function cargarDatosDashboard() {
    try {
        const res = await fetch(`${API_BASE}/reportes/dashboard`);
        if (!res.ok) {
            throw new Error(`Error ${res.status}: ${res.statusText}`);
        }

        const data = await res.json();
        const colas = data.colas || [];
        const totales = data.totales || {};

        // KPIs
        document.getElementById("kpiLlamadas").textContent    = (totales.llamadas_totales ?? 0).toLocaleString();
        document.getElementById("kpiRespondidas").textContent = (totales.respondidas ?? 0).toLocaleString();
        document.getElementById("kpiAbandono").textContent    = (totales.pct_abandonadas ?? 0).toFixed(1) + "%";
        document.getElementById("kpiEspera").textContent      = "—";

        // Gráfico
        chartDistribucion.data.labels = colas.map(r => r.cola || 'Desconocida');
        chartDistribucion.data.datasets[0].data = colas.map(r => r.llamadas_totales || 0);

        // Guardamos el detalle completo para el tooltip
        chartDistribucion.data.datasets[0].detalleColas = colas;

        chartDistribucion.update();

    } catch (error) {
        console.error("Error cargando dashboard:", error);
    }
}

document.addEventListener('DOMContentLoaded', () => {
    inicializarGrafico();
    cargarDatosDashboard();
    setInterval(cargarDatosDashboard, 10000);  // refresca cada 10 segundos
});

function irControl() {
    window.location.href = "callcenter.html";
}

function irAgentes() {
    window.location.href = "agentes-tiempo.html";
}