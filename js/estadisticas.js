const btnCargar = document.getElementById("btn-cargar-excel");
const btnLimpiar = document.getElementById("btn-limpiar");
const contenedorGraficos = document.getElementById("contenedorGraficos");
const bienvenida = document.querySelector(".text-muted");

let charts = [];

function destruirGraficosAnteriores() {
    charts.forEach(chart => chart.destroy());
    charts = [];
}

function procesarDatos(estudiantes) {
    destruirGraficosAnteriores();

    let totalGF = 0, totalGP = 0, totalEX = 0;

    estudiantes.forEach(est => {
        totalGF += est["Gestión Formativa"] || 0;
        totalGP += est["Gestión Práctica"] || 0;
        totalEX += est["Examen"] || 0;
    });

    const promedioGF = totalGF / estudiantes.length;
    const promedioGP = totalGP / estudiantes.length;
    const promedioEX = totalEX / estudiantes.length;

    const nombres = estudiantes.map(e => e["Nombres y Apellidos"]);
    const promedios = estudiantes.map(e => e["Promedio"]);

    // Contenedor de gráficos
    contenedorGraficos.style.display = "block";

    if (bienvenida) bienvenida.style.display = "none";

    // Gráfico: Promedios por estudiante
    charts.push(new Chart(document.getElementById("graficoEstudiantes"), {
        type: 'bar',
        data: {
            labels: nombres,
            datasets: [{
                label: 'Promedio',
                data: promedios,
                backgroundColor: 'rgba(54, 162, 235, 0.6)',
                borderColor: 'rgba(54, 162, 235, 1)',
                borderWidth: 1
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { beginAtZero: true, max: 10 },
                y: { ticks: { font: { size: 10 } } }
            },
            plugins: { legend: { display: false } }
        }
    }));

    
    // Gráfico: Promedio por tipo de nota
    charts.push(new Chart(document.getElementById("graficoTipos"), {
        type: 'pie',
        data: {
            labels: ["Gestión Formativa", "Gestión Práctica", "Examen"],
            datasets: [{
                data: [promedioGF, promedioGP, promedioEX],
                backgroundColor: [
                    'rgba(255, 99, 132, 0.6)',
                    'rgba(255, 206, 86, 0.6)',
                    'rgba(75, 192, 192, 0.6)'
                ]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false
        }
    }));

    // Gráfico: Top 5 mejores estudiantes
    const top5 = [...estudiantes].sort((a, b) => b["Promedio"] - a["Promedio"]).slice(0, 5);
    charts.push(new Chart(document.getElementById("graficoTop5"), {
        type: 'bar',
        data: {
            labels: top5.map(e => e["Nombres y Apellidos"]),
            datasets: [{
                label: "Promedio",
                data: top5.map(e => e["Promedio"]),
                backgroundColor: 'rgba(75, 192, 192, 0.7)',
                borderColor: 'rgba(75, 192, 192, 1)',
                borderWidth: 1
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            scales: { x: { beginAtZero: true, max: 10 } }
        }
    }));

    // Gráfico: Aprobados vs Reprobados
    let aprobados = 0;
    promedios.forEach(p => {
        if (p >= 7) aprobados++;
    });
    const reprobados = promedios.length - aprobados;

    charts.push(new Chart(document.getElementById("graficoAprobados"), {
        type: 'doughnut',
        data: {
            labels: ['Aprobados', 'Reprobados'],
            datasets: [{
                data: [aprobados, reprobados],
                backgroundColor: [
                    'rgba(54, 162, 235, 0.7)',
                    'rgba(255, 99, 132, 0.7)'
                ],
                borderColor: [
                    'rgba(54, 162, 235, 1)',
                    'rgba(255, 99, 132, 1)'
                ],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false
        }
    }));
}

btnCargar.addEventListener("click", () => {
    const materia = document.getElementById("materia").value;
    const parcial = document.getElementById("parcial").value;

    if (!materia || !parcial) {
        alert("Seleccione una materia y un parcial");
        return;
    }

    // Ajusta aquí según tu estructura de carpetas y nombres de archivo
    const rutaExcel = `excel/${materia}/${parcial} parcial_${materia.toLowerCase()}.xlsx`;

    console.log("Intentando cargar:", rutaExcel);

    fetch(rutaExcel)
        .then((res) => {
            if (!res.ok) throw new Error("Archivo no encontrado");
            return res.arrayBuffer();
        })
        .then((buffer) => {
            const workbook = XLSX.read(buffer, { type: "buffer" });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const datos = XLSX.utils.sheet_to_json(sheet);
            procesarDatos(datos);
        })
        .catch((error) => {
            console.error("Error al cargar el Excel:", error);
            alert("No se pudo cargar el archivo Excel.");
        });
});

btnLimpiar.addEventListener("click", () => {
    location.reload();
});

// Resaltado dinámico del menú
document.addEventListener('DOMContentLoaded', function () {
    const currentPage = location.pathname.split('/').pop();
    const menuLinks = document.querySelectorAll('.nav-link');

    menuLinks.forEach(link => {
        const linkPage = link.getAttribute('href');
        if (linkPage === currentPage) {
            link.classList.add('active');
        } else {
            link.classList.remove('active');
        }
    });
});