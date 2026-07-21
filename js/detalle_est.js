document.addEventListener('DOMContentLoaded', async () => {
    console.group("🔄 Iniciando carga de detalles del estudiante");

    try {
        //Obtener parámetros de la URL
        const params = new URLSearchParams(window.location.search);
        const nombre = decodeURIComponent(params.get('nombre'));
        const materia = params.get('materia');
        const parcial = params.get('parcial');

        console.log("🔍 Parámetros recibidos:", { nombre, materia, parcial });

        if (!nombre || !materia || !parcial) {
            throw new Error("Faltan parámetros en la URL (nombre, materia o parcial)");
        }

        //Construir ruta del archivo Excel
        const nombreArchivo = `${parcial} parcial_${materia.toLowerCase()}.xlsx`;
        const ruta = `excel/${materia}/${nombreArchivo}`;
        console.log("📁 Intentando cargar archivo:", ruta);

        //Carga archivo Excel
        const response = await fetch(ruta);
        if (!response.ok) {
            throw new Error(`No se encontró el archivo: ${nombreArchivo} (Error ${response.status})`);
        }

        const data = await response.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });

        //Verificar hojas disponibles
        console.log("📑 Hojas en el Excel:", workbook.SheetNames);
        if (workbook.SheetNames.length === 0) {
            throw new Error("El archivo Excel no contiene hojas");
        }

        //Procesar primera hoja 
        const hoja = workbook.Sheets[workbook.SheetNames[0]];
        const datos = XLSX.utils.sheet_to_json(hoja, { defval: null });
        console.log("📋 Primeras filas del Excel:", datos.slice(0, 3));

        // 6. Buscar estudiante (insensible a mayúsculas)
        const estudiante = datos.find(e => {
            const nombreExcel = e['Nombres y Apellidos']?.toString().trim();
            return nombreExcel && nombreExcel.toLowerCase() === nombre.toLowerCase();
        });

        if (!estudiante) {
            const nombresDisponibles = datos.slice(0, 5).map(e => e['Nombres y Apellidos']).join(', ');
            throw new Error(`Estudiante "${nombre}" no encontrado. Primeros nombres disponibles: ${nombresDisponibles}...`);
        }

        console.log("🎓 Estudiante encontrado:", estudiante);

        // Muestra información básica
        document.getElementById('nombre-estudiante').textContent = nombre;

        // Carga imagen (con fallback)
        const imgElement = document.getElementById('foto-estudiante');
        const indexEstudiante = datos.findIndex(e =>
            e['Nombres y Apellidos']?.toString().toLowerCase() === nombre.toLowerCase()
        ) + 1;

        imgElement.src = `images/estudiantes/estudiante_${indexEstudiante}.jpg`;
        imgElement.onerror = () => {
            imgElement.src = 'images/estudiantes/default.jpg';
            console.warn("Usando imagen por defecto");
        };

        // Muestra datos numéricos (con validación)
        const formatNumber = (value) => {
            const num = Number(value);
            return !isNaN(num) ? num.toFixed(2) : 'N/A';
        };

        document.getElementById('gestion-practica').textContent =
            formatNumber(estudiante['Gestión Práctica']);
        document.getElementById('gestion-formativa').textContent =
            formatNumber(estudiante['Gestión Formativa']);
        document.getElementById('examen').textContent =
            formatNumber(estudiante['Examen']);
        document.getElementById('promedio').textContent =
            formatNumber(estudiante['Promedio']);

        // Crea gráficos
        crearGraficos(
            parseFloat(estudiante['Gestión Práctica']) || 0,
            parseFloat(estudiante['Gestión Formativa']) || 0,
            parseFloat(estudiante['Examen']) || 0,
            parseFloat(estudiante['Promedio']) || 0
        );

    } catch (error) {
        console.error("❌ Error crítico:", error);
        alert(`Error: ${error.message}`);
        document.getElementById('nombre-estudiante').textContent = "Error al cargar datos";

        // Muestra mensaje de error en los campos
        ['gestion-practica', 'gestion-formativa', 'examen', 'promedio'].forEach(id => {
            document.getElementById(id).textContent = 'Error';
        });
    } finally {
        console.groupEnd();
    }
});

function crearGraficos(gestionPractica, gestionFormativa, examen, promedio) {
    console.log("📊 Preparando gráficos con datos:", {
        gestionPractica, gestionFormativa, examen, promedio
    });

    // Configuración común
    const datos = {
        labels: ['Práctica', 'Formativa', 'Examen', 'Promedio'],
        valores: [gestionPractica, gestionFormativa, examen, promedio],
        colores: [
            'rgba(54, 162, 235, 0.8)',
            'rgba(75, 192, 192, 0.8)',
            'rgba(255, 159, 64, 0.8)',
            'rgba(153, 102, 255, 0.8)'
        ],
        bordes: [
            'rgba(54, 162, 235, 1)',
            'rgba(75, 192, 192, 1)',
            'rgba(255, 159, 64, 1)',
            'rgba(153, 102, 255, 1)'
        ]
    };

    // 1. Gráfico de Barras (evaluado sobre 10)
    try {
        new Chart(
            document.getElementById('graficoBarras').getContext('2d'),
            {
                type: 'bar',
                data: {
                    labels: datos.labels,
                    datasets: [{
                        label: 'Puntuación (sobre 10)',
                        data: datos.valores,
                        backgroundColor: datos.colores,
                        borderColor: datos.bordes,
                        borderWidth: 2,
                        borderRadius: 6,
                        borderSkipped: false
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label: function (context) {
                                    return `${context.label}: ${context.raw}`;
                                }
                            }
                        }
                    },
                    scales: {
                        y: {
                            beginAtZero: true,
                            max: 10,
                            ticks: { stepSize: 1 },
                            grid: { color: 'rgba(0, 0, 0, 0.05)' }
                        },
                        x: {
                            grid: { display: false }
                        }
                    },
                    animation: {
                        duration: 1000,
                        easing: 'easeOutQuart'
                    }
                }
            }
        );
    } catch (error) {
        console.error("Error al crear gráfico de barras:", error);
    }

    // 2. Gráfico Circular (sobre 100%)
    try {
        new Chart(
            document.getElementById('graficoCircular').getContext('2d'),
            {
                type: 'doughnut',
                data: {
                    labels: datos.labels,
                    datasets: [{
                        data: datos.valores,
                        backgroundColor: datos.colores,
                        borderColor: 'rgba(255, 255, 255, 0.8)',
                        borderWidth: 2,
                        cutout: '70%'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'right',
                            labels: {
                                padding: 15,
                                usePointStyle: true,
                                pointStyle: 'circle',
                                boxWidth: 10
                            }
                        },
                        tooltip: {
                            callbacks: {
                                label: function (context) {
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const porcentaje = ((context.raw / total) * 100).toFixed(1);
                                    return `${context.label}: ${porcentaje}%`;
                                }
                            }
                        }

                    },
                    animation: {
                        animateScale: true,
                        animateRotate: true
                    },
                    hover: {
                        mode: 'nearest',
                        intersect: true
                    }
                }
            }
        );
    } catch (error) {
        console.error("Error al crear gráfico circular:", error);
    }
}