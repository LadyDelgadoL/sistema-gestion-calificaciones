document.addEventListener('DOMContentLoaded', () => {
    const btnCargar = document.getElementById('btn-cargar-excel');
    const btnLimpiar = document.getElementById('btn-limpiar');
    const btnBuscar = document.getElementById('btn-buscar');
    const inputBusqueda = document.getElementById('busqueda-apellido');
    const selectMateria = document.getElementById('materia');
    const selectParcial = document.getElementById('parcial');
    const tablaBody = document.querySelector('#tabla-estudiantes tbody');

    // Cargar datos desde sessionStorage si existen
    let datosActuales = JSON.parse(sessionStorage.getItem('datosEstudiantes')) || [];
    let filtrosActuales = JSON.parse(sessionStorage.getItem('filtrosEstudiantes')) || {};

    // Restaurar filtros si existen
    if (filtrosActuales.materia) {
        selectMateria.value = filtrosActuales.materia;
    }
    if (filtrosActuales.parcial) {
        selectParcial.value = filtrosActuales.parcial;
    }
    if (filtrosActuales.busqueda) {
        inputBusqueda.value = filtrosActuales.busqueda;
    }

    // Función para guardar estado actual
    const guardarEstado = () => {
        sessionStorage.setItem('datosEstudiantes', JSON.stringify(datosActuales));
        sessionStorage.setItem('filtrosEstudiantes', JSON.stringify({
            materia: selectMateria.value,
            parcial: selectParcial.value,
            busqueda: inputBusqueda.value
        }));
    };

    // Función para cargar datos desde Excel
    const cargarDatosExcel = async () => {
        const materia = selectMateria.value;
        const parcial = selectParcial.value;

        if (!materia || !parcial) {
            alert('Por favor seleccione una materia y un parcial.');
            return;
        }

        const nombreArchivo = `${parcial} parcial_${materia.toLowerCase()}.xlsx`;
        const ruta = `excel/${materia}/${nombreArchivo}`;

        console.log("📁 Intentando cargar archivo:", ruta);

        try {
            btnCargar.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Cargando...';
            btnCargar.disabled = true;

            const response = await fetch(ruta);

            if (!response.ok) {
                throw new Error(`No se encontró el archivo: ${nombreArchivo}`);
            }

            const data = await response.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });
            const hoja = workbook.Sheets[workbook.SheetNames[0]];
            datosActuales = XLSX.utils.sheet_to_json(hoja);

            console.log("✅ Datos cargados correctamente:", datosActuales);
            llenarTabla(datosActuales);
            guardarEstado();

        } catch (error) {
            console.error("❌ Error al cargar el archivo Excel:", error);
            alert(`Error al cargar el archivo:\n${error.message}`);
        } finally {
            btnCargar.innerHTML = '<i class="fas fa-file-excel mr-2"></i>Cargar Excel';
            btnCargar.disabled = false;
        }
    };

    // Modifica la función llenarTabla para enviar correctamente los parámetros
    function llenarTabla(estudiantes) {
        console.log("🧪 Datos recibidos para mostrar:", estudiantes);

        const cuerpo = document.querySelector('#tabla-estudiantes tbody');
        cuerpo.innerHTML = '';

        estudiantes.forEach((est, index) => {
            const promedio = est['Promedio'] !== undefined ? Number(est['Promedio']).toFixed(2) : "0.00";
            const nombreCompleto = est['Nombres y Apellidos'] || '';
            const materia = document.getElementById('materia').value;
            const parcial = document.getElementById('parcial').value;

            // Usar encodeURIComponent para nombres con caracteres especiales
            const nombreCodificado = encodeURIComponent(nombreCompleto);

            const fila = document.createElement('tr');
            fila.innerHTML = `
            <td>${nombreCompleto}</td>
            <td>
                ${est['Gestión Formativa'] || 0}
                <a href="detalle-estudiante.html?nombre=${nombreCodificado}&materia=${materia}&parcial=${parcial}&gestion=GF" 
                    class="lupa-detalle" title="Ver detalle GF">
                    <i class="fas fa-search"></i>
                </a>
            </td>
            <td>
                ${est['Gestión Práctica'] || 0}
                <a href="detalle-estudiante.html?nombre=${nombreCodificado}&materia=${materia}&parcial=${parcial}&gestion=GP" 
                    class="lupa-detalle" title="Ver detalle GP">
                    <i class="fas fa-search"></i>
                </a>
            </td>
            <td>${est['Examen'] || 0}</td>
            <td>${promedio}</td>
        `;
            cuerpo.appendChild(fila);
        });
    }

    // Función para limpiar la tabla y filtros
    const limpiarDatos = () => {
        datosActuales = [];
        sessionStorage.removeItem('datosEstudiantes');
        sessionStorage.removeItem('filtrosEstudiantes');
        tablaBody.innerHTML = '';
        selectMateria.value = '';
        selectParcial.value = '';
        inputBusqueda.value = '';

        const fila = document.createElement('tr');
        fila.innerHTML = '<td colspan="5" class="text-center">No hay datos cargados</td>';
        tablaBody.appendChild(fila);

        console.log("🗑️ Todos los datos han sido limpiados");
    };

    // Función para buscar estudiantes por apellido
    const buscarPorApellido = () => {
        if (datosActuales.length === 0) {
            alert("Primero cargue un archivo Excel");
            return;
        }

        const filtro = inputBusqueda.value.trim().toLowerCase();

        if (!filtro) {
            llenarTabla(datosActuales);
            return;
        }

        const resultados = datosActuales.filter(est => {
            const nombreCompleto = (est['Nombres y Apellidos'] || est['Estudiante'] || '').toLowerCase();
            return nombreCompleto.includes(filtro);
        });

        if (resultados.length === 0) {
            alert("No se encontraron estudiantes con ese apellido");
            return;
        }

        llenarTabla(resultados);
        guardarEstado();
        console.log("🔍 Búsqueda realizada:", filtro);
    };

    // Event Listeners
    btnCargar.addEventListener('click', cargarDatosExcel);
    btnLimpiar.addEventListener('click', limpiarDatos);
    btnBuscar.addEventListener('click', buscarPorApellido);

    inputBusqueda.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            buscarPorApellido();
        }
    });

    // Cargar datos almacenados al inicio si existen
    if (datosActuales.length > 0) {
        llenarTabla(datosActuales);
        console.log("🔄 Datos recuperados de sessionStorage");
    } else {
        limpiarDatos();
    }
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