document.addEventListener("DOMContentLoaded", () => {
    const btnCargar = document.getElementById("btn-cargar-excel");
    const btnLimpiar = document.getElementById("btn-limpiar");
    const materiaSelect = document.getElementById("materia");
    const parcialSelect = document.getElementById("parcial");
    const contenedorContenido = document.getElementById("contenido");

    btnCargar.addEventListener("click", () => {
        const materia = materiaSelect.value;
        const parcial = parcialSelect.value;

        if (!materia || !parcial) {
            alert("Seleccione una materia y un parcial");
            return;
        }

        const ruta = `excel/${materia}/${parcial} parcial_${materia.toLowerCase()}.xlsx`;
        console.log("Cargando Excel:", ruta);

        fetch(ruta)
            .then(res => {
                if (!res.ok) throw new Error("Archivo no encontrado");
                return res.arrayBuffer();
            })
            .then(buffer => {
                const workbook = XLSX.read(buffer, { type: "buffer" });
                const datos = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
                mostrarCalificaciones(datos);
                contenedorContenido.style.display = "block"; 
            })
            .catch(err => {
                console.error("Error al cargar Excel:", err);
                alert("No se pudo cargar el archivo.");
            });
    });

    btnLimpiar.addEventListener("click", () => {
        contenedorContenido.style.display = "none"; 
        materiaSelect.value = "";
        parcialSelect.value = "";
    });

    function mostrarCalificaciones(estudiantes) {
        if (estudiantes.length === 0) return;

        const promGF = promedio(estudiantes.map(e => e["Gestión Formativa"]));
        const promGP = promedio(estudiantes.map(e => e["Gestión Práctica"]));
        const promEX = promedio(estudiantes.map(e => e["Examen"]));
        const promGlobal = promedio(estudiantes.map(e => e["Promedio"]));

        document.getElementById("prom-gf").textContent = promGF.toFixed(2);
        document.getElementById("prom-gp").textContent = promGP.toFixed(2);
        document.getElementById("prom-ex").textContent = promEX.toFixed(2);
        document.getElementById("prom-global").textContent = promGlobal.toFixed(2);

        // Top 3
        const top = [...estudiantes].sort((a, b) => b["Promedio"] - a["Promedio"]).slice(0, 3);
        const tablaTop = document.querySelector("#top-primer-parcial tbody");
        tablaTop.innerHTML = "";

        top.forEach((est, i) => {
            const fila = document.createElement("tr");
            fila.className = `puesto-${i + 1}`;
            fila.innerHTML = `
                <td>${i + 1}</td>
                <td>${est["Nombres y Apellidos"]}</td>
                <td>${est["Gestión Formativa"].toFixed(2)}</td>
                <td>${est["Gestión Práctica"].toFixed(2)}</td>
                <td>${est["Examen"].toFixed(2)}</td>
            `;
            tablaTop.appendChild(fila);
        });

        // Análisis y sugerencia automáticos
        const analisis = document.getElementById("texto-analisis");
        const sugerencia = document.getElementById("texto-sugerencia");
        const mayor = Math.max(promGF, promGP, promEX);
        let tipoMayor = promGP === mayor ? "Gestión Práctica" : promGF === mayor ? "Gestión Formativa" : "Examen";
        let tipoMenor = promGP === Math.min(promGF, promGP, promEX) ? "Gestión Práctica" : promGF === Math.min(promGF, promGP, promEX) ? "Gestión Formativa" : "Examen";

        analisis.innerHTML = `
            Se observa que el promedio más alto corresponde a <strong>${tipoMayor} (${mayor.toFixed(2)})</strong>,
            mientras que <strong>${tipoMenor}</strong> tiene el promedio más bajo <strong>(${Math.min(promGF, promGP, promEX).toFixed(2)})</strong>.
            Esto puede sugerir que los estudiantes se desempeñan mejor en <em>${tipoMayor.toLowerCase()}</em> que en <em>${tipoMenor.toLowerCase()}</em>.
        `;

        sugerencia.textContent = `Se recomienda reforzar los contenidos de ${tipoMenor.toLowerCase()}, ya que es el área con menor rendimiento general.`;
    }

    function promedio(arr) {
        return arr.reduce((a, b) => a + parseFloat(b || 0), 0) / arr.length;
    }
});