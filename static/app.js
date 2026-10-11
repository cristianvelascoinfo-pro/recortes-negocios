/* ==============================================================================
   [SECCIÓN 1 - CÁLCULO DE FECHA Y CONFIGURACIÓN INICIAL]
   (Propósito: Calcula el número de recorte exacto según semanas fijadas desde el Lunes 28 de Septiembre 2026)
   ============================================================================== */
document.addEventListener("DOMContentLoaded", () => {
    const numRecorteInput = document.getElementById("numRecorte");
    const btnBuscar = document.getElementById("btnBuscar");
    const btnReset = document.getElementById("btnReset");
    const newsGrid = document.getElementById("newsGrid");
    const resultsPanel = document.getElementById("resultsPanel");
    const tabContent = document.getElementById("tabContent");
    const btnDescargar = document.getElementById("btnDescargar");

    let archivoGeneradoUrl = "";

    // Lunes 28 de Septiembre de 2026 = Recorte N° 1
    // Lunes 05 de Octubre de 2026 al Lunes 12 de Octubre de 2026 = Recorte N° 2
    function calcularNumeroRecorte() {
        const fechaInicioSemana1 = new Date("2026-09-28T00:00:00");
        const fechaActual = new Date();
        
        const diferenciaMilisegundos = fechaActual - fechaInicioSemana1;
        const diasTranscurridos = Math.floor(diferenciaMilisegundos / (1000 * 60 * 60 * 24));
        const numeroRecorte = Math.floor(diasTranscurridos / 7) + 1;

        numRecorteInput.value = numeroRecorte > 0 ? numeroRecorte : 1;
    }
    calcularNumeroRecorte();

    /* ==============================================================================
       [SECCIÓN 2 - PETICIÓN DE BÚSQUEDA Y BOTÓN RESET]
       (Propósito: Llama al backend con opción de refrescar resultados si no convencen)
       ============================================================================== */
    async function realizarBusqueda(esReset = false) {
        btnBuscar.disabled = true;
        btnReset.disabled = true;
        btnBuscar.innerText = esReset ? "⏳ REFRESCANDO OPCIONES..." : "⏳ BUSCANDO NOTICIAS EN VIVO...";
        newsGrid.innerHTML = "";
        resultsPanel.classList.remove("active");
        archivoGeneradoUrl = "";

        try {
            const response = await fetch("/api/buscar", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ reset: esReset })
            });
            const noticias = await response.json();

            btnBuscar.disabled = false;
            btnReset.disabled = false;
            btnReset.style.display = "inline-block";
            btnBuscar.innerText = "🔍 BUSCAR 3 NOTICIAS EN VIVO";
            btnReset.innerText = "🔄 RESETEAR Y BUSCAR NUEVAS";

            if (!noticias || noticias.length === 0) {
                alert("No se encontraron noticias. Intenta de nuevo.");
                return;
            }

            noticias.forEach((noticia, index) => {
                const card = document.createElement("div");
                card.className = "news-card";
                card.id = `card-${index}`;
                card.innerHTML = `
                    <div class="source">${noticia.medio || 'MEDIO DIGITAL'}</div>
                    <h3>${noticia.titulo}</h3>
                    <p>${noticia.resumen_corto}</p>
                    <button class="btn-neon btn-seleccionar" style="width: 100%; margin-top: 1rem;">SELECCIONAR NOTICIA ${index + 1}</button>
                `;

                const btnSel = card.querySelector(".btn-seleccionar");
                btnSel.addEventListener("click", (e) => {
                    e.stopPropagation();
                    seleccionarYProcesar(noticia, index);
                });

                card.addEventListener("click", () => seleccionarYProcesar(noticia, index));
                newsGrid.appendChild(card);
            });

        } catch (error) {
            console.error("Error al buscar:", error);
            btnBuscar.disabled = false;
            btnReset.disabled = false;
            btnBuscar.innerText = "🔍 BUSCAR 3 NOTICIAS EN VIVO";
        }
    }

    btnBuscar.addEventListener("click", () => realizarBusqueda(false));
    btnReset.addEventListener("click", () => realizarBusqueda(true));

    /* ==============================================================================
       [SECCIÓN 3 - SELECCIÓN, OCULTAMIENTO Y DESPLIEGUE DE GUÍA DE SUBRAYADO]
       (Propósito: Oculta las tarjetas no elegidas y muestra la guía visual con glosario)
       ============================================================================== */
    async function seleccionarYProcesar(noticia, indiceElegido) {
        // Desaparecer las otras tarjetas no seleccionadas de la pantalla
        const cards = document.querySelectorAll(".news-card");
        cards.forEach((card, idx) => {
            if (idx !== indiceElegido) {
                card.style.transition = "all 0.5s ease";
                card.style.opacity = "0";
                setTimeout(() => card.remove(), 500);
            }
        });

        const numRecorte = numRecorteInput.value;
        resultsPanel.classList.add("active");
        btnDescargar.disabled = true;
        btnDescargar.style.opacity = "0.5";
        btnDescargar.innerText = "⏳ GENERANDO DOCUMENTO Y GUÍA DE SUBRAYADO...";

        tabContent.innerHTML = "<p style='color: var(--primary-neon);'>⏳ Extrayendo noticia, generando resaltado de frases clave y redactando análisis de 15 líneas...</p>";

        try {
            const response = await fetch("/api/procesar", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    num_recorte: parseInt(numRecorte),
                    url: noticia.url,
                    titulo: noticia.titulo
                })
            });

            const resultado = await response.json();

            if (resultado.exito) {
                archivoGeneradoUrl = resultado.download_url;
                btnDescargar.disabled = false;
                btnDescargar.style.opacity = "1";
                btnDescargar.innerText = "📥 DESCARGAR DOCUMENTO WORD (.DOCX) LISTO PARA IMPRIMIR";

                let glosarioHtml = resultado.glosario.map(g => `
                    <div class="glosario-card">
                        <strong>📌 ${g.palabra}:</strong> ${g.definición}<br>
                        <small style="color: var(--primary-neon);"><em>Sinónimos:</em> ${g.sinonimos}</small><br>
                        <small style="color: #ff416c;"><em>Antónimos:</em> ${g.antonimos}</small>
                    </div>
                `).join("");

                tabContent.innerHTML = `
                    <h3 style="color: var(--primary-neon);">🖍️ 1. GUÍA PARA SUBRAYAR FÍSICAMENTE EN TU HOJA IMPRESA:</h3>
                    <p style="margin-top: 0.5rem; line-height: 1.8; background: rgba(0,0,0,0.3); padding: 1rem; border-radius: 8px;">
                        ${resultado.texto_subrayado}
                    </p>
                    <br>
                    <h3 style="color: var(--primary-neon);">📊 2. ANÁLISIS ACADÉMICO DE LA NOTICIA (15 LÍNEAS / ~177 PALABRAS):</h3>
                    <p style="margin-top: 0.5rem; line-height: 1.6; text-align: justify;">${resultado.analisis}</p>
                    <br>
                    <h3 style="color: var(--primary-neon);">📚 3. GLOSARIO ACADÉMICO (5 PALABRAS CLAVE):</h3>
                    ${glosarioHtml}
                `;
            } else {
                tabContent.innerHTML = `<p style="color: red;">Error al procesar: ${resultado.error}</p>`;
            }

        } catch (error) {
            console.error("Error al procesar:", error);
            tabContent.innerHTML = "<p style='color: red;'>Error en el servidor al generar el informe.</p>";
        }
    }

    /* ==============================================================================
       [SECCIÓN 4 - EVENTO DE DESCARGA DE DOCUMENTO]
       (Propósito: Descarga el archivo Word final estructurado)
       ============================================================================== */
    btnDescargar.addEventListener("click", () => {
        if (archivoGeneradoUrl) {
            window.location.href = archivoGeneradoUrl;
        } else {
            alert("El documento aún se está generando. Por favor espera unos segundos.");
        }
    });
});