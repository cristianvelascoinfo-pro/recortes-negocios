/* ==============================================================================
   [SECCIÓN 1 - CONFIGURACIÓN DE ELEMENTOS Y EVENTOS INICIALES]
   (Propósito: Captura los elementos del DOM y asigna eventos principales)
   ============================================================================== */
document.addEventListener("DOMContentLoaded", () => {
    const btnBuscar = document.getElementById("btnBuscar");
    const newsGrid = document.getElementById("newsGrid");
    const resultsPanel = document.getElementById("resultsPanel");
    const tabContent = document.getElementById("tabContent");
    const btnDescargar = document.getElementById("btnDescargar");

    let archivoGeneradoUrl = "";

    /* ==============================================================================
       [SECCIÓN 2 - PETICIÓN DE BÚSQUEDA DE NOTICIAS AL BACKEND]
       (Propósito: Llama al endpoint /api/buscar y renderiza las 3 tarjetas estilo Google Preview)
       ============================================================================== */
    btnBuscar.addEventListener("click", async () => {
        btnBuscar.disabled = true;
        btnBuscar.innerText = "⏳ BUSCANDO NOTICIAS EN VIVO...";
        newsGrid.innerHTML = "";
        resultsPanel.classList.remove("active");
        archivoGeneradoUrl = "";

        try {
            const response = await fetch("/api/buscar", { method: "POST" });
            const noticias = await response.json();

            btnBuscar.disabled = false;
            btnBuscar.innerText = "🔍 BUSCAR 3 NOTICIAS EN VIVO";

            if (!noticias || noticias.length === 0) {
                alert("No se encontraron noticias recientes. Inténtalo de nuevo.");
                return;
            }

            noticias.forEach((noticia, index) => {
                const card = document.createElement("div");
                card.className = "news-card";
                card.innerHTML = `
                    <div class="source">${noticia.medio || 'MEDIO DIGITAL'}</div>
                    <h3>${noticia.titulo}</h3>
                    <p>${noticia.resumen_corto}</p>
                    <button class="btn-neon btn-seleccionar" style="width: 100%; margin-top: 1rem;">SELECCIONAR NOTICIA ${index + 1}</button>
                `;

                const btnSeleccionar = card.querySelector(".btn-seleccionar");
                btnSeleccionar.addEventListener("click", (e) => {
                    e.stopPropagation();
                    procesarNoticia(noticia);
                });

                card.addEventListener("click", () => procesarNoticia(noticia));
                newsGrid.appendChild(card);
            });

        } catch (error) {
            console.error("Error al buscar noticias:", error);
            alert("Ocurrió un error al conectar con el motor de búsqueda.");
            btnBuscar.disabled = false;
            btnBuscar.innerText = "🔍 BUSCAR 3 NOTICIAS EN VIVO";
        }
    });

    /* ==============================================================================
       [SECCIÓN 3 - PROCESAMIENTO DE NOTICIA SELECCIONADA Y GENERACIÓN WORD]
       (Propósito: Envia la noticia elegida al backend para generar el resumen, análisis y .docx)
       ============================================================================== */
    async function procesarNoticia(noticia) {
        const numRecorte = document.getElementById("numRecorte").value || 2;

        resultsPanel.classList.add("active");
        btnDescargar.disabled = true;
        btnDescargar.style.opacity = "0.5";
        btnDescargar.innerText = "⏳ GENERANDO DOCUMENTO WORD...";

        tabContent.innerHTML = "<p style='color: var(--primary-neon);'>⏳ Procesando noticia con Gemini AI, tomando captura HD del encabezado y creando el documento Word...</p>";

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

                tabContent.innerHTML = `
                    <h3>✅ Vista Previa del Análisis (15 Líneas):</h3>
                    <p style="margin-top: 1rem; line-height: 1.6;">${resultado.analisis}</p>
                `;
            } else {
                tabContent.innerHTML = `<p style="color: red;">Error al procesar: ${resultado.error}</p>`;
                btnDescargar.innerText = "❌ ERROR AL GENERAR ARCHIVO";
            }

        } catch (error) {
            console.error("Error al procesar la noticia:", error);
            tabContent.innerHTML = "<p style='color: red;'>Error en el servidor al generar el informe.</p>";
            btnDescargar.innerText = "❌ ERROR EN EL SERVIDOR";
        }
    }

    /* ==============================================================================
       [SECCIÓN 4 - DESCARGA DEL ARCHIVO COMPILADO .DOCX]
       (Propósito: Redirecciona para descargar el archivo Word listo para imprimir)
       ============================================================================== */
    btnDescargar.addEventListener("click", () => {
        if (archivoGeneradoUrl) {
            window.location.href = archivoGeneradoUrl;
        } else {
            alert("El documento aún se está generando. Por favor espera unos segundos.");
        }
    });
});