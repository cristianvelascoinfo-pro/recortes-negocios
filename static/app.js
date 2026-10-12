/* ==============================================================================
   [SECCIÓN 1 - CONFIGURACIÓN E INICIALIZACIÓN DE EVENTOS]
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
       [SECCIÓN 2 - TARJETAS DE PREVIEW CON ENLACE DIRECTO A LA NOTICIA]
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
                body: JSON.stringify({ reset: esReset, timestamp: Date.now() })
            });
            const noticias = await response.json();

            btnBuscar.disabled = false;
            btnReset.disabled = false;
            btnReset.style.display = "inline-block";
            btnBuscar.innerText = "🔍 BUSCAR 3 NOTICIAS EN VIVO";
            btnReset.innerText = "🔄 RESETEAR Y BUSCAR NUEVAS";

            if (!noticias || noticias.length === 0) {
                alert("No se encontraron noticias. Inténtalo de nuevo.");
                return;
            }

            noticias.forEach((noticia, index) => {
                const card = document.createElement("div");
                card.className = "news-card";
                card.innerHTML = `
                    <img src="${noticia.imagen_url || 'https://images.unsplash.com/photo-1460925895917-afdab827c52f?w=600&auto=format&fit=crop&q=80'}" alt="Imagen noticia">
                    <div class="news-card-content">
                        <div class="source">📰 FUENTE: ${noticia.medio || 'MEDIO DIGITAL'}</div>
                        <h3>${noticia.titulo}</h3>
                        <p>${noticia.resumen_corto}</p>
                        
                        <!-- ENLACE DIRECTO EN LA TARJETA DE PREVIEW -->
                        <div style="margin: 0.75rem 0; padding: 0.5rem; background: rgba(0, 242, 254, 0.08); border: 1px solid var(--primary-neon); border-radius: 6px;">
                            <a href="${noticia.url}" target="_blank" rel="noopener noreferrer" style="color: var(--primary-neon); font-size: 0.82rem; font-weight: bold; text-decoration: underline; display: inline-block;">
                                🌐 Abrir noticia original en vivo ↗
                            </a>
                        </div>

                        <button class="btn-neon btn-seleccionar" style="width: 100%;">SELECCIONAR NOTICIA ${index + 1}</button>
                    </div>
                `;

                const btnSel = card.querySelector(".btn-seleccionar");
                btnSel.addEventListener("click", (e) => {
                    e.stopPropagation();
                    seleccionarYProcesar(noticia, index);
                });

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
       [SECCIÓN 3 - RESULTADOS Y GENERACIÓN]
       ============================================================================== */
    async function seleccionarYProcesar(noticia, indiceElegido) {
        const cards = document.querySelectorAll(".news-card");
        cards.forEach((card, idx) => {
            if (idx !== indiceElegido) {
                card.style.transition = "all 0.4s ease";
                card.style.opacity = "0";
                setTimeout(() => card.remove(), 400);
            }
        });

        const numRecorte = numRecorteInput.value;
        resultsPanel.classList.add("active");
        btnDescargar.disabled = true;
        btnDescargar.style.opacity = "0.5";
        btnDescargar.innerText = "⏳ GENERANDO DOCUMENTO Y GUÍA DE SUBRAYADO...";

        tabContent.innerHTML = "<p style='color: var(--primary-neon);'>⏳ Procesando la noticia seleccionada, aplicando plantilla guía y redactando ~410 palabras...</p>";

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
                    <div style="background: rgba(0, 242, 254, 0.12); border: 1px solid var(--primary-neon); padding: 1rem 1.25rem; border-radius: 10px; margin-bottom: 1.5rem; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px;">
                        <div>
                            <strong style="color: var(--primary-neon); font-size: 0.95rem;">🌐 ENLACE DIRECTO A LA NOTICIA ESPECÍFICA:</strong><br>
                            <span style="color: #ffffff; font-size: 0.88rem; word-break: break-all;">${noticia.url}</span>
                        </div>
                        <a href="${noticia.url}" target="_blank" rel="noopener noreferrer" class="btn-neon" style="font-size: 0.8rem; padding: 0.5rem 1rem; text-decoration: none;">ABRIR ENLACE ↗</a>
                    </div>

                    <h3 style="color: var(--primary-neon);">🖍️ 1. GUÍA PARA SUBRAYAR FÍSICAMENTE EN TU HOJA IMPRESA:</h3>
                    <div id="subrayadoContainer" style="margin-top: 0.5rem; line-height: 1.8; background: rgba(0,0,0,0.35); padding: 1.25rem; border-radius: 8px; text-align: justify; border: 1px solid rgba(255,255,255,0.05);"></div>
                    <br>
                    <h3 style="color: var(--primary-neon);">📊 2. ANÁLISIS ACADÉMICO DE LA NOTICIA (15 LÍNEAS / ~177 PALABRAS):</h3>
                    <p style="margin-top: 0.5rem; line-height: 1.6; text-align: justify; background: rgba(0,0,0,0.2); padding: 1rem; border-radius: 8px;">${resultado.analisis}</p>
                    <br>
                    <h3 style="color: var(--primary-neon);">📚 3. GLOSARIO ACADÉMICO (5 PALABRAS CLAVE - MIN 2 LÍNEAS):</h3>
                    ${glosarioHtml}
                `;

                document.getElementById("subrayadoContainer").innerHTML = resultado.texto_subrayado;

            } else {
                tabContent.innerHTML = `<p style="color: red;">Error al procesar: ${resultado.error}</p>`;
            }

        } catch (error) {
            console.error("Error al procesar:", error);
            tabContent.innerHTML = "<p style='color: red;'>Error en el servidor al generar el informe.</p>";
        }
    }

    btnDescargar.addEventListener("click", () => {
        if (archivoGeneradoUrl) {
            window.location.href = archivoGeneradoUrl;
        } else {
            alert("El documento aún se está generando. Por favor espera unos segundos.");
        }
    });
});