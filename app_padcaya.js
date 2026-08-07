require('dotenv').config(); 
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const express = require('express');
const { GoogleGenAI } = require('@google/genai');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Lectura segura de la API Key desde las Variables de Entorno de Render
const apiKeyApiario = process.env.GEMINI_API_KEY || ""; 
const ai = new GoogleGenAI({ apiKey: apiKeyApiario });

const RUTA_HISTORIAL = path.join(__dirname, 'historial_colmenas.json');

// Crear el archivo JSON de historial si no existe
if (!fs.existsSync(RUTA_HISTORIAL)) {
    fs.writeFileSync(RUTA_HISTORIAL, JSON.stringify({}), 'utf-8');
}

const MAPA_MUNICIPIOS = {
    "1": "1 Calendario Floral Padcaya.xlsx",
    "2": "2 Calendario Tarija Yesera.xlsx",
    "3": "3 Calendario Tarija Monte Cercado.xlsx",
    "4": "4 Calendario San Lorenzo Canasmoro.xlsx",
    "5": "5 Calendario San Lorenzo Carachimayo.xlsx",
    "6": "6 Calendario Floral Caraparí.xlsx",
    "7": "7 Calendario Floral Uriondo.xlsx",
    "8": "8 Calendario Floral Villamontes.xlsx",
    "9": "9 Calendario Yacuiba llanura.xlsx",
    "10": "10 Calendario Yacuiba transicion.xlsx",
    "11": "11 Calendario Yacuiba pie de monte.xlsx",
    "12": "12 Calendario Entre rios Salinas.xlsx",
    "13": "13 Calendario Entre rios Los Campos.xlsx",
    "14": "14 Calendario Entre rios Chaco.xlsx",
    "15": "15 Calendario Entre rios Chiquiaca.xlsx"
};

function obtenerSimulacionSENAMHI(idMunicipio) {
    const heladasSimuladas = {
        "1": { min: -1.5, max: 18, desc: "Riesgo de Helada Ligera", helada: true },
        "2": { min: -2.0, max: 17, desc: "Alerta de Helada Moderada", helada: true },
        "4": { min: 0.5, max: 19, desc: "Frío Severo Sin Helada", helada: false },
        "5": { min: 0.0, max: 20, desc: "Frío de Valle", helada: false },
        "6": { min: 8.0, max: 26, desc: "Clima Templado Mansas Brisas", helada: false },
        "8": { min: 12.0, max: 32, desc: "Caluroso / Seco", helada: false }
    };
    return heladasSimuladas[idMunicipio] || { min: 2.0, max: 20, desc: "Clima Variable de Valles", helada: false };
}

function obtenerDatosFloracion(mesActual, archivoExcel) {
    const rutaExcel = path.join(__dirname, archivoExcel);
    try {
        if (!fs.existsSync(rutaExcel)) return { texto: "No se encuentra el archivo.", valores: [0, 0, 0] };
        const workbook = xlsx.readFile(rutaExcel);
        const datos = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
        
        let indiceMes = -1; let filaEncontrada = -1;
        const prefijoMes = mesActual.toString().toUpperCase().trim().substring(0, 3);

        for (let f = 0; f < Math.min(datos.length, 10); f++) {
            const fila = datos[f] || [];
            for (let c = 0; c < fila.length; c++) {
                if (fila[c] && fila[c].toString().toUpperCase().includes(prefijoMes)) {
                    indiceMes = c; filaEncontrada = f; break;
                }
            }
            if (indiceMes !== -1) break;
        }
        
        if (indiceMes !== -1) {
            let fTotales = -1;
            for (let f = 0; f < datos.length; f++) {
                const txt = String(datos[f]?.[0] || datos[f]?.[1] || datos[f]?.[2] || '').toUpperCase();
                if (txt.includes("TOTAL")) { fTotales = f; break; }
            }
            if (fTotales === -1) fTotales = datos.length - 1; 

            const filaFinal = datos[fTotales] || [];
            return {
                texto: `Curva: 1ra[${filaFinal[indiceMes] || 0}], 2da[${filaFinal[indiceMes+1] || 0}], 3ra[${filaFinal[indiceMes+2] || 0}]`,
                valores: [parseFloat(filaFinal[indiceMes])||0, parseFloat(filaFinal[indiceMes+1])||0, parseFloat(filaFinal[indiceMes+2])||0]
            };
        }
        return { texto: "Mes no encontrado.", valores: [0, 0, 0] };
    } catch (e) { return { texto: "Error Excel.", valores: [0, 0, 0] }; }
}

app.get('/', (req, res) => {
    res.send(`
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no">
        <title>Apiario Tarija Pro v2</title>
        <style>
            * { box-sizing: border-box; margin: 0; padding: 0; font-family: sans-serif; }
            body { background-color: #f3f4f6; padding: 10px; color: #1f2937; }
            header { background: linear-gradient(135deg, #d97706, #f59e0b); color: white; padding: 15px; border-radius: 12px; text-align: center; margin-bottom: 15px; }
            .card { background: white; border-radius: 12px; padding: 15px; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
            label { display: block; font-size: 11px; font-weight: bold; color: #6b7280; margin-bottom: 4px; text-transform: uppercase; }
            input, select, textarea { width: 100%; background: #f9fafb; border: 1px solid #d1d5db; border-radius: 8px; padding: 10px; font-size: 14px; margin-bottom: 10px; color: #1f2937; outline: none; }
            button { width: 100%; background-color: #d97706; color: white; padding: 12px; border: none; border-radius: 8px; font-size: 14px; font-weight: bold; cursor: pointer; margin-bottom: 8px; }
            button.btn-secondary { background-color: #4b5563; }
            button.btn-mic { background-color: #dc2626; color: white; margin-bottom: 10px; font-size: 15px; }
            button.btn-mic.grabando { background-color: #16a34a; animation: pulse 1s infinite; }
            button.btn-audit { background-color: #059669; }
            button.btn-clima { background-color: #2563eb; }
            .response-box { display: none; background-color: #fffbeb; border-left: 4px solid #d97706; border-radius: 8px; padding: 12px; margin-top: 12px; font-size: 13px; }
            .historial-box { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px; max-height: 180px; overflow-y: auto; margin-bottom: 10px; font-size: 12px; }
            .item-historial { border-bottom: 1px dashed #cbd5e1; padding: 6px 0; }
            .item-historial:last-child { border-bottom: none; }
            .canvas-scroll-wrapper { width: 100%; overflow-x: auto; background: #fafafa; border-radius: 8px; margin-bottom: 10px; }
            .canvas-container { width: 1100px; padding: 10px; }
            .clima-card { background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 8px; padding: 12px; margin-bottom: 10px; display: none; }
            .audit-table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 12px; }
            .audit-table th { background-color: #f3f4f6; text-align: left; padding: 6px; }
            .audit-table td { padding: 8px 6px; border-bottom: 1px solid #e5e7eb; }
            @keyframes pulse { 0% { opacity: 1; } 50% { opacity: 0.5; } 100% { opacity: 1; } }
        </style>
    </head>
    <body>
        <header>
            <h2>🐝 Apiario Tarija Pro v2</h2>
            <p>Manejo Inteligente y Bitácora de Voz</p>
        </header>

        <div class="card">
            <form id="formConsultar">
                <label>📍 Seleccionar Zona Apícola</label>
                <select id="conMunicipio">
                    <option value="1">1. Padcaya (Calendario Floral)</option>
                    <option value="2">2. Tarija - Yesera</option>
                    <option value="3">3. Tarija - Monte Cercado</option>
                    <option value="4">4. San Lorenzo - Canasmoro</option>
                    <option value="5">5. San Lorenzo - Carachimayo</option>
                    <option value="6">6. Caraparí (Calendario Floral)</option>
                    <option value="7">7. Uriondo (Calendario Floral)</option>
                    <option value="8">8. Villamontes (Calendario Floral)</option>
                    <option value="9">9. Yacuiba - Llanura</option>
                    <option value="10">10. Yacuiba - Transición</option>
                    <option value="11">11. Yacuiba - Pie de Monte</option>
                    <option value="12">12. Entre Ríos - Salinas</option>
                    <option value="13">13. Entre Ríos - Los Campos</option>
                    <option value="14">14. Entre Ríos - Chaco</option>
                    <option value="15">15. Entre Ríos - Chiquiaca</option>
                </select>

                <div class="canvas-scroll-wrapper">
                    <div class="canvas-container"><canvas id="graficoFloracion" width="1080" height="130"></canvas></div>
                </div>

                <div style="display: flex; gap: 8px; margin-bottom: 10px;">
                    <button type="button" class="btn-audit" id="btnAuditar">🔍 Auditar Flora</button>
                    <button type="button" class="btn-clima" id="btnClima">🌤️ Clima SENAMHI</button>
                </div>

                <div id="panelClima" class="clima-card"><div id="contenidoClima"></div></div>

                <div id="contenedorAuditoria" style="display:none; margin-bottom: 10px; background: #f0fdf4; padding: 10px; border-radius: 8px; border: 1px solid #bbf7d0;">
                    <div id="tablaPropuestas"></div>
                </div>

                <label>Mes de Trabajo</label>
                <select id="conMes">
                    <option value="ENERO">Enero</option>
                    <option value="FEBRERO">Febrero</option>
                    <option value="MARZO">Marzo</option>
                    <option value="ABRIL">Abril</option>
                    <option value="MAYO">Mayo</option>
                    <option value="JUNIO">Junio</option>
                    <option value="JULIO">Julio</option>
                    <option value="AGOSTO">Agosto</option>
                    <option value="SEPTIEMBRE">Septiembre</option>
                    <option value="OCTUBRE">Octubre</option>
                    <option value="NOVIEMBRE">Noviembre</option>
                    <option value="DICIEMBRE">Diciembre</option>
                </select>

                <label>Nº Colmena</label>
                <input type="number" id="conColmena" value="1" min="1">

                <button type="button" class="btn-mic" id="btnMicrofono">🎙️ Toca para Dictar por Voz</button>

                <label>Notas Dictadas / Pregunta</label>
                <textarea id="conPregunta" rows="3" placeholder="Toca el micrófono o escribe aquí el reporte de la colmena..."></textarea>
                
                <button type="button" id="btnGuardarNota" style="background-color:#0284c7;">💾 Guardar Reporte en Bitácora</button>
                <button type="submit" class="btn-secondary">Consultar al Ingeniero Virtual</button>
            </form>
            
            <div class="response-box" id="responseBox"></div>

            <div style="margin-top: 15px;">
                <label>📜 Historial Registrado de esta Colmena</label>
                <div class="historial-box" id="cajaHistorial">Cargando historial...</div>
            </div>
        </div>

        <script>
            let reconocedor = null;
            let grabando = false;

            if ('webkitSpeechRecognition' in window || 'SpeechRecognition' in window) {
                const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
                reconocedor = new SpeechRecognition();
                reconocedor.lang = 'es-BO';
                reconocedor.continuous = false;
                reconocedor.interimResults = false;

                reconocedor.onresult = (e) => {
                    const textoDictado = e.results[0][0].transcript;
                    const txtArea = document.getElementById('conPregunta');
                    txtArea.value = (txtArea.value ? txtArea.value + " " : "") + textoDictado;
                };

                reconocedor.onend = () => {
                    grabando = false;
                    const btn = document.getElementById('btnMicrofono');
                    btn.innerText = '🎙️ Toca para Dictar por Voz';
                    btn.classList.remove('grabando');
                };
            }

            document.getElementById('btnMicrofono').addEventListener('click', () => {
                if (!reconocedor) {
                    alert('Tu navegador no soporta reconocimiento de voz. Usa Chrome en Android.');
                    return;
                }
                const btn = document.getElementById('btnMicrofono');
                if (!grabando) {
                    reconocedor.start();
                    grabando = true;
                    btn.innerText = '🔴 Escuchando... ¡Habla ahora!';
                    btn.classList.add('grabando');
                } else {
                    reconocedor.stop();
                }
            });

            async function cargarHistorialColmena() {
                const colmena = document.getElementById('conColmena').value;
                const caja = document.getElementById('cajaHistorial');
                caja.innerHTML = '<i>Cargando registros...</i>';

                try {
                    const res = await fetch('/api/historial/' + colmena);
                    const data = await res.json();
                    if (!data.registros || data.registros.length === 0) {
                        caja.innerHTML = '<span style="color:#9ca3af;">Sin reportes grabados para esta colmena.</span>';
                        return;
                    }
                    let html = '';
                    data.registros.forEach(r => {
                        html += '<div class="item-historial">' +
                                '<small style="color:#0284c7;font-weight:bold;">[' + r.fecha + ']</small><br>' +
                                r.texto +
                                '</div>';
                    });
                    caja.innerHTML = html;
                } catch(e) {
                    caja.innerHTML = '<span style="color:red;">Error al cargar historial.</span>';
                }
            }

            document.getElementById('btnGuardarNota').addEventListener('click', async () => {
                const colmena = document.getElementById('conColmena').value;
                const texto = document.getElementById('conPregunta').value;
                if (!texto.trim()) return alert('Primero dicta o escribe una nota.');

                const res = await fetch('/api/historial/guardar', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ colmena, texto })
                });
                const data = await res.json();
                if (data.exito) {
                    document.getElementById('conPregunta').value = '';
                    cargarHistorialColmena();
                }
            });

            document.getElementById('btnClima').addEventListener('click', async () => {
                const id = document.getElementById('conMunicipio').value;
                const btn = document.getElementById('btnClima');
                const panel = document.getElementById('panelClima');
                const contenedor = document.getElementById('contenidoClima');
                
                btn.innerText = '📡 Conectando...'; btn.disabled = true;
                panel.style.display = 'block';
                contenedor.innerHTML = '<i>Obteniendo reporte SENAMHI e IA...</i>';

                try {
                    const res = await fetch('/api/alerta-clima?id=' + id);
                    const data = await res.json();
                    
                    let htmlAlert = data.clima.helada 
                        ? '<span style="color:#dc2626;font-weight:bold;">❄️ ALERTA DE HELADA DETECTADA</span>' 
                        : '<span style="color:#16a34a;font-weight:bold;">☀️ Condición Estable</span>';

                    contenedor.innerHTML = 
                        '<div style="margin-bottom:8px;">' +
                            '<b>Estación SENAMHI:</b> ' + data.clima.desc + '<br>' +
                            '<b>Temperaturas:</b> Mín: ' + data.clima.min + '°C | Máx: ' + data.clima.max + '°C<br>' +
                            '<b>Estado:</b> ' + htmlAlert +
                        '</div>' +
                        '<hr style="border:0; border-top:1px solid #bfdbfe; margin:8px 0;">' +
                        '<div style="font-size:12px; color:#1e3a8a;">' +
                            '<b>💡 Diagnóstico del Ingeniero Virtual:</b><br>' +
                            data.recomendacion_ia +
                        '</div>';
                } catch(e) {
                    contenedor.innerHTML = '<span style="color:red;">Error al conectar con la estación meteorológica.</span>';
                } finally {
                    btn.innerText = '🌤️ Clima SENAMHI'; btn.disabled = false;
                }
            });

            document.getElementById('btnAuditar').addEventListener('click', async () => {
                const id = document.getElementById('conMunicipio').value;
                const btn = document.getElementById('btnAuditar');
                const panel = document.getElementById('contenedorAuditoria');
                const tablaDiv = document.getElementById('tablaPropuestas');
                btn.innerText = '⚡ Analizando...'; btn.disabled = true;

                try {
                    const res = await fetch('/api/auditar-flora?id=' + id);
                    const data = await res.json();
                    if(!data.propuestas || data.propuestas.length === 0) {
                        tablaDiv.innerHTML = '<p style="color:#166534;">✅ Curva en total armonía botánica regional.</p>';
                    } else {
                        let html = '<table class="audit-table"><thead><tr><th>Planta</th><th>Ajuste</th><th>Acción</th></tr></thead><tbody>';
                        data.propuestas.forEach(p => {
                            html += '<tr>' +
                                '<td><b>' + p.planta + '</b><br><small style="color:#6b7280">' + p.justificacion + '</small></td>' +
                                '<td style="color:#b45309;font-weight:bold;">' + p.nueva_calificacion + '</td>' +
                                '<td><button style="padding:4px;font-size:10px;margin:0;background:#10b981;color:white;border:none;" type="button" onclick="alert(\\'Sincronización registrada.\\')">OK</button></td>' +
                            '</tr>';
                        });
                        html += '</tbody></table>'; tablaDiv.innerHTML = html;
                    }
                    panel.style.display = 'block';
                } catch(e) { 
                    alert('Error en auditoría.'); 
                } finally { 
                    btn.innerText = '🔍 Auditar Flora'; btn.disabled = false; 
                }
            });

            document.getElementById('formConsultar').addEventListener('submit', async (e) => {
                e.preventDefault();
                const box = document.getElementById('responseBox');
                box.style.display = 'block'; box.innerHTML = '⏳ Consultando al Ingeniero...';
                const res = await fetch('/api/consultar', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        idMunicipio: document.getElementById('conMunicipio').value,
                        mes: document.getElementById('conMes').value,
                        colmena: document.getElementById('conColmena').value,
                        pregunta: document.getElementById('conPregunta').value
                    })
                });
                const data = await res.json(); box.innerHTML = data.respuesta;
            });

            function dibujarGraficoAnual(etiquetas, valores) {
                const canvas = document.getElementById('graficoFloracion');
                const ctx = canvas.getContext('2d');
                ctx.clearRect(0, 0, canvas.width, canvas.height);
                ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
                if(!valores || valores.length === 0) return;
                const maxVal = Math.max(...valores, 10);
                const puntos = valores.map((val, i) => ({ x: 30 + i * 28, y: canvas.height - 30 - (val / maxVal * 80) }));
                ctx.beginPath(); ctx.moveTo(puntos[0].x, puntos[0].y);
                for(let i = 1; i < puntos.length; i++) ctx.lineTo(puntos[i].x, puntos[i].y);
                ctx.strokeStyle = '#d97706'; ctx.lineWidth = 2.5; ctx.stroke();
                puntos.forEach((p, i) => {
                    ctx.beginPath(); ctx.arc(p.x, p.y, 3.5, 0, 2 * Math.PI); ctx.fillStyle = '#b45309'; ctx.fill();
                    if(i % 3 === 0) {
                        ctx.fillStyle = '#4b5563'; ctx.font = '9px sans-serif';
                        ctx.fillText(etiquetas[i].split(' ')[0], p.x, canvas.height - 10);
                    }
                });
            }

            async function cargarCurvaAnual() {
                const id = document.getElementById('conMunicipio').value;
                const res = await fetch('/api/floracion-anual?id=' + id);
                const data = await res.json();
                if(data.valores) dibujarGraficoAnual(data.etiquetas, data.valores);
            }

            document.getElementById('conColmena').addEventListener('change', cargarHistorialColmena);
            document.getElementById('conMunicipio').addEventListener('change', () => {
                document.getElementById('contenedorAuditoria').style.display = 'none';
                document.getElementById('panelClima').style.display = 'none';
                cargarCurvaAnual();
            });

            window.addEventListener('DOMContentLoaded', () => {
                cargarCurvaAnual();
                cargarHistorialColmena();
            });
        </script>
    </body>
    </html>`);
});

app.get('/api/historial/:colmena', (req, res) => {
    try {
        const data = JSON.parse(fs.readFileSync(RUTA_HISTORIAL, 'utf-8'));
        const lista = data[req.params.colmena] || [];
        res.json({ registros: lista.slice().reverse() });
    } catch (e) {
        res.json({ registros: [] });
    }
});

app.post('/api/historial/guardar', (req, res) => {
    try {
        const { colmena, texto } = req.body;
        const data = JSON.parse(fs.readFileSync(RUTA_HISTORIAL, 'utf-8'));
        
        if (!data[colmena]) data[colmena] = [];
        
        const fechaHora = new Date().toLocaleString('es-BO', { timeZone: 'America/La_Paz' });
        data[colmena].push({ fecha: fechaHora, texto: texto });

        fs.writeFileSync(RUTA_HISTORIAL, JSON.stringify(data, null, 2), 'utf-8');
        res.json({ exito: true });
    } catch (e) {
        res.status(500).json({ exito: false, error: e.message });
    }
});

// ALGORITMO BÚSQUEDA DINÁMICA DE CURVA FLORAL ANUAL (Soporta todos los archivos Excel)
app.get('/api/floracion-anual', (req, res) => {
    try {
        const idMunicipio = req.query.id || "1";
        const archivoExcel = MAPA_MUNICIPIOS[idMunicipio];
        const rutaCompleta = path.join(__dirname, archivoExcel);

        if (!fs.existsSync(rutaCompleta)) {
            return res.status(404).json({ error: "Archivo no encontrado" });
        }

        const workbook = xlsx.readFile(rutaCompleta);
        const datos = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });

        const mesesAnio = [
            { n: "ENE", p: "ENE" }, { n: "FEB", p: "FEB" }, { n: "MAR", p: "MAR" },
            { n: "ABR", p: "ABR" }, { n: "MAY", p: "MAY" }, { n: "JUN", p: "JUN" },
            { n: "JUL", p: "JUL" }, { n: "AGO", p: "AGO" }, { n: "SEP", p: "SEP" },
            { n: "OCT", p: "OCT" }, { n: "NOV", p: "NOV" }, { n: "DIC", p: "DIC" }
        ];

        let etiquetas = []; let valores = []; let fTotales = -1;

        // Buscar la fila de TOTALES dinámicamente
        for (let f = 0; f < datos.length; f++) {
            const txt = String(datos[f]?.[0] || datos[f]?.[1] || datos[f]?.[2] || '').toUpperCase();
            if (txt.includes("TOTAL")) { fTotales = f; break; }
        }
        if (fTotales === -1) fTotales = datos.length - 1;
        const filaFinal = datos[fTotales] || [];

        // Buscar en las primeras 10 filas dónde están las cabeceras de los meses
        mesesAnio.forEach(mes => {
            etiquetas.push(`${mes.n} 1`, `${mes.n} 2`, `${mes.n} 3`);
            let cIdx = -1;

            for (let f = 0; f < Math.min(datos.length, 10); f++) {
                const fila = datos[f] || [];
                for (let c = 0; c < fila.length; c++) {
                    if (String(fila[c] || '').toUpperCase().includes(mes.p)) {
                        cIdx = c;
                        break;
                    }
                }
                if (cIdx !== -1) break;
            }

            if (cIdx !== -1) {
                valores.push(
                    parseFloat(filaFinal[cIdx]) || 0,
                    parseFloat(filaFinal[cIdx + 1]) || 0,
                    parseFloat(filaFinal[cIdx + 2]) || 0
                );
            } else { 
                valores.push(0, 0, 0); 
            }
        });

        res.json({ etiquetas, valores });
    } catch (e) { 
        res.status(500).json({ error: e.message }); 
    }
});

app.get('/api/alerta-clima', async (req, res) => {
    try {
        const id = req.query.id || "1";
        const clima = obtenerSimulacionSENAMHI(id);

        const promptClima = `Eres un experto apícola e ingeniero agrónomo en Tarija, Bolivia.
        Se registran estos datos climáticos para la zona:
        - Mínima: ${clima.min}°C, Máxima: ${clima.max}°C.
        - Estado: ${clima.desc}.
        - Helada Activa: ${clima.helada ? 'SÍ' : 'NO'}.

        Proporciona un diagnóstico técnico brevísimo (2 o 3 viñetas) indicando:
        1. Cómo afecta esto el pecoreo y la floración local.
        2. Acción preventiva inmediata requerida en las colmenas.`;

        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: promptClima
        });

        res.json({ clima, recomendacion_ia: response.text });
    } catch (e) {
        res.json({ clima: { min: 0, max: 0, desc: "Error clima", helada: false }, recomendacion_ia: "No se pudo obtener el diagnóstico del Ingeniero Virtual." });
    }
});

app.post('/api/consultar', async (req, res) => {
    const { idMunicipio, mes, pregunta } = req.body;
    const info = obtenerDatosFloracion(mes, MAPA_MUNICIPIOS[idMunicipio || "1"]);
    try {
        const response = await ai.models.generateContent({
            model: 'gemini-1.5-flash',
            contents: [{ role: 'user', parts: [{ text: `Eres experto apícola en Tarija, Bolivia. Info del mes actual: ${info.texto}. Responde de forma clara y directa: ${pregunta}` }] }]
        });
        res.json({ respuesta: response.text });
    } catch (e) { 
        res.json({ respuesta: "Error al conectar con la IA: " + e.message }); 
    }
});

app.get('/api/auditar-flora', async (req, res) => {
    try {
        const id = req.query.id || "1";
        const archivoExcel = MAPA_MUNICIPIOS[id];
        const workbook = xlsx.readFile(path.join(__dirname, archivoExcel));
        const datos = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });

        let plantas = [];
        for (let f = 3; f < datos.length; f++) {
            const nom = datos[f]?.[0] || datos[f]?.[1];
            if (nom && !nom.toUpperCase().includes("TOTAL") && !nom.toUpperCase().includes("MES")) {
                plantas.push(nom.trim());
            }
        }
        plantas = [...new Set(plantas)].slice(0, 4);

        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: [{ role: 'user', parts: [{ text: `Eres un PhD en botánica apícola especialista en el departamento de Tarija, Bolivia. Analiza estas plantas locales extraídas de los registros: ${plantas.join(", ")}. Devuelve única y exclusivamente un arreglo JSON limpio (sin bloques markdown \`\`\`json ni texto introductorio) bajo este formato estricto: [{"planta":"Nombre de la Planta","nueva_calificacion":"Excelente/Óptima/Saturada","justificacion":"Explicación corta de su comportamiento melífero local"}]` }] }]
        });

        let txt = response.text.trim();
        if (txt.startsWith("```json")) txt = txt.substring(7);
        if (txt.endsWith("```")) txt = txt.substring(0, txt.length - 3);

        res.json({ propuestas: JSON.parse(txt.trim()) });
    } catch (e) { 
        res.json({ propuestas: [] }); 
    }
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n🚀 ¡SISTEMA APÍCOLA INTEGRADO Y LISTO! En puerto ${PORT}`);
});