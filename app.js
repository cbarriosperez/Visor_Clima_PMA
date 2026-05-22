/**
 * ==========================================================================
 * LÓGICA GENERAL - VISOR DE CLIMA PMA (DASHBOARD CLIMÁTICO DINÁMICO)
 * ==========================================================================
 */

// Mapeo de localidades a sus archivos Excel correspondientes
const EXCEL_FILES = {
    "Nuevo Paraíso": "Datos_clima/Precipitacion_y_temperatura_Nuevo_Paraiso.xlsx",
    "Santa Sofía": "Datos_clima/Precipitacion_y_temperatura_Santa_Sofia.xlsx",
    "San Miguel": "Datos_clima/Precipitacion_y_temperatura_San_Miguel.xlsx",
    "La Primavera": "Datos_clima/Precipitacion_y_temperatura_La_Primavera.xlsx"
};

// Variables de Estado de la Aplicación
let appState = {
    locations: [],
    selectedLocation: null,
    dailyData: [],      // Datos diarios cargados del Excel actual
    weeklyData: [],     // Datos agregados semanalmente
    monthlyData: [],    // Datos agregados mensualmente
    currentScale: 'daily' // 'daily', 'weekly', 'monthly'
};

// Punto de Entrada Inicial
document.addEventListener("DOMContentLoaded", async () => {
    // 0. Inicializar el Tema
    initTheme();
    
    showLoading(true, "Cargando coordenadas...");
    
    try {
        // 1. Cargar y procesar coordenadas
        appState.locations = await loadCoordinates();
        populateDropdown(appState.locations);
        
        // 2. Inicializar el Mapa
        initMap(appState.locations, handleLocationSelection);
        
        // 3. Seleccionar la primera localidad por defecto que tenga datos
        const defaultLoc = appState.locations.find(loc => {
            const path = EXCEL_FILES[loc.name];
            return path !== undefined && path !== null;
        }) || appState.locations[0];
        
        if (defaultLoc) {
            await handleLocationSelection(defaultLoc.name);
        } else {
            showLoading(false);
        }
        
        // 4. Configurar event listeners para los controles de interfaz
        setupEventListeners();
        
        // Inicializar iconos Lucide
        lucide.createIcons();
    } catch (error) {
        console.error("Error al inicializar la aplicación:", error);
        showLoading(true, `Error al cargar la aplicación: ${error.message}. Asegúrate de ejecutar el servidor local.`);
    }
});

/**
 * ==========================================================================
 * PARSEADORES DE DATOS Y COORDENADAS
 * ==========================================================================
 */

/**
 * Parsea un string en formato DMS (Grados, Minutos, Segundos) a grados decimales.
 * Soporta caracteres extraños de codificación e ignorará caracteres que no sean números.
 * Ejemplo: "3°45'29.6\"S " -> -3.75822
 */
function parseDMS(dmsStr) {
    if (!dmsStr) return null;
    
    const clean = dmsStr.trim();
    
    // Expresión regular robusta para capturar: Grados, Minutos, Segundos y Dirección
    // Captura grupos numéricos consecutivos y la letra de dirección final
    const regex = /(\d+)[^\d\.]+(\d+)[^\d\.]+([\d\.]+)[^\d\.a-zA-Z]*([NSEWnsew])/;
    const match = clean.match(regex);
    
    if (match) {
        const deg = parseFloat(match[1]);
        const min = parseFloat(match[2]);
        const sec = parseFloat(match[3]);
        const dir = match[4].toUpperCase();
        
        let decimal = deg + (min / 60) + (sec / 3600);
        if (dir === 'S' || dir === 'W') {
            decimal = -decimal;
        }
        return decimal;
    }
    
    // Si ya viene en formato decimal (ej. -3.75822)
    const num = parseFloat(clean.replace(',', '.'));
    if (!isNaN(num)) return num;
    
    return null;
}

function normalizeName(name) {
    if (!name) return "";
    const norm = name.trim();
    if (/Para/i.test(norm)) {
        return "Nuevo Paraíso";
    }
    if (/Sof/i.test(norm)) {
        return "Santa Sofía";
    }
    if (/Miguel/i.test(norm)) {
        return "San Miguel";
    }
    if (/Primavera/i.test(norm)) {
        return "La Primavera";
    }
    return norm;
}

/**
 * Lee el archivo Coordenadas.csv y retorna un array de localidades estructurado.
 */
async function loadCoordinates() {
    const response = await fetch('Datos_clima/Coordenadas.csv');
    if (!response.ok) {
        throw new Error("No se pudo cargar el archivo Coordenadas.csv");
    }
    const text = await response.text();
    const lines = text.split('\n');
    const locations = [];
    
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;
        
        // Separador de punto y coma
        const parts = line.split(';');
        if (parts.length < 3) continue;
        
        const rawName = parts[0].replace(/"/g, '').trim();
        const name = normalizeName(rawName);
        const rawLat = parts[1].replace(/"/g, '').trim();
        const rawLon = parts[2].replace(/"/g, '').trim();
        
        const lat = parseDMS(rawLat);
        const lon = parseDMS(rawLon);
        
        if (lat !== null && lon !== null) {
            locations.push({ name, lat, lon, rawLat, rawLon });
        }
    }
    return locations;
}

/**
 * Convierte el número de serie de fecha de Excel a un string formateado yyyy-MM-dd en UTC.
 * Evita desajustes por zonas horarias del navegador.
 */
function excelDateToDateString(serial) {
    if (typeof serial !== 'number') {
        return String(serial);
    }
    // El origen de fecha en Excel es 1900-01-01, y el epoch de Unix es 1970-01-01 (serial 25569)
    const utcDays = serial - 25569;
    const utcValue = utcDays * 86400 * 1000;
    const date = new Date(utcValue);
    
    const yyyy = date.getUTCFullYear();
    const mm = String(date.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(date.getUTCDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

/**
 * Carga binaria del archivo Excel y conversión a JSON usando SheetJS
 */
async function loadClimateData(filePath) {
    const response = await fetch(filePath);
    if (!response.ok) {
        throw new Error(`Error HTTP: ${response.status} - No se pudo descargar el archivo Excel.`);
    }
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: 'array' });
    
    // Tomar la primera hoja
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    
    // Convertir la hoja a JSON
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    
    const processed = [];
    for (const row of rawRows) {
        // Ignorar filas sin Fecha
        if (row.Fecha === undefined || row.Fecha === "") continue;
        
        const serial = parseFloat(row.Fecha);
        if (isNaN(serial)) continue; // Saltarse headers o texto no numérico
        
        const dateStr = excelDateToDateString(serial);
        
        // Limpieza de datos climáticos
        const prec = cleanValue(row.prec);
        const tmin = cleanValue(row.tmin);
        const tmax = cleanValue(row.tmax);
        
        processed.push({
            date: dateStr,
            prec: prec,
            tmin: tmin,
            tmax: tmax
        });
    }
    
    // Asegurar orden cronológico
    processed.sort((a, b) => a.date.localeCompare(b.date));
    return processed;
}

/**
 * Limpia y castea un valor de celda a float o retorna null
 */
function cleanValue(val) {
    if (val === undefined || val === null || val === "" || String(val).trim() === "--") {
        return null;
    }
    const num = parseFloat(String(val).replace(',', '.'));
    return isNaN(num) ? null : num;
}

/**
 * ==========================================================================
 * ALGORITMOS DE AGREGACIÓN TEMPORAL (DAILY, WEEKLY, MONTHLY)
 * ==========================================================================
 */

/**
 * Calcula el lunes de la semana de una fecha en UTC
 */
function getStartOfWeek(dateStr) {
    const date = new Date(dateStr + "T00:00:00Z");
    const day = date.getUTCDay();
    // Ajustar si el día es domingo (0) -> restar 6 días, o si es lunes-sábado -> restar (day-1) días
    const diff = date.getUTCDate() - day + (day === 0 ? -6 : 1);
    const monday = new Date(date.setUTCDate(diff));
    return monday.toISOString().split('T')[0];
}

/**
 * Agrupa datos diarios por semana
 * Suma la precipitación y promedia temperaturas
 */
function aggregateWeekly(dailyData) {
    const groups = {};
    dailyData.forEach(row => {
        const weekKey = getStartOfWeek(row.date);
        if (!groups[weekKey]) {
            groups[weekKey] = { precs: [], tmins: [], tmaxs: [] };
        }
        if (row.prec !== null) groups[weekKey].precs.push(row.prec);
        if (row.tmin !== null) groups[weekKey].tmins.push(row.tmin);
        if (row.tmax !== null) groups[weekKey].tmaxs.push(row.tmax);
    });
    
    return Object.keys(groups).sort().map(weekKey => {
        const g = groups[weekKey];
        const sumPrec = g.precs.reduce((a, b) => a + b, 0);
        const avgTmin = g.tmins.length > 0 ? (g.tmins.reduce((a, b) => a + b, 0) / g.tmins.length) : null;
        const avgTmax = g.tmaxs.length > 0 ? (g.tmaxs.reduce((a, b) => a + b, 0) / g.tmaxs.length) : null;
        return {
            label: `Sem. del ${formatSimpleDate(weekKey)}`,
            prec: sumPrec,
            tmin: avgTmin,
            tmax: avgTmax
        };
    });
}

/**
 * Agrupa datos diarios por mes
 * Suma la precipitación y promedia temperaturas
 */
function aggregateMonthly(dailyData) {
    const groups = {};
    const monthNames = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
    
    dailyData.forEach(row => {
        const monthKey = row.date.substring(0, 7); // e.g. "2026-04"
        if (!groups[monthKey]) {
            groups[monthKey] = { precs: [], tmins: [], tmaxs: [] };
        }
        if (row.prec !== null) groups[monthKey].precs.push(row.prec);
        if (row.tmin !== null) groups[monthKey].tmins.push(row.tmin);
        if (row.tmax !== null) groups[monthKey].tmaxs.push(row.tmax);
    });
    
    return Object.keys(groups).sort().map(monthKey => {
        const g = groups[monthKey];
        const [year, month] = monthKey.split('-');
        const monthName = monthNames[parseInt(month) - 1];
        
        const sumPrec = g.precs.reduce((a, b) => a + b, 0);
        const avgTmin = g.tmins.length > 0 ? (g.tmins.reduce((a, b) => a + b, 0) / g.tmins.length) : null;
        const avgTmax = g.tmaxs.length > 0 ? (g.tmaxs.reduce((a, b) => a + b, 0) / g.tmaxs.length) : null;
        return {
            label: `${monthName} ${year}`,
            prec: sumPrec,
            tmin: avgTmin,
            tmax: avgTmax
        };
    });
}

/**
 * Formateadores de fecha amigables en español
 */
function formatDateFull(dateStr) {
    if (!dateStr) return "--";
    const [year, month, day] = dateStr.split('-');
    const months = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
    return `${parseInt(day)} de ${months[parseInt(month) - 1]}, ${year}`;
}

function formatSimpleDate(dateStr) {
    if (!dateStr) return "--";
    const [year, month, day] = dateStr.split('-');
    const months = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
    return `${parseInt(day)} ${months[parseInt(month) - 1]}`;
}

/**
 * ==========================================================================
 * GESTIÓN DE UI E INTERACCIÓN DEL USUARIO
 * ==========================================================================
 */

/**
 * Muestra/Oculta la pantalla de carga
 */
function showLoading(show, message = "Cargando...") {
    const overlay = document.getElementById('loading-overlay');
    const textEl = overlay.querySelector('p');
    textEl.textContent = message;
    
    if (show) {
        overlay.classList.remove('hidden');
    } else {
        overlay.classList.add('hidden');
    }
}

/**
 * Llena el selector desplegable con las localidades
 */
function populateDropdown(locations) {
    const select = document.getElementById('locality-select');
    select.innerHTML = '<option value="" disabled>Selecciona una localidad...</option>';
    
    locations.forEach(loc => {
        const option = document.createElement('option');
        option.value = loc.name;
        option.textContent = loc.name + (EXCEL_FILES[loc.name] ? "" : " (Sin datos)");
        select.appendChild(option);
    });
}

/**
 * Configura los event listeners de los botones de escala y dropdown
 */
function setupEventListeners() {
    // Dropdown de Localidades
    const select = document.getElementById('locality-select');
    select.addEventListener('change', (e) => {
        handleLocationSelection(e.target.value);
    });
    
    // Botones de Escala Temporal (Diario, Semanal, Mensual)
    const tabButtons = document.querySelectorAll('.tab-btn');
    tabButtons.forEach(btn => {
        btn.addEventListener('click', (e) => {
            tabButtons.forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
            
            const scale = e.target.getAttribute('data-scale');
            appState.currentScale = scale;
            
            // Actualizar la visualización con la nueva escala
            renderDashboardData();
        });
    });
}

/**
 * Procesa la selección de una nueva localidad
 */
async function handleLocationSelection(name) {
    const normalized = normalizeName(name);
    appState.selectedLocation = appState.locations.find(loc => loc.name === normalized);
    
    if (!appState.selectedLocation) {
        showLoading(false);
        return;
    }
    
    // Actualizar selector desplegable
    document.getElementById('locality-select').value = normalized;
    
    // Actualizar encabezados
    document.getElementById('active-locality-name').textContent = normalized;
    document.getElementById('active-locality-coords').textContent = 
        `Lat: ${appState.selectedLocation.lat.toFixed(5)}° | Lon: ${appState.selectedLocation.lon.toFixed(5)}°`;
        
    // Centrar mapa suavemente
    focusMapOnLocation(appState.selectedLocation);
    
    const excelPath = EXCEL_FILES[normalized];
    
    if (!excelPath) {
        // Mostrar advertencia de falta de archivo
        appState.dailyData = [];
        appState.weeklyData = [];
        appState.monthlyData = [];
        showNoDataWarning(true);
        updateQuickStats([]);
        showLoading(false); // Ocultar cargador en localidades sin datos
        return;
    }
    
    showNoDataWarning(false);
    showLoading(true, `Leyendo datos climáticos de ${normalized}...`);
    
    try {
        // Cargar datos climáticos del Excel correspondiente
        appState.dailyData = await loadClimateData(excelPath);
        
        // Generar agregaciones
        appState.weeklyData = aggregateWeekly(appState.dailyData);
        appState.monthlyData = aggregateMonthly(appState.dailyData);
        
        // Actualizar estadísticas basadas en toda la serie de datos diarios
        updateQuickStats(appState.dailyData);
        
        // Renderizar gráficos
        renderDashboardData();
    } catch (err) {
        console.error(err);
        showLoading(true, `Error al cargar Excel: ${err.message}. Asegúrate de que el archivo exista en Datos_clima/.`);
    } finally {
        showLoading(false);
    }
}

/**
 * Controla si se visualiza el aviso de falta de datos climáticos
 */
function showNoDataWarning(show) {
    const warningEl = document.getElementById('no-data-warning');
    const chartsContainer = document.getElementById('charts-grid');
    
    if (show) {
        warningEl.classList.remove('hidden');
        chartsContainer.classList.add('hidden');
    } else {
        warningEl.classList.add('hidden');
        chartsContainer.classList.remove('hidden');
    }
}

/**
 * Renderiza los gráficos en base a la escala seleccionada
 */
function renderDashboardData() {
    let activeDataset = [];
    
    if (appState.currentScale === 'daily') {
        activeDataset = appState.dailyData;
    } else if (appState.currentScale === 'weekly') {
        activeDataset = appState.weeklyData;
    } else if (appState.currentScale === 'monthly') {
        activeDataset = appState.monthlyData;
    }
    
    if (activeDataset.length === 0) {
        return;
    }
    
    // Dibujar gráficos Chart.js
    updateCharts(activeDataset, appState.currentScale);
}

/**
 * Calcula y actualiza las tarjetas de estadísticas rápidas
 */
function updateQuickStats(dailyData) {
    let totalPrec = 0;
    let maxTemp = -Infinity;
    let maxTempDate = "";
    let minTemp = Infinity;
    let minTempDate = "";
    
    let precCount = 0;
    let tmaxCount = 0;
    let tminCount = 0;
    
    dailyData.forEach(row => {
        if (row.prec !== null) {
            totalPrec += row.prec;
            precCount++;
        }
        if (row.tmax !== null) {
            if (row.tmax > maxTemp) {
                maxTemp = row.tmax;
                maxTempDate = formatDateFull(row.date);
            }
            tmaxCount++;
        }
        if (row.tmin !== null) {
            if (row.tmin < minTemp) {
                minTemp = row.tmin;
                minTempDate = formatDateFull(row.date);
            }
            tminCount++;
        }
    });
    
    // Escribir valores en la UI
    document.getElementById('stat-prec-val').textContent = precCount > 0 ? `${totalPrec.toFixed(1)} mm` : "--";
    document.getElementById('stat-prec-sub').textContent = precCount > 0 ? `Lluvia total (${dailyData.length} d)` : "Sin datos";
    
    if (tmaxCount > 0) {
        document.getElementById('stat-tmax-val').textContent = `${maxTemp.toFixed(1)} °C`;
        document.getElementById('stat-tmax-sub').textContent = `Máxima: ${maxTempDate}`;
    } else {
        document.getElementById('stat-tmax-val').textContent = "--";
        document.getElementById('stat-tmax-sub').textContent = "Sin mediciones";
    }
    
    if (tminCount > 0) {
        document.getElementById('stat-tmin-val').textContent = `${minTemp.toFixed(1)} °C`;
        document.getElementById('stat-tmin-sub').textContent = `Mínima: ${minTempDate}`;
    } else {
        document.getElementById('stat-tmin-val').textContent = "--";
        document.getElementById('stat-tmin-sub').textContent = "Sin mediciones";
    }
}

/**
 * ==========================================================================
 * INTEGRACIÓN CON MAPA INTERACTIVO (LEAFLET)
 * ==========================================================================
 */

let map;
let markersLayer;
let mapTileLayer;

/**
 * Inicializa el mapa y los marcadores
 */
function initMap(locations, onSelectLocation) {
    // Coordenadas iniciales para abarcar Putumayo y Amazonas en el visor
    map = L.map('map', {
        zoomControl: true,
        attributionControl: false
    }).setView([-1.5, -73], 6);
    
    // Obtener tema inicial
    const isLight = document.body.classList.contains('light-theme');
    const tileUrl = isLight 
        ? 'https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png' 
        : 'https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png';
        
    // Tile layer premium (CartoDB Dark Matter / Positron)
    mapTileLayer = L.tileLayer(tileUrl, {
        maxZoom: 19
    }).addTo(map);
    
    markersLayer = L.layerGroup().addTo(map);
    
    locations.forEach(loc => {
        const hasData = EXCEL_FILES[loc.name] !== undefined && EXCEL_FILES[loc.name] !== null;
        
        // Marcador circular premium
        const marker = L.circleMarker([loc.lat, loc.lon], {
            radius: 9,
            fillColor: hasData ? 'hsl(210, 100%, 55%)' : 'hsl(0, 84%, 60%)', // azul si tiene datos, rojo si no
            color: '#ffffff',
            weight: 1.5,
            opacity: 1,
            fillOpacity: 0.85
        }).addTo(markersLayer);
        
        // Popup personalizado con variables de color CSS
        marker.bindPopup(`
            <div style="font-family: 'Outfit', sans-serif; padding: 4px; min-width: 160px;">
                <h4 style="margin: 0 0 6px 0; color: var(--popup-title); font-size: 1rem; border-bottom: 1px solid var(--popup-border); padding-bottom: 4px;">${loc.name}</h4>
                <p style="margin: 0 0 6px 0; color: var(--popup-text); font-size: 0.8rem; line-height: 1.3;">
                    <b>Latitud:</b> ${loc.lat.toFixed(5)}°<br>
                    <b>Longitud:</b> ${loc.lon.toFixed(5)}°
                </p>
                <div style="font-size: 0.72rem; font-weight:600; padding: 4px 8px; border-radius: 8px; text-align: center; 
                     background: ${hasData ? 'var(--badge-bg-success)' : 'var(--badge-bg-danger)'}; 
                     color: ${hasData ? 'var(--badge-text-success)' : 'var(--badge-text-danger)'};">
                    ${hasData ? '✓ Datos climáticos listos' : '✗ Archivo Excel ausente'}
                </div>
            </div>
        `);
        
        marker.on('click', () => {
            onSelectLocation(loc.name);
        });
        
        // Guardar referencia al marcador en la lista de localidades para interactuar después
        loc.marker = marker;
    });
}

/**
 * Centra y hace zoom sobre la localidad seleccionada en el mapa
 */
function focusMapOnLocation(loc) {
    if (map && loc) {
        map.setView([loc.lat, loc.lon], 9, {
            animate: true,
            duration: 1.0
        });
        
        // Abrir el popup del marcador
        if (loc.marker) {
            loc.marker.openPopup();
        }
    }
}

/**
 * ==========================================================================
 * INTEGRACIÓN CON GRÁFICOS (CHART.JS)
 * ==========================================================================
 */

let tempChart = null;
let precChart = null;

/**
 * Crea o actualiza los dos gráficos del dashboard
 */
function updateCharts(data, scale) {
    const labels = data.map(r => r.label || formatSimpleDate(r.date));
    const tmaxs = data.map(r => r.tmax);
    const tmins = data.map(r => r.tmin);
    const precs = data.map(r => r.prec);
    
    // Obtener colores del tema dinámico
    const colors = getThemeColors();
    
    // --------------------------------------------------
    // 1. GRÁFICO DE TEMPERATURAS (Línea)
    // --------------------------------------------------
    const ctxTemp = document.getElementById('tempChart').getContext('2d');
    if (tempChart) tempChart.destroy(); // Limpiar gráfico viejo si existe
    
    // Crear degradados degradados de fondo para relleno
    const gradMax = ctxTemp.createLinearGradient(0, 0, 0, 300);
    gradMax.addColorStop(0, 'rgba(249, 115, 22, 0.25)');
    gradMax.addColorStop(1, 'rgba(249, 115, 22, 0.0)');
    
    const gradMin = ctxTemp.createLinearGradient(0, 0, 0, 300);
    gradMin.addColorStop(0, 'rgba(20, 184, 166, 0.15)');
    gradMin.addColorStop(1, 'rgba(20, 184, 166, 0.0)');
    
    tempChart = new Chart(ctxTemp, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Temp Máxima',
                    data: tmaxs,
                    borderColor: 'hsl(14, 95%, 55%)',
                    backgroundColor: gradMax,
                    borderWidth: 2.5,
                    fill: true,
                    tension: 0.35,
                    pointBackgroundColor: 'hsl(14, 95%, 55%)',
                    pointHoverRadius: 6,
                    spanGaps: true
                },
                {
                    label: 'Temp Mínima',
                    data: tmins,
                    borderColor: 'hsl(175, 90%, 45%)',
                    backgroundColor: gradMin,
                    borderWidth: 2.5,
                    fill: true,
                    tension: 0.35,
                    pointBackgroundColor: 'hsl(175, 90%, 45%)',
                    pointHoverRadius: 6,
                    spanGaps: true
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        color: colors.textPrimary,
                        font: { family: 'Outfit', size: 12, weight: 500 }
                    }
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    backgroundColor: 'rgba(15, 23, 42, 0.95)',
                    titleColor: '#fff',
                    titleFont: { family: 'Outfit', size: 13, weight: 600 },
                    bodyFont: { family: 'Outfit' },
                    borderColor: 'rgba(255, 255, 255, 0.1)',
                    borderWidth: 1,
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) label += ': ';
                            if (context.raw !== null) {
                                label += context.raw.toFixed(1) + ' °C';
                            } else {
                                label += 'Sin registro';
                            }
                            return label;
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: { color: colors.gridLines, drawBorder: false },
                    ticks: { color: colors.textSecondary, font: { family: 'Outfit', size: 11 } }
                },
                y: {
                    grid: { color: colors.gridLines, drawBorder: false },
                    ticks: { color: colors.textSecondary, font: { family: 'Outfit', size: 11 } },
                    title: {
                        display: true,
                        text: 'Temperatura (°C)',
                        color: colors.textSecondary,
                        font: { family: 'Outfit', size: 12, weight: 500 }
                    }
                }
            }
        }
    });
    
    // --------------------------------------------------
    // 2. GRÁFICO DE PRECIPITACIONES (Barras)
    // --------------------------------------------------
    const ctxPrec = document.getElementById('precChart').getContext('2d');
    if (precChart) precChart.destroy();
    
    const gradPrec = ctxPrec.createLinearGradient(0, 0, 0, 300);
    gradPrec.addColorStop(0, 'rgba(59, 130, 246, 0.85)');
    gradPrec.addColorStop(1, 'rgba(59, 130, 246, 0.1)');
    
    precChart = new Chart(ctxPrec, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Precipitación',
                data: precs,
                backgroundColor: gradPrec,
                borderColor: 'hsl(210, 100%, 55%)',
                borderWidth: 1.5,
                borderRadius: 4,
                barPercentage: scale === 'daily' ? 0.65 : 0.45
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        color: colors.textPrimary,
                        font: { family: 'Outfit', size: 12, weight: 500 }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(15, 23, 42, 0.95)',
                    titleColor: '#fff',
                    titleFont: { family: 'Outfit', size: 13, weight: 600 },
                    bodyFont: { family: 'Outfit' },
                    borderColor: 'rgba(255, 255, 255, 0.1)',
                    borderWidth: 1,
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) label += ': ';
                            if (context.raw !== null) {
                                label += context.raw.toFixed(1) + ' mm';
                            } else {
                                label += 'Sin registro';
                            }
                            return label;
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: { color: colors.gridLines, drawBorder: false },
                    ticks: { color: colors.textSecondary, font: { family: 'Outfit', size: 11 } }
                },
                y: {
                    grid: { color: colors.gridLines, drawBorder: false },
                    ticks: { color: colors.textSecondary, font: { family: 'Outfit', size: 11 } },
                    title: {
                        display: true,
                        text: 'Precipitación (mm)',
                        color: colors.textSecondary,
                        font: { family: 'Outfit', size: 12, weight: 500 }
                    }
                }
            }
        }
    });
}

/**
 * ==========================================================================
 * GESTIÓN DE TEMAS (DARK / LIGHT MODE)
 * ==========================================================================
 */

/**
 * Inicializa y configura el tema basado en localStorage
 */
function initTheme() {
    const savedTheme = localStorage.getItem('theme') || 'dark'; // 'dark' por defecto
    const isLight = savedTheme === 'light';
    
    if (isLight) {
        document.body.classList.add('light-theme');
    } else {
        document.body.classList.remove('light-theme');
    }
}

/**
 * Registra event listeners y lógica para el botón selector de tema
 */
function setupEventListeners() {
    // Dropdown de Localidades
    const select = document.getElementById('locality-select');
    select.addEventListener('change', (e) => {
        handleLocationSelection(e.target.value);
    });
    
    // Botones de Escala Temporal (Diario, Semanal, Mensual)
    const tabButtons = document.querySelectorAll('.tab-btn');
    tabButtons.forEach(btn => {
        btn.addEventListener('click', (e) => {
            tabButtons.forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
            
            const scale = e.target.getAttribute('data-scale');
            appState.currentScale = scale;
            
            // Actualizar la visualización con la nueva escala
            renderDashboardData();
        });
    });

    // Botón de alternancia de tema (Claro / Oscuro)
    const themeBtn = document.getElementById('theme-toggle-btn');
    if (themeBtn) {
        themeBtn.addEventListener('click', () => {
            const isCurrentlyLight = document.body.classList.contains('light-theme');
            const willBeLight = !isCurrentlyLight;
            
            if (willBeLight) {
                document.body.classList.add('light-theme');
                localStorage.setItem('theme', 'light');
            } else {
                document.body.classList.remove('light-theme');
                localStorage.setItem('theme', 'dark');
            }
            
            // Actualizar mapa y gráficos de forma inmediata
            updateMapTheme(willBeLight);
            updateChartsTheme();
        });
    }
}

/**
 * Retorna la paleta de colores activa según el tema seleccionado
 */
function getThemeColors() {
    const isLight = document.body.classList.contains('light-theme');
    return {
        textPrimary: isLight ? 'hsl(222, 28%, 15%)' : 'hsl(210, 40%, 98%)',
        textSecondary: isLight ? 'hsl(215, 16%, 45%)' : 'hsl(215, 20%, 65%)',
        gridLines: isLight ? 'rgba(0, 0, 0, 0.06)' : 'rgba(255, 255, 255, 0.04)',
    };
}

/**
 * Actualiza la capa de mapa de Leaflet
 */
function updateMapTheme(isLight) {
    if (mapTileLayer) {
        const url = isLight 
            ? 'https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png' 
            : 'https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png';
        mapTileLayer.setUrl(url);
    }
}

/**
 * Actualiza la apariencia y el color de las fuentes/ejes de los gráficos
 */
function updateChartsTheme() {
    if (!tempChart && !precChart) return;
    
    const colors = getThemeColors();
    
    if (tempChart) {
        tempChart.options.plugins.legend.labels.color = colors.textPrimary;
        tempChart.options.scales.x.grid.color = colors.gridLines;
        tempChart.options.scales.x.ticks.color = colors.textSecondary;
        tempChart.options.scales.y.grid.color = colors.gridLines;
        tempChart.options.scales.y.ticks.color = colors.textSecondary;
        tempChart.options.scales.y.title.color = colors.textSecondary;
        tempChart.update();
    }
    
    if (precChart) {
        precChart.options.plugins.legend.labels.color = colors.textPrimary;
        precChart.options.scales.x.grid.color = colors.gridLines;
        precChart.options.scales.x.ticks.color = colors.textSecondary;
        precChart.options.scales.y.grid.color = colors.gridLines;
        precChart.options.scales.y.ticks.color = colors.textSecondary;
        precChart.options.scales.y.title.color = colors.textSecondary;
        precChart.update();
    }
}
