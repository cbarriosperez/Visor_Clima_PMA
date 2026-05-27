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
    currentScale: 'daily', // 'daily', 'weekly', 'monthly'
    filterStartDate: null, // Fecha de inicio del filtro (yyyy-MM-dd) o null
    filterEndDate: null    // Fecha de fin del filtro (yyyy-MM-dd) o null
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
        
        // Reiniciar filtro de fechas y calendario
        appState.filterStartDate = null;
        appState.filterEndDate = null;
        initDateRangePicker();
        
        // Renderizar gráficos (incluye actualización de estadísticas)
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
    const filteredDaily = getFilteredDailyData();
    let activeDataset = [];
    
    if (appState.currentScale === 'daily') {
        activeDataset = filteredDaily;
    } else if (appState.currentScale === 'weekly') {
        activeDataset = aggregateWeekly(filteredDaily);
    } else if (appState.currentScale === 'monthly') {
        activeDataset = aggregateMonthly(filteredDaily);
    }
    
    // Actualizar estadísticas con datos filtrados
    updateQuickStats(filteredDaily);
    
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

    // Botón de colapsar/expandir el banner informativo
    const infoBanner = document.getElementById('info-banner');
    const infoToggleBtn = document.getElementById('info-toggle-btn');
    if (infoBanner && infoToggleBtn) {
        // Restaurar estado guardado
        const bannerCollapsed = localStorage.getItem('infoBannerCollapsed') === 'true';
        if (bannerCollapsed) {
            infoBanner.classList.add('collapsed');
        }
        
        // Hacer clic tanto en el botón como en el encabezado
        const infoHeader = infoBanner.querySelector('.info-banner-header');
        infoHeader.addEventListener('click', () => {
            infoBanner.classList.toggle('collapsed');
            const isCollapsed = infoBanner.classList.contains('collapsed');
            localStorage.setItem('infoBannerCollapsed', isCollapsed);
        });
    }

    // Selector de rango de fechas (Calendario)
    setupCalendarListeners();
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

/**
 * ==========================================================================
 * SELECTOR DE RANGO DE FECHAS (CALENDARIO DUAL ESTILO AGENCIAS DE VIAJE)
 * ==========================================================================
 */

const MONTH_NAMES_FULL = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

const calendarState = {
    baseYear: 2026,
    baseMonth: 0,
    selectionStart: null,
    selectionEnd: null,
    hoveredDate: null,
    isOpen: false,
    dataMinDate: null,
    dataMaxDate: null
};

/**
 * Retorna los datos diarios filtrados por el rango de fechas seleccionado
 */
function getFilteredDailyData() {
    let data = appState.dailyData;
    if (appState.filterStartDate) {
        data = data.filter(r => r.date >= appState.filterStartDate);
    }
    if (appState.filterEndDate) {
        data = data.filter(r => r.date <= appState.filterEndDate);
    }
    return data;
}

/**
 * Inicializa el calendario con los datos de la localidad activa
 */
function initDateRangePicker() {
    if (appState.dailyData.length === 0) return;

    const dates = appState.dailyData.map(r => r.date).sort();
    calendarState.dataMinDate = dates[0];
    calendarState.dataMaxDate = dates[dates.length - 1];

    // Posicionar el calendario para mostrar los últimos dos meses de datos
    const lastDate = new Date(calendarState.dataMaxDate + 'T00:00:00Z');
    calendarState.baseYear = lastDate.getUTCFullYear();
    calendarState.baseMonth = lastDate.getUTCMonth() - 1;
    if (calendarState.baseMonth < 0) {
        calendarState.baseMonth = 11;
        calendarState.baseYear--;
    }

    calendarState.selectionStart = null;
    calendarState.selectionEnd = null;
    calendarState.hoveredDate = null;

    updateTriggerDisplay();
    renderCalendars();

    // Resetear presets
    document.querySelectorAll('.preset-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.dataset.preset === 'all') btn.classList.add('active');
    });
    updateDateRangeInfo('Selecciona la fecha de inicio');
}

/**
 * Actualiza el texto del botón trigger con las fechas activas
 */
function updateTriggerDisplay() {
    const startEl = document.getElementById('trigger-start-date');
    const endEl = document.getElementById('trigger-end-date');
    if (!startEl || !endEl) return;

    if (appState.filterStartDate && appState.filterEndDate) {
        startEl.textContent = formatDateFull(appState.filterStartDate);
        endEl.textContent = formatDateFull(appState.filterEndDate);
    } else if (calendarState.dataMinDate && calendarState.dataMaxDate) {
        startEl.textContent = formatDateFull(calendarState.dataMinDate);
        endEl.textContent = formatDateFull(calendarState.dataMaxDate);
    } else {
        startEl.textContent = '--';
        endEl.textContent = '--';
    }
}

/**
 * Renderiza ambos meses del calendario dual
 */
function renderCalendars() {
    renderMonth('cal-left-days', 'cal-left-title', calendarState.baseYear, calendarState.baseMonth);

    let rightMonth = calendarState.baseMonth + 1;
    let rightYear = calendarState.baseYear;
    if (rightMonth > 11) {
        rightMonth = 0;
        rightYear++;
    }
    renderMonth('cal-right-days', 'cal-right-title', rightYear, rightMonth);
}

/**
 * Renderiza los días de un mes individual en el contenedor indicado
 */
function renderMonth(containerId, titleId, year, month) {
    const container = document.getElementById(containerId);
    const title = document.getElementById(titleId);
    if (!container || !title) return;

    title.textContent = `${MONTH_NAMES_FULL[month]} ${year}`;

    // Calcular offset para empezar en lunes
    const firstDay = new Date(Date.UTC(year, month, 1)).getUTCDay();
    const startOffset = firstDay === 0 ? 6 : firstDay - 1;
    const daysInMonth = new Date(Date.UTC(year, month + 1, 0)).getUTCDate();

    // Determinar rango de selección activo (incluye hover preview)
    const selStart = calendarState.selectionStart;
    const selEnd = calendarState.selectionEnd || calendarState.hoveredDate;
    let rangeStart = selStart;
    let rangeEnd = selEnd;
    if (rangeStart && rangeEnd && rangeStart > rangeEnd) {
        [rangeStart, rangeEnd] = [rangeEnd, rangeStart];
    }

    // Fecha de hoy
    const today = new Date();
    const todayStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;

    let html = '';

    // Celdas vacías de offset
    for (let i = 0; i < startOffset; i++) {
        html += '<span class="cal-day empty"></span>';
    }

    for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        let classes = ['cal-day'];

        // Verificar si la fecha está dentro del rango de datos
        const inDataRange = calendarState.dataMinDate && calendarState.dataMaxDate &&
            dateStr >= calendarState.dataMinDate && dateStr <= calendarState.dataMaxDate;
        if (!inDataRange) {
            classes.push('disabled');
        }

        // Estados de selección
        if (dateStr === selStart || dateStr === (calendarState.selectionEnd || '')) {
            classes.push('selected-endpoint');
        }
        if (rangeStart && rangeEnd && dateStr > rangeStart && dateStr < rangeEnd) {
            classes.push('in-range');
        }
        if (dateStr === rangeStart && rangeEnd) {
            classes.push('range-start');
        }
        if (dateStr === rangeEnd && rangeStart) {
            classes.push('range-end');
        }

        // Hoy
        if (dateStr === todayStr) {
            classes.push('today');
        }

        html += `<span class="${classes.join(' ')}" data-date="${dateStr}">${day}</span>`;
    }

    container.innerHTML = html;
}

/**
 * Maneja el clic en un día del calendario
 */
function handleDayClick(dateStr) {
    if (!calendarState.dataMinDate || dateStr < calendarState.dataMinDate || dateStr > calendarState.dataMaxDate) return;

    if (!calendarState.selectionStart || calendarState.selectionEnd) {
        // Iniciar nueva selección
        calendarState.selectionStart = dateStr;
        calendarState.selectionEnd = null;
        calendarState.hoveredDate = null;
        updateDateRangeInfo('Ahora selecciona la fecha final');
    } else {
        // Completar selección
        if (dateStr < calendarState.selectionStart) {
            calendarState.selectionEnd = calendarState.selectionStart;
            calendarState.selectionStart = dateStr;
        } else {
            calendarState.selectionEnd = dateStr;
        }
        const days = daysBetween(calendarState.selectionStart, calendarState.selectionEnd);
        updateDateRangeInfo(`${formatSimpleDate(calendarState.selectionStart)} → ${formatSimpleDate(calendarState.selectionEnd)} (${days + 1} días)`);
    }

    // Limpiar preset activo
    document.querySelectorAll('.preset-btn').forEach(b => b.classList.remove('active'));
    renderCalendars();
}

/**
 * Maneja hover sobre un día para previsualizar el rango
 */
function handleDayHover(dateStr) {
    if (calendarState.selectionStart && !calendarState.selectionEnd) {
        calendarState.hoveredDate = dateStr;
        renderCalendars();
    }
}

/**
 * Aplica el rango de fechas seleccionado y filtra los datos
 */
function applyDateRange() {
    if (calendarState.selectionStart && calendarState.selectionEnd) {
        appState.filterStartDate = calendarState.selectionStart;
        appState.filterEndDate = calendarState.selectionEnd;
    } else if (calendarState.selectionStart) {
        appState.filterStartDate = calendarState.selectionStart;
        appState.filterEndDate = calendarState.selectionStart;
    } else {
        appState.filterStartDate = null;
        appState.filterEndDate = null;
    }

    updateTriggerDisplay();
    toggleCalendarDropdown(false);
    renderDashboardData();
}

/**
 * Limpia la selección y muestra todos los datos
 */
function clearDateRange() {
    calendarState.selectionStart = null;
    calendarState.selectionEnd = null;
    calendarState.hoveredDate = null;
    appState.filterStartDate = null;
    appState.filterEndDate = null;

    document.querySelectorAll('.preset-btn').forEach(b => {
        b.classList.remove('active');
        if (b.dataset.preset === 'all') b.classList.add('active');
    });

    updateTriggerDisplay();
    updateDateRangeInfo('Selecciona la fecha de inicio');
    renderCalendars();
    renderDashboardData();
}

/**
 * Aplica un preset de rango temporal rápido
 */
function applyPreset(preset) {
    if (!calendarState.dataMaxDate) return;

    if (preset === 'all') {
        clearDateRange();
        return;
    }

    const maxDate = calendarState.dataMaxDate;
    const end = new Date(maxDate + 'T00:00:00Z');
    let startDate;

    if (preset === '7d') {
        const s = new Date(end); s.setUTCDate(s.getUTCDate() - 6);
        startDate = s.toISOString().split('T')[0];
    } else if (preset === '30d') {
        const s = new Date(end); s.setUTCDate(s.getUTCDate() - 29);
        startDate = s.toISOString().split('T')[0];
    } else if (preset === '90d') {
        const s = new Date(end); s.setUTCDate(s.getUTCDate() - 89);
        startDate = s.toISOString().split('T')[0];
    } else if (preset === '6m') {
        const s = new Date(end); s.setUTCMonth(s.getUTCMonth() - 6);
        startDate = s.toISOString().split('T')[0];
    }

    // Limitar al rango mínimo de datos
    if (startDate < calendarState.dataMinDate) {
        startDate = calendarState.dataMinDate;
    }

    calendarState.selectionStart = startDate;
    calendarState.selectionEnd = maxDate;

    // Navegar calendario al final del rango
    const endParsed = new Date(maxDate + 'T00:00:00Z');
    calendarState.baseYear = endParsed.getUTCFullYear();
    calendarState.baseMonth = endParsed.getUTCMonth() - 1;
    if (calendarState.baseMonth < 0) {
        calendarState.baseMonth = 11;
        calendarState.baseYear--;
    }

    // Actualizar botones de preset
    document.querySelectorAll('.preset-btn').forEach(b => {
        b.classList.remove('active');
        if (b.dataset.preset === preset) b.classList.add('active');
    });

    const days = daysBetween(startDate, maxDate);
    updateDateRangeInfo(`${formatSimpleDate(startDate)} → ${formatSimpleDate(maxDate)} (${days + 1} días)`);
    renderCalendars();

    // Auto-aplicar
    appState.filterStartDate = startDate;
    appState.filterEndDate = maxDate;
    updateTriggerDisplay();
    renderDashboardData();
}

/**
 * Abre o cierra el dropdown del calendario
 */
function toggleCalendarDropdown(forceState) {
    const dropdown = document.getElementById('date-range-dropdown');
    const trigger = document.getElementById('date-range-trigger');
    if (!dropdown || !trigger) return;

    calendarState.isOpen = forceState !== undefined ? forceState : !calendarState.isOpen;

    if (calendarState.isOpen) {
        dropdown.classList.add('open');
        trigger.classList.add('active');
        renderCalendars();
    } else {
        dropdown.classList.remove('open');
        trigger.classList.remove('active');
    }
}

function updateDateRangeInfo(text) {
    const el = document.getElementById('date-range-info');
    if (el) el.textContent = text;
}

function daysBetween(dateStr1, dateStr2) {
    const d1 = new Date(dateStr1 + 'T00:00:00Z');
    const d2 = new Date(dateStr2 + 'T00:00:00Z');
    return Math.round(Math.abs((d2 - d1) / (1000 * 60 * 60 * 24)));
}

/**
 * Registra todos los event listeners del calendario
 */
function setupCalendarListeners() {
    // Botón trigger para abrir/cerrar
    const trigger = document.getElementById('date-range-trigger');
    if (trigger) {
        trigger.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleCalendarDropdown();
        });
    }

    // Cerrar al hacer clic fuera
    document.addEventListener('click', (e) => {
        const picker = document.getElementById('date-range-picker');
        if (picker && !picker.contains(e.target) && calendarState.isOpen) {
            toggleCalendarDropdown(false);
        }
    });

    // Navegación entre meses
    const prevBtn = document.getElementById('cal-prev');
    const nextBtn = document.getElementById('cal-next');

    if (prevBtn) {
        prevBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            calendarState.baseMonth--;
            if (calendarState.baseMonth < 0) {
                calendarState.baseMonth = 11;
                calendarState.baseYear--;
            }
            renderCalendars();
        });
    }

    if (nextBtn) {
        nextBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            calendarState.baseMonth++;
            if (calendarState.baseMonth > 11) {
                calendarState.baseMonth = 0;
                calendarState.baseYear++;
            }
            renderCalendars();
        });
    }

    // Delegación de eventos para clics y hover en días del calendario
    ['cal-left-days', 'cal-right-days'].forEach(id => {
        const container = document.getElementById(id);
        if (container) {
            container.addEventListener('click', (e) => {
                const day = e.target.closest('.cal-day:not(.empty):not(.disabled)');
                if (day && day.dataset.date) {
                    handleDayClick(day.dataset.date);
                }
            });
            container.addEventListener('mouseover', (e) => {
                const day = e.target.closest('.cal-day:not(.empty):not(.disabled)');
                if (day && day.dataset.date) {
                    handleDayHover(day.dataset.date);
                }
            });
        }
    });

    // Botones de preset
    document.querySelectorAll('.preset-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            applyPreset(e.currentTarget.dataset.preset);
        });
    });

    // Botones Aplicar y Limpiar
    const applyBtn = document.getElementById('date-range-apply');
    const clearBtn = document.getElementById('date-range-clear');

    if (applyBtn) {
        applyBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            applyDateRange();
        });
    }

    if (clearBtn) {
        clearBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            clearDateRange();
            toggleCalendarDropdown(false);
        });
    }
}
