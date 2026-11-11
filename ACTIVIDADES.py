import os
import pandas as pd
import json

# Rutas de entrada y salida
excel_path = r"C:\Users\edwar.vanegas_PANASA\OneDrive - ManpowerGroup Colombia\Escritorio\PROYECTO\ACTIVIDADES\FOOD\ACTIVIDADES.xlsx"
output_dir = r"C:\Users\edwar.vanegas_PANASA\OneDrive - ManpowerGroup Colombia\Escritorio\PROYECTO\ACTIVIDADES\FOOD"
output_html = os.path.join(output_dir, "index.html")

def build_nestle_gallery():
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    print("Leyendo Excel desde:", excel_path)
    df = pd.read_excel(excel_path)
    df = df.fillna("")
    
    data_records = df.to_dict(orient="records")
    json_data_str = json.dumps(data_records, default=str)

    html_content = f'''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Actividades Nestlé - ManpowerGroup</title>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        :root {{
            --primary-blue: #0070ba;
            --dark-navy: #0f2b48;
            --bg-color: #e8eff7;
            --card-bg: #ffffff;
            --text-main: #1c2d42;
            --text-muted: #627284;
            --border-radius: 12px;
            --shadow: 0 4px 15px rgba(15, 43, 72, 0.06);
        }}

        * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Plus Jakarta Sans', sans-serif; }}

        body {{
            background-color: var(--bg-color);
            color: var(--text-main);
            padding: 18px;
            min-height: 100vh;
        }}

        /* Header Principal */
        .top-header {{
            background: linear-gradient(90deg, #103254 0%, #1e5a8a 35%, #8b2332 75%, #b94025 100%);
            border-radius: 14px;
            padding: 12px 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            color: white;
            box-shadow: 0 6px 20px rgba(0,0,0,0.15);
            margin-bottom: 16px;
        }}

        .brand-title-box {{ display: flex; align-items: center; gap: 14px; }}
        .bars-logo {{ display: flex; gap: 4px; align-items: flex-end; height: 26px; }}
        .bar {{ width: 6px; border-radius: 3px; transform: skewX(-15deg); }}
        .bar-1 {{ height: 16px; background: #2f80ed; }}
        .bar-2 {{ height: 22px; background: #56ccf2; }}
        .bar-3 {{ height: 18px; background: #27ae60; }}
        .bar-4 {{ height: 26px; background: #eb5757; }}
        .bar-5 {{ height: 20px; background: #f2994a; }}

        .header-title {{ font-size: 20px; font-weight: 800; }}

        .btn-download {{
            background: #ffffff; color: #0f2b48; border: none;
            padding: 8px 18px; border-radius: 20px; font-weight: 700;
            font-size: 13px; cursor: pointer; display: flex; align-items: center;
            gap: 8px; box-shadow: 0 3px 8px rgba(0,0,0,0.2); transition: all 0.2s ease;
        }}
        .btn-download:hover {{ transform: translateY(-2px); background: #f8fafc; }}

        /* Pestañas de Cargo */
        .tabs-cargo-container {{
            display: flex; gap: 10px; margin-bottom: 16px; overflow-x: auto; padding-bottom: 6px;
        }}

        .cargo-tab {{
            background: #ffffff; color: var(--text-main); border: 1px solid rgba(0,0,0,0.08);
            padding: 6px 16px 6px 8px; border-radius: 24px; font-size: 13px; font-weight: 700;
            cursor: pointer; transition: all 0.25s ease; white-space: nowrap;
            display: flex; align-items: center; gap: 8px; flex-shrink: 0;
        }}

        .cargo-tab-img {{
            width: 28px; height: 28px; border-radius: 50%; object-fit: cover; border: 1.5px solid var(--primary-blue);
        }}

        .cargo-tab.active {{
            background: var(--primary-blue); color: #ffffff; border-color: var(--primary-blue);
            box-shadow: 0 4px 12px rgba(0, 112, 186, 0.3);
        }}
        .cargo-tab.active .cargo-tab-img {{ border-color: #ffffff; }}

        /* Filtros Interconectados */
        .filters-bar {{
            background: #ffffff; border-radius: 14px; padding: 14px;
            display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 10px;
            margin-bottom: 18px; box-shadow: var(--shadow); align-items: flex-end;
        }}

        .filter-item {{ display: flex; flex-direction: column; gap: 4px; }}
        .filter-label {{ font-size: 10px; font-weight: 800; color: var(--dark-navy); text-transform: uppercase; }}
        .filter-control {{
            width: 100%; padding: 8px 10px; border-radius: 8px; border: 1px solid #dce4ec;
            background: #f8fafc; font-size: 12px; font-weight: 600; color: var(--text-main); outline: none;
        }}

        .btn-reset-filters {{
            background: #f1f5f9; border: 1px solid #cbd5e1; padding: 8px 12px; border-radius: 8px;
            font-size: 12px; font-weight: 700; color: var(--dark-navy); cursor: pointer; height: 35px;
        }}

        /* KPIs en UNA SOLA FILA */
        .kpi-section-title {{
            font-size: 13px; font-weight: 800; color: var(--dark-navy); margin-bottom: 10px; text-transform: uppercase;
        }}

        .kpi-single-row-container {{
            display: flex; gap: 12px; overflow-x: auto; padding-bottom: 10px; margin-bottom: 20px;
        }}

        .kpi-card {{
            background: #ffffff; border-radius: 12px; padding: 12px 16px; text-align: center;
            box-shadow: var(--shadow); border: 1px solid rgba(0, 0, 0, 0.04);
            min-width: 180px; flex: 0 0 auto;
        }}

        .kpi-title {{ font-size: 10px; font-weight: 800; color: var(--text-muted); text-transform: uppercase; margin-bottom: 4px; white-space: nowrap; }}
        .kpi-value {{ font-size: 22px; font-weight: 800; color: var(--primary-blue); }}
        .kpi-subtext {{ font-size: 10px; font-weight: 600; color: var(--text-muted); margin-top: 2px; }}

        /* Galería Grid */
        .gallery-grid {{
            display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 16px;
        }}

        .gal-card {{
            background: #ffffff; border-radius: 14px; overflow: hidden; box-shadow: var(--shadow);
            transition: all 0.3s ease; display: flex; flex-direction: column; border: 1px solid rgba(0,0,0,0.05);
            cursor: pointer;
        }}

        .gal-card:hover {{ transform: translateY(-4px); box-shadow: 0 8px 20px rgba(0,0,0,0.12); }}

        .gal-img-container {{ width: 100%; height: 150px; background: #e2e8f0; position: relative; overflow: hidden; }}
        .gal-img {{ width: 100%; height: 100%; object-fit: cover; }}

        .gal-badge {{
            position: absolute; top: 8px; right: 8px; background: rgba(15, 43, 72, 0.85);
            color: white; font-size: 10px; font-weight: 700; padding: 3px 8px; border-radius: 10px;
        }}

        .gal-content {{ padding: 12px; display: flex; flex-direction: column; gap: 6px; }}
        .gal-tag {{ font-size: 10px; font-weight: 800; color: var(--primary-blue); text-transform: uppercase; }}
        .gal-pdv {{ font-size: 13px; font-weight: 800; color: var(--dark-navy); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .gal-detail {{ font-size: 11px; color: var(--text-muted); display: flex; justify-content: space-between; }}

        /* Visor Lightbox Modal GIGANTE (MAXIMIZADO AL 95%) */
        .lightbox {{
            position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
            background: rgba(15, 23, 42, 0.92); backdrop-filter: blur(10px); z-index: 2000;
            display: flex; align-items: center; justify-content: center; opacity: 0; pointer-events: none;
            transition: opacity 0.3s ease; padding: 15px;
        }}

        .lightbox.active {{ opacity: 1; pointer-events: auto; }}

        .lightbox-card {{
            background: #ffffff; border-radius: 20px; 
            width: 95vw; height: 92vh; max-width: 1400px; max-height: 92vh;
            display: flex; flex-direction: row; overflow: hidden; box-shadow: 0 25px 60px rgba(0,0,0,0.5);
            position: relative;
        }}

        @media (max-width: 900px) {{
            .lightbox-card {{ flex-direction: column; overflow-y: auto; height: 95vh; }}
        }}

        .lightbox-close {{
            position: absolute; top: 14px; right: 20px; font-size: 32px; color: #0f2b48;
            cursor: pointer; font-weight: 800; z-index: 10; background: rgba(255,255,255,0.8);
            width: 40px; height: 40px; border-radius: 50%; display: flex; align-items: center; justify-content: center;
        }}

        .lightbox-main-img {{
            flex: 1.8; background: #080f18; display: flex; flex-direction: column;
            align-items: center; justify-content: center; position: relative; height: 100%; overflow: hidden;
        }}

        .lightbox-main-img img {{
            max-width: 96%; max-height: 88vh; object-fit: contain; border-radius: 8px;
        }}

        .lightbox-nav {{
            position: absolute; top: 50%; transform: translateY(-50%);
            background: rgba(255,255,255,0.25); color: white; width: 48px; height: 48px;
            border-radius: 50%; display: flex; align-items: center; justify-content: center;
            font-size: 24px; cursor: pointer; user-select: none; transition: background 0.2s;
        }}
        .lightbox-nav:hover {{ background: rgba(255,255,255,0.5); }}
        .lightbox-prev {{ left: 16px; }}
        .lightbox-next {{ right: 16px; }}

        .lightbox-info {{
            flex: 1; min-width: 340px; padding: 28px; display: flex; flex-direction: column; justify-content: space-between; overflow-y: auto; background: #ffffff;
        }}

        .info-title {{ font-size: 20px; font-weight: 800; color: var(--dark-navy); margin-bottom: 16px; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; }}

        .info-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 14px; font-size: 13px; margin-bottom: 20px; }}
        .info-item label {{ font-weight: 800; color: var(--text-muted); display: block; font-size: 10px; text-transform: uppercase; margin-bottom: 2px; }}
        .info-item span {{ font-weight: 700; color: var(--text-main); font-size: 13px; word-break: break-word; }}

        .thumbnails-strip {{
            display: flex; gap: 10px; overflow-x: auto; padding: 10px 0; border-top: 1px solid #e2e8f0;
        }}
        .thumb-img {{
            width: 65px; height: 65px; border-radius: 10px; object-fit: cover; cursor: pointer;
            border: 2px solid transparent; opacity: 0.6; transition: all 0.2s;
        }}
        .thumb-img.active {{ border-color: var(--primary-blue); opacity: 1; transform: scale(1.08); }}

        .empty-msg {{ grid-column: 1 / -1; text-align: center; padding: 40px; background: white; border-radius: 14px; color: var(--text-muted); font-weight: 600; }}
    </style>
</head>
<body>

    <!-- Header Principal -->
    <header class="top-header">
        <div class="brand-title-box">
            <div class="bars-logo">
                <div class="bar bar-1"></div>
                <div class="bar bar-2"></div>
                <div class="bar bar-3"></div>
                <div class="bar bar-4"></div>
                <div class="bar bar-5"></div>
            </div>
            <h1 class="header-title">Actividades Nestlé - ManpowerGroup</h1>
        </div>
        <button class="btn-download" onclick="downloadExcel()">📊 Descargar Base Entrada</button>
    </header>

    <!-- NAVEGACIÓN POR CARGO -->
    <nav class="tabs-cargo-container" id="tabs-cargo"></nav>

    <!-- FILTROS INTERCONECTADOS -->
    <div class="filters-bar">
        <div class="filter-item">
            <label class="filter-label">📅 Fecha Desde</label>
            <input type="date" id="f-fecha-desde" class="filter-control" onchange="onFilterChange()">
        </div>
        <div class="filter-item">
            <label class="filter-label">📅 Fecha Hasta</label>
            <input type="date" id="f-fecha-hasta" class="filter-control" onchange="onFilterChange()">
        </div>
        <div class="filter-item">
            <label class="filter-label">👤 Nombre Completo</label>
            <select id="f-nombre" class="filter-control" onchange="onFilterChange()">
                <option value="">Todos</option>
            </select>
        </div>
        <div class="filter-item">
            <label class="filter-label">🎯 Actividad</label>
            <select id="f-actividad" class="filter-control" onchange="onFilterChange()">
                <option value="">Todas</option>
            </select>
        </div>
        <div class="filter-item">
            <label class="filter-label">🎁 Incentivo</label>
            <select id="f-incentivo" class="filter-control" onchange="onFilterChange()">
                <option value="">Todos</option>
            </select>
        </div>
        <div class="filter-item">
            <label class="filter-label">🏪 Punto de Venta</label>
            <select id="f-pdv" class="filter-control" onchange="onFilterChange()">
                <option value="">Todos</option>
            </select>
        </div>
        <div class="filter-item">
            <label class="filter-label">🏙️ Ciudad</label>
            <select id="f-ciudad" class="filter-control" onchange="onFilterChange()">
                <option value="">Todas</option>
            </select>
        </div>
        <button class="btn-reset-filters" onclick="resetFilters()">Limpiar Filtros</button>
    </div>

    <!-- TARJETAS DE INCENTIVOS EN UNA SOLA FILA -->
    <div class="kpi-section-title" id="kpi-section-title">Resumen de Cantidades por Incentivo</div>
    <div class="kpi-single-row-container" id="kpi-row"></div>

    <!-- GALERÍA DE REGISTROS -->
    <main class="gallery-grid" id="gallery-grid"></main>

    <!-- Visor Lightbox Modal GIGANTE -->
    <div class="lightbox" id="lightbox" onclick="closeLightbox(event)">
        <div class="lightbox-card" onclick="event.stopPropagation()">
            <span class="lightbox-close" onclick="closeLightbox(null, true)">&times;</span>
            
            <div class="lightbox-main-img">
                <div class="lightbox-nav lightbox-prev" onclick="changePhoto(-1)">&#10094;</div>
                <img id="lb-img" src="" alt="Ampliada">
                <div class="lightbox-nav lightbox-next" onclick="changePhoto(1)">&#10095;</div>
            </div>

            <div class="lightbox-info">
                <div>
                    <h3 class="info-title" id="lb-title">Detalle del Registro</h3>
                    <div class="info-grid" id="lb-details"></div>
                </div>

                <div>
                    <label style="font-size:11px; font-weight:800; color:var(--dark-navy); text-transform:uppercase;">Fotos del registro (<span id="lb-photo-count">0</span>)</label>
                    <div class="thumbnails-strip" id="lb-thumbnails"></div>
                </div>
            </div>
        </div>
    </div>

    <script>
        const rawData = {json_data_str};
        let selectedCargo = "";
        let currentRecordPhotos = [];
        let currentPhotoIndex = 0;

        // Búsqueda flexible de campos en el Excel
        function getValue(item, keys) {{
            for (let k of keys) {{
                const foundKey = Object.keys(item).find(key => key.trim().toLowerCase() === k.trim().toLowerCase());
                if (foundKey && item[foundKey] !== undefined && item[foundKey] !== null && String(item[foundKey]).trim() !== "") {{
                    return String(item[foundKey]).trim();
                }}
            }}
            return "";
        }}

        function cleanStr(str) {{
            return String(str || "").trim().toLowerCase();
        }}

        function getPhotosList(fotoStr) {{
            if (!fotoStr) return [];
            return String(fotoStr).split(',').map(s => s.trim()).filter(s => s.length > 0);
        }}

        // Mapeo de claves por campo
        const KEYS = {{
            cargo: ["Cargo", "CARGO", "cargo"],
            nombre: ["Nombre completo", "Nombre Completo", "NOMBRE COMPLETO", "Nombre", "NOMBRE"],
            actividad: ["Actividad", "ACTIVIDAD"],
            incentivo: ["¿Qué incentivo recibe?", "Que incentivo recibe?", "Incentivo", "INCENTIVO"],
            pdv: ["Punto de Venta", "PUNTO DE VENTA", "PDV", "Punto de venta"],
            ciudad: ["Ciudad", "CIUDAD"],
            fecha: ["Fecha", "FECHA"],
            cantidad: ["cantidad", "Cantidad", "CANTIDAD"],
            foto: ["Soporte Fotografico", "Soporte Fotográfico", "Foto", "FOTO"]
        }};

        // Inicializar Pestañas con Fotos de cada Cargo
        function initCargoTabs() {{
            const cargoMap = {{}};
            
            rawData.forEach(item => {{
                const cargoOriginal = getValue(item, KEYS.cargo);
                if (cargoOriginal) {{
                    const photos = getPhotosList(getValue(item, KEYS.foto));
                    if (!cargoMap[cargoOriginal] && photos.length > 0) {{
                        cargoMap[cargoOriginal] = photos[0];
                    }} else if (!cargoMap[cargoOriginal]) {{
                        cargoMap[cargoOriginal] = 'https://via.placeholder.com/100?text=Cargo';
                    }}
                }}
            }});

            const cargosSorted = Object.keys(cargoMap).sort();
            const tabsContainer = document.getElementById('tabs-cargo');
            tabsContainer.innerHTML = '';

            cargosSorted.forEach((c, index) => {{
                const btn = document.createElement('button');
                btn.className = `cargo-tab ${{index === 0 ? 'active' : ''}}`;
                btn.innerHTML = `<img class="cargo-tab-img" src="${{cargoMap[c]}}" onerror="this.src='https://via.placeholder.com/100?text=C';"> ${{c}}`;
                btn.onclick = () => selectCargo(c, btn);
                tabsContainer.appendChild(btn);
            }});

            const btnTodos = document.createElement('button');
            btnTodos.className = 'cargo-tab';
            btnTodos.innerHTML = `📂 Todos los Cargos`;
            btnTodos.onclick = () => selectCargo('TODOS', btnTodos);
            tabsContainer.appendChild(btnTodos);

            selectedCargo = cargosSorted.length > 0 ? cargosSorted[0] : 'TODOS';
        }}

        function selectCargo(cargo, elem) {{
            selectedCargo = cargo;
            document.querySelectorAll('.cargo-tab').forEach(t => t.classList.remove('active'));
            elem.classList.add('active');
            
            const titleLabel = cargo === "TODOS" ? "Resumen de Cantidades - Todos los Cargos" : `Resumen de Cantidades por Incentivo (${{cargo}})`;
            document.getElementById('kpi-section-title').textContent = titleLabel;

            onFilterChange();
        }}

        // Lógica de filtrado general
        function getFilteredData(excludeFilterKey = null) {{
            const fDesde = document.getElementById('f-fecha-desde').value;
            const fHasta = document.getElementById('f-fecha-hasta').value;
            const fNom = document.getElementById('f-nombre').value;
            const fAct = document.getElementById('f-actividad').value;
            const fInc = document.getElementById('f-incentivo').value;
            const fPdv = document.getElementById('f-pdv').value;
            const fCiu = document.getElementById('f-ciudad').value;

            return rawData.filter(item => {{
                if (selectedCargo !== "TODOS" && cleanStr(getValue(item, KEYS.cargo)) !== cleanStr(selectedCargo)) return false;

                if (excludeFilterKey !== 'nombre' && fNom && cleanStr(getValue(item, KEYS.nombre)) !== cleanStr(fNom)) return false;
                if (excludeFilterKey !== 'actividad' && fAct && cleanStr(getValue(item, KEYS.actividad)) !== cleanStr(fAct)) return false;
                if (excludeFilterKey !== 'incentivo' && fInc && cleanStr(getValue(item, KEYS.incentivo)) !== cleanStr(fInc)) return false;
                if (excludeFilterKey !== 'pdv' && fPdv && cleanStr(getValue(item, KEYS.pdv)) !== cleanStr(fPdv)) return false;
                if (excludeFilterKey !== 'ciudad' && fCiu && cleanStr(getValue(item, KEYS.ciudad)) !== cleanStr(fCiu)) return false;

                const fecha = String(getValue(item, KEYS.fecha)).substring(0, 10);
                if (fDesde && fecha < fDesde) return false;
                if (fHasta && fecha > fHasta) return false;

                return true;
            }});
        }}

        // Filtros cruzados/conversacionales entre sí
        function updateDropdownOptions() {{
            updateSingleSelect('f-nombre', getFilteredData('nombre'), KEYS.nombre);
            updateSingleSelect('f-actividad', getFilteredData('actividad'), KEYS.actividad);
            updateSingleSelect('f-incentivo', getFilteredData('incentivo'), KEYS.incentivo);
            updateSingleSelect('f-pdv', getFilteredData('pdv'), KEYS.pdv);
            updateSingleSelect('f-ciudad', getFilteredData('ciudad'), KEYS.ciudad);
        }}

        function updateSingleSelect(elemId, data, keys) {{
            const select = document.getElementById(elemId);
            const currentVal = select.value;
            select.innerHTML = '<option value="">Todos</option>';

            const set = new Set();
            data.forEach(item => {{
                const val = getValue(item, keys);
                if (val) set.add(val);
            }});

            Array.from(set).sort().forEach(v => {{
                const opt = document.createElement('option');
                opt.value = v; opt.textContent = v;
                if (v === currentVal) opt.selected = true;
                select.appendChild(opt);
            }});
        }}

        function onFilterChange() {{
            updateDropdownOptions();
            const filtered = getFilteredData();
            renderKPIsSingleRow(filtered);
            renderGallery(filtered);
        }}

        function renderKPIsSingleRow(data) {{
            const kpiRow = document.getElementById('kpi-row');
            kpiRow.innerHTML = '';

            const incentivosMap = {{}};
            let totalCantidadGeneral = 0;

            data.forEach(item => {{
                const inc = getValue(item, KEYS.incentivo) || "Sin Especificar";
                const cant = parseFloat(getValue(item, KEYS.cantidad)) || 0;

                incentivosMap[inc] = (incentivosMap[inc] || 0) + cant;
                totalCantidadGeneral += cant;
            }});

            const totalCard = document.createElement('div');
            totalCard.className = 'kpi-card';
            totalCard.innerHTML = `
                <div class="kpi-title">TOTAL REGISTROS</div>
                <div class="kpi-value">${{totalCantidadGeneral.toLocaleString()}}</div>
                <div class="kpi-subtext">Unidades Registradas</div>
            `;
            kpiRow.appendChild(totalCard);

            Object.keys(incentivosMap).forEach(incKey => {{
                const val = incentivosMap[incKey];
                const pct = totalCantidadGeneral > 0 ? ((val / totalCantidadGeneral) * 100).toFixed(1) : 0;
                
                const card = document.createElement('div');
                card.className = 'kpi-card';
                card.innerHTML = `
                    <div class="kpi-title" title="${{incKey}}">${{incKey}}</div>
                    <div class="kpi-value">${{val.toLocaleString()}}</div>
                    <div class="kpi-subtext">${{pct}}% del total</div>
                `;
                kpiRow.appendChild(card);
            }});
        }}

        function renderGallery(data) {{
            const grid = document.getElementById('gallery-grid');
            grid.innerHTML = '';

            if (data.length === 0) {{
                grid.innerHTML = '<div class="empty-msg">❌ No hay registros que coincidan con la combinación de filtros seleccionada.</div>';
                return;
            }}

            data.forEach(item => {{
                const fotos = getPhotosList(getValue(item, KEYS.foto));
                const imgCover = fotos.length > 0 ? fotos[0] : 'https://via.placeholder.com/300x200?text=Sin+Imagen';

                const card = document.createElement('div');
                card.className = 'gal-card';
                card.onclick = () => openLightbox(item);

                card.innerHTML = `
                    <div class="gal-img-container">
                        <img class="gal-img" src="${{imgCover}}" loading="lazy" onerror="this.src='https://via.placeholder.com/300x200?text=No+Disponible'">
                        ${{fotos.length > 1 ? `<div class="gal-badge">+${{fotos.length - 1}} fotos</div>` : ''}}
                    </div>
                    <div class="gal-content">
                        <div class="gal-tag">${{getValue(item, KEYS.actividad) || 'General'}}</div>
                        <div class="gal-pdv" title="${{getValue(item, KEYS.pdv)}}">${{getValue(item, KEYS.pdv) || 'PDV N/A'}}</div>
                        <div class="gal-detail">
                            <span>👤 ${{getValue(item, KEYS.nombre) || 'N/A'}}</span>
                            <span>🏙️ ${{getValue(item, KEYS.ciudad) || 'N/A'}}</span>
                        </div>
                        <div class="gal-detail" style="margin-top: 4px; font-weight:700; color: var(--dark-navy);">
                            <span>📅 ${{String(getValue(item, KEYS.fecha)).substring(0, 10)}}</span>
                            <span>📦 Cant: ${{getValue(item, KEYS.cantidad) || 0}}</span>
                        </div>
                    </div>
                `;
                grid.appendChild(card);
            }});
        }}

        function openLightbox(item) {{
            currentRecordPhotos = getPhotosList(getValue(item, KEYS.foto));
            if (currentRecordPhotos.length === 0) {{
                currentRecordPhotos = ['https://via.placeholder.com/800x600?text=Sin+Imagen'];
            }}
            currentPhotoIndex = 0;

            document.getElementById('lb-title').textContent = getValue(item, KEYS.pdv) || 'Detalle del Registro';
            document.getElementById('lb-details').innerHTML = `
                <div class="info-item"><label>Nombre Completo</label><span>${{getValue(item, KEYS.nombre) || 'N/A'}}</span></div>
                <div class="info-item"><label>Cargo</label><span>${{getValue(item, KEYS.cargo) || 'N/A'}}</span></div>
                <div class="info-item"><label>Ciudad</label><span>${{getValue(item, KEYS.ciudad) || 'N/A'}}</span></div>
                <div class="info-item"><label>Actividad</label><span>${{getValue(item, KEYS.actividad) || 'N/A'}}</span></div>
                <div class="info-item"><label>Incentivo</label><span>${{getValue(item, KEYS.incentivo) || 'N/A'}}</span></div>
                <div class="info-item"><label>Cantidad</label><span>${{getValue(item, KEYS.cantidad) || 0}}</span></div>
                <div class="info-item"><label>Fecha</label><span>${{String(getValue(item, KEYS.fecha)).substring(0, 10)}}</span></div>
            `;

            document.getElementById('lb-photo-count').textContent = currentRecordPhotos.length;

            const thumbContainer = document.getElementById('lb-thumbnails');
            thumbContainer.innerHTML = '';
            currentRecordPhotos.forEach((url, i) => {{
                const img = document.createElement('img');
                img.className = `thumb-img ${{i === 0 ? 'active' : ''}}`;
                img.src = url;
                img.onclick = () => setLightboxPhoto(i);
                img.onerror = () => {{ img.src = 'https://via.placeholder.com/100?text=Error'; }};
                thumbContainer.appendChild(img);
            }});

            setLightboxPhoto(0);
            document.getElementById('lightbox').classList.add('active');
        }}

        function setLightboxPhoto(index) {{
            currentPhotoIndex = index;
            document.getElementById('lb-img').src = currentRecordPhotos[currentPhotoIndex];
            
            const thumbs = document.querySelectorAll('.thumb-img');
            thumbs.forEach((t, i) => {{
                if (i === index) t.classList.add('active');
                else t.classList.remove('active');
            }});
        }}

        function changePhoto(step) {{
            currentPhotoIndex = (currentPhotoIndex + step + currentRecordPhotos.length) % currentRecordPhotos.length;
            setLightboxPhoto(currentPhotoIndex);
        }}

        function closeLightbox(e, force = false) {{
            if (force || (e && e.target.id === 'lightbox')) {{
                document.getElementById('lightbox').classList.remove('active');
            }}
        }}

        function resetFilters() {{
            document.getElementById('f-fecha-desde').value = '';
            document.getElementById('f-fecha-hasta').value = '';
            document.getElementById('f-nombre').value = '';
            document.getElementById('f-actividad').value = '';
            document.getElementById('f-incentivo').value = '';
            document.getElementById('f-pdv').value = '';
            document.getElementById('f-ciudad').value = '';
            onFilterChange();
        }}

        function downloadExcel() {{
            const worksheet = XLSX.utils.json_to_sheet(rawData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Actividades Nestlé");
            XLSX.writeFile(workbook, "Base_Entrada_Actividades_Nestle.xlsx");
        }}

        // Inicialización
        initCargoTabs();
        onFilterChange();
    </script>
</body>
</html>'''

    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    print("✅ Dashboard generado exitosamente en:", output_html)

if __name__ == "__main__":
    build_nestle_gallery()