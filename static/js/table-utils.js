/**
 * SISGEMOSI - Utilidades de Tabla y Sistema
 * Funciones compartidas para ordenamiento, filtrado, búsqueda y Modo Diurno/Nocturno.
 */

// ===== MODO DIURNO Y NOCTURNO =====
function toggleTheme() {
    const currentTheme = document.documentElement.getAttribute('data-bs-theme') || 'light';
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-bs-theme', newTheme);
    localStorage.setItem('sisgemosi_theme', newTheme);
    actualizarIconoTema(newTheme);
}

function actualizarIconoTema(theme) {
    const icon = document.getElementById('themeIcon');
    const btn = document.getElementById('btnThemeToggle');
    if (!icon) return;
    if (theme === 'dark') {
        icon.className = 'bi bi-sun-fill text-warning';
        if (btn) {
            btn.setAttribute('title', 'Cambiar a Modo Diurno (Claro)');
            btn.classList.remove('btn-outline-light');
            btn.classList.add('btn-outline-warning');
        }
    } else {
        icon.className = 'bi bi-moon-stars-fill text-white';
        if (btn) {
            btn.setAttribute('title', 'Cambiar a Modo Nocturno (Oscuro)');
            btn.classList.remove('btn-outline-warning');
            btn.classList.add('btn-outline-light');
        }
    }
}

// ===== ORDENAMIENTO DE TABLAS =====
function ordenarTabla(tableId, colIndex, tipo) {
    const tabla = document.getElementById(tableId);
    if (!tabla) return;
    const tbody = tabla.querySelector('tbody');
    if (!tbody) return;
    
    const filas = Array.from(tbody.querySelectorAll('tr'));
    const th = tabla.querySelectorAll('thead th')[colIndex];
    
    // Determinar dirección
    let ascendente = true;
    if (th && th.dataset.sortDir === 'asc') {
        ascendente = false;
        th.dataset.sortDir = 'desc';
    } else if (th) {
        th.dataset.sortDir = 'asc';
    }
    
    // Limpiar indicadores de otras columnas
    tabla.querySelectorAll('thead th').forEach(h => {
        if (h !== th) delete h.dataset.sortDir;
        const arrow = h.querySelector('.sort-arrow');
        if (arrow && h !== th) arrow.textContent = '';
    });
    
    // Actualizar flecha visual
    if (th) {
        let arrow = th.querySelector('.sort-arrow');
        if (!arrow) {
            arrow = document.createElement('span');
            arrow.className = 'sort-arrow ms-1';
            th.appendChild(arrow);
        }
        arrow.textContent = ascendente ? '▲' : '▼';
    }
    
    filas.sort((a, b) => {
        let valA = a.cells[colIndex]?.textContent?.trim() || '';
        let valB = b.cells[colIndex]?.textContent?.trim() || '';
        
        if (tipo === 'numero') {
            valA = parseFloat(valA.replace(/[^\d.-]/g, '')) || 0;
            valB = parseFloat(valB.replace(/[^\d.-]/g, '')) || 0;
        } else if (tipo === 'fecha') {
            // DD/MM/YYYY
            const pA = valA.split('/'); const pB = valB.split('/');
            valA = pA.length === 3 ? new Date(pA[2], pA[1]-1, pA[0]).getTime() : 0;
            valB = pB.length === 3 ? new Date(pB[2], pB[1]-1, pB[0]).getTime() : 0;
        } else {
            valA = valA.toLowerCase();
            valB = valB.toLowerCase();
        }
        
        if (valA < valB) return ascendente ? -1 : 1;
        if (valA > valB) return ascendente ? 1 : -1;
        return 0;
    });
    
    filas.forEach(fila => tbody.appendChild(fila));
}

// ===== FILTRADO DE TABLAS =====
function filtrarTabla(tableId, inputId, columnas) {
    const input = document.getElementById(inputId);
    const tabla = document.getElementById(tableId);
    if (!input || !tabla) return;
    
    const texto = input.value.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const filas = tabla.querySelectorAll('tbody tr');
    let visibles = 0;
    
    filas.forEach(fila => {
        let coincide = false;
        const cols = columnas || Array.from({length: fila.cells.length}, (_, i) => i);
        
        for (const col of cols) {
            const celda = fila.cells[col];
            if (celda) {
                const contenido = celda.textContent.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
                if (contenido.includes(texto)) {
                    coincide = true;
                    break;
                }
            }
        }
        
        fila.style.display = coincide ? '' : 'none';
        if (coincide) visibles++;
    });
    
    // Actualizar contador si existe
    const counter = document.getElementById(tableId + '-count');
    if (counter) counter.textContent = visibles;
}

// ===== FILTRADO POR SELECT =====
function filtrarPorSelect(tableId, selectId, colIndex) {
    const select = document.getElementById(selectId);
    const tabla = document.getElementById(tableId);
    if (!select || !tabla) return;
    
    const valor = select.value.toLowerCase();
    const filas = tabla.querySelectorAll('tbody tr');
    
    filas.forEach(fila => {
        if (!valor) {
            fila.style.display = '';
            return;
        }
        const celda = fila.cells[colIndex];
        const contenido = celda ? celda.textContent.toLowerCase().trim() : '';
        fila.style.display = contenido.includes(valor) ? '' : 'none';
    });
}

// ===== INICIALIZACIÓN AL CARGAR =====
document.addEventListener('DOMContentLoaded', function() {
    // Sincronizar icono del tema
    const savedTheme = localStorage.getItem('sisgemosi_theme') || 'light';
    actualizarIconoTema(savedTheme);

    // Inicializar tooltips de Bootstrap
    const tooltipTriggerList = document.querySelectorAll('[data-bs-toggle="tooltip"]');
    tooltipTriggerList.forEach(el => new bootstrap.Tooltip(el));
});
