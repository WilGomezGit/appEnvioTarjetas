let loadedData = [];
let searchData = [];
let searchResults = [];
let currentPage = 1;
const rowsPerPage = 50;

// Procesar el archivo de datos
function processDataFile() {
    const fileInput = document.getElementById('fileInputData');
    const file = fileInput.files[0];

    if (!file) {
        alert("Por favor selecciona un archivo de datos");
        return;
    }

    // Validar el tipo de archivo
    const allowedExtensions = ['xlsx', 'xls', 'ods'];
    const fileExtension = file.name.split('.').pop().toLowerCase();
    if (!allowedExtensions.includes(fileExtension)) {
        alert('Debes cargar un archivo en Excel (.xlsx, .xls) o en OpenDocument (.ods)');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        loadedData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

        // Remover encabezado
        loadedData.shift();

        displayPage(1); // Mostrar la primera página de los datos cargados
    };
    reader.readAsArrayBuffer(file);
}

// Procesar el archivo de búsqueda
function processSearchFile() {
    const fileInput = document.getElementById('fileInputSearch');
    const file = fileInput.files[0];

    if (!file) {
        alert("Por favor selecciona un archivo de búsqueda");
        return;
    }

    // Validar el tipo de archivo
    const allowedExtensions = ['xlsx', 'xls', 'ods'];
    const fileExtension = file.name.split('.').pop().toLowerCase();
    if (!allowedExtensions.includes(fileExtension)) {
        alert('Debes cargar un archivo en Excel (.xlsx, .xls) o en OpenDocument (.ods)');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        searchData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        searchResults = searchData.map(row => {
            const result = loadedData.find(dataRow => dataRow[0] === row[0]);
            return result ? result : [row[0], '', '', '', '', '', '', '', '', '', '', 'NO ESTÁ EN LA BASE DE DATOS'];
        });

        // Exportar los resultados directamente sin mostrar en la página
        exportToExcel('resultados.xlsx');
    };
    reader.readAsArrayBuffer(file);
}

// Limpiar la tabla
function clearTable() {
    const tableBody = document.getElementById('dataBody');
    tableBody.innerHTML = '';
    loadedData = [];
    searchResults = [];
    currentPage = 1;
    displayPagination();
}

// Cancelar selección de archivo de búsqueda
function cancelSearchFile() {
    const fileInput = document.getElementById('fileInputSearch');
    fileInput.value = '';
    searchData = [];
    searchResults = [];
    const tableBody = document.getElementById('dataBody');
    tableBody.innerHTML = '';
}

// Mostrar página específica
function displayPage(page) {
    const tableBody = document.getElementById('dataBody');
    tableBody.innerHTML = '';
    currentPage = page;

    const start = (page - 1) * rowsPerPage;
    const end = start + rowsPerPage;

    const pageData = loadedData.slice(start, end);

    pageData.forEach((row, index) => {
        const tableRow = document.createElement('tr');
        tableRow.innerHTML = `
            <td>${start + index + 1}</td>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${row[2]}</td>
            <td>${row[3]}</td>
            <td>${row[4]}</td>
            <td>${row[5]}</td>
            <td>${row[6]}</td>
            <td>${row[7]}</td>
            <td>${row[8]}</td>
            <td>${row[9]}</td>
            <td>${row[10]}</td>
            <td>${row[11]}</td>
        `;
        tableBody.appendChild(tableRow);
    });

    displayPagination();
}

// Mostrar resultados de búsqueda
function displaySearchResults() {
    const tableBody = document.getElementById('dataBody');
    tableBody.innerHTML = '';

    searchResults.forEach((row, index) => {
        const tableRow = document.createElement('tr');
        tableRow.innerHTML = `
            <td>${index + 1}</td>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${row[2]}</td>
            <td>${row[3]}</td>
            <td>${row[4]}</td>
            <td>${row[5]}</td>
            <td>${row[6]}</td>
            <td>${row[7]}</td>
            <td>${row[8]}</td>
            <td>${row[9]}</td>
            <td>${row[10]}</td>
            <td>${row[11]}</td>
            <td>${row[12]}</td> <!-- Mensaje adicional en la última columna -->
        `;
        tableBody.appendChild(tableRow);
    });

    displayPagination();
}

// Mostrar paginación
function displayPagination() {
    const paginationDiv = document.getElementById('pagination');
    paginationDiv.innerHTML = '';

    const totalPages = Math.ceil(loadedData.length / rowsPerPage);

    if (totalPages > 1) {
        // Botón de primera página
        const firstButton = document.createElement('button');
        firstButton.textContent = 'Primera';
        firstButton.addEventListener('click', () => displayPage(1));
        paginationDiv.appendChild(firstButton);

        // Botones numerados
        for (let i = 1; i <= totalPages; i++) {
            const pageButton = document.createElement('button');
            pageButton.textContent = i;
            pageButton.className = i === currentPage ? 'active' : '';
            pageButton.addEventListener('click', () => displayPage(i));
            paginationDiv.appendChild(pageButton);
        }

        // Botón de última página
        const lastButton = document.createElement('button');
        lastButton.textContent = 'Última';
        lastButton.addEventListener('click', () => displayPage(totalPages));
        paginationDiv.appendChild(lastButton);
    }
}

// Exportar a Excel
function exportToExcel(filename = '') {
    if (searchResults.length === 0) {
        alert('No hay datos de búsqueda para exportar.');
        return;
    }

    // Añadir consecutivo en la primera columna
    const wsData = [
        ['No', 'NIT', 'Raz. Social', 'Dirección Local', 'Municipio Local', 'Celular Local', 'Teléfono Local', 'Zona', 'Dirección Principal', 'Municipio Principal', 'Celular Principal', 'Teléfono Principal', 'Email Principal', 'Estado'],
        ...searchResults.map((row, index) => [index + 1, ...row]) // Añadir consecutivo
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'Resultados');
    XLSX.writeFile(wb, filename);
}

// Funciones para los botones de desplazamiento
function scrollToTop() {
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function scrollToBottom() {
    window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
}
