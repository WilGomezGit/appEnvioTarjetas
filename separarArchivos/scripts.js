function processFile() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        alert("Por favor selecciona un archivo");
        return;
    }

    // Validar que el nombre del archivo comience con "CB636876"
    if (!file.name.startsWith("CB636876")) {
        alert("*** ¡Error! *** Por favor, cargue el archivo CB correcto.");
        return;
    }

    const reader = new FileReader();
    reader.onload = function(event) {
        const fileContent = event.target.result;
        const lines = fileContent.split('\n');
        const tableBody = document.getElementById('dataBody');
        tableBody.innerHTML = ''; // Limpiar filas existentes

        const empresaCounts = {}; // Objeto para almacenar el conteo de cada empresa

        // Primera pasada para contar ocurrencias de cada empresa
        lines.forEach((line) => {
            if (line.trim()) {
                const empresa = line.substring(72, 113).trim();
                if (empresa) {
                    empresaCounts[empresa] = (empresaCounts[empresa] || 0) + 1;
                }
            }
        });

        // Segunda pasada para crear las filas de la tabla
        lines.forEach((line, index) => {
            if (line.trim()) {
                const cedula = line.substring(42, 54).trim().replace(/^0+/, '');
                const tarjeta = line.substring(56, 72).trim();
                const empresa = line.substring(72, 113).trim();
                const cant = empresaCounts[empresa] || 0;

                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${index + 1}</td>
                    <td>${cedula}</td>
                    <td>${tarjeta}</td>
                    <td>${empresa}</td>
                    <td>${cant}</td>
                `;
                tableBody.appendChild(row);
            }
        });
    };
    reader.readAsText(file);
}

function clearTable() {
    const tableBody = document.getElementById('dataBody');
    tableBody.innerHTML = '';
}

document.getElementById('exportButton').addEventListener('click', function () {
    exportToExcel('archivo.xlsx');
});

function exportToExcel(filename = '') {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        alert("Por favor selecciona un archivo");
        return;
    }

    const reader = new FileReader();
    reader.onload = function(event) {
        const fileContent = event.target.result;
        const lines = fileContent.split('\n');

        // Process the original table
        const originalData = [];
        const empresaCounts = {};
        lines.forEach((line, index) => {
            if (line.trim()) {
                const cedula = line.substring(42, 54).trim().replace(/^0+/, '');
                const tarjeta = line.substring(56, 72).trim();
                const empresa = line.substring(72, 113).trim();
                empresaCounts[empresa] = (empresaCounts[empresa] || 0) + 1;

                originalData.push([index + 1, cedula, tarjeta, empresa, empresaCounts[empresa]]);
            }
        });

        // Process the unique table
        const uniqueData = [];
        const seenEmpresas = new Set();
        lines.forEach((line, index) => {
            if (line.trim()) {
                const cedula = line.substring(42, 54).trim().replace(/^0+/, '');
                const tarjeta = line.substring(56, 72).trim();
                const empresa = line.substring(72, 113).trim();

                if (!seenEmpresas.has(empresa)) {
                    const cant = empresaCounts[empresa] || 0;
                    uniqueData.push([index + 1, cedula, tarjeta, empresa, cant]);
                    seenEmpresas.add(empresa);
                }
            }
        });

        // Create workbook and worksheets
        const wb = XLSX.utils.book_new();

        // Define worksheet for original data
        const wsOriginal = XLSX.utils.aoa_to_sheet([['No', 'CEDULA', 'TARJETA', 'EMPRESA', 'CANT'], ...originalData]);

        // Define worksheet for unique data
        const wsUnique = XLSX.utils.aoa_to_sheet([['No', 'CEDULA', 'TARJETA', 'EMPRESA', 'CANT'], ...uniqueData]);

        // Center the headers in both worksheets
        const centerAlign = { alignment: { horizontal: 'center' } };
        const headerRange = { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } };

        // Apply the styles to header cells
        for (let i = headerRange.s.c; i <= headerRange.e.c; i++) {
            const cell = wsOriginal[XLSX.utils.encode_cell({ r: 0, c: i })];
            const uniqueCell = wsUnique[XLSX.utils.encode_cell({ r: 0, c: i })];
            if (cell) cell.s = { ...cell.s, ...centerAlign };
            if (uniqueCell) uniqueCell.s = { ...uniqueCell.s, ...centerAlign };
        }

        XLSX.utils.book_append_sheet(wb, wsOriginal, 'Original');
        XLSX.utils.book_append_sheet(wb, wsUnique, 'SinDuplicados');

        XLSX.writeFile(wb, filename);
    };
    reader.readAsText(file);
}
