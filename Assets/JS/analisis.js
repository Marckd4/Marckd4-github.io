let consolidatedWorkbook;

function processFile(file, callback) {
    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        const emisor = jsonData[2][1];
        const numeroOrden = jsonData[3][1];

        let df = [];
        let currentItem = null;

        // Iterar sobre las filas del Excel a partir de la fila 17
        jsonData.slice(17).forEach(row => {
            if (row[0] && !String(row[0]).includes('Descripción')) {
                // Es una línea nueva de producto
                if (currentItem) {
                    df.push(currentItem);
                }
                currentItem = {
                    línea: row[0],
                    código: row[1],
                    descripción: "", // Se llenará después
                    cantidad: row[6],
                    unidadEmpacada: row[8]
                };
            } else if (String(row[0]).includes('Descripción')) {
                // Es la descripción extendida del producto
                currentItem.descripción = row[2];
            }
        });

        // Añadir el último item al dataframe
        if (currentItem) {
            df.push(currentItem);
        }

        // Filtro para eliminar filas no deseadas
        df = df.filter(row => 
            !String(row.descripción).includes('Etapa de Transporte') &&
            !String(row.descripción).includes('Modo de Transporte') &&
            !String(row.descripción).includes('Quién Transporta')
        );

        callback(emisor, numeroOrden, df);
    };
    reader.readAsArrayBuffer(file);
}

function saveConsolidatedFile(emisor, numeroOrden, df) {
    let consolidatedData = [
        ['Emisor:', emisor],
        ['Número de Orden de Compra:', numeroOrden],
        [],
        ['Linea', 'Cod. UPC', 'Descripción', 'Cantidad', 'Unid/Emp.']
    ];

    df.forEach(row => {
        consolidatedData.push([row.línea, row.código, row.descripción, row.cantidad, row.unidadEmpacada]);
    });

    consolidatedWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.aoa_to_sheet(consolidatedData);
    XLSX.utils.book_append_sheet(consolidatedWorkbook, newSheet, 'Consolidado');

    document.getElementById('downloadBtn').classList.remove('hidden');
}

function selectFile() {
    const fileInput = document.getElementById('fileInput');
    if (fileInput.files.length === 0) {
        alert('Por favor, selecciona un archivo Excel.');
        return;
    }
    const file = fileInput.files[0];

    processFile(file, (emisor, numeroOrden, df) => {
        saveConsolidatedFile(emisor, numeroOrden, df);
        document.getElementById('output').innerHTML = `<pre>${JSON.stringify(df, null, 2)}</pre>`;
    });
}

function downloadFile() {
    if (consolidatedWorkbook) {
        XLSX.writeFile(consolidatedWorkbook, 'consolidated_file.xlsx');
        alert('Archivo descargado con éxito');
    } else {
        alert('No hay archivo consolidado para descargar');
    }
}
