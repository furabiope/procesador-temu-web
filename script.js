document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const processBtn = document.getElementById('processBtn');
    const fileName = document.getElementById('fileName');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    const resultDiv = document.getElementById('result');
    const output = document.getElementById('output');
    const downloadBtn = document.getElementById('downloadBtn');

    let processedData = null;

    // Evento al seleccionar archivo
    // Función para eliminar columnas no deseadas
    function eliminarColumnas(jsonData, columnasAEliminar) {
        return jsonData.map(row => {
            const nuevoRow = {...row};
            columnasAEliminar.forEach(col => delete nuevoRow[col]);
            return nuevoRow;
        });
    }
    
    // Modifica el evento de clic del botón "Procesar"
    processBtn.addEventListener('click', () => {
        const file = fileInput.files[0];
        if (!file) return;
    
        updateProgress(10, "Leyendo archivo...");
        
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                updateProgress(30, "Procesando datos...");
                
                // Leer el archivo Excel
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                let jsonData = XLSX.utils.sheet_to_json(firstSheet);
                
                // Columnas a eliminar (ajusta según tus necesidades)
                const columnasAEliminar = ["DNI", "RUC", "Names and Surnames"]; 
                
                // Eliminar columnas
                jsonData = eliminarColumnas(jsonData, columnasAEliminar);
                
                // Crear nuevo libro de trabajo sin las columnas
                const newWorkbook = XLSX.utils.book_new();
                const newWorksheet = XLSX.utils.json_to_sheet(jsonData);
                XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Datos Procesados");
                
                // Mostrar preview (primeras 5 filas)
                updateProgress(70, "Preparando resultados...");
                output.textContent = JSON.stringify(jsonData.slice(0, 5), null, 2);
                
                // Guardar datos para descarga
                processedData = newWorkbook;
                
                updateProgress(100, "¡Completado!");
                resultDiv.classList.remove('hidden');
                downloadBtn.classList.remove('hidden');
            } catch (error) {
                output.textContent = `Error: ${error.message}`;
                updateProgress(0, "Error en el procesamiento");
                resultDiv.classList.remove('hidden');
            }
        };
        reader.readAsArrayBuffer(file);
    });

    // Descargar resultados
    downloadBtn.addEventListener('click', () => {
        if (!processedData) return;
        
        XLSX.writeFile(processedData, 'procesado_temu.xlsx');
    });

    // Actualizar barra de progreso
    function updateProgress(value, text) {
        progressBar.value = value;
        progressText.textContent = text;
    }
});
