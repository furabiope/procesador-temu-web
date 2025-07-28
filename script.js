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
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            fileName.textContent = e.target.files[0].name;
            processBtn.disabled = false;
            resultDiv.classList.add('hidden');
        }
    });

    // Procesar archivo
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
                
                // Simular procesamiento (DEMO)
                setTimeout(() => {
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                    
                    // Mostrar primeras 5 filas como demo
                    const demoData = jsonData.slice(0, 5).map(row => row.join('\t')).join('\n');
                    output.textContent = demoData;
                    
                    // Guardar datos para descarga
                    processedData = workbook;
                    
                    updateProgress(100, "¡Completado!");
                    resultDiv.classList.remove('hidden');
                    downloadBtn.classList.remove('hidden');
                }, 1000);
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