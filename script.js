// script.js
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const processBtn = document.getElementById('processBtn');
    const fileName = document.getElementById('fileName');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    const resultDiv = document.getElementById('result');
    const output = document.getElementById('output');
    const downloadBtn = document.getElementById('downloadBtn');

    // Datos de provincias/distritos (reemplaza con tu JSON real)
    const distritosProvincias = {
        "LIMA": "LIMA",
        "CALLAO": "CALLAO",
        // Agrega más mapeos según tu infoprovin.json
    };
    const listaProvincias = [...new Set(Object.values(distritosProvincias))];

    class ProcesadorTEMU {
        constructor() {
            this.COLUMNAS_NECESARIAS = [
                "HAWB (EER Guide)", "DNI", "RUC", "Address1", "Department", 
                "Description", "Weight (kgs)", "Names and Surnames", "FOB (US$.)", 
                "Quantity of products", "Invoice No.", "Ruc Sunat", "Nombres Sunat", 
                "Condición Sunat", "Estado Sunat"
            ];
            this.NOMBRE_PROCESADOR = "temu";
            this.distritosProvincias = distritosProvincias;
            this.listaProvincias = listaProvincias;
        }

        procesarDepartment(valor) {
            valor = String(valor).toUpperCase().trim();
            
            // 1. Verificar si es un distrito conocido
            if (this.distritosProvincias[valor]) {
                return this.distritosProvincias[valor];
            }
            
            // 2. Verificar patrones como "MUNICIPALIDAD DE [PROVINCIA]"
            for (const provincia of this.listaProvincias) {
                const provinciaUpper = String(provincia).toUpperCase().trim();
                const regex1 = new RegExp(`MUNICIPALIDAD\\s*(?:DE\\s*)?${provinciaUpper}`, 'i');
                const regex2 = new RegExp(`PROVINCIA\\s*(?:DE\\s*)?${provinciaUpper}`, 'i');
                const regex3 = new RegExp(`REGION\\s*(?:DE\\s*)?${provinciaUpper}`, 'i');
                
                if (regex1.test(valor) || regex2.test(valor) || regex3.test(valor)) {
                    return provinciaUpper;
                }
            }
            
            // 3. Valor por defecto
            return "LIMA";
        }

        determinarTipoDocumento(ruc) {
            const rucStr = String(ruc);
            if (rucStr.length === 11) return "4";
            if (rucStr.length === 8) return "3";
            return "";
        }

        ajustarPesos(df, pesoTotalInicial) {
            const columna = "Weight (kgs)";
            const nuevaCol = "nuevopeso";
            df[nuevaCol] = df[columna].map(val => parseFloat(val) || 0);

            let pesoActual = df[columna].reduce((sum, val) => sum + (parseFloat(val) || 0), 0);
            let pesoFaltante = pesoTotalInicial - pesoActual;

            const umbrales = [1.0, 0.9, 0.8, 0.7, 0.6, 0.5, 0.4];
            const maxIncremento = 1.89;

            for (const umbral of umbrales) {
                const seleccion = df.filter(row => parseFloat(row[columna]) > umbral);
                const n = seleccion.length;
                if (n === 0) continue;

                const incremento = parseFloat((pesoFaltante / n).toFixed(3));
                if (incremento > maxIncremento) continue;

                seleccion.forEach((row, idx) => {
                    const index = df.indexOf(row);
                    df[index][nuevaCol] = parseFloat((df[index][nuevaCol] + incremento).toFixed(3));
                });

                pesoActualizado = df.reduce((sum, row) => sum + (parseFloat(row[nuevaCol]) || 0), 0);
                pesoFaltante = parseFloat((pesoTotalInicial - pesoActualizado).toFixed(6));

                if (Math.abs(pesoFaltante) < 0.0005) break;
            }

            // Ajuste fino final
            const diferenciaFinal = parseFloat((pesoTotalInicial - df.reduce((sum, row) => 
                sum + (parseFloat(row[nuevaCol]) || 0), 0)).toFixed(6));
            
            if (Math.abs(diferenciaFinal) <= 0.7 && diferenciaFinal !== 0) {
                let maxPeso = 0;
                let idxMax = 0;
                df.forEach((row, idx) => {
                    const peso = parseFloat(row[nuevaCol]);
                    if (peso > maxPeso) {
                        maxPeso = peso;
                        idxMax = idx;
                    }
                });
                df[idxMax][nuevaCol] = parseFloat((parseFloat(df[idxMax][nuevaCol]) + parseFloat(diferenciaFinal)).toFixed(3));
            }

            return df;
        }

        crearResumen(df, pesoTotalInicial, numGuiasInicial, numGuiasFinal) {
            const conteoRuc = df.filter(row => String(row["Ruc Sunat"]).length === 11).length;
            const conteoDni = df.filter(row => String(row["Ruc Sunat"]).length === 8).length;

            return [{
                "Procesador": this.NOMBRE_PROCESADOR,
                "Peso Total Inicial": pesoTotalInicial,
                "Diferencia": parseFloat((df.reduce((sum, row) => sum + (parseFloat(row["nuevopeso"]) || 0), 0) - pesoTotalInicial).toFixed(3)),
                "Guías Iniciales": numGuiasInicial,
                "Guías Finales": numGuiasFinal,
                "Peso Total Final": parseFloat(df.reduce((sum, row) => sum + (parseFloat(row["nuevopeso"]) || 0), 0).toFixed(3)),
                "RUCs (11 dígitos)": conteoRuc,
                "DNIs (8 dígitos)": conteoDni
            }];
        }

        preprocesarDatos(df) {
            // Insertar ID
            df.forEach((row, index) => row.id = index + 1);

            // Verificar si existe la columna DNI
            const tieneDni = df.some(row => row.hasOwnProperty("DNI"));

            if (tieneDni) {
                df.forEach(row => {
                    const estadoSunat = String(row["Estado Sunat"] || "").toUpperCase().trim();
                    const condicionSunat = String(row["Condición Sunat"] || "").toUpperCase().trim();
                    const dniValido = row["DNI"] !== undefined && row["DNI"] !== null && String(row["DNI"]).trim() !== "";

                    if ((estadoSunat !== "ACTIVO" || condicionSunat !== "HABIDO") && dniValido) {
                        row["Ruc Sunat"] = String(row["DNI"]);
                        row["Nombres Sunat"] = String(row["Names and Surnames"] || "");
                        row["Estado Sunat"] = "ACTIVO";
                        row["Condición Sunat"] = "HABIDO";
                    }
                });
            }

            // Eliminar columnas no necesarias
            const columnasAEliminar = ["DNI", "RUC"];
            if (df.some(row => row.hasOwnProperty("Names and Surnames"))) {
                columnasAEliminar.push("Names and Surnames");
            }

            df.forEach(row => {
                columnasAEliminar.forEach(col => delete row[col]);
            });

            // Normalizar texto
            const textColumns = ["Estado Sunat", "Condición Sunat", "Nombres Sunat"];
            df.forEach(row => {
                textColumns.forEach(col => {
                    if (row[col] !== undefined) {
                        row[col] = String(row[col]).toUpperCase().trim();
                    }
                });
            });

            // Formatear RUC Sunat
            df.forEach(row => {
                if (row["Ruc Sunat"] !== undefined) {
                    row["Ruc Sunat"] = String(row["Ruc Sunat"]).replace(/\.0$/, '').trim();
                }
            });

            return df;
        }

        filtrarRegistros(df) {
            return df.filter(row => {
                const rucSunat = String(row["Ruc Sunat"] || "");
                const estadoSunat = String(row["Estado Sunat"] || "").toUpperCase();
                const condicionSunat = String(row["Condición Sunat"] || "").toUpperCase();

                const condicion = (
                    estadoSunat !== "ACTIVO" ||
                    rucSunat === "12345678" ||
                    rucSunat === "87654321" ||
                    condicionSunat !== "HABIDO" ||
                    rucSunat.trim() === "" ||
                    rucSunat.length < 8 ||
                    (rucSunat.startsWith("20") && rucSunat.length === 11)
                );

                if (condicion) {
                    row["Ruc Sunat"] = "";
                    row["Nombres Sunat"] = "";
                    return false;
                }
                return true;
            });
        }

        crearHojaParaCopiar(df) {
            const fechaActual = new Date();
            const fechaEmision = `${fechaActual.getDate().toString().padStart(2, '0')}/${(fechaActual.getMonth() + 1).toString().padStart(2, '0')}/${fechaActual.getFullYear()}`;
            
            const fechaFactura = new Date(fechaActual);
            fechaFactura.setDate(fechaActual.getDate() - 15);
            const fechaFacturaStr = `${fechaFactura.getDate().toString().padStart(2, '0')}/${(fechaFactura.getMonth() + 1).toString().padStart(2, '0')}/${fechaFactura.getFullYear()}`;

            return df.map(row => ({
                "GUIA HIJA": row["Invoice No."],
                "FECHA EMISION": fechaEmision,
                "CATEGORIA": "02",
                "RUC TRANSPORTISTA": "",
                "RUC DTEER": "20492060977",
                "TIPO_FACTURA": "03",
                "FACTURA": row["HAWB (EER Guide)"],
                "FECHA FACTURA": fechaFacturaStr,
                "RIESGO IMPORTADOR": "1",
                "INDICADOR DIVER. GUIAS": "",
                "INDICADOR ENV FRONTERA": "",
                "INDICADOR RECONOCIMIENTO FISICO": "N",
                "VIGENCIA DE MERCADERIA": "N",
                "TIPO_DOCUMENTO CONSIGNATARIO": this.determinarTipoDocumento(row["Ruc Sunat"]),
                "NRO DOCUMENTO CONSIGNATARIO": row["Ruc Sunat"],
                "NOMBRE COMPLETO CONSIGNATARIO": row["Nombres Sunat"],
                "DIRECCION": row["Address1"],
                "CIUDAD": row["Department"],
                "PAIS": "PE",
                "TELEFONO": "",
                "CORREO": "",
                "TIPO DOCUMENTO EMBARCADOR": "",
                "NRO DOCUMENTO EMBARCADOR": "",
                "NOMBRE COMPLETO EMBARCADOR": "YUEMA EXPRESS VENDOR",
                "TIPO_DOCUMENTO REMITENTE": "",
                "NRO DOCUMENTO REMITENTE": "",
                "NOMBRE COMPLETO REMITENTE": "TEMU",
                "DIRECCION1": "",
                "CIUDAD1": "",
                "PAIS1": "",
                "ITEM": "0001",
                "PARTIDA ARANCELARIA": "9809000020",
                "TNAN": "",
                "PAIS ORIGEN": "CN",
                "VALOR FOB": row["FOB (US$.)"],
                "AJUSTE": "",
                "MONEDA FLETE": "USD",
                "FLETE": 7.000,
                "SEGURO": "",
                "PESO BRUTO": row["nuevopeso"],
                "CANT BULTOS": 1.000,
                "UNIDAD FISICA": row["Quantity of products"],
                "TIPO UNIDAD COMERCIAL": "U",
                "UNIDAD COMERCIAL": row["Quantity of products"],
                "DESCRIPCION MERCANCIA": row["Description"],
                "DESCRIPCION COMERCIAL": row["Description"],
                "MARCA": "S/MARCA",
                "MODELO": "S/MODELO",
                "ESTADO": "10",
                "MERCANCIA RESTRINGIDA": "",
                "CATEGORIA RIESGO": "",
                "CONTENIDO": row["Description"],
                "TIPO DE SEGURO": "1",
                "REFERENCIA": "TEMU"
            }));
        }

        procesarArchivo(jsonData) {
            try {
                // Validar columnas necesarias
                const columnasFaltantes = this.COLUMNAS_NECESARIAS.filter(col => 
                    !jsonData[0] || !jsonData[0].hasOwnProperty(col)
                );
                
                if (columnasFaltantes.length > 0) {
                    throw new Error(`Faltan columnas requeridas: ${columnasFaltantes.join(", ")}`);
                }

                const numGuiasInicial = jsonData.length;
                const pesoTotalInicial = jsonData.reduce((sum, row) => 
                    sum + (parseFloat(row["Weight (kgs)"]) || 0), 0);

                // Preprocesamiento
                let datosProcesados = this.preprocesarDatos(jsonData);

                // Filtrar registros
                datosProcesados = this.filtrarRegistros(datosProcesados);
                const numGuiasFinal = datosProcesados.length;

                // Ajustar pesos
                datosProcesados = this.ajustarPesos(datosProcesados, pesoTotalInicial);

                // Determinar tipos de documento
                datosProcesados.forEach(row => {
                    row["Tipo_De_Documento"] = this.determinarTipoDocumento(row["Ruc Sunat"]);
                });

                // Procesar departamentos
                datosProcesados.forEach(row => {
                    row["Department"] = this.procesarDepartment(row["Department"]);
                });

                // Limitar longitud de campos
                datosProcesados.forEach(row => {
                    if (row["Address1"]) row["Address1"] = String(row["Address1"]).substring(0, 59).toUpperCase();
                    if (row["Description"]) row["Description"] = String(row["Description"]).substring(0, 149);
                    if (row["HAWB (EER Guide)"]) row["HAWB (EER Guide)"] = String(row["HAWB (EER Guide)"]).replace(/\.0$/, '');
                });

                // Crear hojas de salida
                const paraCopiar = this.crearHojaParaCopiar(datosProcesados);
                const resumen = this.crearResumen(datosProcesados, pesoTotalInicial, numGuiasInicial, numGuiasFinal);

                return {
                    success: true,
                    message: "Archivo TEMU procesado correctamente",
                    datosProcesados,
                    paraCopiar,
                    resumen
                };
            } catch (error) {
                return {
                    success: false,
                    message: `Error al procesar archivo TEMU: ${error.message}`
                };
            }
        }
    }

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
        if (!file) {
            output.textContent = "❌ No se seleccionó ningún archivo";
            resultDiv.classList.remove('hidden');
            return;
        }

        updateProgress(10, "Leyendo archivo...");
        
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                updateProgress(30, "Procesando datos...");
                
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                let jsonData = XLSX.utils.sheet_to_json(firstSheet);
                
                // Procesar con nuestra clase
                const procesador = new ProcesadorTEMU();
                const resultado = procesador.procesarArchivo(jsonData);
                
                if (!resultado.success) {
                    throw new Error(resultado.message);
                }
                
                // Crear nuevo libro de Excel con las 3 hojas
                const newWorkbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(
                    newWorkbook, 
                    XLSX.utils.json_to_sheet(resultado.datosProcesados), 
                    "Datos Procesados"
                );
                XLSX.utils.book_append_sheet(
                    newWorkbook, 
                    XLSX.utils.json_to_sheet(resultado.paraCopiar), 
                    "Para Copiar"
                );
                XLSX.utils.book_append_sheet(
                    newWorkbook, 
                    XLSX.utils.json_to_sheet(resultado.resumen), 
                    "Resumen"
                );
                
                updateProgress(70, "Preparando resultados...");
                output.textContent = JSON.stringify(resultado.datosProcesados.slice(0, 5), null, 2);
                
                processedData = newWorkbook;
                updateProgress(100, "¡Completado!");
                resultDiv.classList.remove('hidden');
                downloadBtn.classList.remove('hidden');
            } catch (error) {
                output.textContent = `❌ Error: ${error.message}`;
                updateProgress(0, "Error en el procesamiento");
                resultDiv.classList.remove('hidden');
            }
        };
        reader.onerror = () => {
            output.textContent = "❌ Error al leer el archivo";
            updateProgress(0, "Error");
            resultDiv.classList.remove('hidden');
        };
        reader.readAsArrayBuffer(file);
    });

    // Descargar resultados
    downloadBtn.addEventListener('click', () => {
        if (!processedData) return;
        XLSX.writeFile(processedData, 'procesado_temu.xlsx');
    });

    function updateProgress(value, text) {
        progressBar.value = value;
        progressText.textContent = text;
    }

    // Variable para almacenar los datos procesados
    let processedData = null;
});
