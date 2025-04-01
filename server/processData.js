import { parse } from 'csv-parse';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export async function processCSVFile(filename, importesReferencia) {
    try {
        const filePath = path.join(__dirname, '../uploads', filename);
        const fileContent = await fs.readFile(filePath, 'utf-8');

        const records = await new Promise((resolve, reject) => {
            parse(fileContent, {
                columns: true,
                skip_empty_lines: true
            }, (err, records) => {
                if (err) reject(err);
                else resolve(records);
            });
        });

        // Procesar los datos
        const resultados = [];
        const cuits = [...new Set(records.map(row => row.cuit))];

        for (const cuit of cuits) {
            const registrosCuit = records.filter(row => row.cuit === cuit);
            const diferenciaPorCuil = new Map(); // Para almacenar las diferencias por CUIL y año

            for (const registro of registrosCuit) {
                const cuil = registro.cuil;
                const anio = parseInt(registro.anio);
                const key = `${cuil}-${anio}`;

                if (!diferenciaPorCuil.has(key)) {
                    diferenciaPorCuil.set(key, 0);
                }

                const meses = [
                    'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
                    'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
                ];

                meses.forEach((mes, index) => {
                    const nombreColumna = `aporte_${mes}`;
                    if (registro[nombreColumna] !== undefined) {
                        const aporte = parseFloat(registro[nombreColumna]);

                        // Buscar el importe de referencia para este mes y año
                        const importeRef = importesReferencia.find(imp =>
                            imp.anio === anio &&
                            imp.mes === (index + 1)
                        );

                        if (importeRef && aporte < importeRef.importe) {
                            const diferencia = importeRef.importe - aporte;
                            diferenciaPorCuil.set(key, diferenciaPorCuil.get(key) + diferencia);
                        }
                    }
                });
            }

            // Agrupar las diferencias por CUIL y año
            const diferenciasDetalladas = Array.from(diferenciaPorCuil.entries()).map(([key, diferencia]) => {
                const [cuil, anio] = key.split('-');
                return {
                    cuil,
                    anio: parseInt(anio),
                    diferencia
                };
            });

            // Calcular la diferencia total para este CUIT
            const diferenciaTotal = diferenciasDetalladas.reduce((total, item) => total + item.diferencia, 0);

            resultados.push({
                cuit,
                diferenciasDetalladas,
                diferenciaTotal
            });
        }

        return resultados;
    } catch (error) {
        console.error('Error procesando archivo:', error);
        throw error;
    }
} 