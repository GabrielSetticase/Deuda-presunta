import odbc from 'odbc';
import path from 'path';
import { getDatabase } from './database.js';
import XLSX from 'xlsx';

// Función para obtener la fecha de hace 120 meses desde la fecha actual
function getFechaInicio() {
    const fechaActual = new Date();
    return new Date(
        fechaActual.getFullYear(),
        fechaActual.getMonth() - 120,
        1
    );
}

async function obtenerUltimosPeriodos(actasPath) {
    try {
        console.log('Conectando a base de datos de actas...');
        const connectionString = `Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=${actasPath}`;
        const connection = await odbc.connect(connectionString);

        // Consulta para obtener el último período por CUIT desde la tabla correcta
        const query = `
            SELECT CUIT, MAX(PERIODO_HASTA) as ultimo_periodo
            FROM VWRPT_ACTAS_INSPECTOR_IN
            GROUP BY CUIT
        `;

        console.log('Ejecutando consulta de últimos períodos...');
        const rows = await connection.query(query);
        await connection.close();

        // Crear un Map con los últimos períodos
        const ultimosPeriodos = new Map();
        for (const row of rows) {
            if (row.CUIT && row.ultimo_periodo) {
                const fecha = new Date(row.ultimo_periodo);
                // Avanzar al mes siguiente
                fecha.setMonth(fecha.getMonth() + 1);
                // Establecer día 1 del mes
                fecha.setDate(1);
                ultimosPeriodos.set(row.CUIT, fecha);
                console.log(`CUIT ${row.CUIT}: Último período ${row.ultimo_periodo}, comenzar desde ${fecha.toISOString()}`);
            }
        }

        return ultimosPeriodos;
    } catch (error) {
        console.error('Error obteniendo últimos períodos:', error);
        throw error;
    }
}

export async function processODBFile(cuilesPath, importes, actasPath, updateProgress = () => { }) {
    try {
        console.log('Iniciando procesamiento...');
        updateProgress('Conectando a las bases de datos...');

        // Obtener los últimos períodos del archivo de actas
        console.log('Obteniendo últimos períodos...');
        const ultimosPeriodos = await obtenerUltimosPeriodos(actasPath);

        // Obtener la fecha de inicio para CUITs sin período en ACTAS
        const fechaInicioDefault = getFechaInicio();
        console.log(`Fecha de inicio por defecto (120 meses atrás): ${fechaInicioDefault.toISOString()}`);

        // Conectar a la base de datos de CUILES
        console.log('Conectando a la base de datos de CUILES...');
        const connectionString = `Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=${cuilesPath}`;
        const connection = await odbc.connect(connectionString);

        // Intentar diferentes nombres de tabla posibles para los aportes
        const posiblesTablasVistas = [
            'VW_DJBENEF_ANNIO_AP_ADIC_DEL',
            'VW_DDJJF931_ANIO_AP_CUIT',
            'DDJJF931_ANIO_AP_CUIT',
            'VW_DDJJ_ANIO_AP_CUIT',
            'DDJJ_ANIO_AP_CUIT',
            'DDJJ_APORTES'
        ];

        let tableName = null;
        for (const tabla of posiblesTablasVistas) {
            try {
                console.log(`Intentando acceder a la tabla/vista: ${tabla}`);
                const testQuery = `SELECT TOP 1 * FROM ${tabla}`;
                await connection.query(testQuery);
                tableName = tabla;
                console.log(`Tabla encontrada: ${tabla}`);
                break;
            } catch (err) {
                console.log(`La tabla ${tabla} no existe o no es accesible`);
            }
        }

        if (!tableName) {
            throw new Error('No se encontró ninguna tabla válida en la base de datos');
        }

        // Procesar registros
        const BATCH_SIZE = 500;
        let procesados = 0;
        const resultados = new Map();

        while (true) {
            const batchQuery = `
                SELECT TOP ${BATCH_SIZE} *
                FROM ${tableName}
                ORDER BY CUIT, CUIL, ANIO
            `;

            const rows = await connection.query(batchQuery);
            if (!rows || rows.length === 0) break;

            for (const row of rows) {
                const cuit = row.CUIT;
                const cuil = row.CUIL;
                const anio = parseInt(row.ANIO);

                // Obtener el período desde el cual debemos procesar
                let fechaInicio = ultimosPeriodos.get(cuit);
                if (!fechaInicio) {
                    console.log(`No se encontró período para CUIT ${cuit}, usando fecha de inicio por defecto (120 meses atrás)`);
                    fechaInicio = fechaInicioDefault;
                }

                // Inicializar el registro de resultados para este CUIT si no existe
                if (!resultados.has(cuit)) {
                    resultados.set(cuit, {
                        cuit,
                        diferenciasDetalladas: new Map(),
                        diferenciaTotal: 0,
                        ultimoPeriodo: fechaInicio.toISOString()
                    });
                }

                const registro = resultados.get(cuit);
                const key = `${cuil}-${anio}`;

                if (!registro.diferenciasDetalladas.has(key)) {
                    registro.diferenciasDetalladas.set(key, 0);
                }

                // Procesar cada mes
                const meses = [
                    'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
                    'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'
                ];

                meses.forEach((mes, index) => {
                    const fechaActual = new Date(anio, index, 1);
                    // Solo procesar si la fecha es posterior al último período
                    if (fechaActual < fechaInicio) {
                        return;
                    }

                    const aporteColumn = `APORTE_${mes}`;
                    const aporteAdicColumn = `APORTE_ADIC_OS_${mes}`;

                    // Verificar si el campo existe y no está vacío
                    const tieneAporte = row[aporteColumn] !== undefined &&
                        row[aporteColumn] !== null &&
                        row[aporteColumn] !== '' &&
                        !isNaN(parseFloat(row[aporteColumn]));

                    if (tieneAporte) {
                        const aporte = parseFloat(row[aporteColumn] || 0);
                        const aporteAdic = parseFloat(row[aporteAdicColumn] || 0);
                        const aporteTotal = aporte + aporteAdic;

                        const importeRef = importes.find(imp =>
                            imp.anio === anio &&
                            imp.mes === (index + 1)
                        );

                        if (importeRef && aporteTotal < importeRef.importe) {
                            const diferencia = importeRef.importe - aporteTotal;
                            const diferenciaActual = registro.diferenciasDetalladas.get(key);
                            registro.diferenciasDetalladas.set(key, diferenciaActual + diferencia);
                            registro.diferenciaTotal += diferencia;

                            console.log(`Diferencia encontrada - CUIT: ${cuit}, CUIL: ${cuil}, Período: ${mes}/${anio}, Aporte: ${aporteTotal}, Referencia: ${importeRef.importe}, Diferencia: ${diferencia}`);
                        }
                    }
                });

                if (procesados % 100 === 0) {
                    updateProgress(`Procesando registros... ${procesados}`, cuit, procesados);
                }
                procesados++;
            }

            if (rows.length < BATCH_SIZE) break;
        }

        await connection.close();

        // Convertir resultados a formato final
        const resultadosFinales = Array.from(resultados.values())
            .filter(r => r.diferenciaTotal > 0) // Solo incluir CUITs con diferencias
            .map(registro => ({
                cuit: registro.cuit,
                ultimoPeriodo: registro.ultimoPeriodo,
                diferenciasDetalladas: Array.from(registro.diferenciasDetalladas.entries())
                    .filter(([_, diferencia]) => diferencia > 0) // Solo incluir detalles con diferencias
                    .map(([key, diferencia]) => {
                        const [cuil, anio] = key.split('-');
                        return {
                            cuil,
                            anio: parseInt(anio),
                            diferencia
                        };
                    }),
                diferenciaTotal: registro.diferenciaTotal
            }));

        console.log('Procesamiento completado');
        updateProgress(`Procesamiento completado. Total: ${procesados} registros procesados`);
        return resultadosFinales;

    } catch (error) {
        console.error('Error en processODBFile:', error);
        updateProgress(`Error: ${error.message}`);
        throw error;
    }
}

export async function processSueldosODB(filePath) {
    try {
        // Determinar el tipo de archivo
        const fileType = getFileType(filePath);
        const conn = await getConnection(filePath, fileType);

        console.log('Procesando archivo:', filePath);

        // Obtener la conexión a nuestra base SQLite
        const db = await getDatabase();
        console.log('Conexión a SQLite establecida');

        // Limpiar la tabla actual
        await db.run('DELETE FROM importes_referencia');
        console.log('Tabla importes_referencia limpiada');

        // Procesar cada registro
        const query = `SELECT * FROM "Sueldos"`;
        const sueldos = await conn.query(query);
        console.log('Registros encontrados:', sueldos.length);

        for (const sueldo of sueldos) {
            const anioMes = sueldo.AnioMes.toString();
            const anio = parseInt(anioMes.substring(0, 4));
            const mes = parseInt(anioMes.substring(4, 6));
            const sueldoBase = parseFloat(sueldo.Sueldo);
            const adicional = parseFloat(sueldo.Adicional);

            // Calcular el aporte del 2.55%
            const aporte255 = sueldoBase * 0.0255;

            // Insertar en nuestra base de datos
            await db.run(
                'INSERT INTO importes_referencia (anio, mes, remuneracion, apExtraordinario) VALUES (?, ?, ?, ?)',
                [anio, mes, sueldoBase, adicional]
            );
        }

        console.log('Proceso completado exitosamente');

        await conn.close();
        return {
            message: 'Importación completada exitosamente',
            registrosProcesados: sueldos.length
        };
    } catch (error) {
        console.error('Error en processSueldosODB:', error);
        throw error;
    }
} 