import odbc from 'odbc';
import path from 'path';
import { getDatabase } from './database.js';
import XLSX from 'xlsx';
import { obtenerDatosEmpresas } from './empresasProcessor.js';

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

export async function processODBFile(cuilesPath, importes, actasPath, empresasPath, updateProgress) {
    if (typeof updateProgress !== 'function') {
        updateProgress = () => {};
    }
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

        // Definir los meses que vamos a procesar
        const meses = [
            'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
            'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'
        ];

        // Obtener todos los registros de una vez
        let query = `
            SELECT CUIT, CUIL, ANIO, ${meses.map(mes =>
            `REMUNERACION_${mes}, APORTE_${mes}, APORTE_ADIC_OS_${mes}`
        ).join(', ')}
                FROM ${tableName}
            WHERE ANIO > 0
            ORDER BY CUIT, CUIL, ANIO`;

        console.log('Ejecutando consulta:', query);
        const allRows = await connection.query(query);
        console.log(`Total de registros encontrados: ${allRows.length}`);

        // Procesar cada registro
        for (const row of allRows) {
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
            let diferenciaAnual = 0;

            // Procesar cada mes
            for (let index = 0; index < meses.length; index++) {
                const mes = meses[index];
                const fechaActual = new Date(anio, index, 1);

                // Solo procesar si la fecha es posterior al último período
                if (fechaActual < fechaInicio) {
                    continue;
                }

                const remuneracionColumn = `REMUNERACION_${mes}`;
                const aporteColumn = `APORTE_${mes}`;
                const aporteAdicColumn = `APORTE_ADIC_OS_${mes}`;

                // Verificar si existe remuneración para el mes
                const tieneRemuneracion = row[remuneracionColumn] !== undefined &&
                    row[remuneracionColumn] !== null &&
                    row[remuneracionColumn] !== '' &&
                    !isNaN(parseFloat(row[remuneracionColumn])) &&
                    parseFloat(row[remuneracionColumn]) > 0;

                // Solo procesar si hay remuneración y aporte informados (no vacíos)
                const tieneAporte = row[aporteColumn] !== undefined &&
                    row[aporteColumn] !== null &&
                    row[aporteColumn] !== '' &&
                    !isNaN(parseFloat(row[aporteColumn]));

                // Solo procesar si hay remuneración válida y mayor a cero
                if (tieneRemuneracion && tieneAporte) {
                    console.log(`Procesando ${mes}/${anio} - Remuneración encontrada: ${row[remuneracionColumn]}`);
                    const aporte = parseFloat(row[aporteColumn] || 0);
                    const aporteTotal = aporte; // Solo consideramos el aporte base, sin sumar APORTE_ADIC_OS
                    const remuneracion = parseFloat(row[remuneracionColumn]);

                    const importeRef = importes.find(imp =>
                        imp.anio === anio &&
                        imp.mes === (index + 1)
                    );

                    if (importeRef) {
                        const importeReferencia = parseFloat(importeRef.remuneracion) * 0.0255 + parseFloat(importeRef.apExtraordinario || 85);

                        console.log(`\nAnalizando - CUIT: ${cuit}, CUIL: ${cuil}, Período: ${mes}/${anio}`);
                        console.log(`Remuneración informada: ${remuneracion.toFixed(2)}`);
                        console.log(`Aporte Total: ${aporteTotal.toFixed(2)} (Aporte: ${aporte.toFixed(2)})`);
                        console.log(`Importe Referencia: ${importeReferencia.toFixed(2)} (Base: ${importeRef.remuneracion}, Extra: ${importeRef.apExtraordinario})`);

                        if (aporteTotal < importeReferencia) {
                            const diferencia = importeReferencia - aporteTotal;
                            diferenciaAnual += diferencia;
                            console.log(`Diferencia encontrada: ${diferencia.toFixed(2)}`);
                            console.log(`Diferencia acumulada en el año: ${diferenciaAnual.toFixed(2)}`);
                        } else {
                            console.log('No hay diferencia a cobrar - El aporte es mayor o igual al importe de referencia');
                        }
                    }
                } else {
                    console.log(`No hay remuneración informada para CUIL ${cuil} en ${mes}/${anio}, se omite el cálculo`);
                }
            }

            // Solo actualizamos el registro si hay diferencia positiva en el año
            if (diferenciaAnual > 0) {
                registro.diferenciasDetalladas.set(key, diferenciaAnual);
                registro.diferenciaTotal += diferenciaAnual;
                console.log(`\nDiferencia total para CUIL ${cuil} en ${anio}: ${diferenciaAnual.toFixed(2)}`);
                console.log(`Diferencia total acumulada para CUIT ${cuit}: ${registro.diferenciaTotal.toFixed(2)}\n`);
            }

            procesados++;
            if (procesados % 100 === 0) {
                updateProgress(`Procesando registros... ${procesados}`, cuit, procesados);
            }
        }

        await connection.close();

        // Limpiar y convertir resultados
        const resultadosFinales = Array.from(resultados.values())
            .map(registro => {
                // Filtrar las diferencias detalladas para mantener solo las positivas
                const diferenciasPositivas = Array.from(registro.diferenciasDetalladas.entries())
                    .filter(([_, diferencia]) => diferencia > 0)
                    .map(([key, diferencia]) => {
                        const [cuil, anio] = key.split('-');
                        return {
                            cuil,
                            anio: parseInt(anio),
                            diferencia
                        };
                    });

                // Recalcular la diferencia total
                const diferenciaTotal = diferenciasPositivas.reduce((sum, det) => sum + det.diferencia, 0);

                return {
                    cuit: registro.cuit,
                    primerPeriodoAVerificar: registro.ultimoPeriodo,
                    diferenciasDetalladas: diferenciasPositivas,
                    diferenciaTotal
                };
            })
            .filter(r => r.diferenciasDetalladas.length > 0 && r.diferenciaTotal > 0);

        // Obtener los datos de las empresas para los CUITs con diferencias
        const cuitsConDiferencias = resultadosFinales.map(r => r.cuit);
        const datosEmpresas = await obtenerDatosEmpresas(empresasPath, cuitsConDiferencias);

        // Integrar los datos de las empresas con los resultados
        const resultadosConEmpresas = resultadosFinales.map(resultado => {
            const datosEmpresa = datosEmpresas.get(resultado.cuit) || {
                razonSocial: 'No disponible',
                calle: 'No disponible',
                numero: 'No disponible',
                localidad: 'No disponible',
                ultimoNroActa: 'No disponible'
            };

            // Asegurar que todos los campos numéricos estén formateados correctamente
            const diferenciasFormateadas = resultado.diferenciasDetalladas.map(diff => ({
                ...diff,
                diferencia: parseFloat(diff.diferencia.toFixed(2))
            }));

            return {
                ...resultado,
                ...datosEmpresa,
                diferenciasDetalladas: diferenciasFormateadas,
                diferenciaTotal: parseFloat(resultado.diferenciaTotal.toFixed(2))
            };
        });

        // Verificar que hay resultados para enviar
        if (resultadosConEmpresas.length === 0) {
            console.log('No se encontraron diferencias para reportar');
            return [];
        }

        console.log('Resultados procesados:', resultadosConEmpresas);

        console.log('Procesamiento completado');
        updateProgress(`Procesamiento completado. Total: ${procesados} registros procesados`);
        return resultadosConEmpresas;

    } catch (error) {
        console.error('Error en processODBFile:', error);
        updateProgress(`Error: ${error.message}`);
        throw error;
    }
}

function getFileType(filePath) {
    const extension = path.extname(filePath).toLowerCase();
    switch (extension) {
        case '.mdb':
            return 'access';
        case '.accdb':
            return 'access';
        case '.odb':
            return 'openoffice';
        default:
            throw new Error(`Tipo de archivo no soportado: ${extension}`);
    }
}

import fs from 'fs';

async function getConnection(filePath, fileType) {
    // Verificar si el archivo existe
    if (!fs.existsSync(filePath)) {
        throw new Error(`No se encontró el archivo de la base de datos en la ruta: ${filePath}`);
    }

    if (fileType === 'access') {
        const connectionString = `Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=${filePath}`;
        try {
            return await odbc.connect(connectionString);
        } catch (error) {
            if (error.odbcErrors) {
                const odbcError = error.odbcErrors[0];
                throw new Error(`Error de conexión a la base de datos Access: ${odbcError.message}`);
            }
            throw error;
        }
    } else if (fileType === 'openoffice') {
        // Para bases de datos OpenOffice
        const connectionString = `Driver={LibreOffice Base Driver};DBQ=${filePath}`;
        try {
            return await odbc.connect(connectionString);
        } catch (error) {
            if (error.odbcErrors) {
                const odbcError = error.odbcErrors[0];
                throw new Error(`Error de conexión a la base de datos OpenOffice: ${odbcError.message}`);
            }
            throw error;
        }
    } else {
        throw new Error(`Tipo de conexión no soportado: ${fileType}`);
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