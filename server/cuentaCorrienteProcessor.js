import odbc from 'odbc';
import fs from 'fs/promises';

export async function obtenerDatosCuentaCorriente(dbPath, cuits) {
    try {
        // Verificar si el archivo existe y tiene una extensión válida
        console.log('Intentando acceder al archivo:', dbPath);
        try {
            await fs.access(dbPath);
        } catch (error) {
            throw new Error(`No se encontró el archivo de la base de datos en la ruta: ${dbPath}`);
        }
        const fileExtension = dbPath.toLowerCase().split('.').pop();
        console.log('Extensión del archivo:', fileExtension);
        if (fileExtension !== 'mdb' && fileExtension !== 'accdb') {
            throw new Error(`El archivo debe tener extensión .mdb o .accdb. Extensión actual: .${fileExtension}`);
        }

        console.log('Conectando a base de datos de cuenta corriente...');
        const connectionString = `Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=${dbPath}`;
        const connection = await odbc.connect(connectionString);

        // Intentar diferentes nombres de tabla/vista posibles
        const posiblesTablas = [
            'VW_CTACTE_EMPRESAS_PROCESO',
            'CTACTE_EMPRESAS_PROCESO',
            'vw_CtaCteEmpresasProceso',
            'CtaCteEmpresasProceso'
        ];

        let tableName = null;
        for (const tabla of posiblesTablas) {
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
            throw new Error('No se encontró ninguna tabla válida de cuenta corriente en la base de datos');
        }

        // Crear la consulta con los CUITs proporcionados
        if (!cuits || cuits.length === 0) {
            console.log('No se proporcionaron CUITs para la consulta');
            return new Map();
        }
        const cuitsString = cuits.map(cuit => String(cuit).replace(/\D/g, '')).join(',');
        const query = `
            SELECT ACTIVIDAD, ANIO, 
                   APORTE_381_1, APORTE_381_2, APORTE_381_3, APORTE_381_4, 
                   APORTE_381_5, APORTE_381_6, APORTE_381_7, APORTE_381_8, 
                   APORTE_381_9, APORTE_381_10, APORTE_381_11, APORTE_381_12
            FROM ${tableName}
            WHERE CUIT IN (${cuitsString})
        `;
        console.log('Query SQL:', query);

        console.log('Ejecutando consulta de datos de cuenta corriente...');
        const rows = await connection.query(query);
        console.log('Resultados encontrados:', rows.length);
        if (rows.length > 0) {
            console.log('Primer resultado:', rows[0]);
        } else {
            console.log('No se encontraron resultados para los CUITs proporcionados');
        }
        await connection.close();

        // Crear un Map con los datos de cuenta corriente
        const datosCuentaCorriente = new Map();
        for (const row of rows) {
            if (row.CUIT) {
                const procesarValor = (valor) => {
                    if (valor === null || valor === undefined || valor === '') return '0';
                    return typeof valor === 'string' ? valor.trim() : String(valor).trim();
                };
                
                const datosCuenta = {
                    actividad: procesarValor(row.ACTIVIDAD),
                    anio: procesarValor(row.ANIO),
                    aportes: {
                        mes1: procesarValor(row.APORTE_381_1),
                        mes2: procesarValor(row.APORTE_381_2),
                        mes3: procesarValor(row.APORTE_381_3),
                        mes4: procesarValor(row.APORTE_381_4),
                        mes5: procesarValor(row.APORTE_381_5),
                        mes6: procesarValor(row.APORTE_381_6),
                        mes7: procesarValor(row.APORTE_381_7),
                        mes8: procesarValor(row.APORTE_381_8),
                        mes9: procesarValor(row.APORTE_381_9),
                        mes10: procesarValor(row.APORTE_381_10),
                        mes11: procesarValor(row.APORTE_381_11),
                        mes12: procesarValor(row.APORTE_381_12)
                    }
                };
                
                datosCuentaCorriente.set(row.CUIT, datosCuenta);
                
                if (typeof row.CUIT === 'number') {
                    datosCuentaCorriente.set(String(row.CUIT), datosCuenta);
                }
            }
        }

        return datosCuentaCorriente;
    } catch (error) {
        console.error('Error al obtener datos de cuenta corriente:', error);
        throw error;
    }
}