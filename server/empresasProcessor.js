import odbc from 'odbc';
import fs from 'fs/promises';

export async function obtenerDatosEmpresas(empresasPath, cuits) {
    try {
        // Verificar si el archivo existe y tiene una extensión válida
        console.log('Intentando acceder al archivo:', empresasPath);
        try {
            await fs.access(empresasPath);
        } catch (error) {
            throw new Error(`No se encontró el archivo de la base de datos en la ruta: ${empresasPath}`);
        }
        const fileExtension = empresasPath.toLowerCase().split('.').pop();
        console.log('Extensión del archivo:', fileExtension);
        if (fileExtension !== 'mdb' && fileExtension !== 'accdb') {
            throw new Error(`El archivo debe tener extensión .mdb o .accdb. Extensión actual: .${fileExtension}`);
        }

        console.log('Conectando a base de datos de empresas...');
        const connectionString = `Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=${empresasPath}`;
        const connection = await odbc.connect(connectionString);

        // Intentar diferentes nombres de tabla/vista posibles
        const posiblesTablas = [
            'vw_EmpresasInterior',
            'EmpresasInterior',
            'VW_EMPRESAS_INTERIOR',
            'EMPRESAS_INTERIOR',
            'Empresas',
            'EMPRESAS'
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
            throw new Error('No se encontró ninguna tabla válida de empresas en la base de datos');
        }

        // Crear la consulta con los CUITs proporcionados, tratándolos como texto
        if (!cuits || cuits.length === 0) {
            console.log('No se proporcionaron CUITs para la consulta');
            return new Map();
        }
        const cuitsString = cuits.map(cuit => String(cuit).replace(/\D/g, '')).join(',');
        const query = `
            SELECT CUIT, RAZONSOCIAL, CALLE, NUMERO, LOCALIDAD, ULTIMO_NRO_ACTA
            FROM ${tableName}
            WHERE CUIT IN (${cuitsString})
        `;
        console.log('Query SQL:', query);

        console.log('Ejecutando consulta de datos de empresas...');
        const rows = await connection.query(query);
        console.log('Resultados encontrados:', rows.length);
        if (rows.length > 0) {
            console.log('Primer resultado:', rows[0]);
        } else {
            console.log('No se encontraron resultados para los CUITs proporcionados');
        }
        await connection.close();

        // Crear un Map con los datos de las empresas y asegurar que los datos estén disponibles para el frontend
        const datosEmpresas = new Map();
        for (const row of rows) {
            if (row.CUIT) {
                const procesarValor = (valor) => {
                    if (valor === null || valor === undefined || valor === '') return 'No disponible';
                    return typeof valor === 'string' ? valor.trim() : String(valor).trim();
                };
                
                const datosEmpresa = {
                    razonSocial: procesarValor(row.RAZONSOCIAL),
                    calle: procesarValor(row.CALLE),
                    numero: procesarValor(row.NUMERO),
                    localidad: procesarValor(row.LOCALIDAD),
                    ultimoNroActa: procesarValor(row.ULTIMO_NRO_ACTA)
                };
                
                // Agregar los datos al Map
                datosEmpresas.set(row.CUIT, datosEmpresa);
                
                // Asegurar que los datos estén disponibles para el frontend
                if (typeof row.CUIT === 'number') {
                    datosEmpresas.set(String(row.CUIT), datosEmpresa);
                }
            }
        }

        return datosEmpresas;
    } catch (error) {
        console.error('Error obteniendo datos de empresas:', error);
        // Si el error es específico de la conexión ODBC, proporcionar un mensaje más descriptivo
        if (error.odbcErrors) {
            const odbcError = error.odbcErrors[0];
            throw new Error(`Error de conexión a la base de datos: ${odbcError.message}`);
        }
        throw error;
    }
}