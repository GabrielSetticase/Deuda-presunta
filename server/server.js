import express from 'express';
import multer from 'multer';
import cors from 'cors';
import { processODBFile, processSueldosODB } from './odbProcessor.js';
import path from 'path';
import fs from 'fs/promises';
import { fileURLToPath } from 'url';
import { initializeDatabase, getDatabase } from './database.js';
import http from 'http';
import { Server } from 'socket.io';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Asegurar que exista la carpeta uploads
const uploadsDir = path.join(__dirname, '../uploads');
try {
    await fs.access(uploadsDir);
} catch (error) {
    await fs.mkdir(uploadsDir, { recursive: true });
    console.log('Carpeta uploads creada');
}

const app = express();
const server = http.createServer(app);
const io = new Server(server, {
    cors: {
        origin: "http://localhost:5173",
        methods: ["GET", "POST"]
    }
});

// Middleware para manejar el estado de procesamiento
let isProcessing = false;
let currentStatus = '';
let currentCUIT = '';
let processedCount = 0;

io.on('connection', (socket) => {
    console.log('Cliente conectado');

    // Si hay un proceso en curso, enviar el estado actual
    if (isProcessing) {
        socket.emit('processingUpdate', {
            status: currentStatus,
            cuit: currentCUIT,
            count: processedCount,
            isProcessing: true
        });
    }

    socket.on('requestStatus', () => {
        if (isProcessing) {
            socket.emit('processingUpdate', {
                status: currentStatus,
                cuit: currentCUIT,
                count: processedCount,
                isProcessing: true
            });
        }
    });

    socket.on('disconnect', () => {
        console.log('Cliente desconectado');
    });
});

// Configuración de CORS más específica
app.use(cors({
    origin: ['http://localhost:5173', 'http://127.0.0.1:5173'],
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With'],
    credentials: true
}));

// Middleware para manejar preflight requests
app.options('*', cors());

app.use(express.json());

// Inicializar la base de datos
try {
    await initializeDatabase();
    console.log('Base de datos inicializada correctamente');
} catch (error) {
    console.error('Error al inicializar la base de datos:', error);
    process.exit(1);
}

// Configurar multer para guardar archivos
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, uploadsDir);
    },
    filename: function (req, file, cb) {
        cb(null, file.originalname);
    }
});

const upload = multer({ storage: storage });

// Middleware para logging
app.use((req, res, next) => {
    console.log(`${req.method} ${req.url}`);
    if (req.body && Object.keys(req.body).length > 0) {
        console.log('Body:', req.body);
    }
    next();
});

// Rutas para importes de referencia
app.get('/api/importes-referencia', async (req, res) => {
    try {
        const db = await getDatabase();
        const importes = await db.all('SELECT * FROM importes_referencia ORDER BY anio DESC, mes DESC');

        const importesConCalculos = importes.map(importe => ({
            ...importe,
            aporte255: (parseFloat(importe.remuneracion) * 0.0255).toFixed(2),
            aporteTotal: (parseFloat(importe.remuneracion) * 0.0255 + parseFloat(importe.apExtraordinario || 85)).toFixed(2)
        }));

        res.json(importesConCalculos);
    } catch (error) {
        console.error('Error al obtener importes:', error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/importes-referencia', async (req, res) => {
    try {
        console.log('\n--- INICIO POST /api/importes-referencia ---');
        console.log('Recibiendo datos:', req.body);
        let { anio, mes, remuneracion, apExtraordinario } = req.body;

        // Convertir a números y validar
        console.log('Valores originales:', { anio, mes, remuneracion, apExtraordinario });

        anio = parseInt(anio);
        mes = parseInt(mes);
        remuneracion = parseFloat(remuneracion);
        apExtraordinario = parseFloat(apExtraordinario || 85);

        console.log('Valores convertidos:', { anio, mes, remuneracion, apExtraordinario });

        // Validación más estricta
        if (isNaN(anio) || isNaN(mes) || isNaN(remuneracion)) {
            const error = 'Los valores deben ser números válidos';
            console.error('Error de validación:', error, { anio, mes, remuneracion, apExtraordinario });
            return res.status(400).json({ error });
        }

        if (anio < 2000 || anio > 2100 || mes < 1 || mes > 12 || remuneracion < 0) {
            const error = 'Valores fuera de rango permitido';
            console.error('Error de rango:', error, { anio, mes, remuneracion });
            return res.status(400).json({ error });
        }

        console.log('Obteniendo conexión a la base de datos...');
        const db = await getDatabase();
        console.log('Conexión obtenida, preparando consulta SQL');

        const sql = 'INSERT INTO importes_referencia (anio, mes, remuneracion, apExtraordinario) VALUES (?, ?, ?, ?)';
        const params = [anio, mes, remuneracion, apExtraordinario];
        console.log('SQL:', sql);
        console.log('Parámetros:', params);

        console.log('Ejecutando consulta...');
        const result = await db.run(sql, params);
        console.log('Resultado de la inserción:', result);

        if (!result || !result.lastID) {
            throw new Error('No se pudo insertar el registro');
        }

        console.log('Registro insertado con ID:', result.lastID);
        console.log('Recuperando registro insertado...');

        const newImporte = await db.get('SELECT * FROM importes_referencia WHERE id = ?', result.lastID);
        console.log('Nuevo importe recuperado:', newImporte);

        if (!newImporte) {
            throw new Error('No se pudo recuperar el importe recién creado');
        }

        const importeConCalculos = {
            ...newImporte,
            aporte255: (parseFloat(newImporte.remuneracion) * 0.0255).toFixed(2),
            aporteTotal: (parseFloat(newImporte.remuneracion) * 0.0255 + parseFloat(newImporte.apExtraordinario || 85)).toFixed(2)
        };

        console.log('Enviando respuesta exitosa:', importeConCalculos);
        console.log('--- FIN POST /api/importes-referencia ---\n');
        res.status(201).json(importeConCalculos);
    } catch (error) {
        console.error('\nError completo al crear importe:', error);
        console.error('Stack trace:', error.stack);
        console.error('--- FIN POST CON ERROR ---\n');
        res.status(500).json({
            error: 'Error guardando importe',
            message: error.message,
            stack: error.stack
        });
    }
});

app.put('/api/importes-referencia/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const { apExtraordinario } = req.body;
        const db = await getDatabase();

        await db.run(
            'UPDATE importes_referencia SET apExtraordinario = ? WHERE id = ?',
            [apExtraordinario, id]
        );

        const updatedImporte = await db.get('SELECT * FROM importes_referencia WHERE id = ?', id);
        res.json(updatedImporte);
    } catch (error) {
        console.error('Error al actualizar importe:', error);
        res.status(500).json({ error: error.message });
    }
});

app.delete('/api/importes-referencia/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const db = await getDatabase();
        await db.run('DELETE FROM importes_referencia WHERE id = ?', id);
        res.json({ message: 'Importe eliminado exitosamente' });
    } catch (error) {
        console.error('Error al eliminar importe:', error);
        res.status(500).json({ error: error.message });
    }
});

// Ruta para subir archivos
app.post('/upload', upload.single('file'), (req, res) => {
    try {
        res.json({ message: 'Archivo subido exitosamente' });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Ruta para procesar archivos
app.post('/process', async (req, res) => {
    try {
        const { actasFile, aportesFile, importes } = req.body;

        const actasPath = path.join(__dirname, '../uploads', actasFile);
        const aportesPath = path.join(__dirname, '../uploads', aportesFile);

        // Verificar que ambos archivos existen
        await fs.access(actasPath);
        await fs.access(aportesPath);

        const resultados = await processODBFile(aportesPath, importes, actasPath);
        res.json(resultados);
    } catch (error) {
        console.error('Error al procesar archivos:', error);
        res.status(500).json({ error: error.message });
    }
});

// Ruta para limpiar la base de datos
app.post('/clear-database', async (req, res) => {
    try {
        const uploadsDir = path.join(__dirname, '../uploads');
        const files = await fs.readdir(uploadsDir);

        for (const file of files) {
            if (file.endsWith('.mdb') || file.endsWith('.accdb') || file.endsWith('.odb')) {
                await fs.unlink(path.join(uploadsDir, file));
            }
        }

        res.json({ message: 'Base de datos limpiada exitosamente' });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Ruta para procesar archivo de sueldos
app.post('/api/importar-sueldos', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No se ha subido ningún archivo' });
        }

        // Verificar la extensión del archivo
        const fileExt = path.extname(req.file.originalname).toLowerCase();
        if (!['.mdb', '.accdb', '.odb'].includes(fileExt)) {
            return res.status(400).json({ error: 'Formato de archivo no soportado. Use .mdb, .accdb o .odb' });
        }

        console.log('Archivo recibido:', req.file);
        const resultado = await processSueldosODB(req.file.path);

        // Eliminar el archivo después de procesarlo
        await fs.unlink(req.file.path);

        res.json(resultado);
    } catch (error) {
        console.error('Error procesando archivo de sueldos:', error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/procesar', async (req, res) => {
    if (isProcessing) {
        return res.status(409).json({ error: 'Ya hay un proceso en ejecución' });
    }

    try {
        isProcessing = true;
        currentStatus = 'Iniciando procesamiento...';
        currentCUIT = '';
        processedCount = 0;
        const { actasPath, cuilesPath } = req.body;

        // Construir rutas absolutas
        const actasFullPath = path.join(process.cwd(), 'uploads', actasPath);
        const cuilesFullPath = path.join(process.cwd(), 'uploads', cuilesPath);

        // Verificar que los archivos existen
        try {
            await fs.access(actasFullPath);
            await fs.access(cuilesFullPath);
        } catch (error) {
            throw new Error(`No se pueden encontrar los archivos: ${error.message}`);
        }

        // Verificar las extensiones de los archivos
        const actasExt = path.extname(actasPath).toLowerCase();
        const cuilesExt = path.extname(cuilesPath).toLowerCase();

        if (!['.mdb', '.accdb', '.odb'].includes(actasExt) || !['.mdb', '.accdb', '.odb'].includes(cuilesExt)) {
            throw new Error('Formato de archivo no soportado. Use .mdb, .accdb o .odb');
        }

        // Función para enviar actualizaciones de progreso
        const updateProgress = (status, cuit = null, count = null) => {
            console.log('Enviando progreso:', { status, cuit, count });
            currentStatus = status;
            if (cuit) currentCUIT = cuit;
            if (count !== undefined) processedCount = count;

            io.emit('processingUpdate', {
                status: currentStatus,
                cuit: currentCUIT,
                count: processedCount,
                isProcessing: true
            });
        };

        updateProgress('Iniciando procesamiento de archivos...');
        console.log('Rutas de archivos:', { actasFullPath, cuilesFullPath });

        const resultados = await processODBFile(cuilesFullPath, [], actasFullPath, updateProgress);

        updateProgress('Procesamiento completado', null, null);

        setTimeout(() => {
            isProcessing = false;
            currentStatus = '';
            currentCUIT = '';
            processedCount = 0;
            io.emit('processingUpdate', {
                status: 'Proceso finalizado',
                isProcessing: false
            });
        }, 2000);

        res.json(resultados);
    } catch (error) {
        console.error('Error en /api/procesar:', error);
        isProcessing = false;
        currentStatus = '';
        currentCUIT = '';
        processedCount = 0;
        io.emit('processingUpdate', {
            status: `Error: ${error.message}`,
            isProcessing: false
        });
        res.status(500).json({ error: error.message });
    }
});

const PORT = 3001;
server.listen(PORT, () => {
    console.log(`Servidor corriendo en http://localhost:${PORT}`);
}); 