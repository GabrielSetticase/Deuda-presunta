import sqlite3 from 'sqlite3';
import { open } from 'sqlite';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs/promises';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

let db = null;

export async function initializeDatabase() {
    try {
        if (!db) {
            console.log('Inicializando base de datos...');

            const dbPath = path.join(__dirname, '../database.sqlite');
            console.log('Ruta absoluta de la base de datos:', path.resolve(dbPath));

            // Verificar que podemos escribir en el directorio
            const dbDir = path.dirname(dbPath);
            try {
                await fs.access(dbDir, fs.constants.W_OK);
                console.log('El directorio tiene permisos de escritura');
            } catch (error) {
                console.error('Error de permisos en el directorio:', error);
                throw new Error('No hay permisos de escritura en el directorio');
            }

            // Habilitar verbose logging para SQLite
            const verbose = sqlite3.verbose();
            console.log('SQLite verbose habilitado');

            try {
                db = await open({
                    filename: dbPath,
                    driver: sqlite3.Database,
                    mode: sqlite3.OPEN_READWRITE | sqlite3.OPEN_CREATE
                });
                console.log('Conexión a la base de datos establecida');
            } catch (error) {
                console.error('Error al abrir la base de datos:', error);
                throw error;
            }

            console.log('Base de datos conectada, verificando tabla...');

            // Verificar si la tabla existe
            const tableExists = await db.get("SELECT name FROM sqlite_master WHERE type='table' AND name='importes_referencia'");

            if (!tableExists) {
                console.log('La tabla no existe, creándola...');
                // Crear la tabla solo si no existe
                await db.exec(`
                    CREATE TABLE importes_referencia (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        anio INTEGER NOT NULL,
                        mes INTEGER NOT NULL,
                        remuneracion REAL NOT NULL,
                        apExtraordinario REAL DEFAULT 85.0
                    )
                `);
                console.log('Tabla creada exitosamente');
            } else {
                console.log('La tabla ya existe, no es necesario crearla');
            }

            // Verificar la estructura de la tabla
            const tableInfo = await db.all("PRAGMA table_info(importes_referencia)");
            console.log('Estructura de la tabla:', tableInfo);
        }
        return db;
    } catch (error) {
        console.error('Error al inicializar la base de datos:', error);
        throw new Error(`Error al inicializar la base de datos: ${error.message}`);
    }
}

export async function getDatabase() {
    try {
        if (!db) {
            await initializeDatabase();
        }
        return db;
    } catch (error) {
        console.error('Error al obtener la conexión de la base de datos:', error);
        throw new Error(`Error al obtener la conexión de la base de datos: ${error.message}`);
    }
} 