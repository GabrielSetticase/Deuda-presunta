import { useState, useEffect } from 'react';
import {
    Container,
    Box,
    Typography,
    Button,
    Paper,
    CircularProgress,
    Table,
    TableBody,
    TableCell,
    TableContainer,
    TableHead,
    TableRow,
    TextField,
    Dialog,
    DialogTitle,
    DialogContent,
    DialogActions,
    Grid,
    IconButton,
    Alert,
    Backdrop
} from '@mui/material';
import axios from 'axios';
import './styles.css';
import io from 'socket.io-client';
import * as XLSX from 'xlsx';

function App() {
    const [importes, setImportes] = useState([]);
    const [selectedFile, setSelectedFile] = useState(null);
    const [selectedActasFile, setSelectedActasFile] = useState(null);
    const [isLoading, setIsLoading] = useState(false);
    const [message, setMessage] = useState('');
    const [showModal, setShowModal] = useState(false);
    const [nuevoImporte, setNuevoImporte] = useState({
        anio: '',
        mes: '',
        remuneracion: '',
        apExtraordinario: '85'
    });
    const [modalOpen, setModalOpen] = useState(false);
    const [isProcessing, setIsProcessing] = useState(false);
    const [resultados, setResultados] = useState([]);
    const [selectedActas, setSelectedActas] = useState('');
    const [selectedCuiles, setSelectedCuiles] = useState('');
    const [processingStatus, setProcessingStatus] = useState('');
    const [currentCUIT, setCurrentCUIT] = useState('');
    const [processedCount, setProcessedCount] = useState(0);
    const [socket, setSocket] = useState(null);
    const [cuitEspecifico, setCuitEspecifico] = useState('');
    const [files, setFiles] = useState({ actas: null, cuiles: null });

    useEffect(() => {
        cargarImportes();

        // Crear conexión websocket con mejor manejo de reconexión
        const newSocket = io('http://localhost:3001', {
            reconnection: true,
            reconnectionDelay: 1000,
            reconnectionDelayMax: 5000,
            reconnectionAttempts: 5,
            timeout: 20000,
            autoConnect: true
        });

        newSocket.on('connect', () => {
            console.log('Conectado al servidor websocket');
            // Restaurar estado si es necesario
            if (isProcessing) {
                newSocket.emit('requestStatus');
            }
        });

        newSocket.on('disconnect', () => {
            console.log('Desconectado del servidor websocket');
            // No reseteamos el estado de procesamiento aquí
            setProcessingStatus('Reconectando al servidor...');
        });

        newSocket.on('connect_error', (error) => {
            console.error('Error de conexión:', error);
            setProcessingStatus('Error de conexión. Intentando reconectar...');
        });

        newSocket.on('processingUpdate', (data) => {
            console.log('Actualización recibida:', data);

            // Actualizar estado de procesamiento
            setIsProcessing(data.isProcessing);

            // Actualizar mensaje de estado
            setProcessingStatus(data.status);

            // Actualizar CUIT actual si existe
            if (data.cuit) {
                setCurrentCUIT(data.cuit);
            }

            // Actualizar contador si existe
            if (data.count !== undefined) {
                setProcessedCount(data.count);
            }

            // Si el proceso ha finalizado, limpiar estados después de un delay
            if (!data.isProcessing) {
                setTimeout(() => {
                    setCurrentCUIT('');
                    setProcessedCount(0);
                    setProcessingStatus('');
                }, 2000);
            }
        });

        setSocket(newSocket);

        return () => {
            if (newSocket) {
                newSocket.disconnect();
            }
        };
    }, []);

    const cargarImportes = async () => {
        try {
            const response = await axios.get('http://localhost:3001/api/importes-referencia');
            const importesConValorPorDefecto = response.data.map(importe => ({
                ...importe,
                apExtraordinario: importe.apExtraordinario || '85'
            }));
            setImportes(importesConValorPorDefecto);
        } catch (error) {
            console.error('Error cargando importes:', error);
            setMessage('Error al cargar importes: ' + (error.response?.data?.error || error.message));
        }
    };

    const calcularAporte255 = (remuneracion) => {
        return (parseFloat(remuneracion) * 0.0255).toFixed(2);
    };

    const calcularAporteTotal = (remuneracion, apExtraordinario) => {
        const aporte255 = parseFloat(calcularAporte255(remuneracion));
        const extraordinario = parseFloat(apExtraordinario);
        return (aporte255 + extraordinario).toFixed(2);
    };

    const handleSaveImporte = async () => {
        try {
            setMessage('Guardando importe...');

            // Convertir y validar los valores
            const anioNum = parseInt(nuevoImporte.anio);
            const mesNum = parseInt(nuevoImporte.mes);
            const remuneracionNum = parseFloat(nuevoImporte.remuneracion);
            const apExtraordinarioNum = parseFloat(nuevoImporte.apExtraordinario || 85);

            // Validar que los valores sean números válidos
            if (isNaN(anioNum) || isNaN(mesNum) || isNaN(remuneracionNum)) {
                setMessage('Error: Los valores deben ser números válidos');
                return;
            }

            const importeData = {
                anio: anioNum,
                mes: mesNum,
                remuneracion: remuneracionNum,
                apExtraordinario: apExtraordinarioNum
            };

            console.log('Enviando datos:', importeData);

            const response = await axios.post(
                'http://localhost:3001/api/importes-referencia',
                importeData,
                {
                    headers: {
                        'Content-Type': 'application/json',
                        'Accept': 'application/json'
                    },
                    timeout: 5000, // 5 segundos de timeout
                    validateStatus: function (status) {
                        return status >= 200 && status < 300;
                    }
                }
            );

            console.log('Respuesta del servidor:', response.data);

            if (response.data) {
                await cargarImportes();
                setShowModal(false);
                setNuevoImporte({
                    anio: '',
                    mes: '',
                    remuneracion: '',
                    apExtraordinario: '85'
                });
                setMessage('¡Importe agregado exitosamente!');
            }
        } catch (error) {
            console.error('Error completo:', error);
            let errorMessage = 'Error al guardar el importe';

            if (error.response) {
                // El servidor respondió con un código de error
                errorMessage += ': ' + (error.response.data?.error || error.response.data?.message || error.message);
                console.error('Error response:', error.response.data);
            } else if (error.request) {
                // La petición fue hecha pero no se recibió respuesta
                errorMessage += ': No se pudo conectar con el servidor';
                console.error('Error request:', error.request);
            } else {
                // Error al configurar la petición
                errorMessage += ': ' + error.message;
            }

            setMessage(errorMessage);
        }
    };

    const handleUpdateApExtraordinario = async (id, newValue) => {
        try {
            const apExtraordinarioNum = parseFloat(newValue);
            if (isNaN(apExtraordinarioNum)) {
                setMessage('Error: El valor debe ser un número válido');
                return;
            }

            const importe = importes.find(imp => imp.id === id);
            const updatedImporte = {
                ...importe,
                apExtraordinario: apExtraordinarioNum
            };

            await axios.put(
                `http://localhost:3001/api/importes-referencia/${id}`,
                updatedImporte,
                {
                    headers: {
                        'Content-Type': 'application/json'
                    }
                }
            );
            await cargarImportes();
        } catch (error) {
            console.error('Error al actualizar:', error);
            setMessage('Error al actualizar el aporte extraordinario: ' + error.message);
        }
    };

    const handleActasFileSelect = (event) => {
        const file = event.target.files[0];
        if (file) {
            setSelectedActasFile(file);
            setMessage('');
        }
    };

    const handleFileSelect = (event) => {
        const file = event.target.files[0];
        if (file) {
            setSelectedFile(file);
            setMessage('');
        }
    };

    const handleUpload = async () => {
        if (!selectedFile || !selectedActasFile) {
            setMessage('Por favor seleccione ambos archivos (Actas y Aportes)');
            return;
        }

        setIsLoading(true);
        setMessage('Procesando archivos...');

        try {
            // Primero subimos el archivo de actas
            const actasFormData = new FormData();
            actasFormData.append('file', selectedActasFile);
            await axios.post('http://localhost:3001/upload', actasFormData);

            // Luego subimos el archivo de aportes
            const aportesFormData = new FormData();
            aportesFormData.append('file', selectedFile);
            await axios.post('http://localhost:3001/upload', aportesFormData);

            // Procesar los archivos
            const processResponse = await axios.post('http://localhost:3001/process', {
                actasFile: selectedActasFile.name,
                aportesFile: selectedFile.name,
                importes
            });

            setMessage('Archivos procesados exitosamente');
            console.log('Resultados:', processResponse.data);
        } catch (error) {
            setMessage(`Error: ${error.response?.data?.error || error.message}`);
        } finally {
            setIsLoading(false);
        }
    };

    const handleClearDatabase = async () => {
        try {
            await axios.post('http://localhost:3001/clear-database');
            setMessage('Base de datos limpiada exitosamente');
            setSelectedFile(null);
            setSelectedActasFile(null);
        } catch (error) {
            setMessage(`Error al limpiar la base de datos: ${error.message}`);
        }
    };

    const handleDeleteImporte = async (id) => {
        try {
            await axios.delete(`http://localhost:3001/api/importes-referencia/${id}`);
            await cargarImportes();
            setMessage('Importe eliminado exitosamente');
        } catch (error) {
            setMessage('Error al eliminar el importe: ' + error.message);
        }
    };

    const handleProcesar = async () => {
        if (!files.actas || !files.cuiles) {
            setMessage('Por favor seleccione ambos archivos antes de procesar');
            return;
        }

        try {
            setIsProcessing(true);
            setMessage('');
            setProcessingStatus('Preparando archivos...');
            setCurrentCUIT('');
            setProcessedCount(0);

            // Primero, subir los archivos
            const actasFormData = new FormData();
            actasFormData.append('file', files.actas);
            actasFormData.append('fileType', 'actas');

            const cuilesFormData = new FormData();
            cuilesFormData.append('file', files.cuiles);
            cuilesFormData.append('fileType', 'cuiles');

            setProcessingStatus('Subiendo archivos...');

            // Subir ambos archivos
            await Promise.all([
                axios.post('http://localhost:3001/upload', actasFormData),
                axios.post('http://localhost:3001/upload', cuilesFormData)
            ]);

            setProcessingStatus('Iniciando procesamiento...');

            // Procesar los archivos
            const response = await axios.post('http://localhost:3001/api/procesar', {
                actasPath: files.actas.name,
                cuilesPath: files.cuiles.name
            });

            console.log('Resultados recibidos:', response.data);

            if (Array.isArray(response.data) && response.data.length > 0) {
                setResultados(response.data);
                let totalDeuda = response.data.reduce((sum, r) => sum + r.diferenciaTotal, 0);
                setMessage(`Procesamiento completado. Se encontraron diferencias en ${response.data.length} CUITs. Deuda total: $${totalDeuda.toFixed(2)}`);
            } else {
                setMessage('El procesamiento finalizó pero no se encontraron diferencias.');
            }

        } catch (error) {
            console.error('Error:', error);
            setMessage('Error al procesar archivos: ' + (error.response?.data?.error || error.message));
        } finally {
            setIsProcessing(false);
            setProcessingStatus('');
        }
    };

    const handleFileChange = (event, type) => {
        const file = event.target.files[0];
        if (file) {
            setFiles({ ...files, [type]: file });
            setMessage('');
        }
    };

    return (
        <div className="container">
            <h1>Sistema de cálculo de deuda presunta</h1>

            <div className="section">
                <h2>Importes de Referencia</h2>
                <div className="actions-container">
                    <button
                        className="add-button"
                        onClick={() => setModalOpen(true)}
                    >
                        + AGREGAR MANUAL
                    </button>
                    <div className="import-section">
                        <input
                            type="file"
                            id="file-upload"
                            accept=".odb,.mdb,.accdb,.xlsx,.xls"
                            onChange={async (e) => {
                                const file = e.target.files[0];
                                if (file) {
                                    setMessage('Importando archivo de sueldos...');
                                    const formData = new FormData();
                                    formData.append('file', file);
                                    try {
                                        const response = await axios.post('http://localhost:3001/api/importar-sueldos', formData);
                                        setMessage(`Importación exitosa: ${response.data.registrosProcesados} registros procesados`);
                                        await cargarImportes();
                                    } catch (error) {
                                        console.error('Error importando sueldos:', error);
                                        setMessage('Error al importar archivo: ' + (error.response?.data?.error || error.message));
                                    }
                                }
                            }}
                            style={{ display: 'none' }}
                        />
                        <button
                            className="import-button"
                            onClick={() => document.getElementById('file-upload').click()}
                        >
                            📁 IMPORTAR DESDE ARCHIVO
                        </button>
                        <span className="file-types">(.odb, .mdb, .accdb, .xlsx, .xls)</span>
                    </div>
                </div>

                {message && (
                    <div className={`message ${message.includes('Error') ? 'error' : 'success'}`}>
                        {message}
                    </div>
                )}

                <table>
                    <thead>
                        <tr>
                            <th>Año</th>
                            <th>Mes</th>
                            <th>Remuneración</th>
                            <th>Aporte 2.55%</th>
                            <th>Ap Extraordinario</th>
                            <th>Aporte Total</th>
                            <th>Acciones</th>
                        </tr>
                    </thead>
                    <tbody>
                        {importes.map((importe, index) => (
                            <tr key={importe.id || index}>
                                <td>{importe.anio}</td>
                                <td>{importe.mes}</td>
                                <td>${importe.remuneracion}</td>
                                <td>${importe.aporte255}</td>
                                <td>
                                    <input
                                        type="number"
                                        className="ap-extraordinario-input"
                                        value={importe.apExtraordinario}
                                        onChange={(e) => handleUpdateApExtraordinario(importe.id, e.target.value)}
                                    />
                                </td>
                                <td>${importe.aporteTotal}</td>
                                <td>
                                    <button onClick={() => handleDeleteImporte(importe.id)}>×</button>
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>

            <div className="section">
                <h2>Procesar Base de Datos</h2>
                <div className="file-upload-section">
                    <div className="file-input-container">
                        <label className="file-input-label">
                            <input
                                type="file"
                                accept=".mdb,.accdb"
                                onChange={(e) => handleFileChange(e, 'actas')}
                                className="file-input"
                            />
                            Seleccionar ACTAS
                        </label>
                        <span className="file-name">{files.actas ? files.actas.name : 'Ningún archivo seleccionado'}</span>
                    </div>

                    <div className="file-input-container">
                        <label className="file-input-label">
                            <input
                                type="file"
                                accept=".mdb,.accdb"
                                onChange={(e) => handleFileChange(e, 'cuiles')}
                                className="file-input"
                            />
                            Seleccionar CUILES
                        </label>
                        <span className="file-name">{files.cuiles ? files.cuiles.name : 'Ningún archivo seleccionado'}</span>
                    </div>

                    <button
                        className="process-button"
                        onClick={handleProcesar}
                        disabled={!files.actas || !files.cuiles}
                    >
                        {isProcessing ? 'PROCESANDO...' : 'PROCESAR ARCHIVOS'}
                    </button>

                    <button className="clear-button" onClick={handleClearDatabase}>
                        LIMPIAR BASE DE DATOS
                    </button>
                </div>
            </div>

            {showModal && (
                <div className="modal">
                    <div className="modal-content">
                        <h2>Agregar Nuevo Importe</h2>
                        <div className="modal-input-group">
                            <div className="input-wrapper">
                                <label>Año:</label>
                                <input
                                    type="number"
                                    value={nuevoImporte.anio}
                                    onChange={(e) => setNuevoImporte({ ...nuevoImporte, anio: e.target.value })}
                                    min="2000"
                                    max="2100"
                                    required
                                />
                            </div>
                            <div className="input-wrapper">
                                <label>Mes:</label>
                                <input
                                    type="number"
                                    value={nuevoImporte.mes}
                                    onChange={(e) => setNuevoImporte({ ...nuevoImporte, mes: e.target.value })}
                                    min="1"
                                    max="12"
                                    required
                                />
                            </div>
                            <div className="input-wrapper">
                                <label>Remuneración:</label>
                                <input
                                    type="number"
                                    value={nuevoImporte.remuneracion}
                                    onChange={(e) => setNuevoImporte({ ...nuevoImporte, remuneracion: e.target.value })}
                                    min="0"
                                    step="0.01"
                                    required
                                />
                            </div>
                            <div className="input-wrapper">
                                <label>Aporte Extraordinario:</label>
                                <input
                                    type="number"
                                    value={nuevoImporte.apExtraordinario}
                                    onChange={(e) => setNuevoImporte({ ...nuevoImporte, apExtraordinario: e.target.value })}
                                    min="0"
                                    step="0.01"
                                    required
                                />
                            </div>
                        </div>
                        <div className="modal-preview">
                            {nuevoImporte.remuneracion && (
                                <>
                                    <p>Aporte 2.55%: ${calcularAporte255(nuevoImporte.remuneracion)}</p>
                                    <p>Aporte Total: ${calcularAporteTotal(nuevoImporte.remuneracion, nuevoImporte.apExtraordinario)}</p>
                                </>
                            )}
                        </div>
                        <div className="modal-buttons">
                            <button onClick={() => {
                                setShowModal(false);
                                setNuevoImporte({
                                    anio: '',
                                    mes: '',
                                    remuneracion: '',
                                    apExtraordinario: '85'
                                });
                            }}>Cancelar</button>
                            <button
                                onClick={handleSaveImporte}
                                disabled={!nuevoImporte.anio || !nuevoImporte.mes || !nuevoImporte.remuneracion}
                            >
                                Guardar
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {isProcessing && (
                <div className="processing-status">
                    <div className="spinner"></div>
                    <div className="processing-details">
                        <p className="status-message">{processingStatus || 'Procesando...'}</p>
                        {currentCUIT && (
                            <p className="progress-detail">
                                Procesando CUIT: {currentCUIT}
                            </p>
                        )}
                        {processedCount > 0 && (
                            <p className="progress-detail">
                                Registros procesados: {processedCount}
                            </p>
                        )}
                    </div>
                </div>
            )}

            {resultados && resultados.length > 0 && (
                <div className="resultados">
                    <h2>Resultados del Procesamiento</h2>
                    <div className="resultados-actions">
                        <button
                            className="action-button print-button"
                            onClick={() => {
                                const printContent = document.createElement('div');
                                printContent.innerHTML = `
                                    <h2>Resultados del Procesamiento de Deuda Presunta</h2>
                                    <table border="1" style="border-collapse: collapse; width: 100%;">
                                        <thead>
                                            <tr>
                                                <th>CUIT</th>
                                                <th>Último Período</th>
                                                <th>Deuda Total</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            ${resultados.map(resultado => `
                                                <tr>
                                                    <td>${resultado.cuit}</td>
                                                    <td>${resultado.primerPeriodoAVerificar ? new Date(resultado.primerPeriodoAVerificar).toLocaleDateString('es-AR') : 'No disponible'}</td>
                                                    <td>$${resultado.diferenciaTotal.toFixed(2)}</td>
                                                </tr>
                                            `).join('')}
                                        </tbody>
                                    </table>
                                `;
                                const printWindow = window.open('', '_blank');
                                printWindow.document.write('<html><head><title>Deuda Presunta - Resultados</title>');
                                printWindow.document.write('<style>body { font-family: Arial, sans-serif; } table { width: 100%; border-collapse: collapse; } th, td { padding: 8px; border: 1px solid #ddd; text-align: left; }</style>');
                                printWindow.document.write('</head><body>');
                                printWindow.document.write(printContent.innerHTML);
                                printWindow.document.write('</body></html>');
                                printWindow.document.close();
                                printWindow.print();
                            }}
                        >
                            🖨️ Imprimir Resultados
                        </button>
                        <button
                            className="action-button export-button"
                            onClick={() => {
                                const wb = XLSX.utils.book_new();
                                const data = resultados.map(resultado => ({
                                    CUIT: resultado.cuit,
                                    'Último Período': resultado.primerPeriodoAVerificar ? new Date(resultado.primerPeriodoAVerificar).toLocaleDateString('es-AR') : 'No disponible',
                                    'Deuda Total': resultado.diferenciaTotal.toFixed(2)
                                }));
                                const ws = XLSX.utils.json_to_sheet(data);
                                XLSX.utils.book_append_sheet(wb, ws, 'Resultados');
                                XLSX.writeFile(wb, 'deuda_presunta_resultados.xlsx');
                            }}
                        >
                            📊 Exportar a Excel
                        </button>
                    </div>
                    <table>
                        <thead>
                            <tr>
                                <th>CUIT</th>
                                <th>Último Período</th>
                                <th>Deuda Total</th>
                                <th>Detalle</th>
                            </tr>
                        </thead>
                        <tbody>
                            {resultados.map((resultado, index) => (
                                <tr key={index}>
                                    <td>{resultado.cuit}</td>
                                    <td>{resultado.primerPeriodoAVerificar ? new Date(resultado.primerPeriodoAVerificar).toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' }) : 'No disponible'}</td>
                                    <td>${resultado.diferenciaTotal.toFixed(2)}</td>
                                    <td>
                                        <details>
                                            <summary>Ver detalle</summary>
                                            <table className="detalle-table">
                                                <thead>
                                                    <tr>
                                                        <th>CUIL</th>
                                                        <th>Año</th>
                                                        <th>Diferencia</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {resultado.diferenciasDetalladas.map((detalle, idx) => (
                                                        <tr key={idx}>
                                                            <td>{detalle.cuil}</td>
                                                            <td>{detalle.anio}</td>
                                                            <td>${detalle.diferencia.toFixed(2)}</td>
                                                        </tr>
                                                    ))}
                                                </tbody>
                                            </table>
                                        </details>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            )}
        </div>
    );
}

export default App;