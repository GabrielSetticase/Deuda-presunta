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
    const [showAllPeriods, setShowAllPeriods] = useState(false);
    const [files, setFiles] = useState({ actas: null, cuiles: null, empresas: null });

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

            // Validar rango de año y mes
            const currentYear = new Date().getFullYear();
            if (anioNum < 2000 || anioNum > currentYear + 10) {
                setMessage(`Error: El año debe estar entre 2000 y ${currentYear + 10}`);
                return;
            }
            if (mesNum < 1 || mesNum > 12) {
                setMessage('Error: El mes debe estar entre 1 y 12');
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
        if (!files.actas || !files.cuiles || !files.empresas) {
            setMessage('Por favor seleccione todos los archivos (ACTAS, CUILES y EMPRESAS) antes de procesar');
            return;
        }

        try {
            setIsProcessing(true);
            setMessage('');
            setProcessingStatus('Preparando archivos...');
            setCurrentCUIT('');
            setProcessedCount(0);

            // Preparar FormData para todos los archivos
            const actasFormData = new FormData();
            actasFormData.append('file', files.actas);
            actasFormData.append('fileType', 'actas');

            const cuilesFormData = new FormData();
            cuilesFormData.append('file', files.cuiles);
            cuilesFormData.append('fileType', 'cuiles');

            const empresasFormData = new FormData();
            empresasFormData.append('file', files.empresas);
            empresasFormData.append('fileType', 'empresas');

            setProcessingStatus('Subiendo archivos...');

            // Subir todos los archivos
            await Promise.all([
                axios.post('http://localhost:3001/upload', actasFormData),
                axios.post('http://localhost:3001/upload', cuilesFormData),
                axios.post('http://localhost:3001/upload', empresasFormData)
            ]);

            setProcessingStatus('Iniciando procesamiento...');

            // Procesar los archivos
            const response = await axios.post('http://localhost:3001/api/procesar', {
                actasPath: files.actas.name,
                cuilesPath: files.cuiles.name,
                empresasPath: files.empresas.name
            });

            console.log('Resultados recibidos:', response.data);

            if (Array.isArray(response.data) && response.data.length > 0) {
                // Asegurarse de que los valores numéricos sean números
                const resultadosFormateados = response.data.map(r => ({
                    ...r,
                    diferenciaTotal: parseFloat(r.diferenciaTotal) || 0
                }));
                
                setResultados(resultadosFormateados);
                let totalDeuda = resultadosFormateados.reduce((sum, r) => sum + r.diferenciaTotal, 0);
                setMessage(`Procesamiento completado. Se encontraron diferencias en ${resultadosFormateados.length} CUITs. Deuda total: $${totalDeuda.toFixed(2)}`);
            } else {
                setResultados([]);
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
                            onClick={async () => {
                                try {
                                    setIsLoading(true);
                                    let empresasTotal = 0;
                                    let totalGeneral = 0;

                                    const data = resultados.map(r => {
                                        empresasTotal++;
                                        totalGeneral += parseFloat(r.diferenciaTotal);
                                        return {
                                            'CUIT': r.cuit,
                                            'Razón Social': r.razonSocial || 'No disponible',
                                            'Deuda Total': r.diferenciaTotal.toFixed(2)
                                        };
                                    });

                                    data.push({
                                        'CUIT': '',
                                        'Razón Social': 'TOTAL GENERAL',
                                        'Cantidad de Empresas': empresasTotal,
                                        'Deuda Total': totalGeneral.toFixed(2)
                                    });

                                    setMessage('Generando archivo Excel...');
                                    await new Promise(resolve => setTimeout(resolve, 0));

                                    const generarExcel = async () => {
                                        const wb = XLSX.utils.book_new();
                                        const ws = XLSX.utils.json_to_sheet(data);
                                        XLSX.utils.book_append_sheet(wb, ws, 'Deuda Presunta');
                                        const buffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
                                        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                                        const url = URL.createObjectURL(blob);
                                        const a = document.createElement('a');
                                        a.href = url;
                                        a.download = 'deuda_presunta_resultados.xlsx';
                                        document.body.appendChild(a);
                                        a.click();
                                        document.body.removeChild(a);
                                        URL.revokeObjectURL(url);
                                    };

                                    await generarExcel();
                                    setMessage('Archivo Excel generado exitosamente');
                                } catch (error) {
                                    console.error('Error al exportar a Excel:', error);
                                    setMessage('Error al generar el archivo Excel: ' + error.message);
                                } finally {
                                    setIsLoading(false);
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

                <div className="toggle-view-container" style={{ marginBottom: '1rem' }}>
                    <button
                        className="toggle-view-button"
                        onClick={() => setShowAllPeriods(!showAllPeriods)}
                    >
                        {showAllPeriods ? '▼ Mostrar últimos 3 meses' : '▶ Mostrar todos los períodos'}
                    </button>
                </div>

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
                        {importes
                            .sort((a, b) => {
                                if (a.anio !== b.anio) return b.anio - a.anio;
                                return b.mes - a.mes;
                            })
                            .slice(0, showAllPeriods ? undefined : 3)
                            .map((importe, index) => (
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

                    <div className="file-input-container">
                        <label className="file-input-label">
                            <input
                                type="file"
                                accept=".mdb,.accdb"
                                onChange={(e) => handleFileChange(e, 'empresas')}
                                className="file-input"
                            />
                            Seleccionar EMPRESAS
                        </label>
                        <span className="file-name">{files.empresas ? files.empresas.name : 'Ningún archivo seleccionado'}</span>
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

            <Dialog open={modalOpen} onClose={() => setModalOpen(false)}>
                <DialogTitle>Agregar Nuevo Importe</DialogTitle>
                <DialogContent>
                    <Grid container spacing={2} sx={{ mt: 1 }}>
                        <Grid item xs={6}>
                            <TextField
                                fullWidth
                                label="Año"
                                type="number"
                                value={nuevoImporte.anio}
                                onChange={(e) => setNuevoImporte({ ...nuevoImporte, anio: e.target.value })}
                                inputProps={{ min: "2000", max: "2035" }}
                                required
                            />
                        </Grid>
                        <Grid item xs={6}>
                            <TextField
                                fullWidth
                                label="Mes"
                                type="number"
                                value={nuevoImporte.mes}
                                onChange={(e) => setNuevoImporte({ ...nuevoImporte, mes: e.target.value })}
                                inputProps={{ min: "1", max: "12" }}
                                required
                            />
                        </Grid>
                        <Grid item xs={6}>
                            <TextField
                                fullWidth
                                label="Remuneración"
                                type="number"
                                value={nuevoImporte.remuneracion}
                                onChange={(e) => setNuevoImporte({ ...nuevoImporte, remuneracion: e.target.value })}
                                inputProps={{ min: "0", step: "0.01" }}
                                required
                            />
                        </Grid>
                        <Grid item xs={6}>
                            <TextField
                                fullWidth
                                label="Aporte Extraordinario"
                                type="number"
                                value={nuevoImporte.apExtraordinario}
                                onChange={(e) => setNuevoImporte({ ...nuevoImporte, apExtraordinario: e.target.value })}
                                inputProps={{ min: "0", step: "0.01" }}
                                required
                            />
                        </Grid>
                    </Grid>
                    {nuevoImporte.remuneracion && (
                        <Box sx={{ mt: 2 }}>
                            <Typography>Aporte 2.55%: ${calcularAporte255(nuevoImporte.remuneracion)}</Typography>
                            <Typography>Aporte Total: ${calcularAporteTotal(nuevoImporte.remuneracion, nuevoImporte.apExtraordinario)}</Typography>
                        </Box>
                    )}
                </DialogContent>
                <DialogActions>
                    <Button 
                        onClick={() => {
                            setModalOpen(false);
                            setNuevoImporte({
                                anio: '',
                                mes: '',
                                remuneracion: '',
                                apExtraordinario: '85'
                            });
                        }}
                    >
                        Cancelar
                    </Button>
                    <Button
                        onClick={handleSaveImporte}
                        disabled={!nuevoImporte.anio || !nuevoImporte.mes || !nuevoImporte.remuneracion}
                        variant="contained"
                        color="primary"
                    >
                        Guardar
                    </Button>
                </DialogActions>
            </Dialog>

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
                            onClick={async () => {
                                const resultadosOrdenados = [...resultados].sort((a, b) => {
                                    const localidadComparison = (a.localidad || '').localeCompare(b.localidad || '');
                                    if (localidadComparison === 0) {
                                        return b.diferenciaTotal - a.diferenciaTotal;
                                    }
                                    return localidadComparison;
                                });

                                const subtotalesPorLocalidad = resultadosOrdenados.reduce((acc, curr) => {
                                    const localidad = curr.localidad || 'No disponible';
                                    if (!acc[localidad]) {
                                        acc[localidad] = {
                                            total: 0,
                                            cantidad: 0
                                        };
                                    }
                                    acc[localidad].total += parseFloat(curr.diferenciaTotal) || 0;
                                    acc[localidad].cantidad += 1;
                                    return acc;
                                }, {});

                                let currentLocalidad = '';
                                const printContent = document.createElement('div');
                                printContent.innerHTML = `
                                    <h2>Resultados del Procesamiento de Deuda Presunta</h2>
                                    <table border="1" style="border-collapse: collapse; width: 100%;">
                                        <thead>
                                            <tr>
                                                <th>CUIT</th>
                                                <th>Razón Social</th>
                                                <th>Calle</th>
                                                <th>Número</th>
                                                <th>Localidad</th>
                                                <th>Último N° Acta</th>
                                                <th>Primer Período Verificado</th>
                                                <th>Deuda Total</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            ${resultadosOrdenados.map((resultado, index) => {
                                                const localidad = resultado.localidad || 'No disponible';
                                                let subtotalRow = '';
                                                if (localidad !== currentLocalidad) {
                                                    if (currentLocalidad !== '') {
                                                        subtotalRow = `<tr style="background-color: #f0f0f0;">
                                                            <td colspan="7"><strong>Subtotal ${currentLocalidad} (${subtotalesPorLocalidad[currentLocalidad].cantidad} empresas)</strong></td>
                                                            <td><strong>$${subtotalesPorLocalidad[currentLocalidad].total.toFixed(2)}</strong></td>
                                                        </tr>
                                                        <tr>
                                                            <th>CUIT</th>
                                                            <th>Razón Social</th>
                                                            <th>Calle</th>
                                                            <th>Número</th>
                                                            <th>Localidad</th>
                                                            <th>Último N° Acta</th>
                                                            <th>Primer Período Verificado</th>
                                                            <th>Deuda Total</th>
                                                        </tr>`;
                                                    }
                                                    currentLocalidad = localidad;
                                                }
                                                return `${subtotalRow}
                                                <tr>
                                                    <td>${resultado.cuit}</td>
                                                    <td>${resultado.razonSocial || 'No disponible'}</td>
                                                    <td>${resultado.calle || 'No disponible'}</td>
                                                    <td>${resultado.numero || 'No disponible'}</td>
                                                    <td>${resultado.localidad || 'No disponible'}</td>
                                                    <td>${resultado.ultimoNroActa || 'No disponible'}</td>
                                                    <td>${resultado.primerPeriodoAVerificar ? new Date(resultado.primerPeriodoAVerificar).toLocaleDateString('es-AR') : 'No disponible'}</td>
                                                    <td>$${resultado.diferenciaTotal.toFixed(2)}</td>
                                                </tr>`;
                                            }).join('')}
                                            <tr style="background-color: #f0f0f0;">
                                                <td colspan="7"><strong>Subtotal ${currentLocalidad} (${subtotalesPorLocalidad[currentLocalidad].cantidad} empresas)</strong></td>
                                                <td><strong>$${subtotalesPorLocalidad[currentLocalidad].total.toFixed(2)}</strong></td>
                                            </tr>
                                            <tr style="background-color: #e0e0e0;">
                                                <td colspan="7"><strong>TOTAL GENERAL (${Object.values(subtotalesPorLocalidad).reduce((a, b) => a + b.cantidad, 0)} empresas)</strong></td>
                                                <td><strong>$${Object.values(subtotalesPorLocalidad).reduce((a, b) => a + b.total, 0).toFixed(2)}</strong></td>
                                            </tr>
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
                            }}>
                                🖨️ Imprimir Resultados
                            </button>
                            <button className="action-button export-button" onClick={async () => {
                                try {
                                    setIsLoading(true);
                                    const resultadosOrdenados = [...resultados].sort((a, b) => {
                                        const localidadComparison = (a.localidad || '').localeCompare(b.localidad || '');
                                        return localidadComparison || b.diferenciaTotal - a.diferenciaTotal;
                                    });

                                    const subtotalesPorLocalidad = {};
                                    const data = [];
                                    let currentLocalidad = '';
                                    let totalGeneral = 0;
                                    let empresasTotal = 0;

                                    // Procesar los datos en lotes para evitar bloquear el navegador
                                    const procesarLote = (inicio, fin) => {
                                        for (let i = inicio; i < fin && i < resultadosOrdenados.length; i++) {
                                            const resultado = resultadosOrdenados[i];
                                            const localidad = resultado.localidad || 'No disponible';

                                            // Inicializar subtotales si es necesario
                                            if (!subtotalesPorLocalidad[localidad]) {
                                                subtotalesPorLocalidad[localidad] = {
                                                    total: 0,
                                                    cantidad: 0
                                                };
                                            }

                                            // Actualizar subtotales
                                            const deudaTotal = parseFloat(resultado.diferenciaTotal) || 0;
                                            subtotalesPorLocalidad[localidad].total += deudaTotal;
                                            subtotalesPorLocalidad[localidad].cantidad += 1;
                                            totalGeneral += deudaTotal;
                                            empresasTotal += 1;

                                            // Agregar encabezados si cambia la localidad
                                            if (localidad !== currentLocalidad) {
                                                if (currentLocalidad !== '') {
                                                    // Agregar subtotal de la localidad anterior
                                                    data.push({
                                                        'CUIT': '',
                                                        'Razón Social': `Subtotal ${currentLocalidad}`,
                                                        'Cantidad de Empresas': subtotalesPorLocalidad[currentLocalidad].cantidad,
                                                        'Deuda Total': subtotalesPorLocalidad[currentLocalidad].total.toFixed(2)
                                                    });
                                                    // Agregar fila en blanco y encabezados para la nueva localidad
                                                    data.push({});
                                                }
                                                data.push({
                                                    'CUIT': 'CUIT',
                                                    'Razón Social': 'Razón Social',
                                                    'Calle': 'Calle',
                                                    'Número': 'Número',
                                                    'Localidad': 'Localidad',
                                                    'Último N° Acta': 'Último N° Acta',
                                                    'Primer Período Verificado': 'Primer Período Verificado',
                                                    'Deuda Total': 'Deuda Total'
                                                });
                                                currentLocalidad = localidad;
                                            }

                                            // Agregar datos de la empresa
                                            data.push({
                                                'CUIT': resultado.cuit,
                                                'Razón Social': resultado.razonSocial || 'No disponible',
                                                'Calle': resultado.calle || 'No disponible',
                                                'Número': resultado.numero || 'No disponible',
                                                'Localidad': resultado.localidad || 'No disponible',
                                                'Último N° Acta': resultado.ultimoNroActa || 'No disponible',
                                                'Primer Período Verificado': resultado.primerPeriodoAVerificar ? 
                                                    new Date(resultado.primerPeriodoAVerificar).toLocaleDateString('es-AR') : 'No disponible',
                                                'Deuda Total': deudaTotal.toFixed(2)
                                            });
                                        }
                                    };

                                    // Procesar todos los datos en lotes de 100
                                    const tamañoLote = 100;
                                    for (let i = 0; i < resultadosOrdenados.length; i += tamañoLote) {
                                        procesarLote(i, i + tamañoLote);
                                    }

                                    // Agregar último subtotal y total general
                                    if (currentLocalidad) {
                                        data.push({
                                            'CUIT': '',
                                            'Razón Social': `Subtotal ${currentLocalidad}`,
                                            'Cantidad de Empresas': subtotalesPorLocalidad[currentLocalidad].cantidad,
                                            'Deuda Total': subtotalesPorLocalidad[currentLocalidad].total.toFixed(2)
                                        });
                                    }

                                    data.push({});
                                    data.push({
                                        'CUIT': '',
                                        'Razón Social': 'TOTAL GENERAL',
                                        'Cantidad de Empresas': empresasTotal,
                                        'Deuda Total': totalGeneral.toFixed(2)
                                    });

                                    // Crear y descargar el archivo Excel de manera asíncrona
                                    setMessage('Generando archivo Excel...');
                                    await new Promise(resolve => setTimeout(resolve, 0)); // Permitir que la UI se actualice

                                    // Dividir el proceso en tareas más pequeñas
                                    const generarExcel = async () => {
                                        const wb = XLSX.utils.book_new();
                                        const ws = XLSX.utils.json_to_sheet([], {skipHeader: true});
                                        
                                        // Agregar datos en lotes más pequeños
                                        const tamañoLoteExcel = 50;
                                        for (let i = 0; i < data.length; i += tamañoLoteExcel) {
                                            const lote = data.slice(i, i + tamañoLoteExcel);
                                            if (i === 0) {
                                                XLSX.utils.sheet_add_json(ws, lote, {skipHeader: true, origin: 0});
                                            } else {
                                                XLSX.utils.sheet_add_json(ws, lote, {skipHeader: true, origin: -1});
                                            }
                                            // Permitir que el navegador respire entre lotes
                                            await new Promise(resolve => setTimeout(resolve, 10));
                                        }

                                        XLSX.utils.book_append_sheet(wb, ws, 'Resultados');
                                        
                                        // Generar el archivo en un blob para mejor manejo de memoria
                                        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                                        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                                        
                                        // Descargar usando URL.createObjectURL
                                        const url = URL.createObjectURL(blob);
                                        const a = document.createElement('a');
                                        a.href = url;
                                        a.download = 'deuda_presunta_resultados.xlsx';
                                        document.body.appendChild(a);
                                        a.click();
                                        document.body.removeChild(a);
                                        URL.revokeObjectURL(url);
                                    };

                                    await generarExcel();
                                    setMessage('Archivo Excel generado exitosamente');
                                } catch (error) {
                                    console.error('Error al exportar a Excel:', error);
                                    setMessage('Error al generar el archivo Excel: ' + error.message);
                                } finally {
                                    setIsLoading(false);
                                }
                            }}>
                                📊 Exportar a Excel
                            </button>
                        </div>
                        <table>
                            <thead>
                                <tr>
                                    <th>CUIT</th>
                                    <th>Razón Social</th>
                                    <th>Calle</th>
                                    <th>Número</th>
                                    <th>Localidad</th>
                                    <th>Último N° Acta</th>
                                    <th>Primer Período Verificado</th>
                                    <th>Deuda Total</th>
                                    <th>Detalle</th>
                                </tr>
                            </thead>
                            <tbody>
                                {(() => {
                                    const resultadosOrdenados = [...resultados].sort((a, b) => {
                                        const localidadComparison = (a.localidad || '').localeCompare(b.localidad || '');
                                        if (localidadComparison === 0) {
                                            return b.diferenciaTotal - a.diferenciaTotal;
                                        }
                                        return localidadComparison;
                                    });

                                    const subtotalesPorLocalidad = resultadosOrdenados.reduce((acc, curr) => {
                                        const localidad = curr.localidad || 'No disponible';
                                        if (!acc[localidad]) {
                                            acc[localidad] = {
                                                total: 0,
                                                cantidad: 0
                                            };
                                        }
                                        acc[localidad].total += curr.diferenciaTotal;
                                        acc[localidad].cantidad += 1;
                                        return acc;
                                    }, {});

                                    let currentLocalidad = '';
                                    const rows = [];

                                    resultadosOrdenados.forEach((resultado, index) => {
                                        const localidad = resultado.localidad || 'No disponible';
                                        if (localidad !== currentLocalidad) {
                                            if (currentLocalidad !== '') {
                                                rows.push(
                                                    <tr key={`subtotal-${currentLocalidad}`} style={{ backgroundColor: '#f0f0f0' }}>
                                                        <td colSpan="7"><strong>Subtotal {currentLocalidad}</strong></td>
                                                        <td><strong>${subtotalesPorLocalidad[currentLocalidad].total.toFixed(2)}</strong></td>
                                                        <td></td>
                                                    </tr>
                                                );
                                            }
                                            currentLocalidad = localidad;
                                        }

                                        rows.push(
                                            <tr key={index}>
                                                <td>{resultado.cuit}</td>
                                                <td>{resultado.razonSocial || 'No disponible'}</td>
                                                <td>{resultado.calle || 'No disponible'}</td>
                                                <td>{resultado.numero || 'No disponible'}</td>
                                                <td>{resultado.localidad || 'No disponible'}</td>
                                                <td>{resultado.ultimoNroActa || 'No disponible'}</td>
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
                                        );
                                    });

                                    rows.push(
                                        <tr key={`subtotal-${currentLocalidad}`} style={{ backgroundColor: '#f0f0f0' }}>
                                            <td colSpan="7"><strong>Subtotal {currentLocalidad} ({subtotalesPorLocalidad[currentLocalidad].cantidad} empresas)</strong></td>
                                            <td><strong>${subtotalesPorLocalidad[currentLocalidad].total.toFixed(2)}</strong></td>
                                            <td></td>
                                        </tr>
                                    );

                                    rows.push(
                                        <tr key="total-general" style={{ backgroundColor: '#e0e0e0' }}>
                                            <td colSpan="7"><strong>TOTAL GENERAL (${Object.values(subtotalesPorLocalidad).reduce((a, b) => a + b.cantidad, 0)} empresas)</strong></td>
                                            <td><strong>${Object.values(subtotalesPorLocalidad).reduce((a, b) => a + b.total, 0).toFixed(2)}</strong></td>
                                            <td></td>
                                        </tr>
                                    );

                                    return rows;
                                })()} 
                            </tbody>
                        </table>
                    </div>
            )}
        </div>
    );
}

export default App;