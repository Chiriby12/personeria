const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');
const fs = require('fs');

let mainWindow;

// Ruta del archivo Excel en el escritorio
const excelPath = path.join(app.getPath('desktop'), 'solicitudes.xlsx');

// CONTRASEÑA PARA EDITAR (cámbiala por la que desees)
const PASSWORD_EDICION = 'admin123';

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1200,
        height: 900,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolation: true,
            nodeIntegration: false
        }
    });
    mainWindow.loadFile('index.html');
}

// Inicializar archivo Excel si no existe
async function initializeExcel() {
    if (!fs.existsSync(excelPath)) {
        const workbook = new ExcelJS.Workbook();
        
        // HOJA 1: Solicitudes (datos principales)
        const worksheet = workbook.addWorksheet('Solicitudes');
        
        // Definir columnas con anchos apropiados
        worksheet.columns = [
            { header: 'Fecha y Hora', key: 'fecha', width: 20 },
            { header: 'Nombre', key: 'nombre', width: 20 },
            { header: 'Apellido', key: 'apellido', width: 20 },
            { header: 'Tipo Documento', key: 'tipoDocumento', width: 15 },
            { header: 'Número Documento', key: 'numeroDocumento', width: 18 },
            { header: 'Tipo Necesidad', key: 'tipoNecesidad', width: 25 },
            { header: 'Descripción', key: 'descripcion', width: 50 }
        ];

        // Estilo del encabezado
        worksheet.getRow(1).font = { bold: true, size: 12 };
        worksheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4472C4' }
        };
        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

        // PROTEGER LA HOJA - Solo lectura sin contraseña
        await worksheet.protect(PASSWORD_EDICION, {
            selectLockedCells: true,
            selectUnlockedCells: true,
            formatCells: false,
            formatColumns: false,
            formatRows: false,
            insertColumns: false,
            insertRows: false,
            insertHyperlinks: false,
            deleteColumns: false,
            deleteRows: false,
            sort: false,
            autoFilter: false,
            pivotTables: false
        });

        // HOJA 2: Búsqueda
        const busquedaSheet = workbook.addWorksheet('Búsqueda');
        
        // Título
        busquedaSheet.mergeCells('A1:G1');
        busquedaSheet.getCell('A1').value = 'BÚSQUEDA DE SOLICITUDES';
        busquedaSheet.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
        busquedaSheet.getCell('A1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF2E7D32' }
        };
        busquedaSheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
        busquedaSheet.getRow(1).height = 30;

        // Instrucciones
        busquedaSheet.mergeCells('A3:G3');
        busquedaSheet.getCell('A3').value = 'INSTRUCCIONES: Use Ctrl+B para buscar en esta hoja';
        busquedaSheet.getCell('A3').font = { italic: true, size: 11, bold: true };
        busquedaSheet.getCell('A3').alignment = { horizontal: 'center' };
        busquedaSheet.getCell('A3').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' }
        };

        busquedaSheet.mergeCells('A4:G4');
        busquedaSheet.getCell('A4').value = 'Presione Ctrl+B, escriba el nombre, apellido o documento y Excel lo encontrará automáticamente';
        busquedaSheet.getCell('A4').font = { italic: true, size: 10 };
        busquedaSheet.getCell('A4').alignment = { horizontal: 'center' };

        // Encabezados
        busquedaSheet.getRow(6).values = ['Fecha y Hora', 'Nombre', 'Apellido', 'Tipo Doc', 'Nro. Documento', 'Tipo Necesidad', 'Descripción'];
        busquedaSheet.getRow(6).font = { bold: true, size: 11 };
        busquedaSheet.getRow(6).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD3D3D3' }
        };

        // Ajustar anchos
        busquedaSheet.getColumn(1).width = 20;
        busquedaSheet.getColumn(2).width = 20;
        busquedaSheet.getColumn(3).width = 20;
        busquedaSheet.getColumn(4).width = 12;
        busquedaSheet.getColumn(5).width = 18;
        busquedaSheet.getColumn(6).width = 25;
        busquedaSheet.getColumn(7).width = 50;

        // PROTEGER LA HOJA DE BÚSQUEDA
        await busquedaSheet.protect(PASSWORD_EDICION, {
            selectLockedCells: true,
            selectUnlockedCells: true,
            formatCells: false,
            formatColumns: false,
            formatRows: false,
            insertColumns: false,
            insertRows: false,
            insertHyperlinks: false,
            deleteColumns: false,
            deleteRows: false,
            sort: false,
            autoFilter: false,
            pivotTables: false
        });

        await workbook.xlsx.writeFile(excelPath);
        console.log('Archivo Excel creado en:', excelPath);
    }
}

// Manejar el guardado de solicitudes
ipcMain.handle('guardar-solicitud', async (event, data) => {
    try {
        const workbook = new ExcelJS.Workbook();
        
        // Intentar abrir el archivo con reintentos
        let attempts = 0;
        const maxAttempts = 3;
        let loaded = false;

        while (attempts < maxAttempts && !loaded) {
            try {
                await workbook.xlsx.readFile(excelPath);
                loaded = true;
            } catch (error) {
                attempts++;
                if (attempts < maxAttempts) {
                    await new Promise(resolve => setTimeout(resolve, 1000));
                } else {
                    throw new Error('El archivo Excel está en uso. Por favor ciérrelo e intente nuevamente.');
                }
            }
        }

        const solicitudesSheet = workbook.getWorksheet('Solicitudes');
        const busquedaSheet = workbook.getWorksheet('Búsqueda');

        // DESPROTEGER TEMPORALMENTE PARA EDITAR
        await solicitudesSheet.unprotect(PASSWORD_EDICION);
        await busquedaSheet.unprotect(PASSWORD_EDICION);

        // Crear el nuevo registro
        const newRow = {
            fecha: new Date().toLocaleString('es-CO', {
                timeZone: 'America/Bogota',
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            }),
            nombre: data.nombre,
            apellido: data.apellido,
            tipoDocumento: data.tipoDocumento,
            numeroDocumento: data.numeroDocumento,
            tipoNecesidad: data.tipoNecesidad,
            descripcion: data.descripcion
        };

        // Agregar a la hoja Solicitudes
        solicitudesSheet.addRow([
            newRow.fecha,
            newRow.nombre,
            newRow.apellido,
            newRow.tipoDocumento,
            newRow.numeroDocumento,
            newRow.tipoNecesidad,
            newRow.descripcion
        ]);

        // Agregar también a la hoja Búsqueda
        busquedaSheet.addRow([
            newRow.fecha,
            newRow.nombre,
            newRow.apellido,
            newRow.tipoDocumento,
            newRow.numeroDocumento,
            newRow.tipoNecesidad,
            newRow.descripcion
        ]);

        // VOLVER A PROTEGER LAS HOJAS
        await solicitudesSheet.protect(PASSWORD_EDICION, {
            selectLockedCells: true,
            selectUnlockedCells: true,
            formatCells: false,
            formatColumns: false,
            formatRows: false,
            insertColumns: false,
            insertRows: false,
            insertHyperlinks: false,
            deleteColumns: false,
            deleteRows: false,
            sort: false,
            autoFilter: false,
            pivotTables: false
        });

        await busquedaSheet.protect(PASSWORD_EDICION, {
            selectLockedCells: true,
            selectUnlockedCells: true,
            formatCells: false,
            formatColumns: false,
            formatRows: false,
            insertColumns: false,
            insertRows: false,
            insertHyperlinks: false,
            deleteColumns: false,
            deleteRows: false,
            sort: false,
            autoFilter: false,
            pivotTables: false
        });

        await workbook.xlsx.writeFile(excelPath);

        return {
            success: true,
            message: 'Solicitud guardada exitosamente',
            filePath: excelPath
        };
    } catch (error) {
        console.error('Error al guardar:', error);
        return {
            success: false,
            message: error.message || 'Error al guardar la solicitud'
        };
    }
});

// Abrir el archivo Excel
ipcMain.handle('abrir-excel', async () => {
    const { shell } = require('electron');
    try {
        await shell.openPath(excelPath);
        return { success: true };
    } catch (error) {
        return { success: false, message: 'No se pudo abrir el archivo Excel' };
    }
});

// Obtener la ruta del archivo Excel
ipcMain.handle('obtener-ruta-excel', async () => {
    return excelPath;
});

app.whenReady().then(async () => {
    await initializeExcel();
    createWindow();

    app.on('activate', function () {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
});

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') app.quit();
});