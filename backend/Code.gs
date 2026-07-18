// ==================== MAIN ENTRY POINT ====================

function doGet() {
    ensureSheetsExist();
    return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('ASPIRE - Sistem Informasi Pengelola Aduan DPU')
        .setWidth(1400)
        .setHeight(900)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
    ensureSheetsExist();
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📋 ASPIRE')
        .addItem('➕ Buka Web App', 'openWebApp')
        .addItem('📊 Lihat Laporan', 'showAllReports')
        .addSeparator()
        .addItem('🔄 Resync ke BINAMARGA (Batch)', 'resyncAllToBinamargaBatch')
        .addItem('🔄 Resync ke Backup', 'resyncAllToBackup')
        .addSeparator()
        .addItem('⚙️ Reset Struktur', 'resetAllSheets')
        .addToUi();
}

function openWebApp() {
    const html = HtmlService.createHtmlOutput('<html><script>window.open("' + ScriptApp.getService().getUrl() + '", "_blank");google.script.host.close();</script></html>')
        .setWidth(1)
        .setHeight(1);
    SpreadsheetApp.getUi().showModalDialog(html, 'Membuka ASPIRE...');
}

function showAllReports() {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_LAPORAN);
    if (sheet) sheet.activate();
}

function resetAllSheets() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Konfirmasi', 'Apakah Anda yakin ingin mereset semua sheet?', ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        [SHEET_LAPORAN, SHEET_BINAMARGA, SHEET_MASTER_STATUS, SHEET_MASTER_DISPOSISI, 
         SHEET_MASTER_SUMBER, SHEET_WA_CONFIG, SHEET_USER_LOGIN, SHEET_KEYWORD]
            .forEach(sheetName => {
                const sheet = ss.getSheetByName(sheetName);
                if (sheet) ss.deleteSheet(sheet);
            });
        ensureSheetsExist();
        ui.alert('✅ Reset berhasil!');
    }
}

function getManifest() {
    return HtmlService.createHtmlOutputFromFile('manifest.json')
        .setMimeType('application/json');
}

function getServiceWorker() {
    return HtmlService.createHtmlOutputFromFile('sw.js')
        .setMimeType('application/javascript');
}

function deployPWA() {
    var url = ScriptApp.getService().getUrl();
    Logger.log('✅ ASPIRE PWA URL: ' + url);
    return url;
}

function exportLaporanToNewSheet(filter) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sourceSheet = ss.getSheetByName(SHEET_LAPORAN);
        if (!sourceSheet) {
            return { success: false, message: 'Sheet LaporanAduan tidak ditemukan' };
        }
        
        const data = sourceSheet.getDataRange().getValues();
        const headers = data[0];
        let filteredData = [];
        
        // Filter data
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            let include = true;
            
            // Filter by year
            if (filter.year) {
                const tanggal = new Date(row[0]);
                if (tanggal.getFullYear() !== filter.year) include = false;
            }
            
            // Filter by date range
            if (filter.startDate) {
                const tanggal = new Date(row[0]);
                const startDate = new Date(filter.startDate);
                if (tanggal < startDate) include = false;
            }
            
            if (filter.endDate) {
                const tanggal = new Date(row[0]);
                const endDate = new Date(filter.endDate);
                endDate.setHours(23, 59, 59);
                if (tanggal > endDate) include = false;
            }
            
            // Filter by disposisi
            if (filter.disposisi && row[3] !== filter.disposisi) {
                include = false;
            }
            
            // Filter by status
            if (filter.status && row[7] !== filter.status) {
                include = false;
            }
            
            if (include) filteredData.push(row);
        }
        
        if (filteredData.length === 0) {
            return { success: false, message: 'Tidak ada data yang sesuai dengan filter' };
        }
        
        // Buat sheet baru
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
        const newSheetName = 'EXPORT_' + timestamp;
        const newSheet = ss.insertSheet(newSheetName);
        
        // Tulis headers
        newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        newSheet.getRange(1, 1, 1, headers.length)
            .setFontWeight('bold')
            .setBackground('#4F46E5')
            .setFontColor('white');
        
        // Tulis data
        if (filteredData.length > 0) {
            newSheet.getRange(2, 1, filteredData.length, headers.length).setValues(filteredData);
        }
        
        newSheet.autoResizeColumns(1, headers.length);
        newSheet.setFrozenRows(1);
        
        const spreadsheetUrl = ss.getUrl();
        
        return {
            success: true,
            message: 'Export berhasil! ' + filteredData.length + ' data diekspor.',
            totalData: filteredData.length,
            sheetName: newSheetName,
            spreadsheetUrl: spreadsheetUrl
        };
        
    } catch (e) {
        console.error('❌ Error export:', e.message);
        return { success: false, message: 'Error: ' + e.message };
    }
}
