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
        .addItem('🔄 Resync ke BINAMARGA (Batch)', 'resyncAllToBinamargaBatchWithProgress')
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