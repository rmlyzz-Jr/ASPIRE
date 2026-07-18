// ==================== KONFIGURASI SPREADSHEET ====================
const SPREADSHEET_ID = '1p6KsVDerY2eogOI9JwF4Ja2q2M-0yADL5IZ4OAF8Jos';
const SPREADSHEET_BACKUP_ID = '1F6FIhJXnKHbcj7aMwrtNGRnnWLwptNlzZtot-Qd97Jk';

// ==================== NAMA SHEET ====================
const SHEET_LAPORAN = 'LaporanAduan';
const SHEET_BINAMARGA = 'BINAMARGA';
const SHEET_MASTER_STATUS = 'master status';
const SHEET_MASTER_DISPOSISI = 'master disposisi';
const SHEET_MASTER_SUMBER = 'mastersumber';
const SHEET_WA_CONFIG = 'MASTER_WA_CONFIG';
const SHEET_USER_LOGIN = 'USER_LOGIN';
const SHEET_KEYWORD = 'MASTER_KEYWORD';

// ==================== KOLOM SHEET (0-index) ====================
const COL_TANGGAL = 0;
const COL_KODE_LAPORAN = 1;
const COL_DESKRIPSI = 2;
const COL_DISPOSISI = 3;
const COL_PEMOHON = 4;
const COL_DETAIL_LOKASI = 5;
const COL_ASAL_MEDIA = 6;
const COL_STATUS = 7;
const COL_NOTES = 8;
const COL_GAMBAR = 9;
const COL_USER = 10;
const COL_JUDUL = 11;

// ==================== KONFIGURASI WAsender API ====================
const WASENDER_API_KEY = '351e732e407183e784178e3861a44ae26896b10c41811fb5ffd4fd6bb4570168';
const WASENDER_API_URL = 'https://www.wasenderapi.com/api/send-message';

// ==================== KONFIGURASI GROUP TUJUAN ====================
const WA_GROUP_ID = '120363283721432243@g.us';
const DEFAULT_MENTION = "6285741448845@s.whatsapp.net";

// ==================== DISPOSISI BINAMARGA LIST ====================
const DISPOSISI_BINAMARGA_LIST = [
    'BINAMARGA',
    'Binamarga',
    'binamarga',
    'BINA MARGA',
    'Bina Marga',
    'bina marga',
    '1108'
];

// ========================================================================
// ==================== HELPER FUNCTIONS ==================================
// ========================================================================

function isDisposisiBinamarga(disposisiValue) {
    if (!disposisiValue) return false;
    const cleanValue = disposisiValue.toString().toLowerCase().trim();
    for (let i = 0; i < DISPOSISI_BINAMARGA_LIST.length; i++) {
        if (cleanValue === DISPOSISI_BINAMARGA_LIST[i].toLowerCase().trim()) {
            return true;
        }
    }
    if (cleanValue.includes('binamarga') || 
        cleanValue.includes('bina marga') || 
        cleanValue.includes('1108')) {
        return true;
    }
    return false;
}

function formatTanggalHelper(tanggalValue) {
    try {
        if (!tanggalValue) return '-';
        var tgl = new Date(tanggalValue);
        if (isNaN(tgl.getTime())) return '-';
        var tahun = tgl.getFullYear();
        var bulan = String(tgl.getMonth() + 1).padStart(2, '0');
        var hari = String(tgl.getDate()).padStart(2, '0');
        return tahun + '-' + bulan + '-' + hari;
    } catch(e) {
        return '-';
    }
}

function formatTanggalIndonesia(tanggalValue) {
    try {
        if (!tanggalValue) return '-';
        var tgl = new Date(tanggalValue);
        if (isNaN(tgl.getTime())) return '-';
        var days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
        var months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 
                      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        return days[tgl.getDay()] + ', ' + tgl.getDate() + ' ' + months[tgl.getMonth()] + ' ' + tgl.getFullYear();
    } catch(e) {
        return '-';
    }
}

function generateKodeLaporan() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_LAPORAN);
        const date = new Date();
        const todayStr = date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
        let count = 1;
        if (sheet && sheet.getLastRow() > 1) {
            const data = sheet.getRange(2, COL_TANGGAL + 1, sheet.getLastRow() - 1, 1).getValues();
            for (let i = 0; i < data.length; i++) {
                if (data[i][0]) {
                    const cellDate = new Date(data[i][0]);
                    if (!isNaN(cellDate.getTime())) {
                        const cellDateStr = cellDate.getFullYear() + '-' + String(cellDate.getMonth() + 1).padStart(2, '0') + '-' + String(cellDate.getDate()).padStart(2, '0');
                        if (cellDateStr === todayStr) count++;
                    }
                }
            }
        }
        return 'SIPU/' + todayStr.replace(/-/g, '') + '/' + String(count).padStart(3, '0');
    } catch (error) {
        const date = new Date();
        return 'SIPU/' + date.getFullYear() + String(date.getMonth() + 1).padStart(2, '0') + String(date.getDate()).padStart(2, '0') + '/001';
    }
}

function findColumnIndex(headers, columnNames) {
    if (typeof columnNames === 'string') {
        columnNames = [columnNames];
    }
    for (let i = 0; i < headers.length; i++) {
        const header = headers[i] ? headers[i].toString().trim() : '';
        for (let j = 0; j < columnNames.length; j++) {
            const searchName = columnNames[j] ? columnNames[j].toString().trim() : '';
            if (header.toLowerCase() === searchName.toLowerCase()) {
                return i;
            }
        }
    }
    return -1;
}

function deleteRowByKode(sheet, kodeLaporan) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const kodeIndex = findColumnIndex(headers, ['KODE LAPORAN', 'KODE_LAPORAN', 'kodeLaporan']);
    if (kodeIndex === -1) return;
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][kodeIndex] === kodeLaporan) {
            sheet.deleteRow(i + 1);
            break;
        }
    }
}

function normalizeMentionsArray(mentionsInput) {
    if (!mentionsInput) return [];
    if (Array.isArray(mentionsInput)) {
        return mentionsInput.map(n => ('' + n).replace(/@s\.whatsapp\.net$/i, '').trim()).filter(n => n);
    }
    const s = '' + mentionsInput;
    return s.split(',').map(n => n.trim()).filter(n => n).map(n => n.replace(/@s\.whatsapp\.net$/i, ''));
}

// ========================================================================
// ==================== KEYWORD FUNCTIONS ================================
// ========================================================================

function loadKeywordsFromSheet() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let sheet = ss.getSheetByName(SHEET_KEYWORD);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_KEYWORD);
            sheet.getRange(1, 1, 1, 4).setValues([['KEYWORD', 'JUDUL', 'ICON', 'PRIORITY']]);
            sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#F59E0B').setFontColor('white');
            sheet.setColumnWidths(1, 4, 200);
            const defaultData = [
                ['jalan berlubang', 'JALAN BERLUBANG', '🚧', 1],
                ['banjir', 'BANJIR & GENANGAN', '🌊', 1],
                ['longsor', 'LONGSOR', '⛰️', 1],
                ['pohon tumbang', 'POHON TUMBANG', '🌳', 1],
                ['drainase', 'DRAINASE TERSUMBAT', '💧', 1],
                ['sampah menumpuk', 'SAMPAH MENUMPUK', '🗑️', 1],
                ['kebakaran', 'KEBAKARAN', '🔥', 1],
            ];
            if (defaultData.length) sheet.getRange(2, 1, defaultData.length, 4).setValues(defaultData);
            sheet.setFrozenRows(1);
            sheet.autoResizeColumns(1, 4);
        }
        if (sheet.getLastRow() <= 1) return [];
        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
        const keywords = [];
        for (let i = 0; i < data.length; i++) {
            const keyword = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
            const judul = data[i][1] ? data[i][1].toString().trim() : '';
            const icon = data[i][2] ? data[i][2].toString().trim() : '📋';
            const priority = data[i][3] ? parseInt(data[i][3]) : 1;
            if (keyword && judul) keywords.push({ keyword, judul, icon, priority });
        }
        keywords.sort((a, b) => a.priority - b.priority);
        return keywords;
    } catch (e) {
        console.error('Error loadKeywordsFromSheet:', e);
        return [];
    }
}

function getKeywordDataWeb() {
    try {
        const keywords = loadKeywordsFromSheet();
        return { success: true, data: keywords };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function detectKeyword(deskripsi) {
    if (!deskripsi) return { title: "📋 ADUAN MASYARAKAT", icon: "📋", rawTitle: "ADUAN MASYARAKAT", keyword: null };
    const lowerDesc = deskripsi.toLowerCase();
    const keywords = loadKeywordsFromSheet();
    if (keywords.length === 0) return { title: "📋 ADUAN MASYARAKAT", icon: "📋", rawTitle: "ADUAN MASYARAKAT", keyword: null };
    for (const item of keywords) {
        if (lowerDesc.includes(item.keyword)) {
            return { title: item.icon + ' ' + item.judul, icon: item.icon, rawTitle: item.judul, keyword: item.keyword };
        }
    }
    const words = deskripsi.trim().split(/\s+/);
    let fallbackTitle = '';
    if (words.length <= 3) fallbackTitle = deskripsi.trim();
    else fallbackTitle = words.slice(0, 5).join(' ') + (words.length > 5 ? '...' : '');
    fallbackTitle = fallbackTitle.charAt(0).toUpperCase() + fallbackTitle.slice(1);
    const locationKeywords = ["jalan", "jl", "perumahan", "kecamatan", "kelurahan", "desa", "rt", "rw", "gang"];
    const hasLocation = locationKeywords.some(loc => lowerDesc.includes(loc));
    const icon = hasLocation ? "📍" : "📋";
    return { title: icon + ' ' + fallbackTitle, icon: icon, rawTitle: fallbackTitle, keyword: null };
}

// ========================================================================
// ==================== IMAGE UPLOAD =====================================
// ========================================================================

function uploadImageToDrive(base64Image, kodeLaporan) {
    try {
        if (!base64Image) return null;
        console.log('📸 Memulai upload gambar untuk:', kodeLaporan);
        const folderName = "SIPU_Laporan_Aduan_Gambar";
        let folder;
        const folderIterator = DriveApp.getFoldersByName(folderName);
        if (folderIterator.hasNext()) folder = folderIterator.next();
        else folder = DriveApp.createFolder(folderName);
        let base64Data = base64Image;
        let contentType = 'image/jpeg';
        const match = base64Image.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
        if (match) { contentType = match[1]; base64Data = match[2]; }
        const raw = Utilities.base64Decode(base64Data);
        const ext = (contentType.split('/')[1] || 'jpg').split('+')[0];
        const fileName = kodeLaporan + '_' + new Date().getTime() + '.' + ext;
        const imageBlob = Utilities.newBlob(raw, contentType, fileName);
        const file = folder.createFile(imageBlob);
        file.setDescription('Gambar laporan aduan SIPU - ' + kodeLaporan);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
        const fileId = file.getId();
        const imageUrl = 'https://drive.google.com/uc?export=download&id=' + fileId;
        console.log('✅ Gambar berhasil diupload. ID:', fileId);
        return imageUrl;
    } catch (e) {
        console.error('❌ Error upload gambar ke Drive:', e.toString());
        return null;
    }
}

function getImagesFromDriveWeb(data) {
    try {
        const folderName = "SIPU_Laporan_Aduan_Gambar";
        let folder;
        const folderIterator = DriveApp.getFoldersByName(folderName);
        if (!folderIterator.hasNext()) return { success: false, message: 'Folder tidak ditemukan' };
        folder = folderIterator.next();
        const files = folder.getFiles();
        const imageUrls = [];
        while (files.hasNext()) {
            const file = files.next();
            if (file.getName().startsWith(data.kodeLaporan)) {
                const fileId = file.getId();
                const imageUrl = 'https://drive.google.com/uc?export=download&id=' + fileId;
                imageUrls.push(imageUrl);
            }
        }
        if (imageUrls.length === 0) return { success: false, message: 'Gambar tidak ditemukan' };
        return { success: true, imageUrls: imageUrls, totalImages: imageUrls.length };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// ========================================================================
// ==================== WHATSAPP FUNCTIONS ===============================
// ========================================================================

function getNamaFromWAConfig(nomorWA) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_WA_CONFIG);
        if (!sheet || sheet.getLastRow() <= 1) return null;
        let cleanNumber = nomorWA.replace(/@s\.whatsapp\.net$/i, '').trim();
        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
        for (let i = 0; i < data.length; i++) {
            const nomorSheet = data[i][1] ? data[i][1].toString().trim() : '';
            const nama = data[i][2] ? data[i][2].toString().trim() : '';
            if (nomorSheet === cleanNumber) return nama;
        }
        return null;
    } catch (e) {
        console.error('Error getNamaFromWAConfig:', e);
        return null;
    }
}

function getMentionListByDisposisi(disposisiValue) {
    try {
        if (!disposisiValue) return [];
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_WA_CONFIG);
        if (!sheet || sheet.getLastRow() <= 1) return [];
        
        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
        const mentionList = [];
        const disposisiClean = disposisiValue.toString().toUpperCase().trim();
        
        for (let i = 0; i < data.length; i++) {
            const rowDisposisi = data[i][0] ? data[i][0].toString().toUpperCase().trim() : '';
            const nomorWA = data[i][1] ? data[i][1].toString().trim() : '';
            if (rowDisposisi === disposisiClean && nomorWA) {
                mentionList.push(nomorWA);
            }
        }
        return mentionList;
    } catch (e) {
        console.error('Error getMentionListByDisposisi:', e);
        return [];
    }
}

function formatWAMessageWithMention(data, kodeLaporan, keyword, mentionNumbers) {
    if (!data) {
        return {
            pesan: '⚠️ Error: Data tidak tersedia',
            mentionList: [DEFAULT_MENTION],
            nama: 'Unknown',
            cleanNumber: '0'
        };
    }
    if (!data.deskripsi) data.deskripsi = '[Deskripsi tidak tersedia]';

    var numbers = normalizeMentionsArray(mentionNumbers);
    var nama = 'Unknown';
    if (numbers.length > 0) {
        var tmp = getNamaFromWAConfig(numbers[0]);
        nama = tmp || numbers[0];
    } else {
        var tmp2 = getNamaFromWAConfig(DEFAULT_MENTION.replace(/@s\.whatsapp\.net$/i, ''));
        nama = tmp2 || DEFAULT_MENTION.replace(/@s\.whatsapp\.net$/i, '');
    }

    var mentionList = numbers.map(function(n) { return n + '@s.whatsapp.net'; });
    if (mentionList.length === 0) mentionList = [DEFAULT_MENTION];

    var deskripsiText = data.deskripsi || '-';
    var finalKeyword = keyword || '📋 ADUAN MASYARAKAT';
    var finalDeskripsi = deskripsiText;
    
    var doubledPattern = new RegExp('^(' + finalKeyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')\\s*\\n\\s*\\1', 'i');
    if (finalDeskripsi.match(doubledPattern)) {
        finalDeskripsi = finalDeskripsi.replace(new RegExp('^(' + finalKeyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')\\s*\\n\\s*'), '');
    }

    var tanggalFormatted = formatTanggalIndonesia(data.tanggal);
    var pesan = '';
    
    if (finalDeskripsi.indexOf(finalKeyword) === -1) {
        pesan = finalKeyword + '\n\n';
    }
    
    pesan += '*KODE LAPORAN:* ' + (kodeLaporan || '-') + '\n';
    pesan += '*TANGGAL:* ' + tanggalFormatted + '\n\n';
    pesan += '*DESKRIPSI ADUAN:*\n';
    pesan += finalDeskripsi + '\n\n';
    pesan += '*LOKASI ADUAN:* \n' + (data.detailLokasi || '-') + '\n\n';
    pesan += '*SUMBER LAPORAN:* ' + (data.asalMedia || '-') + '\n\n';
    pesan += '*PELAPOR:* ' + (data.pemohon || 'Anonim') + '\n\n';
    pesan += '*CATATAN:* ' + (data.notes || '-') + '\n\n';
    
    if (data.disposisi) {
        pesan += '*BIDANG:* ' + data.disposisi + '\n\n';
    }
    
    pesan += '---\n';
    pesan += 'Mohon ditindaklanjuti oleh: ' + mentionList.map(function(m) {
        return '@' + m.replace(/@s\.whatsapp\.net$/, '');
    }).join(', ') + '\n\n';
    pesan += '#' + (data.asalMedia || 'LaporanAduan');

    var cleanNumber = (numbers.length > 0) ? numbers[0] : (DEFAULT_MENTION.replace(/@s\.whatsapp\.net$/i, ''));
    return { 
        pesan: pesan, 
        mentionList: mentionList, 
        nama: nama, 
        cleanNumber: cleanNumber 
    };
}

function sendToWhatsAppGroup(groupId, pesan, mentionList, imageUrl) {
    if (!WASENDER_API_KEY || WASENDER_API_KEY === 'YOUR_WASENDER_API_KEY_HERE') {
        console.warn('⚠️ API Key WaSender belum dikonfigurasi.');
        return { success: false, message: 'API Key belum diisi' };
    }
    
    const payload = {
        to: groupId,
        text: pesan,
        mentions: mentionList || []
    };
    
    if (imageUrl && imageUrl !== '' && imageUrl !== 'null') {
        payload.imageUrl = imageUrl;
    }
    
    const options = {
        method: 'post',
        headers: {
            'Authorization': 'Bearer ' + WASENDER_API_KEY,
            'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };
    
    try {
        const response = UrlFetchApp.fetch(WASENDER_API_URL, options);
        const result = JSON.parse(response.getContentText());
        console.log('✅ WaSender Response:', JSON.stringify(result, null, 2));
        
        if (result.success === true) {
            return { success: true, message: 'Terkirim', data: result };
        } else {
            return { success: false, message: result.message || result.reason || 'Gagal terkirim', data: result };
        }
    } catch (e) {
        console.error('❌ Error kirim WA:', e.message);
        return { success: false, message: e.toString() };
    }
}

// ========================================================================
// ==================== SINKRONISASI DATA ================================
// ========================================================================

function syncToBackupBinamarga(rowData, headers, kodeLaporan) {
    try {
        const backupSs = SpreadsheetApp.openById(SPREADSHEET_BACKUP_ID);
        let binSheet = backupSs.getSheetByName(SHEET_BINAMARGA);
        
        const desiredHeaders = [
            'TANGGAL', 'KODE LAPORAN', 'DESKRIPSI', 'PEMOHON', 
            'DETAIL LOKASI', 'ASAL MEDIA', 'STATUS', 'NOTES', 'GAMBAR'
        ];
        
        if (!binSheet) {
            binSheet = backupSs.insertSheet(SHEET_BINAMARGA);
            binSheet.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
            binSheet.getRange(1, 1, 1, desiredHeaders.length)
                .setFontWeight('bold')
                .setBackground('#10B981')
                .setFontColor('white');
            binSheet.setFrozenRows(1);
        }
        
        const binHeaders = binSheet.getRange(1, 1, 1, binSheet.getLastColumn()).getValues()[0];
        const binKodeCol = binHeaders.indexOf('KODE LAPORAN') + 1;
        if (binKodeCol === 0) return;
        
        const binData = binSheet.getDataRange().getValues();
        let targetRow = -1;
        for (let i = 1; i < binData.length; i++) {
            if (binData[i][binKodeCol - 1] === kodeLaporan) {
                targetRow = i + 1;
                break;
            }
        }
        
        const newRow = [];
        for (let i = 0; i < desiredHeaders.length; i++) {
            const colName = desiredHeaders[i];
            const sourceIndex = headers.indexOf(colName);
            if (sourceIndex !== -1 && sourceIndex < rowData.length) {
                newRow.push(rowData[sourceIndex] !== undefined ? rowData[sourceIndex] : '');
            } else {
                newRow.push('');
            }
        }
        
        if (targetRow !== -1) {
            for (let i = 0; i < newRow.length; i++) {
                binSheet.getRange(targetRow, i + 1).setValue(newRow[i]);
            }
        } else {
            binSheet.appendRow(newRow);
        }
        
        binSheet.autoResizeColumns(1, desiredHeaders.length);
    } catch (e) {
        console.error('❌ Error syncToBackupBinamarga:', e.message);
    }
}

function syncToBinamarga(ss, rowData, headers, kodeLaporan) {
    const disposisiIndex = headers.indexOf('DISPOSISI');
    let disposisiValue = '';
    if (disposisiIndex !== -1 && disposisiIndex < rowData.length) {
        disposisiValue = rowData[disposisiIndex] ? rowData[disposisiIndex].toString() : '';
    }
    
    if (!isDisposisiBinamarga(disposisiValue)) {
        console.log('⏭️ Dilewati (bukan Binamarga): ' + kodeLaporan);
        return;
    }
    
    const desiredHeaders = [
        'TANGGAL', 'KODE LAPORAN', 'DESKRIPSI', 'PEMOHON', 
        'DETAIL LOKASI', 'ASAL MEDIA', 'STATUS', 'NOTES', 'GAMBAR'
    ];
    
    let binSheet = ss.getSheetByName(SHEET_BINAMARGA);
    if (!binSheet) {
        binSheet = ss.insertSheet(SHEET_BINAMARGA);
        binSheet.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
        binSheet.getRange(1, 1, 1, desiredHeaders.length)
            .setFontWeight('bold')
            .setBackground('#10B981')
            .setFontColor('white');
        binSheet.setFrozenRows(1);
    }
    
    const binHeaders = binSheet.getRange(1, 1, 1, binSheet.getLastColumn()).getValues()[0];
    const binKodeCol = binHeaders.indexOf('KODE LAPORAN') + 1;
    if (binKodeCol === 0) return;
    
    const binData = binSheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < binData.length; i++) {
        if (binData[i][binKodeCol - 1] === kodeLaporan) {
            targetRow = i + 1;
            break;
        }
    }
    
    const newRow = [];
    for (let i = 0; i < desiredHeaders.length; i++) {
        const colName = desiredHeaders[i];
        const sourceIndex = headers.indexOf(colName);
        if (sourceIndex !== -1 && sourceIndex < rowData.length) {
            newRow.push(rowData[sourceIndex] !== undefined ? rowData[sourceIndex] : '');
        } else {
            newRow.push('');
        }
    }
    
    if (targetRow !== -1) {
        for (let i = 0; i < newRow.length; i++) {
            binSheet.getRange(targetRow, i + 1).setValue(newRow[i]);
        }
    } else {
        binSheet.appendRow(newRow);
    }
    
    binSheet.autoResizeColumns(1, desiredHeaders.length);
    
    syncToBackupBinamarga(rowData, headers, kodeLaporan);
}

function deleteFromBackup(kodeLaporan) {
    try {
        const backupSs = SpreadsheetApp.openById(SPREADSHEET_BACKUP_ID);
        const backupBinSheet = backupSs.getSheetByName(SHEET_BINAMARGA);
        if (backupBinSheet) {
            deleteRowByKode(backupBinSheet, kodeLaporan);
            console.log('🗑️ Dihapus dari BACKUP: ' + kodeLaporan);
        }
    } catch (e) {
        console.error('❌ Gagal hapus dari backup:', e.message);
    }
}

// ========================================================================
// ==================== DATABASE - SHEET MANAGEMENT =======================
// ========================================================================

function ensureSheetsExist() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        
        // LaporanAduan
        let sheet = ss.getSheetByName(SHEET_LAPORAN);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_LAPORAN);
            const headers = ['TANGGAL', 'KODE LAPORAN', 'DESKRIPSI', 'DISPOSISI', 
                           'PEMOHON', 'DETAIL LOKASI', 'ASAL MEDIA', 'STATUS', 
                           'NOTES', 'GAMBAR', 'USER', 'JUDUL'];
            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            sheet.getRange(1, 1, 1, headers.length)
                .setFontWeight('bold')
                .setBackground('#4F46E5')
                .setFontColor('white');
            sheet.setFrozenRows(1);
        }
        
        // BINAMARGA
        let binSheet = ss.getSheetByName(SHEET_BINAMARGA);
        if (!binSheet) {
            binSheet = ss.insertSheet(SHEET_BINAMARGA);
            const headers = ['TANGGAL', 'KODE LAPORAN', 'DESKRIPSI', 'PEMOHON', 
                           'DETAIL LOKASI', 'ASAL MEDIA', 'STATUS', 'NOTES', 'GAMBAR'];
            binSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            binSheet.getRange(1, 1, 1, headers.length)
                .setFontWeight('bold')
                .setBackground('#10B981')
                .setFontColor('white');
            binSheet.setFrozenRows(1);
        }
        
        // Master Status
        sheet = ss.getSheetByName(SHEET_MASTER_STATUS);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_MASTER_STATUS);
            sheet.getRange(1, 1).setValue('STATUS')
                .setFontWeight('bold')
                .setBackground('#4F46E5')
                .setFontColor('white');
            ['Menunggu', 'Proses', 'Selesai', 'Ditolak'].forEach((v, i) => 
                sheet.getRange(i + 2, 1).setValue(v)
            );
        }
        
        // Master Disposisi
        sheet = ss.getSheetByName(SHEET_MASTER_DISPOSISI);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_MASTER_DISPOSISI);
            sheet.getRange(1, 1).setValue('DISPOSISI')
                .setFontWeight('bold')
                .setBackground('#4F46E5')
                .setFontColor('white');
            ['Bidang Pelayanan', 'Bidang Teknis', 'Bidang Administrasi', 'BINAMARGA']
                .forEach((v, i) => sheet.getRange(i + 2, 1).setValue(v));
        }
        
        // Master Sumber
        sheet = ss.getSheetByName(SHEET_MASTER_SUMBER);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_MASTER_SUMBER);
            sheet.getRange(1, 1).setValue('SUMBER')
                .setFontWeight('bold')
                .setBackground('#4F46E5')
                .setFontColor('white');
            ['WhatsApp', 'Telepon', 'Email', 'Website']
                .forEach((v, i) => sheet.getRange(i + 2, 1).setValue(v));
        }
        
        // WA Config
        sheet = ss.getSheetByName(SHEET_WA_CONFIG);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_WA_CONFIG);
            sheet.getRange(1, 1, 1, 3).setValues([['DISPOSISI', 'NOMOR_WA', 'KETERANGAN']]);
            sheet.getRange(1, 1, 1, 3)
                .setFontWeight('bold')
                .setBackground('#25D366')
                .setFontColor('white');
            sheet.setFrozenRows(1);
        }
        
        // USER_LOGIN
        sheet = ss.getSheetByName(SHEET_USER_LOGIN);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_USER_LOGIN);
            sheet.getRange(1, 1, 1, 6).setValues([
                ['USERNAME', 'PASSWORD', 'NAMA_LENGKAP', 'ROLE', 'EMAIL', 'LAST_LOGIN']
            ]);
            sheet.getRange(1, 1, 1, 6)
                .setFontWeight('bold')
                .setBackground('#4F46E5')
                .setFontColor('white');
            const defaultUsers = [
                ['admin', 'admin123', 'Administrator SIPU', 'ADMIN', 'admin@sipu.com', ''],
                ['romly', 'romly123', 'Romly Wahyu', 'ADMIN', 'romlywa@gmail.com', ''],
                ['user1', 'user123', 'Petugas Lapangan', 'USER', 'user1@sipu.com', '']
            ];
            if (defaultUsers.length) {
                sheet.getRange(2, 1, defaultUsers.length, 6).setValues(defaultUsers);
            }
            sheet.setFrozenRows(1);
        }
        
        // MASTER_KEYWORD
        sheet = ss.getSheetByName(SHEET_KEYWORD);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_KEYWORD);
            sheet.getRange(1, 1, 1, 4).setValues([['KEYWORD', 'JUDUL', 'ICON', 'PRIORITY']]);
            sheet.getRange(1, 1, 1, 4)
                .setFontWeight('bold')
                .setBackground('#F59E0B')
                .setFontColor('white');
            const defaultKeywords = [
                ['jalan berlubang', 'JALAN BERLUBANG', '🚧', 1],
                ['banjir', 'BANJIR & GENANGAN', '🌊', 1],
                ['longsor', 'LONGSOR', '⛰️', 1],
                ['pohon tumbang', 'POHON TUMBANG', '🌳', 1],
                ['drainase', 'DRAINASE TERSUMBAT', '💧', 1],
                ['sampah menumpuk', 'SAMPAH MENUMPUK', '🗑️', 1],
                ['kebakaran', 'KEBAKARAN', '🔥', 1],
            ];
            if (defaultKeywords.length) {
                sheet.getRange(2, 1, defaultKeywords.length, 4).setValues(defaultKeywords);
            }
            sheet.setFrozenRows(1);
        }
        
        return true;
    } catch (error) {
        console.error('Error ensuring sheets exist:', error);
        return false;
    }
}

// ========================================================================
// ==================== DATABASE - CRUD OPERATIONS =======================
// ========================================================================

function getMasterDataWeb() {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const masters = { status: [], disposisi: [], sumber: [] };
        
        let sheet = ss.getSheetByName(SHEET_MASTER_STATUS);
        if (sheet && sheet.getLastRow() > 1) {
            const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
            for (let i = 0; i < data.length; i++) {
                if (data[i][0] && data[i][0].toString().trim()) {
                    masters.status.push(data[i][0].toString().trim());
                }
            }
        }
        if (masters.status.length === 0) masters.status = ['Menunggu', 'Proses', 'Selesai', 'Ditolak'];
        
        sheet = ss.getSheetByName(SHEET_MASTER_DISPOSISI);
        if (sheet && sheet.getLastRow() > 1) {
            const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
            for (let i = 0; i < data.length; i++) {
                if (data[i][0] && data[i][0].toString().trim()) {
                    masters.disposisi.push(data[i][0].toString().trim());
                }
            }
        }
        if (masters.disposisi.length === 0) masters.disposisi = ['Bidang Pelayanan', 'Bidang Teknis', 'BINAMARGA'];
        if (!masters.disposisi.includes('BINAMARGA')) masters.disposisi.push('BINAMARGA');
        
        sheet = ss.getSheetByName(SHEET_MASTER_SUMBER);
        if (sheet && sheet.getLastRow() > 1) {
            const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
            for (let i = 0; i < data.length; i++) {
                if (data[i][0] && data[i][0].toString().trim()) {
                    masters.sumber.push(data[i][0].toString().trim());
                }
            }
        }
        if (masters.sumber.length === 0) masters.sumber = ['WhatsApp', 'Telepon', 'Email', 'Website'];
        
        return masters;
    } catch (error) {
        return {
            status: ['Menunggu', 'Proses', 'Selesai', 'Ditolak'],
            disposisi: ['Bidang Pelayanan', 'Bidang Teknis', 'BINAMARGA'],
            sumber: ['WhatsApp', 'Telepon']
        };
    }
}

function getAllReportsSimple() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_LAPORAN);
        if (!sheet) return { success: false, error: "Sheet tidak ditemukan", data: [] };
        if (sheet.getLastRow() <= 1) return { success: true, data: [], message: "Belum ada data" };
        
        const allData = sheet.getDataRange().getValues();
        const results = [];
        for (let i = 1; i < allData.length; i++) {
            const row = allData[i];
            results.push({
                tanggal: row[0] ? formatTanggalHelper(row[0]) : '-',
                kodeLaporan: row[1] ? row[1].toString() : '',
                deskripsi: row[2] ? row[2].toString() : '',
                disposisi: row[3] ? row[3].toString() : '',
                pemohon: row[4] ? row[4].toString() : '-',
                detailLokasi: row[5] ? row[5].toString() : '',
                asalMedia: row[6] ? row[6].toString() : '',
                status: row[7] ? row[7].toString() : 'Menunggu',
                notes: row[8] ? row[8].toString() : '',
                gambar: row[9] ? row[9].toString() : '',
                user: row[10] ? row[10].toString() : '-',
                judul: row[11] ? row[11].toString() : ''
            });
        }
        return { success: true, data: results, total: results.length };
    } catch(e) {
        return { success: false, error: e.toString(), data: [] };
    }
}

function saveLaporanWeb(data) {
    try {
        console.log('=== saveLaporanWeb START ===');
        if (!data) return { success: false, message: 'Data tidak tersedia!' };
        if (!data.deskripsi || data.deskripsi.trim() === '') {
            return { success: false, message: 'Deskripsi harus diisi!' };
        }
        ensureSheetsExist();
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(SHEET_LAPORAN);
        if (!sheet) ensureSheetsExist();
        
        var kodeLaporan = data.kodeLaporan;
        if (!kodeLaporan || kodeLaporan.trim() === '') {
            kodeLaporan = generateKodeLaporan();
        } else {
            var s = ss.getSheetByName(SHEET_LAPORAN);
            if (s && s.getLastRow() > 1) {
                var existingData = s.getRange(2, COL_KODE_LAPORAN + 1, s.getLastRow() - 1, 1).getValues();
                for (var i = 0; i < existingData.length; i++) {
                    if (existingData[i][0] === kodeLaporan) {
                        return { success: false, message: 'Kode laporan sudah ada! Gunakan kode lain.' };
                    }
                }
            }
        }
        
        var userName = data.userName || 'Unknown';
        var disposisiValue = data.disposisi || '';
        var mainSheet = ss.getSheetByName(SHEET_LAPORAN);
        
        var tanggalValue = data.tanggal || new Date().toISOString().split('T')[0];
        var tanggalFormatted = formatTanggalIndonesia(tanggalValue);
        
        var keyword = data.keyword || '';
        var judul = '';
        if (!keyword || keyword.trim() === '') {
            var keywordObj = detectKeyword(data.deskripsi);
            keyword = keywordObj.title;
            judul = keywordObj.rawTitle || data.deskripsi.substring(0, 50);
        } else {
            judul = keyword.replace(/^[^\s]+\s/, '').trim() || keyword;
        }
        
        var deskripsiToSave = data.deskripsi || '';
        
        var gambarLinks = [];
        var imagesToSend = [];
        if (data.imageBase64Array && Array.isArray(data.imageBase64Array) && data.imageBase64Array.length > 0) {
            imagesToSend = data.imageBase64Array;
        } else if (data.imageBase64 && typeof data.imageBase64 === 'string') {
            imagesToSend = [data.imageBase64];
        }
        
        for (var i = 0; i < imagesToSend.length; i++) {
            var imageUrl = uploadImageToDrive(imagesToSend[i], kodeLaporan + '_img' + (i+1));
            if (imageUrl) gambarLinks.push(imageUrl);
        }
        var gambarLink = gambarLinks.join('; ');
        
        mainSheet.appendRow([
            tanggalValue,
            kodeLaporan,
            deskripsiToSave,
            disposisiValue,
            data.pemohon || '-',
            data.detailLokasi || '',
            data.asalMedia || '',
            data.status || "Menunggu",
            data.notes || "",
            gambarLink,
            userName,
            judul
        ]);
        
        if (disposisiValue && isDisposisiBinamarga(disposisiValue)) {
            var headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
            var lastRow = mainSheet.getLastRow();
            var savedData = mainSheet.getRange(lastRow, 1, 1, mainSheet.getLastColumn()).getValues()[0];
            syncToBinamarga(ss, savedData, headers, kodeLaporan);
        }
        
        var mentionsArray = normalizeMentionsArray(data.mentions);
        
        var waData = {
            tanggal: tanggalFormatted,
            deskripsi: data.deskripsi || '',
            detailLokasi: data.detailLokasi || '',
            pemohon: data.pemohon || 'Anonim',
            asalMedia: data.asalMedia || '',
            notes: data.notes || '',
            disposisi: disposisiValue
        };
        
        var fmt = formatWAMessageWithMention(waData, kodeLaporan, keyword, mentionsArray);
        var waTextResult = { success: false, message: 'Not sent' };
        try { waTextResult = sendToWhatsAppGroup(WA_GROUP_ID, fmt.pesan, fmt.mentionList, null); } catch (e) {}
        
        console.log('⏳ Menunggu 6 detik sebelum mengirim gambar...');
        Utilities.sleep(6000);
        
        var waImageResults = [];
        if (gambarLinks.length > 0) {
            try {
                for (var j = 0; j < gambarLinks.length; j++) {
                    console.log('📸 Mengirim gambar ' + (j+1) + '/' + gambarLinks.length + ' untuk:', kodeLaporan);
                    var imageCaption = keyword + '\n*KODE:* ' + kodeLaporan + '\n*GAMBAR:* ' + (j+1) + '/' + gambarLinks.length;
                    var result = sendToWhatsAppGroup(WA_GROUP_ID, imageCaption, fmt.mentionList, gambarLinks[j]);
                    waImageResults.push({ index: j + 1, success: result.success, message: result.message || 'OK' });
                    if (j < gambarLinks.length - 1) {
                        console.log('⏳ Menunggu 6 detik sebelum gambar ' + (j+2) + '...');
                        Utilities.sleep(6000);
                    }
                }
            } catch (e) {
                waImageResults.push({ success: false, message: e.toString() });
            }
        }
        
        var successCount = waImageResults.filter(function(r) { return r.success; }).length;
        var totalImages = gambarLinks.length;
        
        return {
            success: true,
            kodeLaporan: kodeLaporan,
            message: '✅ Laporan berhasil disimpan ke SIPU! ' + (totalImages > 0 ? successCount + '/' + totalImages + ' gambar terkirim' : ''),
            waTextSent: !!waTextResult.success,
            waImageResults: waImageResults,
            totalImages: totalImages,
            successImages: successCount,
            binamargaSynced: disposisiValue && isDisposisiBinamarga(disposisiValue),
            gambarLink: gambarLink,
            judul: judul
        };
    } catch (error) {
        console.error('❌ Error saveLaporanWeb:', error.message);
        return { success: false, message: error.toString() };
    }
}

function updateLaporanWithImagesWeb(data) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_LAPORAN);
        if (!sheet) throw new Error('Sheet LaporanAduan tidak ditemukan');

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        const kodeIndex = findColumnIndex(headers, ['KODE LAPORAN', 'KODE_LAPORAN', 'kodeLaporan', 'KodeLaporan']);
        if (kodeIndex === -1) {
            throw new Error('Kolom KODE_LAPORAN tidak ditemukan');
        }
        
        const kodeLaporan = data.KODE_LAPORAN || data.kodeLaporan || data.kode || data.kode_laporan;
        if (!kodeLaporan) throw new Error('Kode laporan tidak ditemukan dalam data');
        
        const allData = sheet.getDataRange().getValues();
        let rowIndex = -1;
        let oldDisposisi = '';
        let existingData = null;
        
        for (let i = 1; i < allData.length; i++) {
            if (allData[i][kodeIndex] === kodeLaporan) {
                rowIndex = i + 1;
                existingData = allData[i];
                const disposisiCol = findColumnIndex(headers, ['DISPOSISI']) + 1;
                if (disposisiCol > 0) oldDisposisi = allData[i][disposisiCol - 1] || '';
                break;
            }
        }
        
        if (rowIndex === -1) throw new Error('Data dengan kode ' + kodeLaporan + ' tidak ditemukan');
        
        const disposisiCol = findColumnIndex(headers, ['DISPOSISI']) + 1;
        const statusCol = findColumnIndex(headers, ['STATUS']) + 1;
        const deskripsiCol = findColumnIndex(headers, ['DESKRIPSI', 'Deskripsi', 'deskripsi']) + 1;
        const notesCol = findColumnIndex(headers, ['NOTES']) + 1;
        const gambarCol = findColumnIndex(headers, ['GAMBAR']) + 1;
        const judulCol = findColumnIndex(headers, ['JUDUL']) + 1;
        
        if (disposisiCol > 0) sheet.getRange(rowIndex, disposisiCol).setValue(data.disposisi);
        if (statusCol > 0) sheet.getRange(rowIndex, statusCol).setValue(data.status);
        if (deskripsiCol > 0) sheet.getRange(rowIndex, deskripsiCol).setValue(data.deskripsi);
        if (notesCol > 0) sheet.getRange(rowIndex, notesCol).setValue(data.notes);
        
        if (gambarCol > 0 && data.imageUrls && data.imageUrls.length > 0) {
            var gambarLink = data.imageUrls.join('; ');
            sheet.getRange(rowIndex, gambarCol).setValue(gambarLink);
        }
        
        if (judulCol > 0 && data.judul) {
            sheet.getRange(rowIndex, judulCol).setValue(data.judul);
        }
        
        let waSent = false;
        let imageSentCount = 0;
        let imageFailedCount = 0;
        let disposisiChanged = (oldDisposisi !== data.disposisi);
        
        let mentionList = [];
        
        if (disposisiChanged) {
            var tanggalValue = existingData ? existingData[0] : new Date().toISOString().split('T')[0];
            var tanggalFormatted = formatTanggalIndonesia(tanggalValue);
            const pemohon = existingData ? existingData[4] : 'Anonim';
            const detailLokasi = existingData ? existingData[5] : '';
            const asalMedia = existingData ? existingData[6] : 'Unknown';
            const notes = data.notes || '';
            
            var keyword = data.keyword || '';
            if (!keyword || keyword.trim() === '') {
                var keywordObj = detectKeyword(data.deskripsi || (existingData ? existingData[2] : ''));
                keyword = keywordObj.title;
            }
            
            if (data.mentions && data.mentions.trim() !== '') {
                mentionList = data.mentions.split(',').map(function(m) { return m.trim(); }).filter(function(m) { return m !== ''; });
                console.log('📋 Mention dari frontend:', mentionList);
            }
            
            if (mentionList.length === 0 && data.disposisi) {
                mentionList = getMentionListByDisposisi(data.disposisi);
                console.log('📋 Mention dari WA_CONFIG:', mentionList);
            }
            
            if (mentionList.length === 0) {
                mentionList = [DEFAULT_MENTION];
                console.log('📋 Menggunakan DEFAULT_MENTION:', DEFAULT_MENTION);
            }
            
            const fmt = formatWAMessageWithMention({
                tanggal: tanggalFormatted,
                deskripsi: data.deskripsi || (existingData ? existingData[2] : ''),
                detailLokasi: detailLokasi,
                pemohon: pemohon,
                asalMedia: asalMedia,
                notes: notes,
                disposisi: data.disposisi
            }, kodeLaporan, keyword, mentionList);
            
            console.log('📤 Mengirim pesan teks dengan mention:', mentionList.join(', '));
            const textResult = sendToWhatsAppGroup(WA_GROUP_ID, fmt.pesan, fmt.mentionList, null);
            waSent = textResult.success;
            console.log('✅ Pesan teks:', waSent ? 'Terkirim' : 'Gagal');
            
            if (data.imageUrls && data.imageUrls.length > 0 && waSent) {
                console.log('⏳ Menunggu 6 detik sebelum mengirim ' + data.imageUrls.length + ' gambar...');
                Utilities.sleep(6000);
                for (let i = 0; i < data.imageUrls.length; i++) {
                    const imageUrl = data.imageUrls[i];
                    console.log('📸 Mengirim gambar ' + (i+1) + '/' + data.imageUrls.length + '...');
                    const caption = keyword + '\n*KODE:* ' + kodeLaporan + '\n*GAMBAR:* ' + (i+1) + '/' + data.imageUrls.length + '\n*STATUS:* Disposisi diubah ke ' + data.disposisi;
                    const imageResult = sendToWhatsAppGroup(WA_GROUP_ID, caption, fmt.mentionList, imageUrl);
                    if (imageResult.success) { imageSentCount++; }
                    else { imageFailedCount++; }
                    if (i < data.imageUrls.length - 1) { Utilities.sleep(6000); }
                }
            }
        }
        
        if (data.disposisi && isDisposisiBinamarga(data.disposisi)) {
            const updatedRowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
            syncToBinamarga(ss, updatedRowData, headers, kodeLaporan);
        } else if (oldDisposisi && isDisposisiBinamarga(oldDisposisi) && !isDisposisiBinamarga(data.disposisi)) {
            const binSheet = ss.getSheetByName(SHEET_BINAMARGA);
            if (binSheet) {
                deleteRowByKode(binSheet, kodeLaporan);
                deleteFromBackup(kodeLaporan);
            }
        }
        
        return {
            success: true,
            message: 'Laporan berhasil diperbarui',
            waSent: waSent,
            imageSentCount: imageSentCount,
            imageFailedCount: imageFailedCount,
            totalImages: data.imageUrls ? data.imageUrls.length : 0,
            disposisiChanged: disposisiChanged,
            mentionList: mentionList
        };
    } catch (e) {
        console.error('❌ Error updateLaporanWithImagesWeb:', e.message);
        return { success: false, message: e.toString() };
    }
}

function updateLaporanWithoutWAWeb(data) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_LAPORAN);
        if (!sheet) throw new Error('Sheet LaporanAduan tidak ditemukan');
        
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        const kodeIndex = findColumnIndex(headers, ['KODE LAPORAN', 'KODE_LAPORAN', 'kodeLaporan', 'KodeLaporan']);
        if (kodeIndex === -1) {
            throw new Error('Kolom KODE_LAPORAN tidak ditemukan');
        }
        
        const kodeLaporan = data.KODE_LAPORAN || data.kodeLaporan || data.kode || data.kode_laporan;
        if (!kodeLaporan) throw new Error('Kode laporan tidak ditemukan dalam data');
        
        const allData = sheet.getDataRange().getValues();
        let rowIndex = -1;
        let oldDisposisi = '';
        
        for (let i = 1; i < allData.length; i++) {
            if (allData[i][kodeIndex] === kodeLaporan) {
                rowIndex = i + 1;
                const disposisiCol = findColumnIndex(headers, ['DISPOSISI']) + 1;
                if (disposisiCol > 0) oldDisposisi = allData[i][disposisiCol - 1] || '';
                break;
            }
        }
        
        if (rowIndex === -1) throw new Error('Data dengan kode ' + kodeLaporan + ' tidak ditemukan');
        
        const disposisiCol = findColumnIndex(headers, ['DISPOSISI']) + 1;
        const statusCol = findColumnIndex(headers, ['STATUS']) + 1;
        const deskripsiCol = findColumnIndex(headers, ['DESKRIPSI', 'Deskripsi', 'deskripsi']) + 1;
        const notesCol = findColumnIndex(headers, ['NOTES']) + 1;
        const gambarCol = findColumnIndex(headers, ['GAMBAR']) + 1;
        const judulCol = findColumnIndex(headers, ['JUDUL']) + 1;
        
        if (disposisiCol > 0) sheet.getRange(rowIndex, disposisiCol).setValue(data.disposisi);
        if (statusCol > 0) sheet.getRange(rowIndex, statusCol).setValue(data.status);
        if (deskripsiCol > 0) sheet.getRange(rowIndex, deskripsiCol).setValue(data.deskripsi);
        if (notesCol > 0) sheet.getRange(rowIndex, notesCol).setValue(data.notes);
        
        if (gambarCol > 0 && data.imageUrls && data.imageUrls.length > 0) {
            var gambarLink = data.imageUrls.join('; ');
            sheet.getRange(rowIndex, gambarCol).setValue(gambarLink);
        }
        
        if (judulCol > 0 && data.judul) {
            sheet.getRange(rowIndex, judulCol).setValue(data.judul);
        }
        
        if (data.disposisi && isDisposisiBinamarga(data.disposisi)) {
            const updatedRowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
            syncToBinamarga(ss, updatedRowData, headers, kodeLaporan);
        } else if (oldDisposisi && isDisposisiBinamarga(oldDisposisi) && !isDisposisiBinamarga(data.disposisi)) {
            const binSheet = ss.getSheetByName(SHEET_BINAMARGA);
            if (binSheet) {
                deleteRowByKode(binSheet, kodeLaporan);
                deleteFromBackup(kodeLaporan);
            }
        }
        
        return { success: true, message: 'Laporan berhasil diperbarui' };
    } catch (e) {
        console.error('❌ Error updateLaporanWithoutWAWeb:', e.message);
        return { success: false, message: e.toString() };
    }
}

function updateLaporanWeb(data) {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_LAPORAN);
        if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "Tidak ada data!" };
        const sheetData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
        for (let i = 0; i < sheetData.length; i++) {
            if (sheetData[i][COL_KODE_LAPORAN] === data.kodeLaporan) {
                const rowNum = i + 2;
                if (data.deskripsi !== undefined) sheet.getRange(rowNum, COL_DESKRIPSI + 1).setValue(data.deskripsi);
                if (data.disposisi !== undefined) {
                    const oldDisposisi = sheetData[i][COL_DISPOSISI];
                    sheet.getRange(rowNum, COL_DISPOSISI + 1).setValue(data.disposisi);
                    if (data.disposisi && isDisposisiBinamarga(data.disposisi)) {
                        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
                        const updatedRowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
                        syncToBinamarga(ss, updatedRowData, headers, data.kodeLaporan);
                    } else if (oldDisposisi && isDisposisiBinamarga(oldDisposisi) && !isDisposisiBinamarga(data.disposisi)) {
                        const binSheet = ss.getSheetByName(SHEET_BINAMARGA);
                        if (binSheet) {
                            deleteRowByKode(binSheet, data.kodeLaporan);
                            deleteFromBackup(data.kodeLaporan);
                        }
                    }
                }
                if (data.status !== undefined) sheet.getRange(rowNum, COL_STATUS + 1).setValue(data.status);
                if (data.notes !== undefined) sheet.getRange(rowNum, COL_NOTES + 1).setValue(data.notes);
                if (data.gambar !== undefined) sheet.getRange(rowNum, COL_GAMBAR + 1).setValue(data.gambar);
                if (data.judul !== undefined) sheet.getRange(rowNum, COL_JUDUL + 1).setValue(data.judul);
                return { success: true, message: "Laporan berhasil diupdate!" };
            }
        }
        return { success: false, message: "Laporan tidak ditemukan!" };
    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

function deleteLaporanWeb(data) {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_LAPORAN);
        if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "Tidak ada data!" };
        const sheetData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
        for (let i = 0; i < sheetData.length; i++) {
            if (sheetData[i][COL_KODE_LAPORAN] === data.kodeLaporan) {
                const binSheet = ss.getSheetByName(SHEET_BINAMARGA);
                if (binSheet) {
                    deleteRowByKode(binSheet, data.kodeLaporan);
                    deleteFromBackup(data.kodeLaporan);
                }
                sheet.deleteRow(i + 2);
                return { success: true, message: "Laporan berhasil dihapus!" };
            }
        }
        return { success: false, message: "Laporan tidak ditemukan!" };
    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

function updateLaporanWithStatusCheck(data) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_LAPORAN);
        if (!sheet) throw new Error('Sheet LaporanAduan tidak ditemukan');
        
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        const kodeIndex = findColumnIndex(headers, ['KODE LAPORAN', 'KODE_LAPORAN', 'kodeLaporan', 'KodeLaporan']);
        if (kodeIndex === -1) {
            throw new Error('Kolom KODE_LAPORAN tidak ditemukan');
        }
        
        const kodeLaporan = data.KODE_LAPORAN || data.kodeLaporan || data.kode || data.kode_laporan;
        if (!kodeLaporan) throw new Error('Kode laporan tidak ditemukan');
        
        const allData = sheet.getDataRange().getValues();
        let rowIndex = -1;
        let oldStatus = '';
        let oldDisposisi = '';
        let existingData = null;
        
        for (let i = 1; i < allData.length; i++) {
            if (allData[i][kodeIndex] === kodeLaporan) {
                rowIndex = i + 1;
                existingData = allData[i];
                const statusCol = findColumnIndex(headers, ['STATUS']) + 1;
                if (statusCol > 0) oldStatus = allData[i][statusCol - 1] || '';
                const disposisiCol = findColumnIndex(headers, ['DISPOSISI']) + 1;
                if (disposisiCol > 0) oldDisposisi = allData[i][disposisiCol - 1] || '';
                break;
            }
        }
        
        if (rowIndex === -1) throw new Error('Data dengan kode ' + kodeLaporan + ' tidak ditemukan');
        
        const disposisiCol = findColumnIndex(headers, ['DISPOSISI']) + 1;
        const statusCol = findColumnIndex(headers, ['STATUS']) + 1;
        const deskripsiCol = findColumnIndex(headers, ['DESKRIPSI', 'Deskripsi', 'deskripsi']) + 1;
        const notesCol = findColumnIndex(headers, ['NOTES']) + 1;
        const gambarCol = findColumnIndex(headers, ['GAMBAR']) + 1;
        const judulCol = findColumnIndex(headers, ['JUDUL']) + 1;
        
        if (disposisiCol > 0) sheet.getRange(rowIndex, disposisiCol).setValue(data.disposisi);
        if (statusCol > 0) sheet.getRange(rowIndex, statusCol).setValue(data.status);
        if (deskripsiCol > 0) sheet.getRange(rowIndex, deskripsiCol).setValue(data.deskripsi);
        if (notesCol > 0) sheet.getRange(rowIndex, notesCol).setValue(data.notes);
        if (gambarCol > 0 && data.imageUrls && data.imageUrls.length > 0) {
            sheet.getRange(rowIndex, gambarCol).setValue(data.imageUrls.join('; '));
        }
        if (judulCol > 0 && data.judul) {
            sheet.getRange(rowIndex, judulCol).setValue(data.judul);
        }
        
        const statusChanged = (oldStatus !== data.status);
        const disposisiChanged = (oldDisposisi !== data.disposisi);
        
        if (statusChanged || disposisiChanged) {
            console.log('📊 Status berubah: ' + oldStatus + ' → ' + data.status);
            
            if (data.disposisi && isDisposisiBinamarga(data.disposisi)) {
                const updatedRowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
                const headers2 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
                syncToBinamarga(ss, updatedRowData, headers2, kodeLaporan);
            } else if (oldDisposisi && isDisposisiBinamarga(oldDisposisi) && !isDisposisiBinamarga(data.disposisi)) {
                const binSheet = ss.getSheetByName(SHEET_BINAMARGA);
                if (binSheet) {
                    deleteRowByKode(binSheet, kodeLaporan);
                    deleteFromBackup(kodeLaporan);
                }
            }
        }
        
        return {
            success: true,
            message: 'Laporan berhasil diperbarui',
            statusChanged: statusChanged,
            disposisiChanged: disposisiChanged,
            oldStatus: oldStatus,
            newStatus: data.status,
            oldDisposisi: oldDisposisi,
            newDisposisi: data.disposisi
        };
    } catch (e) {
        console.error('❌ Error updateLaporanWithStatusCheck:', e.message);
        return { success: false, message: e.toString() };
    }
}

// ========================================================================
// ==================== USER AUTHENTICATION ==============================
// ========================================================================

function verifyLoginWeb(data) {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_USER_LOGIN);
        if (!sheet || sheet.getLastRow() <= 1) {
            return { success: false, message: 'Sistem login belum dikonfigurasi' };
        }
        const username = data.username || '';
        const password = data.password || '';
        const users = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
        for (let i = 0; i < users.length; i++) {
            if (users[i][0] === username && users[i][1] === password) {
                sheet.getRange(i + 2, 6).setValue(new Date().toISOString());
                return {
                    success: true,
                    user: {
                        username: users[i][0],
                        nama: users[i][2],
                        role: users[i][3],
                        email: users[i][4]
                    }
                };
            }
        }
        return { success: false, message: 'Username atau password salah!' };
    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

function getUserByUsernameWeb(data) {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_USER_LOGIN);
        if (!sheet || sheet.getLastRow() <= 1) {
            return { success: false, message: 'Data user tidak ditemukan' };
        }
        
        const username = data.username || '';
        if (!username) {
            return { success: false, message: 'Username tidak boleh kosong' };
        }
        
        const users = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
        for (let i = 0; i < users.length; i++) {
            if (users[i][0] === username) {
                return {
                    success: true,
                    user: {
                        username: users[i][0],
                        password: users[i][1],
                        nama: users[i][2],
                        role: users[i][3],
                        email: users[i][4],
                        lastLogin: users[i][5]
                    }
                };
            }
        }
        return { success: false, message: 'User tidak ditemukan' };
    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

function updateLastLoginWeb(data) {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_USER_LOGIN);
        if (!sheet || sheet.getLastRow() <= 1) {
            return { success: false, message: 'Data user tidak ditemukan' };
        }
        
        const username = data.username || '';
        if (!username) {
            return { success: false, message: 'Username tidak boleh kosong' };
        }
        
        const users = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
        for (let i = 0; i < users.length; i++) {
            if (users[i][0] === username) {
                const rowNum = i + 2;
                const now = new Date().toISOString();
                sheet.getRange(rowNum, 6).setValue(now);
                return { success: true, message: 'Last login updated' };
            }
        }
        return { success: false, message: 'User tidak ditemukan' };
    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

// ========================================================================
// ==================== USER MANAGEMENT - ADMIN ==========================
// ========================================================================

function getAllUsersWeb() {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_USER_LOGIN);
        if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };
        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
        const users = [];
        for (let i = 0; i < data.length; i++) {
            users.push({
                username: data[i][0],
                password: data[i][1],
                nama: data[i][2],
                role: data[i][3],
                email: data[i][4],
                lastLogin: data[i][5]
            });
        }
        return { success: true, data: users };
    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

function addUserWeb(data) {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_USER_LOGIN);
        const users = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < users.length; i++) {
            if (users[i][0] === data.username) {
                return { success: false, message: 'Username sudah digunakan!' };
            }
        }
        sheet.appendRow([
            data.username,
            data.password,
            data.namaLengkap,
            data.role || 'USER',
            data.email,
            new Date().toISOString()
        ]);
        return { success: true, message: 'User berhasil ditambahkan!' };
    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

function updateUserWeb(data) {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_USER_LOGIN);
        const users = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
        for (let i = 0; i < users.length; i++) {
            if (users[i][0] === data.username) {
                const rowNum = i + 2;
                sheet.getRange(rowNum, 3).setValue(data.namaLengkap);
                sheet.getRange(rowNum, 4).setValue(data.role);
                sheet.getRange(rowNum, 5).setValue(data.email);
                if (data.newPassword && data.newPassword !== '') {
                    sheet.getRange(rowNum, 2).setValue(data.newPassword);
                }
                return { success: true, message: 'User berhasil diupdate!' };
            }
        }
        return { success: false, message: 'User tidak ditemukan!' };
    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

function deleteUserWeb(data) {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_USER_LOGIN);
        const users = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < users.length; i++) {
            if (users[i][0] === data.username) {
                sheet.deleteRow(i + 2);
                return { success: true, message: 'User berhasil dihapus!' };
            }
        }
        return { success: false, message: 'User tidak ditemukan!' };
    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

function resetPasswordWeb(data) {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_USER_LOGIN);
        const users = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
        for (let i = 0; i < users.length; i++) {
            if (users[i][0] === data.username) {
                sheet.getRange(i + 2, 2).setValue(data.newPassword);
                return { success: true, message: 'Password berhasil direset!' };
            }
        }
        return { success: false, message: 'User tidak ditemukan!' };
    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

// ========================================================================
// ==================== WA CONFIG ========================================
// ========================================================================

function getAllWAConfigWeb() {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_WA_CONFIG);
        if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };
        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
        const list = [];
        for (let i = 0; i < data.length; i++) {
            if (data[i][0] && data[i][0].toString().trim() !== '') {
                list.push({
                    disposisi: data[i][0].toString().trim(),
                    nomorWA: data[i][1] ? data[i][1].toString().trim() : '',
                    keterangan: data[i][2] ? data[i][2].toString().trim() : ''
                });
            }
        }
        return { success: true, data: list };
    } catch (e) {
        return { success: false, data: [], message: e.toString() };
    }
}

// ========================================================================
// ==================== NOTIFICATION ======================================
// ========================================================================

function sendUrgentNotificationWeb(data) {
    try {
        const pesan = data.pesan || '';
        if (!pesan) return { success: false, message: 'Pesan tidak boleh kosong' };
        const result = sendToWhatsAppGroup(WA_GROUP_ID, pesan, [DEFAULT_MENTION], null);
        console.log('📨 Notifikasi URGENT dikirim:', result);
        return {
            success: true,
            waSent: result.success,
            message: result.message || 'OK',
            data: result
        };
    } catch (e) {
        console.error('❌ Error sendUrgentNotificationWeb:', e.message);
        return { success: false, message: e.toString() };
    }
}

// ========================================================================
// ==================== RESYNC FUNCTIONS ==================================
// ========================================================================

function resyncAllToBinamargaBatch() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sourceSheet = ss.getSheetByName(SHEET_LAPORAN);
        if (!sourceSheet) throw new Error('Sheet LaporanAduan tidak ditemukan');
        
        const BATCH_SIZE = 50;
        const PAUSE_MS = 5000;
        
        let binSheet = ss.getSheetByName(SHEET_BINAMARGA);
        if (!binSheet) {
            binSheet = ss.insertSheet(SHEET_BINAMARGA);
            const headers = ['TANGGAL', 'KODE LAPORAN', 'DESKRIPSI', 'PEMOHON', 
                           'DETAIL LOKASI', 'ASAL MEDIA', 'STATUS', 'NOTES', 'GAMBAR'];
            binSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            binSheet.getRange(1, 1, 1, headers.length)
                .setFontWeight('bold')
                .setBackground('#10B981')
                .setFontColor('white');
            binSheet.setFrozenRows(1);
        }
        
        const srcHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
        const srcData = sourceSheet.getDataRange().getValues();
        
        const binamargaData = [];
        const disposisiIndex = srcHeaders.indexOf('DISPOSISI');
        const srcKodeIndex = srcHeaders.indexOf('KODE LAPORAN');
        
        for (let i = 1; i < srcData.length; i++) {
            const row = srcData[i];
            let disposisiValue = '';
            if (disposisiIndex !== -1 && disposisiIndex < row.length) {
                disposisiValue = row[disposisiIndex] ? row[disposisiIndex].toString() : '';
            }
            
            if (isDisposisiBinamarga(disposisiValue)) {
                const kodeLaporan = row[srcKodeIndex] ? row[srcKodeIndex].toString() : '';
                if (kodeLaporan) {
                    binamargaData.push({ row: row, kode: kodeLaporan });
                }
            }
        }
        
        const totalData = binamargaData.length;
        if (totalData === 0) {
            return { success: true, message: 'Tidak ada data BINAMARGA' };
        }
        
        let totalNew = 0;
        let totalUpdated = 0;
        let totalErrors = 0;
        const totalBatches = Math.ceil(totalData / BATCH_SIZE);
        
        for (let batch = 0; batch < totalBatches; batch++) {
            const start = batch * BATCH_SIZE;
            const end = Math.min(start + BATCH_SIZE, totalData);
            const batchData = binamargaData.slice(start, end);
            
            for (let j = 0; j < batchData.length; j++) {
                try {
                    const item = batchData[j];
                    const rowData = item.row;
                    const kodeLaporan = item.kode;
                    
                    const binHeaders = binSheet.getRange(1, 1, 1, binSheet.getLastColumn()).getValues()[0];
                    const binKodeCol = binHeaders.indexOf('KODE LAPORAN') + 1;
                    
                    let targetRow = -1;
                    if (binKodeCol > 0) {
                        const existingData = binSheet.getDataRange().getValues();
                        for (let k = 1; k < existingData.length; k++) {
                            if (existingData[k][binKodeCol - 1] === kodeLaporan) {
                                targetRow = k + 1;
                                break;
                            }
                        }
                    }
                    
                    const desiredHeaders = ['TANGGAL', 'KODE LAPORAN', 'DESKRIPSI', 
                        'PEMOHON', 'DETAIL LOKASI', 'ASAL MEDIA', 'STATUS', 'NOTES', 'GAMBAR'];
                    
                    const newRow = [];
                    for (let k = 0; k < desiredHeaders.length; k++) {
                        const colName = desiredHeaders[k];
                        const srcIndex = srcHeaders.indexOf(colName);
                        if (srcIndex !== -1 && srcIndex < rowData.length) {
                            newRow.push(rowData[srcIndex] !== undefined ? rowData[srcIndex] : '');
                        } else {
                            newRow.push('');
                        }
                    }
                    
                    if (targetRow !== -1) {
                        for (let k = 0; k < newRow.length; k++) {
                            binSheet.getRange(targetRow, k + 1).setValue(newRow[k]);
                        }
                        totalUpdated++;
                    } else {
                        binSheet.appendRow(newRow);
                        totalNew++;
                    }
                    
                    syncToBackupBinamarga(rowData, srcHeaders, kodeLaporan);
                    
                } catch (e) {
                    totalErrors++;
                    console.error('❌ Error:', e.message);
                }
            }
            
            if (batch < totalBatches - 1) {
                Utilities.sleep(PAUSE_MS);
            }
        }
        
        return {
            success: true,
            message: '✅ Resync selesai!\n📥 Data baru: ' + totalNew + '\n🔄 Data diupdate: ' + totalUpdated + '\n❌ Error: ' + totalErrors,
            newData: totalNew,
            updatedData: totalUpdated,
            errors: totalErrors,
            totalData: totalData
        };
        
    } catch (e) {
        return { success: false, message: '❌ Error: ' + e.toString() };
    }
}

function resyncAllToBackup() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sourceSheet = ss.getSheetByName(SHEET_BINAMARGA);
        
        if (!sourceSheet) {
            return { success: false, message: 'Sheet BINAMARGA tidak ditemukan' };
        }
        
        if (sourceSheet.getLastRow() <= 1) {
            return { success: true, message: 'Tidak ada data di BINAMARGA' };
        }
        
        const data = sourceSheet.getDataRange().getValues();
        const headers = data[0];
        let count = 0;
        
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const kodeLaporan = row[1];
            if (kodeLaporan) {
                syncToBackupBinamarga(row, headers, kodeLaporan);
                count++;
            }
        }
        
        return {
            success: true,
            message: '✅ ' + count + ' data berhasil disinkronisasi ke BACKUP',
            totalData: count
        };
        
    } catch (e) {
        console.error('❌ Error resyncAllToBackup:', e.message);
        return { success: false, message: 'Error: ' + e.message };
    }
}

function resyncAllToBinamargaBatchWithProgress() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
        '⚠️ Konfirmasi Resync Batch',
        'Proses ini akan menyinkronkan semua data BINAMARGA secara bertahap (50 data per batch).\n\n' +
        '⚠️ Proses bisa memakan waktu 10-20 menit tergantung jumlah data.\n' +
        '⚠️ Jangan tutup spreadsheet selama proses berjalan.\n\n' +
        'Lanjutkan?',
        ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
        return { success: false, message: 'Dibatalkan oleh user' };
    }
    
    const result = resyncAllToBinamargaBatch();
    ui.alert('✅ Hasil Resync', result.message, ui.ButtonSet.OK);
    return result;
}

function resyncAllToBackupBatchWithProgress() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
        '⚠️ Konfirmasi Resync Backup',
        'Proses ini akan menyinkronkan semua data BINAMARGA ke BACKUP secara bertahap (50 data per batch).\n\n' +
        '⚠️ Proses bisa memakan waktu 5-15 menit tergantung jumlah data.\n' +
        '⚠️ Jangan tutup spreadsheet selama proses berjalan.\n\n' +
        'Lanjutkan?',
        ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
        return { success: false, message: 'Dibatalkan oleh user' };
    }
    
    const result = resyncAllToBackup();
    ui.alert('✅ Hasil Resync Backup', result.message, ui.ButtonSet.OK);
    return result;
}

// ========================================================================
// ==================== EXPORT FUNCTIONS ==================================
// ========================================================================

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

// ========================================================================
// ==================== WEB APP ENTRY POINT ==============================
// ========================================================================

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
