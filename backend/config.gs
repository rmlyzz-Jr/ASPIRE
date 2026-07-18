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

// ==================== KOLOM SHEET LaporanAduan (0-index) ====================
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

function formatTanggalPendek(tanggalValue) {
    try {
        if (!tanggalValue) return '-';
        var tgl = new Date(tanggalValue);
        if (isNaN(tgl.getTime())) return '-';
        var months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Ags', 'Sep', 'Okt', 'Nov', 'Des'];
        return tgl.getDate() + ' ' + months[tgl.getMonth()] + ' ' + tgl.getFullYear();
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

function getSheetData(sheetName) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return null;
        if (sheet.getLastRow() <= 1) return [];
        return sheet.getDataRange().getValues();
    } catch (e) {
        console.error('Error getSheetData:', e);
        return null;
    }
}

function appendToSheet(sheetName, rowData) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            throw new Error('Sheet ' + sheetName + ' tidak ditemukan');
        }
        sheet.appendRow(rowData);
        return true;
    } catch (e) {
        console.error('Error appendToSheet:', e);
        return false;
    }
}

function updateSheetRow(sheetName, rowIndex, colIndex, value) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            throw new Error('Sheet ' + sheetName + ' tidak ditemukan');
        }
        sheet.getRange(rowIndex, colIndex).setValue(value);
        return true;
    } catch (e) {
        console.error('Error updateSheetRow:', e);
        return false;
    }
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
    if (keywords.length === 0) {
        return { title: "📋 ADUAN MASYARAKAT", icon: "📋", rawTitle: "ADUAN MASYARAKAT", keyword: null };
    }
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
