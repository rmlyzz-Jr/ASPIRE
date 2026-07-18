// ==================== KONFIGURASI ====================
const SPREADSHEET_ID = '1p6KsVDerY2eogOI9JwF4Ja2q2M-0yADL5IZ4OAF8Jos';
const SPREADSHEET_BACKUP_ID = '1F6FIhJXnKHbcj7aMwrtNGRnnWLwptNlzZtot-Qd97Jk';
const WASENDER_API_KEY = '351e732e407183e784178e3861a44ae26896b10c41811fb5ffd4fd6bb4570168';
const WASENDER_API_URL = 'https://www.wasenderapi.com/api/send-message';
const WA_GROUP_ID = '120363283721432243@g.us';
const DEFAULT_MENTION = "6285741448845@s.whatsapp.net";

// Sheet names
const SHEET_LAPORAN = 'LaporanAduan';
const SHEET_BINAMARGA = 'BINAMARGA';
const SHEET_MASTER_STATUS = 'master status';
const SHEET_MASTER_DISPOSISI = 'master disposisi';
const SHEET_MASTER_SUMBER = 'mastersumber';
const SHEET_WA_CONFIG = 'MASTER_WA_CONFIG';
const SHEET_USER_LOGIN = 'USER_LOGIN';
const SHEET_KEYWORD = 'MASTER_KEYWORD';

// Column indexes (0-index)
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

// Disposisi Binamarga
const DISPOSISI_BINAMARGA_LIST = [
    'BINAMARGA', 'Binamarga', 'binamarga',
    'BINA MARGA', 'Bina Marga', 'bina marga', '1108'
];

function isDisposisiBinamarga(disposisiValue) {
    if (!disposisiValue) return false;
    const cleanValue = disposisiValue.toString().toLowerCase().trim();
    for (let i = 0; i < DISPOSISI_BINAMARGA_LIST.length; i++) {
        if (cleanValue === DISPOSISI_BINAMARGA_LIST[i].toLowerCase().trim()) return true;
    }
    return cleanValue.includes('binamarga') || 
           cleanValue.includes('bina marga') || 
           cleanValue.includes('1108');
}

function formatTanggalHelper(tanggalValue) {
    try {
        if (!tanggalValue) return '-';
        const tgl = new Date(tanggalValue);
        if (isNaN(tgl.getTime())) return '-';
        return `${tgl.getFullYear()}-${String(tgl.getMonth()+1).padStart(2,'0')}-${String(tgl.getDate()).padStart(2,'0')}`;
    } catch(e) { return '-'; }
}

function formatTanggalIndonesia(tanggalValue) {
    try {
        if (!tanggalValue) return '-';
        const tgl = new Date(tanggalValue);
        if (isNaN(tgl.getTime())) return '-';
        const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
        const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 
                       'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        return `${days[tgl.getDay()]}, ${tgl.getDate()} ${months[tgl.getMonth()]} ${tgl.getFullYear()}`;
    } catch(e) { return '-'; }
}

function generateKodeLaporan() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_LAPORAN);
        const date = new Date();
        const todayStr = `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
        let count = 1;
        if (sheet && sheet.getLastRow() > 1) {
            const data = sheet.getRange(2, COL_TANGGAL + 1, sheet.getLastRow() - 1, 1).getValues();
            for (let i = 0; i < data.length; i++) {
                if (data[i][0]) {
                    const cellDate = new Date(data[i][0]);
                    if (!isNaN(cellDate.getTime())) {
                        const cellDateStr = `${cellDate.getFullYear()}-${String(cellDate.getMonth()+1).padStart(2,'0')}-${String(cellDate.getDate()).padStart(2,'0')}`;
                        if (cellDateStr === todayStr) count++;
                    }
                }
            }
        }
        return `SIPU/${todayStr.replace(/-/g, '')}/${String(count).padStart(3, '0')}`;
    } catch(error) {
        const date = new Date();
        return `SIPU/${date.getFullYear()}${String(date.getMonth()+1).padStart(2,'0')}${String(date.getDate()).padStart(2,'0')}/001`;
    }
}