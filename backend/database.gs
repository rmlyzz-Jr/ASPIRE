// ==================== DATABASE OPERATIONS ====================

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
                ['romly', 'romly123', 'Romly Wahyu', 'ADMIN', 'romlywa@gmail.com', '']
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

function getMasterDataWeb() {
    try {
        ensureSheetsExist();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const masters = { status: [], disposisi: [], sumber: [] };
        
        // Status
        let sheet = ss.getSheetByName(SHEET_MASTER_STATUS);
        if (sheet && sheet.getLastRow() > 1) {
            const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
            for (let i = 0; i < data.length; i++) {
                if (data[i][0] && data[i][0].toString().trim()) {
                    masters.status.push(data[i][0].toString().trim());
                }
            }
        }
        if (masters.status.length === 0) {
            masters.status = ['Menunggu', 'Proses', 'Selesai', 'Ditolak'];
        }
        
        // Disposisi
        sheet = ss.getSheetByName(SHEET_MASTER_DISPOSISI);
        if (sheet && sheet.getLastRow() > 1) {
            const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
            for (let i = 0; i < data.length; i++) {
                if (data[i][0] && data[i][0].toString().trim()) {
                    masters.disposisi.push(data[i][0].toString().trim());
                }
            }
        }
        if (masters.disposisi.length === 0) {
            masters.disposisi = ['Bidang Pelayanan', 'Bidang Teknis', 'BINAMARGA'];
        }
        if (!masters.disposisi.includes('BINAMARGA')) {
            masters.disposisi.push('BINAMARGA');
        }
        
        // Sumber
        sheet = ss.getSheetByName(SHEET_MASTER_SUMBER);
        if (sheet && sheet.getLastRow() > 1) {
            const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
            for (let i = 0; i < data.length; i++) {
                if (data[i][0] && data[i][0].toString().trim()) {
                    masters.sumber.push(data[i][0].toString().trim());
                }
            }
        }
        if (masters.sumber.length === 0) {
            masters.sumber = ['WhatsApp', 'Telepon', 'Email', 'Website'];
        }
        
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