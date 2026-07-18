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

// ==================== GET MASTER DATA ====================
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

// ==================== GET ALL REPORTS ====================
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

// ==================== SAVE LAPORAN ====================
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

// ==================== UPDATE LAPORAN DENGAN GAMBAR ====================
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

// ==================== UPDATE LAPORAN TANPA WA ====================
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

// ==================== UPDATE LAPORAN (LEGACY) ====================
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

// ==================== DELETE LAPORAN ====================
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

// ==================== UPDATE LAPORAN DENGAN STATUS CHECK ====================
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
