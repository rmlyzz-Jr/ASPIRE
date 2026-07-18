// ==================== WHATSAPP SENDER ====================

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

// ==================== IMAGE UPLOAD FUNCTIONS ====================

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
