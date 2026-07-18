// ==================== SINKRONISASI DATA ====================

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
        console.log(`⏭️ Dilewati (bukan Binamarga): ${kodeLaporan}`);
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
                    console.error(`❌ Error:`, e.message);
                }
            }
            
            if (batch < totalBatches - 1) {
                Utilities.sleep(PAUSE_MS);
            }
        }
        
        return {
            success: true,
            message: `✅ Resync selesai!\n📥 Data baru: ${totalNew}\n🔄 Data diupdate: ${totalUpdated}\n❌ Error: ${totalErrors}`,
            newData: totalNew,
            updatedData: totalUpdated,
            errors: totalErrors,
            totalData: totalData
        };
        
    } catch (e) {
        return { success: false, message: '❌ Error: ' + e.toString() };
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