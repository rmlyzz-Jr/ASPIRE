// ==================== USER AUTHENTICATION ====================

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

// ==================== USER MANAGEMENT (ADMIN) ====================

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