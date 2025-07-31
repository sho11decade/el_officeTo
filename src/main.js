const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const path = require('path');
const fs = require('fs');
const { extractImages } = require('./utils/imageExtractor');

let mainWindow;

// アプリケーションが準備完了になったときの処理
app.whenReady().then(() => {
    createWindow();
    
    // macOSでは、ウィンドウが全て閉じられてもアプリケーションは終了しない
    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

// 全てのウィンドウが閉じられたときの処理
app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

function createWindow() {
    // メインウィンドウを作成
    mainWindow = new BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            enableRemoteModule: false,
            preload: path.join(__dirname, 'preload.js')
        },
        icon: path.join(__dirname, '../assets/icon.png'),
        show: false // 初期化完了まで非表示
    });

    // HTMLファイルを読み込み
    mainWindow.loadFile(path.join(__dirname, 'renderer/index.html'));

    // ウィンドウが準備完了したら表示
    mainWindow.once('ready-to-show', () => {
        mainWindow.show();
    });

    // 開発モードでは開発者ツールを開く
    if (process.argv.includes('--dev')) {
        mainWindow.webContents.openDevTools();
    }

    createMenu();
}

function createMenu() {
    const template = [
        {
            label: 'ファイル',
            submenu: [
                {
                    label: 'ファイルを開く',
                    accelerator: 'CmdOrCtrl+O',
                    click: () => {
                        openFileDialog();
                    }
                },
                { type: 'separator' },
                {
                    label: '終了',
                    accelerator: process.platform === 'darwin' ? 'Cmd+Q' : 'Ctrl+Q',
                    click: () => {
                        app.quit();
                    }
                }
            ]
        },
        {
            label: '表示',
            submenu: [
                { role: 'reload', label: '再読み込み' },
                { role: 'forceReload', label: '強制再読み込み' },
                { role: 'toggleDevTools', label: '開発者ツール' },
                { type: 'separator' },
                { role: 'resetZoom', label: 'ズームリセット' },
                { role: 'zoomIn', label: 'ズームイン' },
                { role: 'zoomOut', label: 'ズームアウト' },
                { type: 'separator' },
                { role: 'togglefullscreen', label: 'フルスクリーン切替' }
            ]
        },
        {
            label: 'ヘルプ',
            submenu: [
                {
                    label: 'このアプリについて',
                    click: () => {
                        dialog.showMessageBox(mainWindow, {
                            type: 'info',
                            title: 'Office Image Extractor について',
                            message: 'Office Image Extractor v1.0.0',
                            detail: 'Microsoft Officeファイルから画像を抽出・表示・保存するアプリケーションです。'
                        });
                    }
                }
            ]
        }
    ];

    const menu = Menu.buildFromTemplate(template);
    Menu.setApplicationMenu(menu);
}

async function openFileDialog() {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openFile', 'multiSelections'],
        filters: [
            { name: 'Office Files', extensions: ['docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt'] },
            { name: 'Word Documents', extensions: ['docx', 'doc'] },
            { name: 'Excel Files', extensions: ['xlsx', 'xls'] },
            { name: 'PowerPoint Files', extensions: ['pptx', 'ppt'] },
            { name: 'All Files', extensions: ['*'] }
        ]
    });

    if (!result.canceled && result.filePaths.length > 0) {
        // ファイルが選択された場合、レンダラープロセスに通知
        mainWindow.webContents.send('files-selected', result.filePaths);
    }
}

// レンダラープロセスからのIPC通信を処理
ipcMain.handle('extract-images', async (event, filePaths) => {
    try {
        const results = [];
        for (const filePath of filePaths) {
            const images = await extractImages(filePath);
            results.push({
                filePath,
                fileName: path.basename(filePath),
                images
            });
        }
        return { success: true, data: results };
    } catch (error) {
        console.error('Image extraction error:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('save-image', async (event, imageData, defaultName) => {
    try {
        const result = await dialog.showSaveDialog(mainWindow, {
            defaultPath: defaultName,
            filters: [
                { name: 'PNG Images', extensions: ['png'] },
                { name: 'JPEG Images', extensions: ['jpg', 'jpeg'] },
                { name: 'All Files', extensions: ['*'] }
            ]
        });

        if (!result.canceled && result.filePath) {
            // Base64データをBufferに変換
            const base64Data = imageData.replace(/^data:image\/\w+;base64,/, '');
            const buffer = Buffer.from(base64Data, 'base64');
            
            fs.writeFileSync(result.filePath, buffer);
            return { success: true, filePath: result.filePath };
        }
        
        return { success: false, error: 'Save cancelled' };
    } catch (error) {
        console.error('Save image error:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('save-all-images', async (event, images) => {
    try {
        const result = await dialog.showOpenDialog(mainWindow, {
            properties: ['openDirectory']
        });

        if (!result.canceled && result.filePaths.length > 0) {
            const outputDir = result.filePaths[0];
            const savedFiles = [];

            for (let i = 0; i < images.length; i++) {
                const image = images[i];
                const fileName = `image_${i + 1}.${image.format || 'png'}`;
                const filePath = path.join(outputDir, fileName);
                
                const base64Data = image.data.replace(/^data:image\/\w+;base64,/, '');
                const buffer = Buffer.from(base64Data, 'base64');
                
                fs.writeFileSync(filePath, buffer);
                savedFiles.push(filePath);
            }

            return { success: true, savedFiles };
        }
        
        return { success: false, error: 'Save cancelled' };
    } catch (error) {
        console.error('Save all images error:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('show-open-dialog', async () => {
    return await openFileDialog();
});
