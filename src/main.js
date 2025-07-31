const { app, BrowserWindow, ipcMain, dialog, Menu, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const yazl = require('yazl');
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
    // メインウィンドウを作成（セキュリティ強化）
    mainWindow = new BrowserWindow({
        width: 1200,
        height: 800,
        minWidth: 800,
        minHeight: 600,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            enableRemoteModule: false,
            allowRunningInsecureContent: false,
            experimentalFeatures: false,
            webSecurity: true,
            preload: path.join(__dirname, 'preload.js'),
            sandbox: false // IPCを使用するためfalse
        },
        icon: path.join(__dirname, '../assets/icon.png'),
        show: false // 初期化完了まで非表示
    });

    // セキュリティ強化
    mainWindow.webContents.on('will-navigate', (event, navigationUrl) => {
        // 外部URLへのナビゲーションを防止
        const parsedUrl = new URL(navigationUrl);
        if (parsedUrl.origin !== 'file://') {
            event.preventDefault();
        }
    });

    mainWindow.webContents.setWindowOpenHandler(() => {
        // 新しいウィンドウの開放を防止
        return { action: 'deny' };
    });

    // HTMLファイルを読み込み
    mainWindow.loadFile(path.join(__dirname, 'renderer/index.html'));

    // ウィンドウが準備完了したら表示
    mainWindow.once('ready-to-show', () => {
        mainWindow.show();
    });

    // 開発モードでは開発者ツールを開く
    if (process.argv.includes('--dev') || process.env.NODE_ENV === 'development') {
        mainWindow.webContents.openDevTools();
    } else {
        // プロダクションモードでは開発者ツールを無効化
        mainWindow.webContents.on('before-input-event', (event, input) => {
            // Ctrl+Shift+I, F12, Ctrl+Shift+J を無効化
            if ((input.control && input.shift && input.key.toLowerCase() === 'i') ||
                input.key.toLowerCase() === 'f12' ||
                (input.control && input.shift && input.key.toLowerCase() === 'j')) {
                event.preventDefault();
            }
        });
        
        // 右クリックメニューを無効化
        mainWindow.webContents.on('context-menu', (event) => {
            event.preventDefault();
        });
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
                    label: 'GitHubリポジトリ',
                    click: () => {
                        shell.openExternal('https://github.com/sho11decade/el_officeTo');
                    }
                },
                { type: 'separator' },
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

// レンダラープロセスからのIPC通信を処理（最適化）
ipcMain.handle('extract-images', async (event, filePaths) => {
    try {
        if (!Array.isArray(filePaths) || filePaths.length === 0) {
            throw new Error('有効なファイルパスが指定されていません');
        }
        
        // 同時処理数を制限してメモリ使用量を抑制
        const MAX_CONCURRENT = 3;
        const results = [];
        
        for (let i = 0; i < filePaths.length; i += MAX_CONCURRENT) {
            const batch = filePaths.slice(i, i + MAX_CONCURRENT);
            const batchPromises = batch.map(async (filePath) => {
                try {
                    // ファイルパスのログ出力（デバッグ用）
                    console.log(`Processing file: "${filePath}"`);
                    console.log(`File path length: ${filePath.length}`);
                    console.log(`File path bytes:`, Buffer.from(filePath, 'utf8'));
                    
                    // ファイル存在チェック（文字エンコーディング対応強化版）
                    let actualFilePath = filePath;
                    let fileExists = false;
                    let fileStats = null;
                    
                    // 複数の方法でファイル存在をチェック
                    const pathVariants = [
                        filePath,
                        path.resolve(filePath),
                        path.normalize(filePath),
                        // Windows特有の問題に対応
                        filePath.replace(/\//g, '\\'),
                        filePath.replace(/\\/g, '/')
                    ];
                    
                    for (const pathVariant of pathVariants) {
                        try {
                            fileStats = await fs.promises.stat(pathVariant);
                            if (fileStats.isFile()) {
                                fileExists = true;
                                actualFilePath = pathVariant;
                                console.log(`File found with path variant: "${pathVariant}"`);
                                break;
                            }
                        } catch (statError) {
                            // 次のバリアントを試行
                            continue;
                        }
                    }
                    
                    if (!fileExists) {
                        const errorMsg = `ファイルが見つかりません: ${path.basename(filePath)}`;
                        console.error(errorMsg);
                        console.error('Tried path variants:', pathVariants);
                        throw new Error(errorMsg);
                    }
                    
                    console.log(`File validation successful: ${path.basename(actualFilePath)} (${fileStats.size} bytes)`);
                    
                    const images = await extractImages(actualFilePath, {
                        compress: true,
                        maxDimension: 4096,
                        onProgress: (processed, total) => {
                            // 進捗をレンダラーに送信
                            mainWindow.webContents.send('extraction-progress', {
                                file: path.basename(filePath),
                                processed,
                                total
                            });
                        }
                    });
                    
                    
                    return {
                        filePath: actualFilePath,
                        fileName: path.basename(actualFilePath),
                        images: images || [],
                        success: true
                    };
                } catch (error) {
                    const errorMsg = `ファイル処理エラー: ${path.basename(filePath)} - ${error.message}`;
                    console.error(errorMsg, error);
                    return {
                        filePath,
                        fileName: path.basename(filePath),
                        images: [],
                        success: false,
                        error: error.message
                    };
                }
            });
            
            const batchResults = await Promise.all(batchPromises);
            results.push(...batchResults);
            
            // ガベージコレクションを促進
            if (global.gc) {
                global.gc();
            }
        }
        
        return { success: true, data: results };
    } catch (error) {
        console.error('Image extraction error:', error);
        return { success: false, error: error.message };
    }
});

// ファイルバッファ処理のハンドラー（ドラッグ&ドロップ用）
ipcMain.handle('process-file-buffer', async (event, fileData) => {
    try {
        console.log('ファイルバッファ処理開始:', fileData.name);
        
        if (!fileData || !fileData.buffer) {
            return { success: false, error: 'ファイルデータが無効です' };
        }

        // 一時ファイルを作成してバッファを書き込み
        const tempDir = os.tmpdir();
        const tempFileName = `temp_${Date.now()}_${fileData.name.replace(/[^\w\-_.]/g, '_')}`;
        const tempFilePath = path.join(tempDir, tempFileName);
        
        try {
            // ArrayBuffer を Buffer に変換
            const buffer = Buffer.from(fileData.buffer);
            await fs.promises.writeFile(tempFilePath, buffer);
            
            console.log(`一時ファイル作成: ${tempFilePath} (${buffer.length} bytes)`);
            
            // 画像抽出処理
            const images = await extractImages(tempFilePath, {
                compress: true,
                maxDimension: 4096,
                onProgress: (processed, total) => {
                    event.sender.send('extraction-progress', {
                        file: fileData.name,
                        processed,
                        total
                    });
                }
            });
            
            console.log(`バッファから ${images?.length || 0}個の画像を抽出`);
            
            return {
                success: true,
                data: {
                    fileName: fileData.name,
                    images: images || []
                }
            };
            
        } finally {
            // 一時ファイルを削除
            try {
                await fs.promises.unlink(tempFilePath);
                console.log(`一時ファイル削除: ${tempFilePath}`);
            } catch (unlinkError) {
                console.warn(`一時ファイル削除失敗: ${tempFilePath}`, unlinkError);
            }
        }
        
    } catch (error) {
        console.error('ファイルバッファ処理エラー:', error);
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

ipcMain.handle('export-images-zip', async (event, images, sourceFileName) => {
    try {
        const defaultFileName = `${sourceFileName || 'extracted_images'}_${new Date().toISOString().slice(0, 10)}.zip`;
        
        const result = await dialog.showSaveDialog(mainWindow, {
            defaultPath: defaultFileName,
            filters: [
                { name: 'ZIP Archives', extensions: ['zip'] },
                { name: 'All Files', extensions: ['*'] }
            ]
        });

        if (!result.canceled && result.filePath) {
            const zipFile = new yazl.ZipFile();
            const exportedCount = await createZipFile(zipFile, images, result.filePath);
            
            return { 
                success: true, 
                filePath: result.filePath,
                exportedCount: exportedCount
            };
        }
        
        return { success: false, error: 'Export cancelled' };
    } catch (error) {
        console.error('Export ZIP error:', error);
        return { success: false, error: error.message };
    }
});

// ZIPファイル作成のヘルパー関数
function createZipFile(zipFile, images, outputPath) {
    return new Promise((resolve, reject) => {
        let addedCount = 0;
        let processedCount = 0;
        
        // 同名ファイルを避けるためのカウンター
        const fileNameCounts = {};
        
        images.forEach((image, index) => {
            try {
                // Base64データをBufferに変換
                const base64Data = image.data.replace(/^data:image\/\w+;base64,/, '');
                const buffer = Buffer.from(base64Data, 'base64');
                
                // ファイル名の重複を避ける
                let fileName = image.name || `image_${index + 1}.${image.format || 'png'}`;
                const baseName = path.parse(fileName).name;
                const extension = path.parse(fileName).ext || `.${image.format || 'png'}`;
                
                if (fileNameCounts[fileName]) {
                    fileNameCounts[fileName]++;
                    fileName = `${baseName}_${fileNameCounts[fileName]}${extension}`;
                } else {
                    fileNameCounts[fileName] = 1;
                }
                
                // ZIPファイルにエントリを追加
                zipFile.addBuffer(buffer, fileName, {
                    mtime: new Date(),
                    mode: parseInt('0644', 8) // ファイル権限
                });
                
                addedCount++;
            } catch (error) {
                console.error(`Error processing image ${index}:`, error);
            }
            
            processedCount++;
            
            // 全ての画像を処理完了したらZIPファイルを終了
            if (processedCount === images.length) {
                zipFile.end();
            }
        });
        
        // 空の場合の処理
        if (images.length === 0) {
            zipFile.end();
        }
        
        // ZIPファイルの出力ストリーム作成
        const outputStream = fs.createWriteStream(outputPath);
        
        zipFile.outputStream.pipe(outputStream);
        
        zipFile.outputStream.on('end', () => {
            resolve(addedCount);
        });
        
        zipFile.outputStream.on('error', (error) => {
            reject(error);
        });
        
        outputStream.on('error', (error) => {
            reject(error);
        });
    });
}
