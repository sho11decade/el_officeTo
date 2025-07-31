const { contextBridge, ipcRenderer } = require('electron');

// レンダラープロセスに安全なAPIを公開
contextBridge.exposeInMainWorld('electronAPI', {
    // ファイル操作
    extractImages: (filePaths) => ipcRenderer.invoke('extract-images', filePaths),
    processFileBuffer: (fileData) => ipcRenderer.invoke('process-file-buffer', fileData),
    saveImage: (imageData, defaultName) => ipcRenderer.invoke('save-image', imageData, defaultName),
    saveAllImages: (images) => ipcRenderer.invoke('save-all-images', images),
    exportImagesZip: (images, sourceFileName) => ipcRenderer.invoke('export-images-zip', images, sourceFileName),
    showOpenDialog: () => ipcRenderer.invoke('show-open-dialog'),
    
    // イベントリスナー
    onFilesSelected: (callback) => {
        ipcRenderer.on('files-selected', (event, filePaths) => callback(filePaths));
    },
    
    // 進捗イベントリスナー
    onExtractionProgress: (callback) => {
        ipcRenderer.on('extraction-progress', (event, progressData) => callback(progressData));
    },
    
    // イベントリスナーの削除
    removeAllListeners: (channel) => {
        ipcRenderer.removeAllListeners(channel);
    },
    
    // 進捗リスナーの削除
    removeExtractionProgressListener: () => {
        ipcRenderer.removeAllListeners('extraction-progress');
    }
});
