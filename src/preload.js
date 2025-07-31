const { contextBridge, ipcRenderer } = require('electron');

// レンダラープロセスに安全なAPIを公開
contextBridge.exposeInMainWorld('electronAPI', {
    // ファイル操作
    extractImages: (filePaths) => ipcRenderer.invoke('extract-images', filePaths),
    saveImage: (imageData, defaultName) => ipcRenderer.invoke('save-image', imageData, defaultName),
    saveAllImages: (images) => ipcRenderer.invoke('save-all-images', images),
    showOpenDialog: () => ipcRenderer.invoke('show-open-dialog'),
    
    // イベントリスナー
    onFilesSelected: (callback) => {
        ipcRenderer.on('files-selected', (event, filePaths) => callback(filePaths));
    },
    
    // イベントリスナーの削除
    removeAllListeners: (channel) => {
        ipcRenderer.removeAllListeners(channel);
    }
});
