// DOM要素の取得
const elements = {
    openFileBtn: document.getElementById('openFileBtn'),
    saveAllBtn: document.getElementById('saveAllBtn'),
    fileList: document.getElementById('fileList'),
    fileStats: document.getElementById('fileStats'),
    fileCount: document.getElementById('fileCount'),
    imageCount: document.getElementById('imageCount'),
    imageContainer: document.getElementById('imageContainer'),
    loadingState: document.getElementById('loadingState'),
    gridViewBtn: document.getElementById('gridViewBtn'),
    listViewBtn: document.getElementById('listViewBtn'),
    statusText: document.getElementById('statusText'),
    imageModal: document.getElementById('imageModal'),
    modalBackdrop: document.getElementById('modalBackdrop'),
    modalClose: document.getElementById('modalClose'),
    modalCloseBtn: document.getElementById('modalCloseBtn'),
    modalTitle: document.getElementById('modalTitle'),
    modalImage: document.getElementById('modalImage'),
    modalImageName: document.getElementById('modalImageName'),
    modalImageFormat: document.getElementById('modalImageFormat'),
    modalImageDimensions: document.getElementById('modalImageDimensions'),
    modalImageSize: document.getElementById('modalImageSize'),
    modalSaveBtn: document.getElementById('modalSaveBtn'),
    dragOverlay: document.getElementById('dragOverlay')
};

// アプリケーションの状態
let appState = {
    files: [],
    allImages: [],
    currentView: 'grid',
    selectedImage: null
};

// 初期化
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
    updateStatus('準備完了');
});

// イベントリスナーの初期化
function initializeEventListeners() {
    // ファイル選択ボタン
    elements.openFileBtn.addEventListener('click', () => {
        window.electronAPI.showOpenDialog();
    });

    // 全て保存ボタン
    elements.saveAllBtn.addEventListener('click', () => {
        saveAllImages();
    });

    // ビュー切り替えボタン
    elements.gridViewBtn.addEventListener('click', () => {
        setView('grid');
    });

    elements.listViewBtn.addEventListener('click', () => {
        setView('list');
    });

    // モーダル関連
    elements.modalClose.addEventListener('click', closeModal);
    elements.modalCloseBtn.addEventListener('click', closeModal);
    elements.modalBackdrop.addEventListener('click', closeModal);
    elements.modalSaveBtn.addEventListener('click', saveCurrentImage);

    // ESCキーでモーダルを閉じる
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && elements.imageModal.style.display !== 'none') {
            closeModal();
        }
    });

    // ドラッグ&ドロップ
    document.addEventListener('dragover', (e) => {
        e.preventDefault();
        elements.dragOverlay.style.display = 'flex';
    });

    document.addEventListener('dragleave', (e) => {
        if (e.clientX === 0 && e.clientY === 0) {
            elements.dragOverlay.style.display = 'none';
        }
    });

    document.addEventListener('drop', (e) => {
        e.preventDefault();
        elements.dragOverlay.style.display = 'none';
        
        const files = Array.from(e.dataTransfer.files);
        const filePaths = files.map(file => file.path).filter(path => isOfficeFile(path));
        
        if (filePaths.length > 0) {
            processFiles(filePaths);
        } else {
            updateStatus('対応していないファイル形式です');
        }
    });

    // メインプロセスからのファイル選択イベント
    window.electronAPI.onFilesSelected((filePaths) => {
        processFiles(filePaths);
    });
}

// Officeファイルかどうかを判定
function isOfficeFile(filePath) {
    const supportedExtensions = ['.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt'];
    const ext = filePath.toLowerCase().substring(filePath.lastIndexOf('.'));
    return supportedExtensions.includes(ext);
}

// ファイル処理
async function processFiles(filePaths) {
    try {
        updateStatus('ファイルを処理中...');
        showLoading(true);

        const result = await window.electronAPI.extractImages(filePaths);
        
        if (result.success) {
            appState.files = result.data;
            appState.allImages = [];
            
            // すべての画像を収集
            result.data.forEach(file => {
                file.images.forEach(image => {
                    image.sourceFile = file.fileName;
                    appState.allImages.push(image);
                });
            });

            updateFileList();
            updateImageDisplay();
            updateStats();
            updateStatus(`${appState.files.length}個のファイルから${appState.allImages.length}個の画像を抽出しました`);
        } else {
            updateStatus(`エラー: ${result.error}`);
        }
    } catch (error) {
        console.error('Error processing files:', error);
        updateStatus('ファイルの処理中にエラーが発生しました');
    } finally {
        showLoading(false);
    }
}

// ファイル一覧の更新
function updateFileList() {
    if (appState.files.length === 0) {
        elements.fileList.innerHTML = `
            <div class="empty-state">
                <p>ファイルが選択されていません</p>
                <p class="small">「ファイルを開く」ボタンまたはドラッグ&ドロップでファイルを追加してください</p>
            </div>
        `;
        elements.fileStats.style.display = 'none';
        return;
    }

    elements.fileList.innerHTML = '';
    appState.files.forEach((file, index) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            <div class="file-item-name">${file.fileName}</div>
            <div class="file-item-info">${file.images.length}個の画像</div>
        `;
        
        fileItem.addEventListener('click', () => {
            // ファイル選択時の処理（必要に応じて実装）
            document.querySelectorAll('.file-item').forEach(item => item.classList.remove('selected'));
            fileItem.classList.add('selected');
        });
        
        elements.fileList.appendChild(fileItem);
    });

    elements.fileStats.style.display = 'block';
}

// 画像表示の更新
function updateImageDisplay() {
    if (appState.allImages.length === 0) {
        elements.imageContainer.innerHTML = `
            <div class="empty-state">
                <p>画像がありません</p>
                <p class="small">Officeファイルを読み込んで画像を抽出してください</p>
            </div>
        `;
        elements.saveAllBtn.disabled = true;
        return;
    }

    const containerClass = appState.currentView === 'grid' ? 'image-grid' : 'image-list';
    elements.imageContainer.innerHTML = `<div class="${containerClass}" id="imageDisplayArea"></div>`;
    
    const displayArea = document.getElementById('imageDisplayArea');
    
    appState.allImages.forEach((image, index) => {
        const imageItem = document.createElement('div');
        imageItem.className = 'image-item';
        
        const thumbnailSrc = image.thumbnail || image.data;
        const dimensions = image.width && image.height ? `${image.width} × ${image.height}` : '不明';
        const fileSize = formatFileSize(image.size);
        
        imageItem.innerHTML = `
            <img src="${thumbnailSrc}" alt="${image.name}" class="image-thumbnail">
            <div class="image-info">
                <div class="image-name" title="${image.name}">${image.name}</div>
                <div class="image-details">${dimensions} • ${fileSize}</div>
            </div>
        `;
        
        imageItem.addEventListener('click', () => {
            showImageModal(image);
        });
        
        displayArea.appendChild(imageItem);
    });

    elements.saveAllBtn.disabled = false;
}

// 統計情報の更新
function updateStats() {
    elements.fileCount.textContent = appState.files.length;
    elements.imageCount.textContent = appState.allImages.length;
}

// ビューの切り替え
function setView(view) {
    appState.currentView = view;
    
    elements.gridViewBtn.classList.toggle('active', view === 'grid');
    elements.listViewBtn.classList.toggle('active', view === 'list');
    
    updateImageDisplay();
}

// 画像モーダルの表示
function showImageModal(image) {
    appState.selectedImage = image;
    
    elements.modalTitle.textContent = image.name;
    elements.modalImage.src = image.data;
    elements.modalImageName.textContent = image.name;
    elements.modalImageFormat.textContent = image.format.toUpperCase();
    elements.modalImageDimensions.textContent = image.width && image.height ? 
        `${image.width} × ${image.height} px` : '不明';
    elements.modalImageSize.textContent = formatFileSize(image.size);
    
    elements.imageModal.style.display = 'block';
}

// モーダルを閉じる
function closeModal() {
    elements.imageModal.style.display = 'none';
    appState.selectedImage = null;
}

// 現在の画像を保存
async function saveCurrentImage() {
    if (!appState.selectedImage) return;
    
    try {
        updateStatus('画像を保存中...');
        const result = await window.electronAPI.saveImage(
            appState.selectedImage.data,
            appState.selectedImage.name
        );
        
        if (result.success) {
            updateStatus(`画像を保存しました: ${result.filePath}`);
            closeModal();
        } else {
            updateStatus('画像の保存に失敗しました');
        }
    } catch (error) {
        console.error('Error saving image:', error);
        updateStatus('画像の保存中にエラーが発生しました');
    }
}

// すべての画像を保存
async function saveAllImages() {
    if (appState.allImages.length === 0) return;
    
    try {
        updateStatus('すべての画像を保存中...');
        const result = await window.electronAPI.saveAllImages(appState.allImages);
        
        if (result.success) {
            updateStatus(`${result.savedFiles.length}個の画像を保存しました`);
        } else {
            updateStatus('画像の保存に失敗しました');
        }
    } catch (error) {
        console.error('Error saving all images:', error);
        updateStatus('画像の保存中にエラーが発生しました');
    }
}

// ローディング状態の表示/非表示
function showLoading(show) {
    elements.loadingState.style.display = show ? 'flex' : 'none';
    elements.imageContainer.style.display = show ? 'none' : 'block';
}

// ステータスの更新
function updateStatus(message) {
    elements.statusText.textContent = message;
    console.log('Status:', message);
}

// ファイルサイズのフォーマット
function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

// エラーハンドリング
window.addEventListener('error', (event) => {
    console.error('Unhandled error:', event.error);
    updateStatus('予期しないエラーが発生しました');
});

window.addEventListener('unhandledrejection', (event) => {
    console.error('Unhandled promise rejection:', event.reason);
    updateStatus('非同期処理でエラーが発生しました');
});
