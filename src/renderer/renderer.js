// DOM要素の取得（キャッシュ化）
const elements = (() => {
    const cache = {};
    const getElement = (id) => {
        if (!cache[id]) {
            cache[id] = document.getElementById(id);
        }
        return cache[id];
    };
    
    return {
        get openFileBtn() { return getElement('openFileBtn'); },
        get saveAllBtn() { return getElement('saveAllBtn'); },
        get exportZipBtn() { return getElement('exportZipBtn'); },
        get fileList() { return getElement('fileList'); },
        get fileStats() { return getElement('fileStats'); },
        get fileCount() { return getElement('fileCount'); },
        get imageCount() { return getElement('imageCount'); },
        get imageContainer() { return getElement('imageContainer'); },
        get loadingState() { return getElement('loadingState'); },
        get gridViewBtn() { return getElement('gridViewBtn'); },
        get listViewBtn() { return getElement('listViewBtn'); },
        get statusText() { return getElement('statusText'); },
        get imageModal() { return getElement('imageModal'); },
        get modalBackdrop() { return getElement('modalBackdrop'); },
        get modalClose() { return getElement('modalClose'); },
        get modalCloseBtn() { return getElement('modalCloseBtn'); },
        get modalTitle() { return getElement('modalTitle'); },
        get modalImage() { return getElement('modalImage'); },
        get modalImageName() { return getElement('modalImageName'); },
        get modalImageFormat() { return getElement('modalImageFormat'); },
        get modalImageDimensions() { return getElement('modalImageDimensions'); },
        get modalImageSize() { return getElement('modalImageSize'); },
        get modalSaveBtn() { return getElement('modalSaveBtn'); },
        get modalExportZipBtn() { return getElement('modalExportZipBtn'); },
        get dragOverlay() { return getElement('dragOverlay'); }
    };
})();

// パフォーマンス設定
const PERFORMANCE_CONFIG = {
    VIRTUAL_SCROLL_THRESHOLD: 50, // 50枚以上で仮想スクロール
    LAZY_LOAD_THRESHOLD: 10, // 10枚以上で遅延読み込み
    BATCH_RENDER_SIZE: 20, // バッチレンダリングサイズ
    DEBOUNCE_DELAY: 250 // デバウンス遅延（ミリ秒）
};

// アプリケーションの状態（最適化）
let appState = {
    files: [],
    allImages: [],
    visibleImages: [], // 表示中の画像
    currentView: 'grid',
    selectedImage: null,
    isLoading: false,
    renderQueue: [] // レンダリングキュー
};

// 初期化
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
    initializePerformanceOptimizations();
    updateStatus('準備完了');
});

// パフォーマンス最適化の初期化
function initializePerformanceOptimizations() {
    // Intersection Observer for lazy loading
    if (window.IntersectionObserver) {
        window.imageObserver = new IntersectionObserver(
            debounce(handleImageIntersection, 100),
            { rootMargin: '50px', threshold: 0.1 }
        );
    }
    
    // Virtual scrolling setup
    elements.imageContainer.addEventListener('scroll', 
        throttle(handleVirtualScroll, 16) // 60fps
    );
    
    // Memory cleanup on visibility change
    document.addEventListener('visibilitychange', handleVisibilityChange);
}

// デバウンス関数
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// スロットル関数
function throttle(func, limit) {
    let inThrottle;
    return function() {
        const args = arguments;
        const context = this;
        if (!inThrottle) {
            func.apply(context, args);
            inThrottle = true;
            setTimeout(() => inThrottle = false, limit);
        }
    };
}

// 仮想スクロール処理
function handleVirtualScroll() {
    if (appState.allImages.length < PERFORMANCE_CONFIG.VIRTUAL_SCROLL_THRESHOLD) {
        return; // 閾値以下では通常表示
    }
    
    // 仮想スクロールの実装（簡易版）
    const container = elements.imageContainer;
    const scrollTop = container.scrollTop;
    const containerHeight = container.clientHeight;
    
    // 表示領域の計算
    const itemHeight = 200; // 推定アイテム高さ
    const startIndex = Math.floor(scrollTop / itemHeight);
    const endIndex = Math.min(
        startIndex + Math.ceil(containerHeight / itemHeight) + 2,
        appState.allImages.length
    );
    
    // 表示範囲の更新
    appState.visibleImages = appState.allImages.slice(startIndex, endIndex);
    console.log(`Virtual scroll: showing ${startIndex}-${endIndex} of ${appState.allImages.length}`);
}

// Intersection Observer処理
function handleImageIntersection(entries) {
    entries.forEach(entry => {
        if (entry.isIntersecting) {
            const img = entry.target;
            const dataSrc = img.getAttribute('data-src');
            
            if (dataSrc && !img.src) {
                img.src = dataSrc;
                img.removeAttribute('data-src');
                window.imageObserver?.unobserve(img);
            }
        }
    });
}

// ページ非表示時のメモリクリーンアップ
function handleVisibilityChange() {
    if (document.hidden) {
        // ページが非表示になった場合、メモリ使用量を削減
        const images = document.querySelectorAll('.image-thumbnail');
        images.forEach(img => {
            if (img.src && img.src.startsWith('data:')) {
                img.setAttribute('data-src', img.src);
                img.src = '';
            }
        });
        
        // ガベージコレクションを促進
        if (window.gc) {
            window.gc();
        }
        
        console.log('Memory cleanup performed due to page visibility change');
    }
}

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

    // ZIPエクスポートボタン
    elements.exportZipBtn.addEventListener('click', () => {
        exportAllImagesAsZip();
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
    elements.modalExportZipBtn.addEventListener('click', exportCurrentImageAsZip);

    // ESCキーでモーダルを閉じる
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && elements.imageModal.style.display !== 'none') {
            closeModal();
        }
    });

    // ドラッグ&ドロップ（エラー修正版）
    let dragCounter = 0;
    
    document.addEventListener('dragenter', (e) => {
        e.preventDefault();
        dragCounter++;
        elements.dragOverlay.style.display = 'flex';
    });

    document.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'copy';
    });

    document.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dragCounter--;
        if (dragCounter === 0) {
            elements.dragOverlay.style.display = 'none';
        }
    });

    document.addEventListener('drop', async (e) => {
        e.preventDefault();
        dragCounter = 0;
        elements.dragOverlay.style.display = 'none';
        
        try {
            const files = Array.from(e.dataTransfer.files);
            console.log('Dropped files raw:', files.map(f => ({ 
                name: f.name, 
                path: f.path, 
                type: f.type,
                size: f.size,
                lastModified: f.lastModified
            })));
            
            if (files.length === 0) {
                updateStatus('ファイルがドロップされませんでした');
                return;
            }
            
            updateStatus('ドロップされたファイルを処理中...');
            
            const officeFiles = files.filter(file => isOfficeFile(file.name));
            
            if (officeFiles.length === 0) {
                const fileNames = files.map(f => f.name).join(', ');
                updateStatus(`対応していないファイル形式です: ${fileNames}`);
                return;
            }
            
            // 複数の方法でファイルパスを取得
            const validFilePaths = [];
            const failedFiles = [];
            
            for (const file of officeFiles) {
                let filePath = null;
                
                // 方法1: file.path を直接使用
                if (file.path && file.path.trim() !== '') {
                    filePath = file.path;
                    console.log(`Method 1 - Using file.path: ${filePath}`);
                }
                
                // 方法2: File API を使用してバイナリデータから処理
                else if (file instanceof File) {
                    try {
                        console.log(`Method 2 - Using File API for: ${file.name}`);
                        
                        // FileReader を使用してファイルデータを読み取り
                        const arrayBuffer = await readFileAsArrayBuffer(file);
                        
                        if (arrayBuffer && arrayBuffer.byteLength > 0) {
                            // バイナリデータを直接メインプロセスに送信
                            const result = await window.electronAPI.processFileBuffer({
                                name: file.name,
                                buffer: arrayBuffer,
                                size: file.size,
                                type: file.type,
                                lastModified: file.lastModified
                            });
                            
                            if (result && result.success) {
                                console.log(`File buffer processed successfully: ${file.name}`);
                                
                                // 成功した場合、画像データを直接処理
                                if (result.data && result.data.images) {
                                    const fileData = {
                                        fileName: file.name,
                                        success: true,
                                        images: result.data.images
                                    };
                                    
                                    // 既存の処理フローに統合
                                    appState.files = [fileData];
                                    appState.allImages = [];
                                    
                                    fileData.images.forEach(image => {
                                        image.sourceFile = file.name;
                                        appState.allImages.push(image);
                                    });
                                    
                                    updateFileList();
                                    updateImageDisplay();
                                    updateStats();
                                    updateStatus(`${file.name}から${result.data.images.length}個の画像を抽出しました`);
                                    return; // 成功したので処理終了
                                }
                            }
                        }
                    } catch (bufferError) {
                        console.error(`File buffer processing failed for ${file.name}:`, bufferError);
                    }
                }
                
                if (filePath) {
                    validFilePaths.push(filePath);
                } else {
                    failedFiles.push(file.name);
                    console.warn(`Could not get path for file: ${file.name}`);
                }
            }
            
            // 従来のパス方式で処理可能なファイルがある場合
            if (validFilePaths.length > 0) {
                console.log('Processing files with paths:', validFilePaths);
                await processFiles(validFilePaths);
                
                if (failedFiles.length > 0) {
                    setTimeout(() => {
                        updateStatus(`一部のファイル (${failedFiles.join(', ')}) は処理できませんでした。「ファイルを開く」ボタンを使用してください。`);
                    }, 2000);
                }
            } else if (failedFiles.length > 0) {
                updateStatus(`ファイルパスを取得できませんでした: ${failedFiles.join(', ')}。「ファイルを開く」ボタンを使用してファイルを選択してください。`);
            }
            
        } catch (error) {
            console.error('ドラッグ&ドロップエラー:', error);
            updateStatus('ファイルの読み込み中にエラーが発生しました。「ファイルを開く」ボタンを使用してください。');
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

// FileReader を使用してファイルを ArrayBuffer として読み取り
function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (event) => {
            resolve(event.target.result);
        };
        
        reader.onerror = (error) => {
            console.error('FileReader error:', error);
            reject(error);
        };
        
        reader.readAsArrayBuffer(file);
    });
}

// ファイル処理（エラーハンドリング強化版）
async function processFiles(filePaths) {
    if (!filePaths || filePaths.length === 0) {
        updateStatus('処理するファイルが選択されていません');
        return;
    }

    // ファイルパスの前処理（日本語ファイル名対応）
    const processedFilePaths = filePaths.map(filePath => {
        try {
            // パスの正規化
            const normalizedPath = filePath.replace(/\\/g, '/').replace(/\/+/g, '/');
            console.log(`Original: ${filePath} -> Normalized: ${normalizedPath}`);
            return filePath; // 元のパスを使用（Electronでは正しく処理される）
        } catch (error) {
            console.error(`Path normalization error for ${filePath}:`, error);
            return filePath;
        }
    });

    try {
        updateStatus(`${processedFilePaths.length}個のファイルを処理中...`);
        showLoading(true);
        
        // 進捗表示の初期化
        let processedFiles = 0;
        const totalFiles = processedFilePaths.length;

        // 進捗更新のリスナーを設定
        const progressHandler = (progressData) => {
            const { file, processed, total } = progressData;
            updateStatus(`処理中: ${file} (${processed}/${total})`);
        };

        // 進捗イベントのリスナーを追加
        if (window.electronAPI.onExtractionProgress) {
            window.electronAPI.onExtractionProgress(progressHandler);
        }

        console.log('Sending files to main process:', processedFilePaths);
        const result = await window.electronAPI.extractImages(processedFilePaths);
        
        if (result.success) {
            const successfulFiles = result.data.filter(file => file.success);
            const failedFiles = result.data.filter(file => !file.success);
            
            appState.files = successfulFiles;
            appState.allImages = [];
            
            // 成功したファイルから画像を収集
            successfulFiles.forEach(file => {
                if (file.images && Array.isArray(file.images)) {
                    file.images.forEach(image => {
                        image.sourceFile = file.fileName;
                        appState.allImages.push(image);
                    });
                }
            });

            updateFileList();
            updateImageDisplay();
            updateStats();
            
            // 結果メッセージの生成
            let statusMessage = `${successfulFiles.length}個のファイルから${appState.allImages.length}個の画像を抽出しました`;
            
            if (failedFiles.length > 0) {
                statusMessage += ` (${failedFiles.length}個のファイルで問題が発生)`;
                console.warn('Failed files:', failedFiles.map(f => ({ 
                    fileName: f.fileName, 
                    error: f.error 
                })));
                
                // 失敗したファイルの詳細をユーザーに表示
                const failedFileNames = failedFiles.map(f => f.fileName).join(', ');
                setTimeout(() => {
                    updateStatus(`処理完了。失敗したファイル: ${failedFileNames}`);
                }, 3000);
            }
            
            updateStatus(statusMessage);
        } else {
            updateStatus(`エラー: ${result.error}`);
            console.error('Extraction failed:', result);
        }
    } catch (error) {
        console.error('Error processing files:', error);
        
        // より詳細なエラーメッセージ
        let errorMessage = 'ファイルの処理中にエラーが発生しました';
        
        if (error.message.includes('ENOENT')) {
            errorMessage = 'ファイルが見つかりません（ファイルパスまたは日本語文字に問題がある可能性があります）';
        } else if (error.message.includes('EACCES')) {
            errorMessage = 'ファイルにアクセスできません（権限不足）';
        } else if (error.message.includes('EMFILE')) {
            errorMessage = 'ファイルを開きすぎています（一度に処理するファイル数を減らしてください）';
        } else if (error.message.includes('encoding') || error.message.includes('文字')) {
            errorMessage = '文字エンコーディングの問題が発生しました';
        }
        
        updateStatus(errorMessage);
    } finally {
        showLoading(false);
        
        // 進捗リスナーのクリーンアップ
        if (window.electronAPI.removeExtractionProgressListener) {
            window.electronAPI.removeExtractionProgressListener();
        }
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

// 画像表示の更新（最適化版）
function updateImageDisplay() {
    if (appState.allImages.length === 0) {
        elements.imageContainer.innerHTML = `
            <div class="empty-state">
                <p>画像がありません</p>
                <p class="small">Officeファイルを読み込んで画像を抽出してください</p>
            </div>
        `;
        elements.saveAllBtn.disabled = true;
        elements.exportZipBtn.disabled = true;
        return;
    }

    const containerClass = appState.currentView === 'grid' ? 'image-grid' : 'image-list';
    elements.imageContainer.innerHTML = `<div class="${containerClass}" id="imageDisplayArea"></div>`;
    
    const displayArea = document.getElementById('imageDisplayArea');
    const useVirtualScroll = appState.allImages.length > PERFORMANCE_CONFIG.VIRTUAL_SCROLL_THRESHOLD;
    const useLazyLoad = appState.allImages.length > PERFORMANCE_CONFIG.LAZY_LOAD_THRESHOLD;
    
    // バッチレンダリングの実装
    const renderBatch = (startIndex, endIndex) => {
        const fragment = document.createDocumentFragment();
        
        for (let i = startIndex; i < endIndex && i < appState.allImages.length; i++) {
            const image = appState.allImages[i];
            const imageItem = createImageItem(image, useLazyLoad);
            fragment.appendChild(imageItem);
        }
        
        displayArea.appendChild(fragment);
    };
    
    if (useVirtualScroll) {
        // 仮想スクロール使用時は初期表示分のみレンダリング
        const initialBatchSize = Math.min(PERFORMANCE_CONFIG.BATCH_RENDER_SIZE, appState.allImages.length);
        renderBatch(0, initialBatchSize);
        console.log(`Virtual scrolling enabled for ${appState.allImages.length} images`);
    } else {
        // 通常表示時はバッチでレンダリング
        let currentIndex = 0;
        const renderNextBatch = () => {
            const endIndex = Math.min(currentIndex + PERFORMANCE_CONFIG.BATCH_RENDER_SIZE, appState.allImages.length);
            renderBatch(currentIndex, endIndex);
            currentIndex = endIndex;
            
            if (currentIndex < appState.allImages.length) {
                // 次のバッチを非同期でレンダリング
                requestAnimationFrame(renderNextBatch);
            }
        };
        
        renderNextBatch();
    }

    elements.saveAllBtn.disabled = false;
    elements.exportZipBtn.disabled = false;
}

// 画像アイテムの作成（最適化版）
function createImageItem(image, useLazyLoad) {
    const imageItem = document.createElement('div');
    imageItem.className = 'image-item';
    
    const thumbnailSrc = image.thumbnail || image.data;
    const dimensions = image.width && image.height ? `${image.width} × ${image.height}` : '不明';
    const fileSize = formatFileSize(image.size);
    
    // 遅延読み込み用の属性設定
    const imgSrc = useLazyLoad ? '' : thumbnailSrc;
    const imgDataSrc = useLazyLoad ? thumbnailSrc : '';
    
    imageItem.innerHTML = `
        <img ${imgSrc ? `src="${imgSrc}"` : ''} 
             ${imgDataSrc ? `data-src="${imgDataSrc}"` : ''} 
             alt="${image.name}" 
             class="image-thumbnail ${useLazyLoad ? 'lazy' : ''}"
             loading="lazy">
        <div class="image-info">
            <div class="image-name" title="${image.name}">${image.name}</div>
            <div class="image-details">${dimensions} • ${fileSize}</div>
        </div>
    `;
    
    // 遅延読み込みの設定
    if (useLazyLoad && window.imageObserver) {
        const img = imageItem.querySelector('.image-thumbnail');
        window.imageObserver.observe(img);
    }
    
    imageItem.addEventListener('click', () => {
        showImageModal(image);
    });
    
    return imageItem;
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

// すべての画像をZIPファイルとしてエクスポート
async function exportAllImagesAsZip() {
    if (appState.allImages.length === 0) return;
    
    try {
        updateStatus('ZIPファイルを作成中...');
        
        // 最初のファイル名を取得（複数の場合は混合）
        const sourceFileName = appState.files.length === 1 
            ? path.parse(appState.files[0].fileName).name 
            : 'extracted_images';
            
        const result = await window.electronAPI.exportImagesZip(appState.allImages, sourceFileName);
        
        if (result.success) {
            updateStatus(`${result.exportedCount}個の画像をZIPファイルに保存しました: ${result.filePath}`);
        } else {
            updateStatus('ZIPファイルの作成に失敗しました');
        }
    } catch (error) {
        console.error('Error exporting ZIP:', error);
        updateStatus('ZIPファイルの作成中にエラーが発生しました');
    }
}

// 現在の画像をZIPファイルとしてエクスポート
async function exportCurrentImageAsZip() {
    if (!appState.selectedImage) return;
    
    try {
        updateStatus('ZIPファイルを作成中...');
        
        const imageName = path.parse(appState.selectedImage.name).name;
        const result = await window.electronAPI.exportImagesZip([appState.selectedImage], imageName);
        
        if (result.success) {
            updateStatus(`画像をZIPファイルに保存しました: ${result.filePath}`);
            closeModal();
        } else {
            updateStatus('ZIPファイルの作成に失敗しました');
        }
    } catch (error) {
        console.error('Error exporting ZIP:', error);
        updateStatus('ZIPファイルの作成中にエラーが発生しました');
    }
}

// path モジュールの簡易実装（ブラウザ環境用）
const path = {
    parse: (filePath) => {
        const lastDot = filePath.lastIndexOf('.');
        const lastSlash = Math.max(filePath.lastIndexOf('/'), filePath.lastIndexOf('\\'));
        
        if (lastDot > lastSlash && lastDot !== -1) {
            return {
                name: filePath.substring(lastSlash + 1, lastDot),
                ext: filePath.substring(lastDot)
            };
        } else {
            return {
                name: filePath.substring(lastSlash + 1),
                ext: ''
            };
        }
    }
};

// エラーハンドリング
window.addEventListener('error', (event) => {
    console.error('Unhandled error:', event.error);
    updateStatus('予期しないエラーが発生しました');
});

window.addEventListener('unhandledrejection', (event) => {
    console.error('Unhandled promise rejection:', event.reason);
    updateStatus('非同期処理でエラーが発生しました');
});
