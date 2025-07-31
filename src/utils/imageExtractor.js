const fs = require('fs');
const path = require('path');
const yauzl = require('yauzl');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const sharp = require('sharp');

// 設定定数
const CONFIG = {
    MAX_FILE_SIZE: 100 * 1024 * 1024, // 100MB
    MAX_IMAGE_SIZE: 50 * 1024 * 1024, // 50MB per image
    SUPPORTED_IMAGE_TYPES: ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tiff'],
    THUMBNAIL_SIZE: 200,
    COMPRESSION_QUALITY: 85
};

/**
 * Officeファイルから画像を抽出する（最適化版）
 * @param {string} filePath - ファイルパス
 * @param {Object} options - オプション
 * @returns {Promise<Array>} - 抽出された画像の配列
 */
async function extractImages(filePath, options = {}) {
    // ファイルサイズ検証と正規化されたパスの取得
    const validatedPath = await validateFile(filePath);
    
    const ext = path.extname(validatedPath).toLowerCase();
    
    switch (ext) {
        case '.docx':
            return await extractFromDocx(validatedPath, options);
        case '.doc':
            return await extractFromDoc(validatedPath, options);
        case '.xlsx':
            return await extractFromXlsx(validatedPath, options);
        case '.xls':
            return await extractFromXls(validatedPath, options);
        case '.pptx':
            return await extractFromPptx(validatedPath, options);
        case '.ppt':
            return await extractFromPpt(validatedPath, options);
        default:
            throw new Error(`Unsupported file format: ${ext}`);
    }
}

/**
 * ファイルの検証（日本語ファイル名対応強化版）
 */
async function validateFile(filePath) {
    try {
        // パスの正規化と複数方式でのチェック
        const normalizedPath = path.resolve(filePath);
        console.log(`Validating file: "${normalizedPath}"`);
        console.log(`Original path: "${filePath}"`);
        
        // 文字エンコーディング情報をログ出力
        console.log(`Path encoding check:`, {
            utf8: Buffer.from(filePath, 'utf8').toString('utf8'),
            length: filePath.length,
            bytes: filePath.split('').map(c => c.charCodeAt(0))
        });
        
        const stats = await fs.promises.stat(normalizedPath);
        
        if (stats.size > CONFIG.MAX_FILE_SIZE) {
            throw new Error(`ファイルサイズが制限値 (${CONFIG.MAX_FILE_SIZE / 1024 / 1024}MB) を超えています`);
        }
        
        if (!stats.isFile()) {
            throw new Error('有効なファイルではありません');
        }
        
        console.log(`✓ File validation successful: ${path.basename(normalizedPath)} (${stats.size} bytes)`);
        return normalizedPath; // 正規化されたパスを返す
    } catch (error) {
        console.error(`✗ File validation failed for "${filePath}":`, error);
        
        if (error.code === 'ENOENT') {
            throw new Error(`ファイルが見つかりません: ${path.basename(filePath)}`);
        } else if (error.code === 'EACCES') {
            throw new Error(`ファイルへのアクセスが拒否されました: ${path.basename(filePath)}`);
        } else if (error.code === 'EINVAL') {
            throw new Error(`無効なファイルパスです: ${path.basename(filePath)}`);
        }
        throw error;
    }
}

/**
 * DOCXファイルから画像を抽出（最適化版）
 */
async function extractFromDocx(filePath, options = {}) {
    return new Promise((resolve, reject) => {
        const images = [];
        let processedEntries = 0;
        let totalImageEntries = 0;
        
        // 文字エンコーディング対応のオプション
        const zipOptions = { 
            lazyEntries: true,
            decodeStrings: true // 文字エンコーディング問題の対応
        };
        
        yauzl.open(filePath, zipOptions, (err, zipfile) => {
            if (err) {
                console.error(`YAUZL error for file ${path.basename(filePath)}:`, err);
                reject(new Error(`ZIPファイルの読み込みエラー (${path.basename(filePath)}): ${err.message}`));
                return;
            }
            
            // エントリー数をカウント
            zipfile.entryCount && console.log(`処理対象エントリー数: ${zipfile.entryCount}`);
            
            zipfile.readEntry();
            
            zipfile.on('entry', (entry) => {
                // word/media/ フォルダ内の画像ファイルを探す（最適化）
                if (entry.fileName.startsWith('word/media/') && isImageFile(entry.fileName)) {
                    totalImageEntries++;
                    
                    // 大きすぎる画像をスキップ
                    if (entry.uncompressedSize > CONFIG.MAX_IMAGE_SIZE) {
                        console.warn(`画像が大きすぎるためスキップ: ${entry.fileName} (${entry.uncompressedSize} bytes)`);
                        processedEntries++;
                        zipfile.readEntry();
                        return;
                    }
                    
                    zipfile.openReadStream(entry, (err, readStream) => {
                        if (err) {
                            console.error(`ストリーム読み込みエラー: ${entry.fileName}`, err);
                            processedEntries++;
                            zipfile.readEntry();
                            return;
                        }
                        
                        const chunks = [];
                        let totalSize = 0;
                        
                        readStream.on('data', (chunk) => {
                            totalSize += chunk.length;
                            // メモリ使用量制限
                            if (totalSize > CONFIG.MAX_IMAGE_SIZE) {
                                readStream.destroy();
                                console.warn(`画像サイズ制限超過: ${entry.fileName}`);
                                processedEntries++;
                                zipfile.readEntry();
                                return;
                            }
                            chunks.push(chunk);
                        });
                        
                        readStream.on('end', async () => {
                            const buffer = Buffer.concat(chunks);
                            try {
                                const imageInfo = await processImageBufferOptimized(buffer, entry.fileName, options);
                                if (imageInfo) {
                                    images.push(imageInfo);
                                }
                            } catch (error) {
                                console.error(`画像処理エラー: ${entry.fileName}`, error);
                            }
                            
                            processedEntries++;
                            // 進捗レポート
                            if (options.onProgress) {
                                options.onProgress(processedEntries, totalImageEntries);
                            }
                            
                            zipfile.readEntry();
                        });
                        
                        readStream.on('error', (err) => {
                            console.error(`読み込みエラー: ${entry.fileName}`, err);
                            processedEntries++;
                            zipfile.readEntry();
                        });
                    });
                } else {
                    zipfile.readEntry();
                }
            });
            
            zipfile.on('end', () => {
                resolve(images);
            });
            
            zipfile.on('error', (err) => {
                reject(new Error(`ZIP処理エラー: ${err.message}`));
            });
        });
    });
}

/**
 * DOCファイルから画像を抽出（制限あり）
 */
async function extractFromDoc(filePath) {
    // .docファイルは複雑な形式のため、基本的な対応のみ
    try {
        const result = await mammoth.convertToHtml({ path: filePath });
        const images = [];
        
        // HTMLから画像を抽出する場合の処理
        // 注意: mammoth.jsは埋め込み画像の抽出が限定的
        
        return images;
    } catch (error) {
        throw new Error(`Failed to extract from DOC file: ${error.message}`);
    }
}

/**
 * XLSXファイルから画像を抽出
 */
async function extractFromXlsx(filePath) {
    return new Promise((resolve, reject) => {
        const images = [];
        
        yauzl.open(filePath, { lazyEntries: true }, (err, zipfile) => {
            if (err) {
                reject(err);
                return;
            }
            
            zipfile.readEntry();
            
            zipfile.on('entry', (entry) => {
                // xl/media/ フォルダ内の画像ファイルを探す
                if (entry.fileName.startsWith('xl/media/') && isImageFile(entry.fileName)) {
                    zipfile.openReadStream(entry, (err, readStream) => {
                        if (err) {
                            zipfile.readEntry();
                            return;
                        }
                        
                        const chunks = [];
                        readStream.on('data', (chunk) => chunks.push(chunk));
                        readStream.on('end', async () => {
                            const buffer = Buffer.concat(chunks);
                            try {
                                const imageInfo = await processImageBuffer(buffer, entry.fileName);
                                images.push(imageInfo);
                            } catch (error) {
                                console.error('Error processing image:', error);
                            }
                            zipfile.readEntry();
                        });
                    });
                } else {
                    zipfile.readEntry();
                }
            });
            
            zipfile.on('end', () => {
                resolve(images);
            });
            
            zipfile.on('error', (err) => {
                reject(err);
            });
        });
    });
}

/**
 * XLSファイルから画像を抽出（制限あり）
 */
async function extractFromXls(filePath) {
    // .xlsファイルは複雑な形式のため、基本的な対応のみ
    return [];
}

/**
 * PPTXファイルから画像を抽出
 */
async function extractFromPptx(filePath) {
    return new Promise((resolve, reject) => {
        const images = [];
        
        yauzl.open(filePath, { lazyEntries: true }, (err, zipfile) => {
            if (err) {
                reject(err);
                return;
            }
            
            zipfile.readEntry();
            
            zipfile.on('entry', (entry) => {
                // ppt/media/ フォルダ内の画像ファイルを探す
                if (entry.fileName.startsWith('ppt/media/') && isImageFile(entry.fileName)) {
                    zipfile.openReadStream(entry, (err, readStream) => {
                        if (err) {
                            zipfile.readEntry();
                            return;
                        }
                        
                        const chunks = [];
                        readStream.on('data', (chunk) => chunks.push(chunk));
                        readStream.on('end', async () => {
                            const buffer = Buffer.concat(chunks);
                            try {
                                const imageInfo = await processImageBuffer(buffer, entry.fileName);
                                images.push(imageInfo);
                            } catch (error) {
                                console.error('Error processing image:', error);
                            }
                            zipfile.readEntry();
                        });
                    });
                } else {
                    zipfile.readEntry();
                }
            });
            
            zipfile.on('end', () => {
                resolve(images);
            });
            
            zipfile.on('error', (err) => {
                reject(err);
            });
        });
    });
}

/**
 * PPTファイルから画像を抽出（制限あり）
 */
async function extractFromPpt(filePath) {
    // .pptファイルは複雑な形式のため、基本的な対応のみ
    return [];
}

/**
 * ファイルが画像かどうかを判定（最適化版）
 */
function isImageFile(fileName) {
    const ext = path.extname(fileName).toLowerCase();
    return CONFIG.SUPPORTED_IMAGE_TYPES.includes(ext);
}

/**
 * 最適化された画像バッファ処理
 */
async function processImageBufferOptimized(buffer, fileName, options = {}) {
    try {
        // 空のバッファチェック
        if (!buffer || buffer.length === 0) {
            throw new Error('空の画像バッファです');
        }
        
        const image = sharp(buffer);
        const metadata = await image.metadata();
        
        // 最大寸法チェック
        const maxDimension = options.maxDimension || 4096;
        if (metadata.width > maxDimension || metadata.height > maxDimension) {
            console.warn(`画像サイズが大きすぎます: ${fileName} (${metadata.width}x${metadata.height})`);
        }
        
        // メモリ効率的なサムネイル生成
        const thumbnailSize = options.thumbnailSize || CONFIG.THUMBNAIL_SIZE;
        const thumbnailBuffer = await image
            .resize(thumbnailSize, thumbnailSize, { 
                fit: 'inside', 
                withoutEnlargement: true,
                kernel: sharp.kernel.lanczos3 // 高品質リサイズ
            })
            .jpeg({ quality: CONFIG.COMPRESSION_QUALITY, progressive: true })
            .toBuffer();
        
        // 圧縮された画像データの生成（オプション）
        let compressedBuffer = buffer;
        if (options.compress && metadata.format !== 'svg') {
            compressedBuffer = await image
                .jpeg({ quality: CONFIG.COMPRESSION_QUALITY, progressive: true })
                .toBuffer();
        }
        
        // Base64エンコード（遅延読み込み用にサムネイルのみ即座に生成）
        const thumbnailBase64 = `data:image/jpeg;base64,${thumbnailBuffer.toString('base64')}`;
        
        return {
            id: generateUniqueId(),
            name: path.basename(fileName),
            format: metadata.format || 'unknown',
            width: metadata.width || 0,
            height: metadata.height || 0,
            size: buffer.length,
            compressedSize: compressedBuffer.length,
            thumbnail: thumbnailBase64,
            // 元画像データは遅延読み込み用の関数として保存
            getData: () => `data:image/${metadata.format};base64,${compressedBuffer.toString('base64')}`,
            originalBuffer: buffer // 必要時のみ使用
        };
    } catch (error) {
        console.error(`画像処理エラー: ${fileName}`, error);
        
        // フォールバック処理
        try {
            const base64 = `data:application/octet-stream;base64,${buffer.toString('base64')}`;
            return {
                id: generateUniqueId(),
                name: path.basename(fileName),
                format: 'unknown',
                width: null,
                height: null,
                size: buffer.length,
                compressedSize: buffer.length,
                thumbnail: null,
                getData: () => base64,
                originalBuffer: buffer,
                error: error.message
            };
        } catch (fallbackError) {
            console.error(`フォールバック処理も失敗: ${fileName}`, fallbackError);
            return null; // 完全に失敗した場合はnullを返す
        }
    }
}

/**
 * 画像バッファを処理してメタデータとサムネイルを生成（下位互換用）
 */
async function processImageBuffer(buffer, fileName) {
    const result = await processImageBufferOptimized(buffer, fileName);
    if (!result) return null;
    
    return {
        id: result.id,
        name: result.name,
        format: result.format,
        width: result.width,
        height: result.height,
        size: result.size,
        data: result.getData(),
        thumbnail: result.thumbnail
    };
}

/**
 * 高性能ユニークID生成
 */
function generateUniqueId() {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * ファイルが画像かどうかを判定（最適化版）
 */
function isImageFile(fileName) {
    const ext = path.extname(fileName).toLowerCase();
    return CONFIG.SUPPORTED_IMAGE_TYPES.includes(ext);
}

/**
 * ユニークIDを生成（下位互換用）
 */
function generateId() {
    return generateUniqueId();
}

module.exports = {
    extractImages,
    CONFIG // 設定を外部に公開
};
