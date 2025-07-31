const fs = require('fs');
const path = require('path');
const yauzl = require('yauzl');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const sharp = require('sharp');

/**
 * Officeファイルから画像を抽出する
 * @param {string} filePath - ファイルパス
 * @returns {Promise<Array>} - 抽出された画像の配列
 */
async function extractImages(filePath) {
    const ext = path.extname(filePath).toLowerCase();
    
    switch (ext) {
        case '.docx':
            return await extractFromDocx(filePath);
        case '.doc':
            return await extractFromDoc(filePath);
        case '.xlsx':
            return await extractFromXlsx(filePath);
        case '.xls':
            return await extractFromXls(filePath);
        case '.pptx':
            return await extractFromPptx(filePath);
        case '.ppt':
            return await extractFromPpt(filePath);
        default:
            throw new Error(`Unsupported file format: ${ext}`);
    }
}

/**
 * DOCXファイルから画像を抽出
 */
async function extractFromDocx(filePath) {
    return new Promise((resolve, reject) => {
        const images = [];
        
        yauzl.open(filePath, { lazyEntries: true }, (err, zipfile) => {
            if (err) {
                reject(err);
                return;
            }
            
            zipfile.readEntry();
            
            zipfile.on('entry', (entry) => {
                // word/media/ フォルダ内の画像ファイルを探す
                if (entry.fileName.startsWith('word/media/') && isImageFile(entry.fileName)) {
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
 * ファイルが画像かどうかを判定
 */
function isImageFile(fileName) {
    const imageExtensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.svg', '.webp'];
    const ext = path.extname(fileName).toLowerCase();
    return imageExtensions.includes(ext);
}

/**
 * 画像バッファを処理してメタデータとサムネイルを生成
 */
async function processImageBuffer(buffer, fileName) {
    try {
        const image = sharp(buffer);
        const metadata = await image.metadata();
        
        // サムネイル生成（200x200px）
        const thumbnailBuffer = await image
            .resize(200, 200, { fit: 'inside', withoutEnlargement: true })
            .png()
            .toBuffer();
        
        // 元画像をBase64エンコード
        const originalBase64 = `data:image/${metadata.format};base64,${buffer.toString('base64')}`;
        const thumbnailBase64 = `data:image/png;base64,${thumbnailBuffer.toString('base64')}`;
        
        return {
            id: generateId(),
            name: path.basename(fileName),
            format: metadata.format,
            width: metadata.width,
            height: metadata.height,
            size: buffer.length,
            data: originalBase64,
            thumbnail: thumbnailBase64
        };
    } catch (error) {
        console.error('Error processing image buffer:', error);
        // Sharpで処理できない場合は、そのままBase64として返す
        const base64 = `data:application/octet-stream;base64,${buffer.toString('base64')}`;
        return {
            id: generateId(),
            name: path.basename(fileName),
            format: 'unknown',
            width: null,
            height: null,
            size: buffer.length,
            data: base64,
            thumbnail: null
        };
    }
}

/**
 * ユニークIDを生成
 */
function generateId() {
    return Math.random().toString(36).substr(2, 9);
}

module.exports = {
    extractImages
};
