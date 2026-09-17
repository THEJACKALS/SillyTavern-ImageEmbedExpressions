import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { Jimp } from '../../../../src/jimp.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Patch global fetch to support file:// URLs for WASM binaries in Node.js
const originalFetch = globalThis.fetch;
globalThis.fetch = async (input, init) => {
    const urlStr = String(input);
    if (urlStr.startsWith('file://')) {
        try {
            const filePath = fileURLToPath(urlStr);
            const data = fs.readFileSync(filePath);
            return new Response(data, {
                status: 200,
                headers: { 'Content-Type': 'application/wasm' },
            });
        } catch (err) {
            console.error('Failed to read local file URL:', urlStr, err);
        }
    }
    return originalFetch(input, init);
};

const USER_IMAGES_DIR = path.resolve(__dirname, '../../user/images');
const SETTINGS_FILE = path.resolve(__dirname, '../../settings.json');

console.log('=== ImageEmbedExpressions - Bulk WebP Conversion ===');
console.log(`Scanning directory: ${USER_IMAGES_DIR}`);

async function convertAllExistingToWebP() {
    if (!fs.existsSync(USER_IMAGES_DIR)) {
        console.error(`Directory not found: ${USER_IMAGES_DIR}`);
        return;
    }

    const items = fs.readdirSync(USER_IMAGES_DIR);
    const targetFolders = items.filter(item => {
        const fullPath = path.join(USER_IMAGES_DIR, item);
        return fs.statSync(fullPath).isDirectory() && item.includes('image-embeds-expressions');
    });

    if (targetFolders.length === 0) {
        console.log('No expression image directories found in user/images.');
        return;
    }

    console.log(`Found ${targetFolders.length} expression folder(s):`, targetFolders);

    let convertedCount = 0;
    let skippedCount = 0;
    let errorCount = 0;

    const urlReplacements = new Map();

    for (const folder of targetFolders) {
        const folderPath = path.join(USER_IMAGES_DIR, folder);
        const files = fs.readdirSync(folderPath);

        for (const file of files) {
            const ext = path.extname(file).toLowerCase();
            const filePath = path.join(folderPath, file);

            if (ext === '.webp') {
                skippedCount++;
                continue;
            }

            if (!['.png', '.jpg', '.jpeg', '.gif', '.bmp'].includes(ext)) {
                continue;
            }

            try {
                const fileBuffer = fs.readFileSync(filePath);
                const img = await Jimp.read(fileBuffer);
                const width = img.bitmap.width;
                const height = img.bitmap.height;

                // Obtain WebP buffer preserving 1:1 original resolution
                const webpBuffer = await img.getBuffer('image/webp');

                const newFileName = file.substring(0, file.length - ext.length) + '.webp';
                const newFilePath = path.join(folderPath, newFileName);

                fs.writeFileSync(newFilePath, webpBuffer);
                fs.unlinkSync(filePath);

                // Track replacement URL patterns
                const oldRelUrl = `${folder}/${file}`;
                const newRelUrl = `${folder}/${newFileName}`;
                urlReplacements.set(oldRelUrl, newRelUrl);

                convertedCount++;
                console.log(`[CONVERTED] ${file} -> ${newFileName} (${width}x${height}px)`);
            } catch (err) {
                errorCount++;
                console.error(`[ERROR] Failed to convert ${file}:`, err.stack || err.message);
            }
        }
    }

    console.log('\n--- Conversion Summary ---');
    console.log(`Successfully converted: ${convertedCount}`);
    console.log(`Already WebP (skipped): ${skippedCount}`);
    console.log(`Errors: ${errorCount}`);

    // Update settings.json references if present
    if (fs.existsSync(SETTINGS_FILE) && urlReplacements.size > 0) {
        try {
            console.log('\nUpdating settings.json image references...');
            let settingsText = fs.readFileSync(SETTINGS_FILE, 'utf8');
            let updated = false;

            for (const [oldUrl, newUrl] of urlReplacements.entries()) {
                if (settingsText.includes(oldUrl)) {
                    settingsText = settingsText.split(oldUrl).join(newUrl);
                    updated = true;
                }
            }

            if (updated) {
                fs.writeFileSync(SETTINGS_FILE, settingsText, 'utf8');
                console.log('settings.json updated successfully!');
            } else {
                console.log('No matching URL paths needed updating in settings.json.');
            }
        } catch (err) {
            console.error('Failed to update settings.json:', err.message);
        }
    }
}

convertAllExistingToWebP().catch(err => {
    console.error('Fatal error during WebP conversion:', err);
});
