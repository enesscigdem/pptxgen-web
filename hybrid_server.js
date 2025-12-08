const http = require("http");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execSync } = require("child_process");

function quote(f) {
    const platform = os.platform();
    if (platform === 'win32') {
        // Windows'ta çift tırnak kullan ve içindeki çift tırnakları escape et
        return `"${f.replace(/"/g, '""')}"`;
    } else {
        // Unix/Linux/macOS'ta tek tırnak kullan
        return `'${f.replace(/'/g, `'\''`)}'`;
    }
}

/**
 * LibreOffice binary yolunu bul
 */
function findLibreOffice() {
    const platform = os.platform();
    
    // macOS
    const macPath = "/Applications/LibreOffice.app/Contents/MacOS/soffice";
    if (fs.existsSync(macPath)) {
        return macPath;
    }
    
    // Windows - Standart yolu kontrol et
    if (platform === 'win32') {
        const winPaths = [
            "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
            "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe"
        ];
        
        for (const winPath of winPaths) {
            if (fs.existsSync(winPath)) {
                return winPath;
            }
        }
        
        // PATH'te ara (Windows'ta where.exe kullan)
        try {
            execSync("where.exe soffice", { stdio: 'ignore' });
            return "soffice";
        } catch (e) {
            throw new Error("LibreOffice bulunamadı! Lütfen LibreOffice'i kurun.");
        }
    }
    
    // Linux - PATH'te ara
    try {
        execSync("which libreoffice", { stdio: 'ignore' });
        return "libreoffice";
    } catch (e) {
        try {
            execSync("which soffice", { stdio: 'ignore' });
            return "soffice";
        } catch (e2) {
            throw new Error("LibreOffice bulunamadı! Lütfen LibreOffice'i kurun.");
        }
    }
}

const LIBREOFFICE_CMD = findLibreOffice();

/**
 * PPTX'i PDF'ye çevir (LibreOffice kullanarak)
 */
function convertPptxToPdf(inputPath) {
    const outDir = fs.mkdtempSync(path.join(os.tmpdir(), "lo-pdf-"));
    
    let cmd = `${quote(LIBREOFFICE_CMD)} --headless `;
    cmd += `--convert-to pdf ${quote(inputPath)} --outdir ${quote(outDir)}`;
    
    try {
        execSync(cmd, { stdio: 'pipe' });
    } catch (e) {
        const errorMsg = e.stderr ? e.stderr.toString() : e.message;
        throw new Error(`LibreOffice conversion failed: ${errorMsg}`);
    }
    
    const base = path.basename(inputPath).replace(/\.[^.]+$/, "");
    const expected = path.join(outDir, `${base}.pdf`);
    
    if (fs.existsSync(expected)) return expected;
    
    const files = fs.readdirSync(outDir);
    if (files.length) return path.join(outDir, files[0]);
    
    throw new Error("LibreOffice PDF çıktı üretmedi!");
}

/**
 * PDF'den belirli bir sayfayı image'e çevir (ImageMagick veya pdftoppm kullanarak)
 */
function pdfPageToImage(pdfPath, pageNumber, outputPath) {
    // Önce pdftoppm deneyelim (daha hızlı)
    try {
        const cmd = `pdftoppm -f ${pageNumber} -l ${pageNumber} -png -r 150 ${quote(pdfPath)} ${quote(outputPath.replace('.png', ''))}`;
        execSync(cmd);
        const imagePath = outputPath.replace('.png', `-${String(pageNumber).padStart(2, '0')}-1.png`);
        if (fs.existsSync(imagePath)) {
            // Rename to expected output
            if (imagePath !== outputPath) {
                fs.renameSync(imagePath, outputPath);
            }
            return outputPath;
        }
    } catch (e) {
        console.warn("pdftoppm failed, trying ImageMagick...");
    }
    
    // ImageMagick fallback
    try {
        const cmd = `convert -density 150 ${quote(pdfPath)}[${pageNumber - 1}] ${quote(outputPath)}`;
        execSync(cmd);
        if (fs.existsSync(outputPath)) return outputPath;
    } catch (e) {
        throw new Error(`Image conversion failed: ${e.message}`);
    }
    
    throw new Error("Image conversion failed");
}

function cleanup(dir) {
    if (!fs.existsSync(dir)) return;
    try {
        fs.readdirSync(dir).forEach(f => {
            const full = path.join(dir, f);
            if (fs.statSync(full).isDirectory()) cleanup(full);
            else fs.unlinkSync(full);
        });
        fs.rmdirSync(dir);
    } catch (e) {
        console.warn("Cleanup error:", e.message);
    }
}

const server = http.createServer((req, res) => {
    // CORS headers
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");

    if (req.method === "OPTIONS") {
        res.writeHead(200);
        return res.end();
    }

    if (req.method === "POST" && req.url === "/hybrid/convert-pdf") {
        const chunks = [];
        let bodySize = 0;
        const MAX_BODY_SIZE = 200 * 1024 * 1024; // 200MB limit
        
        req.on("data", d => {
            bodySize += d.length;
            if (bodySize > MAX_BODY_SIZE) {
                res.writeHead(413, {
                    "Content-Type": "application/json"
                });
                res.end(JSON.stringify({ error: "File too large. Maximum size: 200MB" }));
                return;
            }
            chunks.push(d);
        });
        
        req.on("end", () => {
            let tmpDir = null;
            let pdfPath = null;
            try {
                console.log("Request received, body size:", bodySize, "bytes");
                const body = Buffer.concat(chunks).toString('utf8');
                console.log("Body parsed, parsing JSON...");
                const { fileName, fileData } = JSON.parse(body);
                console.log("JSON parsed, fileName:", fileName, "fileData length:", fileData.length);
                
                console.log("Decoding base64...");
                const bytes = Buffer.from(fileData, "base64");
                console.log("Base64 decoded, size:", bytes.length, "bytes");

                tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "hybrid-"));
                const input = path.join(tmpDir, fileName);
                console.log("Writing input file to:", input);
                fs.writeFileSync(input, bytes);
                console.log("Input file written successfully");

                console.log("Starting LibreOffice conversion...");
                pdfPath = convertPptxToPdf(input);
                console.log("PDF conversion completed, path:", pdfPath);
                
                const pdfBuffer = fs.readFileSync(pdfPath);
                console.log("PDF buffer read, size:", pdfBuffer.length, "bytes");

                // Dosya adını güvenli hale getir
                const originalName = fileName.replace(/\.pptx?$/i, '');
                const safeFileName = originalName.replace(/[^\x20-\x7E]/g, '_') + '.pdf';

                console.log("Sending PDF response...");
                
                // Response headers
                res.writeHead(200, {
                    "Content-Type": "application/pdf",
                    "Content-Disposition": `attachment; filename="${safeFileName}"`,
                    "Content-Length": pdfBuffer.length,
                    "Cache-Control": "no-cache"
                });
                
                // Response'u gönder
                res.write(pdfBuffer);
                res.end();
                
                console.log("PDF response sent successfully, size:", pdfBuffer.length);

                cleanup(tmpDir);
            } catch (e) {
                console.error("Conversion error:", e);
                console.error("Error stack:", e.stack);
                if (!res.headersSent) {
                    res.writeHead(500, {
                        "Content-Type": "application/json"
                    });
                    res.end(JSON.stringify({ error: e.message }));
                }
                if (tmpDir) cleanup(tmpDir);
            }
        });
        
        req.on("error", (err) => {
            console.error("Request error:", err);
            if (!res.headersSent) {
                res.writeHead(500, {
                    "Content-Type": "application/json"
                });
                res.end(JSON.stringify({ error: "Request error: " + err.message }));
            }
        });
        return;
    }

    if (req.method === "POST" && req.url === "/hybrid/pdf-page-image") {
        const chunks = [];
        req.on("data", d => chunks.push(d));
        req.on("end", () => {
            let tmpDir = null;
            let pdfPath = null;
            try {
                const body = Buffer.concat(chunks).toString('utf8');
                const { pdfData, pageNumber } = JSON.parse(body);
                const pdfBytes = Buffer.from(pdfData, "base64");

                tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "hybrid-img-"));
                pdfPath = path.join(tmpDir, "input.pdf");
                fs.writeFileSync(pdfPath, pdfBytes);

                const imagePath = path.join(tmpDir, `page-${pageNumber}.png`);
                pdfPageToImage(pdfPath, pageNumber, imagePath);

                const imageBuffer = fs.readFileSync(imagePath);
                const imageBase64 = imageBuffer.toString("base64");

                res.writeHead(200, {
                    "Content-Type": "application/json",
                });
                res.end(JSON.stringify({
                    image: `data:image/png;base64,${imageBase64}`
                }));

                cleanup(tmpDir);
            } catch (e) {
                console.error("Error:", e);
                res.writeHead(500, {
                    "Content-Type": "application/json"
                });
                res.end(JSON.stringify({ error: e.message }));
                if (tmpDir) cleanup(tmpDir);
            }
        });
        return;
    }

    res.writeHead(404, {
        "Content-Type": "application/json"
    });
    res.end(JSON.stringify({ error: "Not found" }));
});

const PORT = process.env.PORT || 3001;
server.timeout = 300000; // 5 dakika timeout
server.keepAliveTimeout = 300000;
server.listen(PORT, () => {
    console.log(`Hybrid server running on port ${PORT}`);
    console.log(`LibreOffice found at: ${LIBREOFFICE_CMD}`);
    console.log("Endpoints:");
    console.log("  POST /hybrid/convert-pdf - Convert PPTX to PDF (max 200MB)");
    console.log("  POST /hybrid/pdf-page-image - Convert PDF page to image");
});

