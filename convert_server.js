const http = require("http");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execSync } = require("child_process");

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

async function convertUsingLO(inputPath, convertTo, opt = {}) {
    const outDir = fs.mkdtempSync(path.join(os.tmpdir(), "lo-"));
    const platform = os.platform();
    const timeout = 120000; // 2 dakika timeout

    let cmd = `${quote(LIBREOFFICE_CMD)} --headless `;
    if (opt.infilter) cmd += `--infilter=${opt.infilter} `;
    cmd += `--convert-to ${convertTo} ${quote(inputPath)} --outdir ${quote(outDir)}`;

    // Windows'ta spawn kullan (timeout için gerekli)
    if (platform === 'win32') {
        return new Promise((resolve, reject) => {
            const { spawn } = require('child_process');
            const startTime = Date.now();
            console.log(`[LibreOffice] Starting conversion: ${path.basename(inputPath)} -> ${convertTo.split(':')[0]}`);
            
            // Windows'ta shell: true ile çalıştır
            const proc = spawn(cmd, [], { 
                stdio: 'pipe',
                shell: true
            });
            
            let stdout = '';
            let stderr = '';
            
            proc.stdout.on('data', (data) => {
                stdout += data.toString();
            });
            
            proc.stderr.on('data', (data) => {
                stderr += data.toString();
            });
            
            const timeoutId = setTimeout(() => {
                try {
                    proc.kill();
                    // Windows'ta process tree'yi de kill et
                    if (proc.pid) {
                        try {
                            require('child_process').exec(`taskkill /F /T /PID ${proc.pid}`, { stdio: 'ignore' });
                        } catch (e) {}
                    }
                } catch (e) {}
                reject(new Error('LibreOffice conversion timeout (2 minutes)'));
            }, timeout);
            
            proc.on('close', (code) => {
                clearTimeout(timeoutId);
                const duration = ((Date.now() - startTime) / 1000).toFixed(2);
                console.log(`[LibreOffice] Process exited with code ${code} in ${duration}s`);
                
                const ext = convertTo.split(":")[0];
                const base = path.basename(inputPath).replace(/\.[^.]+$/, "");
                const expected = path.join(outDir, `${base}.${ext}`);
                
                if (fs.existsSync(expected)) {
                    console.log(`[LibreOffice] Conversion completed: ${path.basename(expected)}`);
                    resolve(expected);
                } else {
                    const files = fs.readdirSync(outDir);
                    if (files.length) {
                        console.log(`[LibreOffice] Conversion completed: ${path.basename(files[0])}`);
                        resolve(path.join(outDir, files[0]));
                    } else {
                        reject(new Error(`LibreOffice çıktı üretmedi! Exit code: ${code}, Error: ${stderr || stdout || 'Unknown'}`));
                    }
                }
            });
            
            proc.on('error', (err) => {
                clearTimeout(timeoutId);
                reject(new Error(`LibreOffice process error: ${err.message}`));
            });
        });
    } else {
        // Unix/Linux/macOS için spawn kullan (timeout için)
        return new Promise((resolve, reject) => {
            const { spawn } = require('child_process');
            const parts = cmd.split(/\s+/);
            const exe = parts[0];
            const args = parts.slice(1);
            
            const proc = spawn(exe, args, { stdio: 'pipe' });
            let stdout = '';
            let stderr = '';
            
            proc.stdout.on('data', (data) => {
                stdout += data.toString();
            });
            
            proc.stderr.on('data', (data) => {
                stderr += data.toString();
            });
            
            const timeoutId = setTimeout(() => {
                try {
                    proc.kill('SIGTERM');
                    setTimeout(() => {
                        if (!proc.killed) proc.kill('SIGKILL');
                    }, 5000);
                } catch (e) {}
                reject(new Error('LibreOffice conversion timeout (2 minutes)'));
            }, timeout);
            
            proc.on('close', (code) => {
                clearTimeout(timeoutId);
                const ext = convertTo.split(":")[0];
                const base = path.basename(inputPath).replace(/\.[^.]+$/, "");
                const expected = path.join(outDir, `${base}.${ext}`);
                
                if (fs.existsSync(expected)) {
                    resolve(expected);
                } else {
                    const files = fs.readdirSync(outDir);
                    if (files.length) {
                        resolve(path.join(outDir, files[0]));
                    } else {
                        reject(new Error(`LibreOffice çıktı üretmedi! Exit code: ${code}, Error: ${stderr || stdout || 'Unknown'}`));
                    }
                }
            });
            
            proc.on('error', (err) => {
                clearTimeout(timeoutId);
                reject(new Error(`LibreOffice process error: ${err.message}`));
            });
        });
    }
}

async function pipeline(input, type) {
    switch (type) {
        case "pptx-to-pdf":
        case "word-to-pdf":
            return await convertUsingLO(input, "pdf");

        case "pdf-to-pptx":
            // PDF'den PPTX'e direkt dönüşüm dene
            try {
                return await convertUsingLO(input, 'pptx:"Impress MS PowerPoint 2007 XML"');
            } catch (e) {
                // Fallback: ODP'ye çevir sonra PPTX'e
                console.log("Direct PDF to PPTX failed, trying ODP intermediate...");
                const odp = await convertUsingLO(input, "odp");
                return await convertUsingLO(odp, 'pptx:"Impress MS PowerPoint 2007 XML"');
            }

        case "pdf-to-word": {
            // PDF'den Word'e dönüşüm
            try {
                const doc = await convertUsingLO(input, "doc", { infilter: "writer_pdf_import" });
                return await convertUsingLO(doc, "docx");
            } catch (e) {
                // Fallback: ODT'ye çevir sonra DOCX'e
                console.log("PDF to DOC failed, trying ODT intermediate...");
                const odt = await convertUsingLO(input, "odt", { infilter: "writer_pdf_import" });
                return await convertUsingLO(odt, "docx");
            }
        }

        case "pptx-to-word": {
            // PPTX -> PDF -> Word (en güvenilir yol)
            const pdf = await convertUsingLO(input, "pdf");
            try {
                const doc = await convertUsingLO(pdf, "doc", { infilter: "writer_pdf_import" });
                return await convertUsingLO(doc, "docx");
            } catch (e) {
                // Fallback: ODT'ye çevir sonra DOCX'e
                console.log("PDF to DOC failed, trying ODT intermediate...");
                const odt = await convertUsingLO(pdf, "odt", { infilter: "writer_pdf_import" });
                return await convertUsingLO(odt, "docx");
            }
        }

        case "word-to-pptx": {
            // Word -> PDF -> PPTX
            const pdf = await convertUsingLO(input, "pdf");
            try {
                return await convertUsingLO(pdf, 'pptx:"Impress MS PowerPoint 2007 XML"');
            } catch (e) {
                // Fallback: ODP'ye çevir sonra PPTX'e
                console.log("PDF to PPTX failed, trying ODP intermediate...");
                const odp = await convertUsingLO(pdf, "odp");
                return await convertUsingLO(odp, 'pptx:"Impress MS PowerPoint 2007 XML"');
            }
        }

        default:
            throw new Error(`Unsupported conversion type: ${type}`);
    }
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

http.createServer((req, res) => {

    // CORS preflight
    if (req.method === "OPTIONS") {
        res.writeHead(200, {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type"
        });
        return res.end();
    }

    if (req.method === "POST" && req.url === "/convert") {
        const chunks = [];
        let bodySize = 0;
        const MAX_BODY_SIZE = 200 * 1024 * 1024; // 200MB limit
        
        req.on("data", d => {
            bodySize += d.length;
            if (bodySize > MAX_BODY_SIZE) {
                res.writeHead(413, {
                    "Access-Control-Allow-Origin": "*",
                    "Content-Type": "application/json"
                });
                res.end(JSON.stringify({ error: "File too large. Maximum size: 200MB" }));
                return;
            }
            chunks.push(d);
        });
        
        req.on("end", async () => {
            let tmpDir = null;
            let outputPath = null;
            try {
                const body = Buffer.concat(chunks).toString('utf8');
                const { fileName, fileData, type } = JSON.parse(body);
                const bytes = Buffer.from(fileData, "base64");

                tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "conv-"));
                const input = path.join(tmpDir, fileName);
                fs.writeFileSync(input, bytes);

                const startTime = Date.now();
                console.log(`[${new Date().toISOString()}] Converting ${fileName} to ${type}...`);
                outputPath = await pipeline(input, type);
                const duration = ((Date.now() - startTime) / 1000).toFixed(2);
                console.log(`[${new Date().toISOString()}] Conversion complete: ${path.basename(outputPath)} (${duration}s)`);

                const buf = fs.readFileSync(outputPath);
                
                // Dosya adını güvenli hale getir
                const originalName = fileName.replace(/\.[^.]+$/, '');
                const extMap = {
                    "pptx-to-pdf": ".pdf",
                    "pdf-to-pptx": ".pptx",
                    "pptx-to-word": ".docx",
                    "word-to-pptx": ".pptx",
                    "pdf-to-word": ".docx",
                    "word-to-pdf": ".pdf",
                };
                const newExt = extMap[type] || path.extname(outputPath);
                const safeFileName = originalName.replace(/[^\x20-\x7E]/g, '_') + newExt;
                
                res.writeHead(200, {
                    "Access-Control-Allow-Origin": "*",
                    "Content-Type": "application/octet-stream",
                    "Content-Disposition": `attachment; filename="${safeFileName}"`,
                    "Content-Length": buf.length
                });
                res.end(buf);

                cleanup(tmpDir);
            }
            catch (e) {
                console.error("Conversion error:", e);
                if (!res.headersSent) {
                    res.writeHead(500, { 
                        "Access-Control-Allow-Origin": "*",
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
                    "Access-Control-Allow-Origin": "*",
                    "Content-Type": "application/json"
                });
                res.end(JSON.stringify({ error: "Request error: " + err.message }));
            }
        });
        return;
    }
    else {
        res.writeHead(404, { "Access-Control-Allow-Origin": "*" });
        res.end("Not found");
    }

}).listen(process.env.PORT || 3002, () => {
    const port = process.env.PORT || 3002;
    console.log(`Format Converter server running on port ${port}`);
    console.log(`LibreOffice found at: ${LIBREOFFICE_CMD}`);
    console.log("Endpoint: POST /convert");
});
