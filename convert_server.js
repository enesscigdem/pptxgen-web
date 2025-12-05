const http = require("http");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execSync } = require("child_process");

function quote(f) {
    return `'${f.replace(/'/g, `'\''`)}'`;
}

function convertUsingLO(inputPath, convertTo, opt = {}) {
    const outDir = fs.mkdtempSync(path.join(os.tmpdir(), "lo-"));

    let cmd = "libreoffice --headless ";
    if (opt.infilter) cmd += `--infilter=${opt.infilter} `;
    cmd += `--convert-to ${convertTo} ${quote(inputPath)} --outdir ${quote(outDir)}`;

    execSync(cmd); // hata olursa try/catch upstream yakalayacak

    const ext = convertTo.split(":")[0];
    const base = path.basename(inputPath).replace(/\.[^.]+$/, "");
    const expected = path.join(outDir, `${base}.${ext}`);

    if (fs.existsSync(expected)) return expected;

    const files = fs.readdirSync(outDir);
    if (files.length) return path.join(outDir, files[0]);

    throw new Error("LibreOffice çıktı üretmedi!");
}

function pipeline(input, type) {
    switch (type) {
        case "pptx-to-pdf":
        case "word-to-pdf":
            return convertUsingLO(input, "pdf");

        case "pdf-to-pptx":
            return convertUsingLO(input, 'pptx:"Impress MS PowerPoint 2007 XML"');

        case "pdf-to-word": {
            const doc = convertUsingLO(input, "doc", { infilter: "writer_pdf_import" });
            return convertUsingLO(doc, "docx");
        }

        case "pptx-to-word": {
            const pdf = convertUsingLO(input, "pdf");
            const doc = convertUsingLO(pdf, "doc", { infilter: "writer_pdf_import" });
            return convertUsingLO(doc, "docx");
        }

        case "word-to-pptx": {
            const pdf = convertUsingLO(input, "pdf");
            return convertUsingLO(pdf, 'pptx:"Impress MS PowerPoint 2007 XML"');
        }

        default:
            throw new Error("unsupported type");
    }
}

function cleanup(dir) {
    if (!fs.existsSync(dir)) return;
    fs.readdirSync(dir).forEach(f => {
        const full = path.join(dir, f);
        if (fs.statSync(full).isDirectory()) cleanup(full);
        else fs.unlinkSync(full);
    });
    fs.rmdirSync(dir);
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
        let body = "";
        req.on("data", d => body += d);
        req.on("end", () => {
            try {
                const { fileName, fileData, type } = JSON.parse(body);
                const bytes = Buffer.from(fileData, "base64");

                const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "conv-"));
                const input = path.join(tmp, fileName);
                fs.writeFileSync(input, bytes);

                const output = pipeline(input, type);

                const buf = fs.readFileSync(output);
                res.writeHead(200, {
                    "Access-Control-Allow-Origin": "*",
                    "Content-Type": "application/octet-stream",
                    "Content-Disposition": `attachment; filename="${path.basename(output)}"`,
                });
                res.end(buf);

                cleanup(path.dirname(input));
                cleanup(path.dirname(output));
            }
            catch (e) {
                res.writeHead(500, { "Access-Control-Allow-Origin": "*" });
                res.end("Error: " + e.message);
            }
        });
    }
    else {
        res.writeHead(404, { "Access-Control-Allow-Origin": "*" });
        res.end("Not found");
    }

}).listen(3000, () => console.log("Converter running on 3000"));
