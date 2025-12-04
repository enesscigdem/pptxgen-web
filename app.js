// EMU -> cm helper
function emuToCm(emu) {
    let v = parseInt(emu, 10);
    if (isNaN(v)) return null;
    // 1 inch = 914400 EMU, 1 inch = 2.54cm → 1cm ≈ 360000 EMU
    return +(v / 360000).toFixed(2);
}
let themeColorMap = {}; // theme / scheme renklerini buraya koyacağız
// Global JSON (tüm sunum)
let presentationJson = null;

let lastJson = localStorage.getItem("presentationJson");
if(lastJson){
    presentationJson = JSON.parse(lastJson);
}

let mediaPath = "ppt/media/";

async function getBase64(zip, path) {
    let file = zip.file(path);
    if (!file) return null;
    let bin = await file.async("base64");
    let ext = path.split(".").pop().toLowerCase();
    return `data:image/${ext};base64,${bin}`;
}

document.getElementById("fileInput").addEventListener("change", async e => {

    let file = e.target.files[0];
    if (!file) return;

    let loading = document.getElementById("loading");
    let progressContainer = document.getElementById("progressContainer");
    let progressBar = document.getElementById("progressBar");
    let progressLabel = document.getElementById("progressLabel");
    let out = document.getElementById("output");
    let globalPanel = document.getElementById("globalPanel");
    let fileInfo = document.getElementById("fileInfo");
    let statsInfo = document.getElementById("statsInfo");

    out.innerHTML = "";
    loading.style.display = "block";
    globalPanel.classList.add("hidden");
    presentationJson = null;

    // show progress bar and reset
    if (progressContainer) progressContainer.style.display = "block";
    if (progressLabel) progressLabel.style.display = "block";
    if (progressBar) progressBar.style.width = "0%";
    if (progressLabel) progressLabel.textContent = "0%";

    let arrayBuffer = await file.arrayBuffer();
    let zip = await JSZip.loadAsync(arrayBuffer);
    let parser = new DOMParser();

    // slide size
    let presXml = await zip.file("ppt/presentation.xml").async("text");
    let presDoc = parser.parseFromString(presXml, "application/xml");
    let sldSz = presDoc.querySelector("p\\:sldSz,sldSz");
    let slideW = +sldSz.getAttribute("cx");
    let slideH = +sldSz.getAttribute("cy");

    // THEME RENKLERINI OKU
    try {
        let presRelsPath = "ppt/_rels/presentation.xml.rels";
        if (zip.file(presRelsPath)) {
            let presRelsXml = await zip.file(presRelsPath).async("text");
            let presRelsDoc = parser.parseFromString(presRelsXml, "application/xml");
            let rels = presRelsDoc.querySelectorAll("Relationship");
            let themeRel = null;
            rels.forEach(r => {
                let type = r.getAttribute("Type") || "";
                if (type.includes("/theme")) {
                    themeRel = r;
                }
            });
            if (themeRel) {
                let target = themeRel.getAttribute("Target"); // örn: "theme/theme1.xml" veya "../theme/theme1.xml"
                let themePath = target.replace("..", "ppt");
                if (!themePath.startsWith("ppt/")) {
                    themePath = "ppt/" + themePath.replace(/^\/+/, "");
                }

                if (zip.file(themePath)) {
                    let themeXml = await zip.file(themePath).async("text");
                    let themeDoc = parser.parseFromString(themeXml, "application/xml");
                    let clrScheme = themeDoc.querySelector("a\\:clrScheme,clrScheme");
                    if (clrScheme) {
                        Array.from(clrScheme.children).forEach(node => {
                            // node localName accent1, accent2, dk1, lt1 vs.
                            let name = node.localName || node.nodeName.split(":").pop();
                            let srgb = node.querySelector("a\\:srgbClr,srgbClr");
                            if (srgb && srgb.getAttribute("val")) {
                                themeColorMap[name] = "#" + srgb.getAttribute("val");
                            }
                        });
                    }
                }
            }
        }
    } catch (e) {
        console.warn("Theme okunamadı:", e);
    }


    let slidePaths = Object.keys(zip.files)
        .filter(p => /ppt\/slides\/slide\d+\.xml$/.test(p))
        .sort((a, b) => parseInt(a.match(/slide(\d+)/)[1]) - parseInt(b.match(/slide(\d+)/)[1]));

    let slidesJson = [];
    let totalElements = 0;

    for (let i = 0; i < slidePaths.length; i++) {
        let p = slidePaths[i];

        // slide xml
        let xml = await zip.file(p).async("text");
        let doc = parser.parseFromString(xml, "application/xml");

        // slide rels (image mapping için)
        let relsPath = p.replace("slides/", "slides/_rels/") + ".rels";
        let relsMap = {};
        if (zip.file(relsPath)) {
            let relsXml = await zip.file(relsPath).async("text");
            let relsDoc = parser.parseFromString(relsXml, "application/xml");
            let rels = relsDoc.querySelectorAll("Relationship");
            rels.forEach(r => {
                let id = r.getAttribute("Id");
                let target = r.getAttribute("Target"); // ../media/image1.png
                relsMap[id] = target;
            });
        }

        let slideNumber = parseInt(p.match(/slide(\d+)/)[1]);
        let slideJson = {
            slideNumber,
            size: {
                widthEmu: slideW,
                heightEmu: slideH,
                widthCm: emuToCm(slideW),
                heightCm: emuToCm(slideH)
            },
            elements: []
        };

        let zIndex = 0;
        // counter for generic shapes (non-text)
        let anonShapeCount = 0;

        // ========== SHAPES ==========
        // Parse all shapes, capturing both text and plain shapes with fill/border
        let shapes = doc.querySelectorAll("p\\:sp,sp");
        shapes.forEach((sp, idxShape) => {
            // Determine geometry for this shape
            let xfrm = sp.querySelector("a\\:xfrm,xfrm");
            let geom = null;
            if (xfrm) {
                let off = xfrm.querySelector("a\\:off,off");
                let ext = xfrm.querySelector("a\\:ext,ext");
                if (off && ext) {
                    let x = +off.getAttribute("x");
                    let y = +off.getAttribute("y");
                    let w = +ext.getAttribute("cx");
                    let h = +ext.getAttribute("cy");
                    geom = {
                        xEmu: x,
                        yEmu: y,
                        wEmu: w,
                        hEmu: h,
                        xCm: emuToCm(x),
                        yCm: emuToCm(y),
                        wCm: emuToCm(w),
                        hCm: emuToCm(h)
                    };
                }
            }

            // base style from spPr (fill/border)
            let spPr = sp.querySelector("a\\:spPr,spPr");
            let baseStyle = {
                fillColor: null,
                outlineColor: null,
                outlineWidth: null
            };

            if (spPr) {
                // 1) FILL COLOR
                let fillClr = spPr.querySelector("a\\:solidFill > a\\:srgbClr, solidFill > srgbClr");
                if (fillClr && fillClr.getAttribute("val")) {
                    baseStyle.fillColor = "#" + fillClr.getAttribute("val");
                } else {
                    // srgb yoksa schemeClr (accent1, accent2 vs.) dene
                    let schemeClr = spPr.querySelector("a\\:solidFill > a\\:schemeClr, solidFill > schemeClr");
                    if (schemeClr && schemeClr.getAttribute("val")) {
                        let key = schemeClr.getAttribute("val");
                        if (themeColorMap[key]) {
                            baseStyle.fillColor = themeColorMap[key];
                        }
                    }
                }

                // 2) OUTLINE (border) COLOR + WIDTH
                let ln = spPr.querySelector("a\\:ln,ln");
                if (ln) {
                    let lnClr = ln.querySelector("a\\:solidFill > a\\:srgbClr, solidFill > srgbClr");
                    if (lnClr && lnClr.getAttribute("val")) {
                        baseStyle.outlineColor = "#" + lnClr.getAttribute("val");
                    } else {
                        let lnScheme = ln.querySelector("a\\:solidFill > a\\:schemeClr, solidFill > schemeClr");
                        if (lnScheme && lnScheme.getAttribute("val")) {
                            let key = lnScheme.getAttribute("val");
                            if (themeColorMap[key]) {
                                baseStyle.outlineColor = themeColorMap[key];
                            }
                        }
                    }
                    if (ln.getAttribute("w")) {
                        baseStyle.outlineWidth = parseInt(ln.getAttribute("w"), 10);
                    }
                }
            }


            // Determine if shape has text
            let txBody = sp.querySelector("a\\:txBody,txBody");
            let hasText = !!txBody;

            // Determine placeholder type
            let ph = sp.querySelector("p\\:ph,ph");
            let placeholderType = ph ? ph.getAttribute("type") || null : null;

            // Determine preset geometry type (e.g. rect, roundRect)
            let prstGeom = sp.querySelector("a\\:prstGeom,prstGeom");
            let prstType = prstGeom ? prstGeom.getAttribute("prst") || null : null;

            if (hasText) {
                // process paragraphs and runs
                let paragraphs = txBody.querySelectorAll("a\\:p,p");
                if (!paragraphs.length) return;

                let allText = [];
                let paraJson = [];
                paragraphs.forEach(pNode => {
                    let pText = [];
                    let runs = [];
                    let pPr = pNode.querySelector("a\\:pPr,pPr");

                    // Bullet type & level
                    let bulletType = "none";
                    let bulletLevel = 0;
                    if (pPr) {
                        let buAuto = pPr.querySelector("a\\:buAutoNum,buAutoNum");
                        let buChar = pPr.querySelector("a\\:buChar,buChar");
                        let buNone = pPr.querySelector("a\\:buNone,buNone");
                        if (buAuto) bulletType = "number";
                        else if (buChar) bulletType = "bullet";
                        else if (buNone) bulletType = "none";
                        let lvl = pPr.getAttribute("lvl");
                        if (lvl) bulletLevel = parseInt(lvl, 10);
                    }
                    // alignment
                    let align = null;
                    if (pPr) {
                        let algnNode = pPr.querySelector("a\\:algn,algn");
                        if (algnNode && algnNode.textContent) align = algnNode.textContent;
                    }
                    let rNodes = pNode.querySelectorAll("a\\:r,r");
                    rNodes.forEach(r => {
                        let tNode = r.querySelector("a\\:t,t");
                        if (!tNode) return;
                        let text = tNode.textContent;
                        if (!text.trim()) return;
                        pText.push(text);
                        let rPr = r.querySelector("a\\:rPr,rPr");
                        let runStyle = {
                            fontFamily: null,
                            fontSize: null,
                            bold: false,
                            italic: false,
                            underline: false,
                            color: null
                        };
                        if (rPr) {
                            if (rPr.getAttribute("sz")) runStyle.fontSize = parseInt(rPr.getAttribute("sz"), 10) / 100;
                            runStyle.bold = rPr.getAttribute("b") === "1" || rPr.getAttribute("b") === "true";
                            runStyle.italic = rPr.getAttribute("i") === "1" || rPr.getAttribute("i") === "true";
                            let u = rPr.getAttribute("u");
                            runStyle.underline = !!u && u !== "none";
                            let latin = rPr.querySelector("a\\:latin,latin");
                            if (latin) runStyle.fontFamily = latin.getAttribute("typeface") || null;
                            let clr = rPr.querySelector("a\\:solidFill > a\\:srgbClr, solidFill > srgbClr");
                            if (clr) runStyle.color = "#" + clr.getAttribute("val");
                            else {
                                // Eğer metin rengi yoksa shape fillColor fallback olsun
                                if (baseStyle.fillColor) {
                                    runStyle.color = baseStyle.fillColor;
                                }
                            }

                        }
                        runs.push({ text, style: runStyle });
                    });
                    if (pText.length) {
                        allText.push(pText.join(""));
                        paraJson.push({
                            text: pText.join(""),
                            bullet: { type: bulletType, level: bulletLevel },
                            align,
                            runs
                        });
                    }
                });
                if (!allText.length) return;
                // Top-level style from first run
                let firstRunStyle = paraJson[0]?.runs[0]?.style || {};
                let topStyle = {
                    fontFamily: firstRunStyle.fontFamily || null,
                    fontSize: firstRunStyle.fontSize || null,
                    bold: firstRunStyle.bold || false,
                    italic: firstRunStyle.italic || false,
                    underline: firstRunStyle.underline || false,
                    align: paraJson[0]?.align || null,
                    color: firstRunStyle.color || null,
                    fillColor: baseStyle.fillColor,
                    outlineColor: baseStyle.outlineColor,
                    outlineWidth: baseStyle.outlineWidth
                };
                // Map placeholder type to text type
                let type = "textbox";
                if (placeholderType === "title" || placeholderType === "ctrTitle") type = "title";
                else if (placeholderType === "body") type = "content";
                else if (placeholderType === "sldNum") type = "slideNumber";
                else if (placeholderType === "dt") type = "date";
                let elementJson = {
                    id: `s${slideNumber}-el${idxShape + 1}`,
                    kind: "shape",
                    type,
                    placeholderType,
                    zIndex: ++zIndex,
                    geometry: geom,
                    content: {
                        text: allText.join("\n"),
                        paragraphs: paraJson
                    },
                    style: topStyle,
                    shapeType: prstType
                };
                slideJson.elements.push(elementJson);
            } else {
                // Plain shape without text
                let elementJson = {
                    id: `s${slideNumber}-sh${++anonShapeCount}`,
                    kind: "shape",
                    type: "shape",
                    placeholderType: null,
                    zIndex: ++zIndex,
                    geometry: geom,
                    content: null,
                    style: {
                        fontFamily: null,
                        fontSize: null,
                        bold: false,
                        italic: false,
                        underline: false,
                        align: null,
                        color: null,
                        fillColor: baseStyle.fillColor,
                        outlineColor: baseStyle.outlineColor,
                        outlineWidth: baseStyle.outlineWidth
                    },
                    shapeType: prstType
                };
                slideJson.elements.push(elementJson);
            }
        });

        // ========== IMAGES ==========
        let pics = doc.querySelectorAll("p\\:pic,pic");

        for (let idxPic = 0; idxPic < pics.length; idxPic++) {
            let pic = pics[idxPic];

            let xfrm = pic.querySelector("a\\:xfrm,xfrm");
            let geom = null;
            if (xfrm) {
                let off = xfrm.querySelector("a\\:off,off");
                let ext = xfrm.querySelector("a\\:ext,ext");
                if (off && ext) {
                    let x = +off.getAttribute("x");
                    let y = +off.getAttribute("y");
                    let w = +ext.getAttribute("cx");
                    let h = +ext.getAttribute("cy");
                    geom = {
                        xEmu: x,
                        yEmu: y,
                        wEmu: w,
                        hEmu: h,
                        xCm: emuToCm(x),
                        yCm: emuToCm(y),
                        wCm: emuToCm(w),
                        hCm: emuToCm(h)
                    };
                }
            }

            // image relationship
            let blip = pic.querySelector("a\\:blipFill > a\\:blip, blipFill > blip");
            let imageRef = null;
            let imageName = null;
            if (blip) {
                let rId = blip.getAttribute("r:embed") || blip.getAttribute("embed");
                if (rId && relsMap[rId]) {
                    let target = relsMap[rId].replace("..", "ppt");
                    imageRef = target;
                    imageName = target.split("/").pop();
                }
            }

            let elem = {
                id: `s${slideNumber}-img${idxPic + 1}`,
                kind: "picture",
                type: "image",
                zIndex: ++zIndex,
                geometry: geom,
                imageRef,
                imageName,
                imageBase64: null
            };

            slideJson.elements.push(elem);

            // Base64'i async olarak ekliyoruz ✔
            if (imageRef) {
                elem.imageBase64 = await getBase64(zip, imageRef);
            }
        }

        // ========== CHART FRAMES ==========
        // Capture chart frames (e.g. donut, bar, pie) which are not treated as images
        let graphics = doc.querySelectorAll("p\\:graphicFrame, graphicFrame");
        graphics.forEach((gf, idxG) => {
            // geometry
            let gfXfrm = gf.querySelector("a\\:xfrm,xfrm");
            let geom = null;
            if (gfXfrm) {
                let off = gfXfrm.querySelector("a\\:off,off");
                let ext = gfXfrm.querySelector("a\\:ext,ext");
                if (off && ext) {
                    let x = +off.getAttribute("x");
                    let y = +off.getAttribute("y");
                    let w = +ext.getAttribute("cx");
                    let h = +ext.getAttribute("cy");
                    geom = {
                        xEmu: x,
                        yEmu: y,
                        wEmu: w,
                        hEmu: h,
                        xCm: emuToCm(x),
                        yCm: emuToCm(y),
                        wCm: emuToCm(w),
                        hCm: emuToCm(h)
                    };
                }
            }
            // find chart id
            let chartNode = gf.querySelector("a\\:graphicData > c\\:chart, graphicData > chart");
            let chartRel = null;
            if (chartNode) {
                chartRel = chartNode.getAttribute("r:id") || chartNode.getAttribute("id");
            }
            slideJson.elements.push({
                id: `s${slideNumber}-chart${idxG + 1}`,
                kind: "chart",
                type: "chart",
                placeholderType: null,
                zIndex: ++zIndex,
                geometry: geom,
                content: null,
                style: {
                    fillColor: null,
                    outlineColor: null,
                    outlineWidth: null
                },
                chartRelId: chartRel
            });
        });


        slidesJson.push(slideJson);
        totalElements += slideJson.elements.length;

        // Update progress bar after each slide processed
        let progressBar = document.getElementById("progressBar");
        let progressLabel = document.getElementById("progressLabel");
        if (progressBar) {
            let percent = Math.round(((i + 1) / slidePaths.length) * 100);
            progressBar.style.width = percent + "%";
            if (progressLabel) progressLabel.textContent = percent + "%";
        }

        // === SLIDE CARD UI ===
        let textCount = slideJson.elements.filter(el => el.type !== "image").length;
        let imgCount = slideJson.elements.filter(el => el.type === "image").length;

        let card = document.createElement("div");
        card.className = "card";

        card.innerHTML = `
    <div class="flex justify-between items-center mb-2 cursor-pointer group" onclick="toggleCard(this)">
        <div class="space-y-1">
            <h2 class="font-bold text-lg flex items-center">
                <span class="mr-2">📌 Slayt ${slideJson.slideNumber}</span>
                <span class="toggle-icon text-gray-400 group-hover:text-gray-600 text-sm">➕</span>
            </h2>
            <div class="space-x-2">
                <span class="badge">🧩 ${slideJson.elements.length} element</span>
                <span class="badge">🔤 ${textCount} text</span>
                <span class="badge">🖼️ ${imgCount} image</span>
            </div>
        </div>

        <button class="copyBtn text-xs bg-gray-100 px-2 py-1 rounded border">
            JSON Kopyala
        </button>
    </div>

    <pre class="slide-json hidden">${JSON.stringify(slideJson, null, 2)}</pre>
`;



        out.appendChild(card);
    }

    // global json obj
    presentationJson = {
        fileName: file.name,
        slideCount: slidesJson.length,
        slideSize: {
            widthEmu: slideW,
            heightEmu: slideH,
            widthCm: emuToCm(slideW),
            heightCm: emuToCm(slideH)
        },
        slides: slidesJson
    };

    loading.style.display = "none";

    // hide progress after complete
    if (progressContainer) progressContainer.style.display = "none";
    if (progressLabel) progressLabel.style.display = "none";

    // GLOBAL PANEL
    fileInfo.textContent = `Dosya: ${file.name}`;
    statsInfo.textContent =
        `Slayt: ${slidesJson.length} · Toplam element: ${totalElements} · Slide boyutu: `
        + `${presentationJson.slideSize.widthCm}cm × ${presentationJson.slideSize.heightCm}cm`;
    globalPanel.classList.remove("hidden");

    // per-slide copy buttons
    document.querySelectorAll(".copyBtn").forEach(btn => {
        btn.onclick = () => {
            let pre = btn.parentElement.nextElementSibling.innerText;
            navigator.clipboard.writeText(pre);
            btn.innerText = "Kopyalandı ✔";
            setTimeout(() => (btn.innerText = "Bu Slayt JSON'unu Kopyala"), 1200);
        };
    });

    // global copy all
    document.getElementById("btnCopyAll").onclick = () => {
        if (!presentationJson) return;
        navigator.clipboard.writeText(JSON.stringify(presentationJson, null, 2));
        let btn = document.getElementById("btnCopyAll");
        btn.textContent = "Kopyalandı ✔";
        setTimeout(() => (btn.textContent = "Tüm JSON'u Kopyala"), 1200);
    };

    // download json
    document.getElementById("btnDownloadJson").onclick = () => {
        if (!presentationJson) return;
        let blob = new Blob(
            [JSON.stringify(presentationJson, null, 2)],
            { type: "application/json" }
        );
        let url = URL.createObjectURL(blob);
        let a = document.createElement("a");
        a.href = url;
        a.download = (presentationJson.fileName || "presentation") + ".json";
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
    };
});
