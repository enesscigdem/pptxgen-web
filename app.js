// EMU → cm helper
function emuToCm(emu) {
    const v = Number.parseInt(emu, 10)
    if (isNaN(v)) return null
    return +(v / 360000).toFixed(2)
}

const themeColorMap = {}
let presentationJson = null
const lastJson = localStorage.getItem("presentationJson")
if (lastJson) presentationJson = JSON.parse(lastJson)

const mediaPath = "ppt/media/"
const JSZip = window.JSZip // Declare JSZip variable

async function getBase64(zip, path) {
    const file = zip.file(path)
    if (!file) return null
    const bin = await file.async("base64")
    const ext = path.split(".").pop().toLowerCase()
    return `data:image/${ext};base64,${bin}`
}

// Drag and drop functionality
const dropZone = document.getElementById("dropZone")
if (dropZone) {
    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault()
        dropZone.classList.add("dragover")
    })

    dropZone.addEventListener("dragleave", () => {
        dropZone.classList.remove("dragover")
    })

    dropZone.addEventListener("drop", (e) => {
        e.preventDefault()
        dropZone.classList.remove("dragover")
        const files = e.dataTransfer.files
        if (files.length > 0 && files[0].name.endsWith(".pptx")) {
            document.getElementById("fileInput").files = files
            processFile(files[0])
        }
    })
}

// Toggle slide card
function toggleCard(cardEl) {
    cardEl.classList.toggle("expanded")
}

document.getElementById("fileInput").addEventListener("change", async (e) => {
    const file = e.target.files[0]
    if (file) processFile(file)
})

async function processFile(file) {
    if (!file) return

    const progressSection = document.getElementById("progressSection")
    const progressBar = document.getElementById("progressBar")
    const progressLabel = document.getElementById("progressLabel")
    const out = document.getElementById("output")
    const globalPanel = document.getElementById("globalPanel")
    const fileInfo = document.getElementById("fileInfo")
    const statsInfo = document.getElementById("statsInfo")

    out.innerHTML = ""
    progressSection.classList.add("visible")
    globalPanel.classList.remove("visible")
    presentationJson = null

    progressBar.style.width = "0%"
    progressLabel.textContent = "0%"

    const arrayBuffer = await file.arrayBuffer()
    const zip = await JSZip.loadAsync(arrayBuffer)
    const parser = new DOMParser()

    // slide size
    const presXml = await zip.file("ppt/presentation.xml").async("text")
    const presDoc = parser.parseFromString(presXml, "application/xml")
    const sldSz = presDoc.querySelector("p\\:sldSz,sldSz")
    const slideW = +sldSz.getAttribute("cx")
    const slideH = +sldSz.getAttribute("cy")

    // read theme colours
    try {
        const presRelsPath = "ppt/_rels/presentation.xml.rels"
        if (zip.file(presRelsPath)) {
            const presRelsXml = await zip.file(presRelsPath).async("text")
            const presRelsDoc = parser.parseFromString(presRelsXml, "application/xml")
            const rels = presRelsDoc.querySelectorAll("Relationship")
            let themeRel = null
            rels.forEach((r) => {
                const type = r.getAttribute("Type") || ""
                if (type.includes("/theme")) themeRel = r
            })
            if (themeRel) {
                const target = themeRel.getAttribute("Target")
                let themePath = target.replace("..", "ppt")
                if (!themePath.startsWith("ppt/")) themePath = "ppt/" + themePath.replace(/^\/+/, "")
                if (zip.file(themePath)) {
                    const themeXml = await zip.file(themePath).async("text")
                    const themeDoc = parser.parseFromString(themeXml, "application/xml")
                    const clrScheme = themeDoc.querySelector("a\\:clrScheme,clrScheme")
                    if (clrScheme) {
                        Array.from(clrScheme.children).forEach((node) => {
                            const name = node.localName || node.nodeName.split(":").pop()
                            const srgb = node.querySelector("a\\:srgbClr,srgbClr")
                            if (srgb && srgb.getAttribute("val")) {
                                themeColorMap[name] = "#" + srgb.getAttribute("val")
                            }
                        })
                    }
                }
            }
        }
    } catch (e) {
        console.warn("Theme okunamadı:", e)
    }

    const slidePaths = Object.keys(zip.files)
        .filter((p) => /ppt\/slides\/slide\d+\.xml$/.test(p))
        .sort((a, b) => Number.parseInt(a.match(/slide(\d+)/)[1]) - Number.parseInt(b.match(/slide(\d+)/)[1]))

    const slidesJson = []
    let totalElements = 0

    for (let i = 0; i < slidePaths.length; i++) {
        const p = slidePaths[i]
        const xml = await zip.file(p).async("text")
        const doc = parser.parseFromString(xml, "application/xml")

        // slide relationships for images
        const relsPath = p.replace("slides/", "slides/_rels/") + ".rels"
        const relsMap = {}
        if (zip.file(relsPath)) {
            const relsXml = await zip.file(relsPath).async("text")
            const relsDoc = parser.parseFromString(relsXml, "application/xml")
            const rels = relsDoc.querySelectorAll("Relationship")
            rels.forEach((r) => {
                const id = r.getAttribute("Id")
                const target = r.getAttribute("Target")
                relsMap[id] = target
            })
        }

        const slideNumber = Number.parseInt(p.match(/slide(\d+)/)[1])
        const slideJson = {
            slideNumber,
            size: {
                widthEmu: slideW,
                heightEmu: slideH,
                widthCm: emuToCm(slideW),
                heightCm: emuToCm(slideH),
            },
            background: null,
            elements: [],
        }

        // read slide background
        try {
            const bgNode = doc.querySelector("p\\:bg,bg")
            if (bgNode) {
                const bgPr = bgNode.querySelector("p\\:bgPr,bgPr")
                const bgObj = {}
                if (bgPr) {
                    const fillClr = bgPr.querySelector("a\\:solidFill > a\\:srgbClr, solidFill > srgbClr")
                    if (fillClr && fillClr.getAttribute("val")) {
                        bgObj.fillColor = "#" + fillClr.getAttribute("val")
                    } else {
                        const schemeClr = bgPr.querySelector("a\\:solidFill > a\\:schemeClr, solidFill > schemeClr")
                        if (schemeClr && schemeClr.getAttribute("val")) {
                            const key = schemeClr.getAttribute("val")
                            if (themeColorMap[key]) {
                                bgObj.fillColor = themeColorMap[key]
                            }
                        }
                    }
                    const blip = bgPr.querySelector("a\\:blip,blip")
                    if (blip) {
                        const rId = blip.getAttribute("r:embed") || blip.getAttribute("embed")
                        if (rId && relsMap[rId]) {
                            const target = relsMap[rId]
                            const targetPath = target.replace(/\.\.\//g, "ppt/")
                            bgObj.imageRef = targetPath
                            try {
                                bgObj.imageBase64 = await getBase64(zip, targetPath)
                            } catch (err) {
                                console.warn("Background image base64 extraction failed", err)
                            }
                        }
                    }
                    if (Object.keys(bgObj).length > 0) {
                        slideJson.background = bgObj
                    }
                }
            }
        } catch (bgErr) {
            console.warn("Slide background parse error", bgErr)
        }

        let zIndex = 0
        let anonShapeCount = 0
        // Flag to disable legacy extraction logic. We will perform our own extraction later.
        const skipOldExtraction = true

        // SHAPES
        // If skipOldExtraction is true, legacy extraction will be skipped. Our custom extractor runs after this block.
        if (!skipOldExtraction) {
            const shapes = doc.querySelectorAll("p\\:sp,sp")
            shapes.forEach((sp, idxShape) => {
                const xfrm = sp.querySelector("a\\:xfrm,xfrm")
                let geom = null
                if (xfrm) {
                    const off = xfrm.querySelector("a\\:off,off")
                    const ext = xfrm.querySelector("a\\:ext,ext")
                    if (off && ext) {
                        const x = +off.getAttribute("x")
                        const y = +off.getAttribute("y")
                        const w = +ext.getAttribute("cx")
                        const h = +ext.getAttribute("cy")
                        let rotVal = 0
                        if (xfrm.getAttribute("rot")) {
                            const rotAttr = Number.parseInt(xfrm.getAttribute("rot"), 10)
                            if (!isNaN(rotAttr)) rotVal = rotAttr / 60000.0
                        }
                        geom = {
                            xEmu: x,
                            yEmu: y,
                            wEmu: w,
                            hEmu: h,
                            xCm: emuToCm(x),
                            yCm: emuToCm(y),
                            wCm: emuToCm(w),
                            hCm: emuToCm(h),
                            rot: rotVal,
                        }
                    }
                }
                // base style from spPr
                const spPr = sp.querySelector("a\\:spPr,spPr")
                const baseStyle = { fillColor: null, outlineColor: null, outlineWidth: null }
                if (spPr) {
                    const fillClr = spPr.querySelector("a\\:solidFill > a\\:srgbClr, solidFill > srgbClr")
                    if (fillClr && fillClr.getAttribute("val")) {
                        baseStyle.fillColor = "#" + fillClr.getAttribute("val")
                    } else {
                        const schemeClr = spPr.querySelector("a\\:solidFill > a\\:schemeClr, solidFill > schemeClr")
                        if (schemeClr && schemeClr.getAttribute("val")) {
                            const key = schemeClr.getAttribute("val")
                            if (themeColorMap[key]) {
                                baseStyle.fillColor = themeColorMap[key]
                            }
                        }
                    }
                    const ln = spPr.querySelector("a\\:ln,ln")
                    if (ln) {
                        const lnClr = ln.querySelector("a\\:solidFill > a\\:srgbClr, solidFill > srgbClr")
                        if (lnClr && lnClr.getAttribute("val")) {
                            baseStyle.outlineColor = "#" + lnClr.getAttribute("val")
                        } else {
                            const lnScheme = ln.querySelector("a\\:solidFill > a\\:schemeClr, solidFill > schemeClr")
                            if (lnScheme && lnScheme.getAttribute("val")) {
                                const key = lnScheme.getAttribute("val")
                                if (themeColorMap[key]) baseStyle.outlineColor = themeColorMap[key]
                            }
                        }
                        if (ln.getAttribute("w")) baseStyle.outlineWidth = Number.parseInt(ln.getAttribute("w"), 10)
                    }
                }

                const txBody = sp.querySelector("a\\:txBody,txBody")
                const hasText = !!txBody
                const ph = sp.querySelector("p\\:ph,ph")
                const placeholderType = ph ? ph.getAttribute("type") || null : null
                const prstGeom = sp.querySelector("a\\:prstGeom,prstGeom")
                const prstType = prstGeom ? prstGeom.getAttribute("prst") || null : null

                if (hasText) {
                    const paragraphs = txBody.querySelectorAll("a\\:p,p")
                    const allText = []
                    const paraJson = []
                    paragraphs.forEach((pNode) => {
                        const pText = []
                        const runs = []
                        const pPr = pNode.querySelector("a\\:pPr,pPr")
                        let bulletType = "none"
                        let bulletLevel = 0
                        if (pPr) {
                            const buAuto = pPr.querySelector("a\\:buAutoNum,buAutoNum")
                            const buChar = pPr.querySelector("a\\:buChar,buChar")
                            const buNone = pPr.querySelector("a\\:buNone,buNone")
                            if (buAuto) bulletType = "number"
                            else if (buChar) bulletType = "bullet"
                            else if (buNone) bulletType = "none"
                            const lvl = pPr.getAttribute("lvl")
                            if (lvl) bulletLevel = Number.parseInt(lvl, 10)
                        }
                        let align = null
                        if (pPr) {
                            const algnNode = pPr.querySelector("a\\:algn,algn")
                            if (algnNode && algnNode.textContent) align = algnNode.textContent
                        }
                        const rNodes = pNode.querySelectorAll("a\\:r,r")
                        rNodes.forEach((r) => {
                            const tNode = r.querySelector("a\\:t,t")
                            if (!tNode) return
                            const text = tNode.textContent
                            if (!text.trim()) return
                            pText.push(text)
                            const rPr = r.querySelector("a\\:rPr,rPr")
                            const runStyle = {
                                fontFamily: null,
                                fontSize: null,
                                bold: false,
                                italic: false,
                                underline: false,
                                color: null,
                            }
                            if (rPr) {
                                if (rPr.getAttribute("sz")) runStyle.fontSize = Number.parseInt(rPr.getAttribute("sz"), 10) / 100
                                runStyle.bold = rPr.getAttribute("b") === "1" || rPr.getAttribute("b") === "true"
                                runStyle.italic = rPr.getAttribute("i") === "1" || rPr.getAttribute("i") === "true"
                                const u = rPr.getAttribute("u")
                                runStyle.underline = !!u && u !== "none"
                                const latin = rPr.querySelector("a\\:latin,latin")
                                if (latin) runStyle.fontFamily = latin.getAttribute("typeface") || null
                                const clr = rPr.querySelector("a\\:solidFill > a\\:srgbClr, solidFill > srgbClr")
                                if (clr) runStyle.color = "#" + clr.getAttribute("val")
                                else if (baseStyle.fillColor) runStyle.color = baseStyle.fillColor
                            }
                            runs.push({ text, style: runStyle })
                        })
                        if (pText.length) {
                            allText.push(pText.join(""))
                            paraJson.push({
                                text: pText.join(""),
                                bullet: { type: bulletType, level: bulletLevel },
                                align,
                                runs,
                            })
                        }
                    })
                    if (allText.length) {
                        const firstRunStyle = paraJson[0]?.runs[0]?.style || {}
                        const topStyle = {
                            fontFamily: firstRunStyle.fontFamily || null,
                            fontSize: firstRunStyle.fontSize || null,
                            bold: firstRunStyle.bold || false,
                            italic: firstRunStyle.italic || false,
                            underline: firstRunStyle.underline || false,
                            align: paraJson[0]?.align || null,
                            color: firstRunStyle.color || null,
                            fillColor: baseStyle.fillColor,
                            outlineColor: baseStyle.outlineColor,
                            outlineWidth: baseStyle.outlineWidth,
                        }
                        let typeVal = "textbox"
                        if (placeholderType === "title" || placeholderType === "ctrTitle") typeVal = "title"
                        else if (placeholderType === "body") typeVal = "content"
                        else if (placeholderType === "sldNum") typeVal = "slideNumber"
                        else if (placeholderType === "dt") typeVal = "date"
                        const elementJson = {
                            id: `s${slideNumber}-el${idxShape + 1}`,
                            kind: "shape",
                            type: typeVal,
                            placeholderType,
                            zIndex: ++zIndex,
                            geometry: geom,
                            content: {
                                text: allText.join("\n"),
                                paragraphs: paraJson,
                            },
                            style: topStyle,
                            shapeType: prstType,
                        }
                        slideJson.elements.push(elementJson)
                        return
                    }
                }

                // plain shape
                const elementJson = {
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
                        outlineWidth: baseStyle.outlineWidth,
                    },
                    shapeType: prstType,
                }
                slideJson.elements.push(elementJson)
            })

            // images
            const pics = doc.querySelectorAll("p\\:pic,pic")
            for (let idxPic = 0; idxPic < pics.length; idxPic++) {
                const pic = pics[idxPic]
                const xfrm = pic.querySelector("a\\:xfrm,xfrm")
                let geom = null
                if (xfrm) {
                    const off = xfrm.querySelector("a\\:off,off")
                    const ext = xfrm.querySelector("a\\:ext,ext")
                    if (off && ext) {
                        const x = +off.getAttribute("x")
                        const y = +off.getAttribute("y")
                        const w = +ext.getAttribute("cx")
                        const h = +ext.getAttribute("cy")
                        let rotVal = 0
                        if (xfrm.getAttribute("rot")) {
                            const rotAttr = Number.parseInt(xfrm.getAttribute("rot"), 10)
                            if (!isNaN(rotAttr)) rotVal = rotAttr / 60000.0
                        }
                        geom = {
                            xEmu: x,
                            yEmu: y,
                            wEmu: w,
                            hEmu: h,
                            xCm: emuToCm(x),
                            yCm: emuToCm(y),
                            wCm: emuToCm(w),
                            hCm: emuToCm(h),
                            rot: rotVal,
                        }
                    }
                }
                const blip = pic.querySelector("a\\:blipFill > a\\:blip, blipFill > blip")
                let imageRef = null
                let imageName = null
                if (blip) {
                    const rId = blip.getAttribute("r:embed") || blip.getAttribute("embed")
                    if (rId && relsMap[rId]) {
                        const target = relsMap[rId].replace("..", "ppt")
                        imageRef = target
                        imageName = target.split("/").pop()
                    }
                }
                const elem = {
                    id: `s${slideNumber}-img${idxPic + 1}`,
                    kind: "picture",
                    type: "image",
                    zIndex: ++zIndex,
                    geometry: geom,
                    imageRef,
                    imageName,
                    imageBase64: null,
                }
                slideJson.elements.push(elem)
                if (imageRef) {
                    elem.imageBase64 = await getBase64(zip, imageRef)
                }
            }

            // charts
            const graphics = doc.querySelectorAll("p\\:graphicFrame, graphicFrame")
            graphics.forEach((gf, idxG) => {
                let geom = null
                const gfXfrm = gf.querySelector("a\\:xfrm,xfrm")
                if (gfXfrm) {
                    const off = gfXfrm.querySelector("a\\:off,off")
                    const ext = gfXfrm.querySelector("a\\:ext,ext")
                    if (off && ext) {
                        const x = +off.getAttribute("x")
                        const y = +off.getAttribute("y")
                        const w = +ext.getAttribute("cx")
                        const h = +ext.getAttribute("cy")
                        let rotVal = 0
                        if (gfXfrm.getAttribute("rot")) {
                            const rotAttr = Number.parseInt(gfXfrm.getAttribute("rot"), 10)
                            if (!isNaN(rotAttr)) rotVal = rotAttr / 60000.0
                        }
                        geom = {
                            xEmu: x,
                            yEmu: y,
                            wEmu: w,
                            hEmu: h,
                            xCm: emuToCm(x),
                            yCm: emuToCm(y),
                            wCm: emuToCm(w),
                            hCm: emuToCm(h),
                            rot: rotVal,
                        }
                    }
                }
                const chartNode = gf.querySelector("a\\:graphicData > c\\:chart, graphicData > chart")
                let chartRel = null
                if (chartNode) {
                    chartRel = chartNode.getAttribute("r:id") || chartNode.getAttribute("id")
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
                        outlineWidth: null,
                    },
                    chartRelId: chartRel,
                })
            })

            // icons inside graphic frames
            try {
                const gfAll = doc.querySelectorAll("p\\:graphicFrame,graphicFrame")
                let iconCounter = 0
                gfAll.forEach((gf) => {
                    const isChart = gf.querySelector("a\\:graphicData > c\\:chart, graphicData > chart")
                    if (isChart) return
                    let geom2 = null
                    const gfXfrm2 = gf.querySelector("a\\:xfrm,xfrm")
                    if (gfXfrm2) {
                        const off2 = gfXfrm2.querySelector("a\\:off,off")
                        const ext2 = gfXfrm2.querySelector("a\\:ext,ext")
                        if (off2 && ext2) {
                            const x = +off2.getAttribute("x")
                            const y = +off2.getAttribute("y")
                            const w = +ext2.getAttribute("cx")
                            const h = +ext2.getAttribute("cy")
                            let rotVal2 = 0
                            if (gfXfrm2.getAttribute("rot")) {
                                const rotAttr2 = Number.parseInt(gfXfrm2.getAttribute("rot"), 10)
                                if (!isNaN(rotAttr2)) rotVal2 = rotAttr2 / 60000.0
                            }
                            geom2 = {
                                xEmu: x,
                                yEmu: y,
                                wEmu: w,
                                hEmu: h,
                                xCm: emuToCm(x),
                                yCm: emuToCm(y),
                                wCm: emuToCm(w),
                                hCm: emuToCm(h),
                                rot: rotVal2,
                            }
                        }
                    }
                    const blip2 = gf.querySelector("a\\:blip,blip")
                    if (blip2) {
                        const rId2 = blip2.getAttribute("r:embed") || blip2.getAttribute("embed")
                        if (rId2 && relsMap[rId2]) {
                            const target2 = relsMap[rId2]
                            const targetPath2 = target2.replace(/\.\.\//g, "ppt/")
                            const elem = {
                                id: `s${slideNumber}-icon${++iconCounter}`,
                                kind: "picture",
                                type: "icon",
                                zIndex: ++zIndex,
                                geometry: geom2,
                                imageRef: targetPath2,
                                imageName: targetPath2.split("/").pop(),
                                imageBase64: null,
                            }
                            slideJson.elements.push(elem)
                            ;(async () => {
                                try {
                                    elem.imageBase64 = await getBase64(zip, targetPath2)
                                } catch (err) {
                                    console.warn("Icon base64 extraction failed", err)
                                }
                            })()
                        }
                    }
                })
            } catch (iconErr) {
                console.warn("Icon extraction error", iconErr)
            }

        } // end legacy extraction guard

        // --- Custom ordered extraction begins here ---
        // Prepare a fresh elements array that preserves the original z-order of slide contents.
        const orderedElements = []
        let orderedZ = 0

        // Helper to extract run style with support for theme scheme colours
        function _getRunStyle(rPr, baseStyle) {
            const runStyle = {
                fontFamily: null,
                fontSize: null,
                bold: false,
                italic: false,
                underline: false,
                color: null,
            }
            if (rPr) {
                if (rPr.getAttribute("sz")) runStyle.fontSize = Number.parseInt(rPr.getAttribute("sz"), 10) / 100
                runStyle.bold = rPr.getAttribute("b") === "1" || rPr.getAttribute("b") === "true"
                runStyle.italic = rPr.getAttribute("i") === "1" || rPr.getAttribute("i") === "true"
                const u = rPr.getAttribute("u")
                runStyle.underline = !!u && u !== "none"
                const latin = rPr.querySelector("a\\:latin,latin")
                if (latin) runStyle.fontFamily = latin.getAttribute("typeface") || null
                let clrNode = rPr.querySelector("a\\:solidFill > a\\:srgbClr, solidFill > srgbClr")
                if (clrNode && clrNode.getAttribute("val")) {
                    runStyle.color = "#" + clrNode.getAttribute("val")
                } else {
                    const schemeClr = rPr.querySelector("a\\:solidFill > a\\:schemeClr, solidFill > schemeClr")
                    if (schemeClr && schemeClr.getAttribute("val")) {
                        const key = schemeClr.getAttribute("val")
                        if (themeColorMap[key]) runStyle.color = themeColorMap[key]
                    } else if (baseStyle.fillColor) {
                        runStyle.color = baseStyle.fillColor
                    }
                }
            }
            return runStyle
        }

        function _getGeometry(xfrm) {
            if (!xfrm) return null
            const off = xfrm.querySelector("a\\:off,off")
            const ext = xfrm.querySelector("a\\:ext,ext")
            if (off && ext) {
                const x = +off.getAttribute("x")
                const y = +off.getAttribute("y")
                const w = +ext.getAttribute("cx")
                const h = +ext.getAttribute("cy")
                let rotVal = 0
                if (xfrm.getAttribute("rot")) {
                    const rotAttr = Number.parseInt(xfrm.getAttribute("rot"), 10)
                    if (!isNaN(rotAttr)) rotVal = rotAttr / 60000.0
                }
                return {
                    xEmu: x,
                    yEmu: y,
                    wEmu: w,
                    hEmu: h,
                    xCm: emuToCm(x),
                    yCm: emuToCm(y),
                    wCm: emuToCm(w),
                    hCm: emuToCm(h),
                    rot: rotVal,
                }
            }
            return null
        }

        function _parseBaseStyle(spPr) {
            const style = { fillColor: null, outlineColor: null, outlineWidth: null }
            if (!spPr) return style
            const fillClr = spPr.querySelector("a\\:solidFill > a\\:srgbClr, solidFill > srgbClr")
            if (fillClr && fillClr.getAttribute("val")) {
                style.fillColor = "#" + fillClr.getAttribute("val")
            } else {
                const schemeClr = spPr.querySelector("a\\:solidFill > a\\:schemeClr, solidFill > schemeClr")
                if (schemeClr && schemeClr.getAttribute("val")) {
                    const key = schemeClr.getAttribute("val")
                    if (themeColorMap[key]) style.fillColor = themeColorMap[key]
                }
            }
            const ln = spPr.querySelector("a\\:ln,ln")
            if (ln) {
                const lnClr = ln.querySelector("a\\:solidFill > a\\:srgbClr, solidFill > srgbClr")
                if (lnClr && lnClr.getAttribute("val")) {
                    style.outlineColor = "#" + lnClr.getAttribute("val")
                } else {
                    const lnScheme = ln.querySelector("a\\:solidFill > a\\:schemeClr, solidFill > schemeClr")
                    if (lnScheme && lnScheme.getAttribute("val")) {
                        const key = lnScheme.getAttribute("val")
                        if (themeColorMap[key]) style.outlineColor = themeColorMap[key]
                    }
                }
                if (ln.getAttribute("w")) style.outlineWidth = Number.parseInt(ln.getAttribute("w"), 10)
            }
            return style
        }

        function _processShape(sp) {
            const xfrm = sp.querySelector("a\\:xfrm,xfrm")
            const geom = _getGeometry(xfrm)
            const spPr = sp.querySelector("a\\:spPr,spPr")
            const baseStyle = _parseBaseStyle(spPr)
            const txBody = sp.querySelector("a\\:txBody,txBody")
            const ph = sp.querySelector("p\\:ph,ph")
            const placeholderType = ph ? ph.getAttribute("type") || null : null
            const prstGeom = sp.querySelector("a\\:prstGeom,prstGeom")
            const prstType = prstGeom ? prstGeom.getAttribute("prst") || null : null
            if (txBody) {
                const paragraphs = txBody.querySelectorAll("a\\:p,p")
                const allText = []
                const paraJson = []
                paragraphs.forEach((pNode, pIndex) => {
                    const pText = []
                    const runs = []
                    const pPr = pNode.querySelector("a\\:pPr,pPr")
                    let bulletType = "none"
                    let bulletLevel = 0
                    let bulletChar = null
                    let bulletStart = null
                    if (pPr) {
                        // Determine bullet/number type and level as well as character and starting index
                        const buAuto = pPr.querySelector("a\\:buAutoNum,buAutoNum")
                        const buChar = pPr.querySelector("a\\:buChar,buChar")
                        const buNone = pPr.querySelector("a\\:buNone,buNone")
                        if (buAuto) {
                            bulletType = "number"
                            const startAttr = buAuto.getAttribute("startAt")
                            if (startAttr) {
                                // Starting index for numbered lists
                                bulletStart = Number.parseInt(startAttr, 10)
                            }
                        } else if (buChar) {
                            bulletType = "bullet"
                            const charAttr = buChar.getAttribute("char") || buChar.getAttribute("char=") || null
                            if (charAttr) {
                                bulletChar = charAttr
                            }
                        } else if (buNone) {
                            bulletType = "none"
                        }
                        const lvl = pPr.getAttribute("lvl")
                        if (lvl) bulletLevel = Number.parseInt(lvl, 10)
                    }
                    let align = null
                    if (pPr) {
                        const algnNode = pPr.querySelector("a\\:algn,algn")
                        if (algnNode && algnNode.textContent) align = algnNode.textContent
                    }
                    const rNodes = pNode.querySelectorAll("a\\:r,r")
                    rNodes.forEach((r) => {
                        const tNode = r.querySelector("a\\:t,t")
                        if (!tNode) return
                        const text = tNode.textContent
                        if (!text.trim()) return
                        pText.push(text)
                        const rPr = r.querySelector("a\\:rPr,rPr")
                        const runStyle = _getRunStyle(rPr, baseStyle)
                        runs.push({ text, style: runStyle })
                    })
                    if (pText.length) {
                        allText.push(pText.join(""))
                        paraJson.push({
                            text: pText.join(""),
                            bullet: {
                                type: bulletType,
                                level: bulletLevel,
                                // char may be undefined for number lists
                                char: typeof bulletChar !== "undefined" ? bulletChar : null,
                                startAt: typeof bulletStart !== "undefined" ? bulletStart : null,
                            },
                            align,
                            runs,
                        })
                    }
                })
                if (allText.length) {
                    const firstRunStyle = paraJson[0]?.runs[0]?.style || {}
                    const topStyle = {
                        fontFamily: firstRunStyle.fontFamily || null,
                        fontSize: firstRunStyle.fontSize || null,
                        bold: firstRunStyle.bold || false,
                        italic: firstRunStyle.italic || false,
                        underline: firstRunStyle.underline || false,
                        align: paraJson[0]?.align || null,
                        color: firstRunStyle.color || null,
                        fillColor: baseStyle.fillColor,
                        outlineColor: baseStyle.outlineColor,
                        outlineWidth: baseStyle.outlineWidth,
                    }
                    let typeVal = "textbox"
                    if (placeholderType === "title" || placeholderType === "ctrTitle") typeVal = "title"
                    else if (placeholderType === "body") typeVal = "content"
                    else if (placeholderType === "sldNum") typeVal = "slideNumber"
                    else if (placeholderType === "dt") typeVal = "date"
                    orderedElements.push({
                        id: `s${slideNumber}-el${orderedElements.length + 1}`,
                        kind: "shape",
                        type: typeVal,
                        placeholderType,
                        zIndex: ++orderedZ,
                        geometry: geom,
                        content: {
                            text: allText.join("\n"),
                            paragraphs: paraJson,
                        },
                        style: topStyle,
                        shapeType: prstType,
                    })
                    return
                }
            }
            // plain shape (no text)
            orderedElements.push({
                id: `s${slideNumber}-sh${orderedElements.length + 1}`,
                kind: "shape",
                type: "shape",
                placeholderType: null,
                zIndex: ++orderedZ,
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
                    outlineWidth: baseStyle.outlineWidth,
                },
                shapeType: prstType,
            })
        }

        async function _processPicture(pic) {
            const xfrm = pic.querySelector("a\\:xfrm,xfrm")
            const geom = _getGeometry(xfrm)
            const blip = pic.querySelector("a\\:blipFill > a\\:blip, blipFill > blip")
            let imageRef = null
            let imageName = null
            if (blip) {
                const rId = blip.getAttribute("r:embed") || blip.getAttribute("embed")
                if (rId && relsMap[rId]) {
                    const tgt = relsMap[rId].replace(/\.\./g, "ppt")
                    imageRef = tgt
                    imageName = tgt.split("/").pop()
                }
            }
            const elem = {
                id: `s${slideNumber}-img${orderedElements.length + 1}`,
                kind: "picture",
                type: "image",
                zIndex: ++orderedZ,
                geometry: geom,
                imageRef,
                imageName,
                imageBase64: null,
            }
            orderedElements.push(elem)
            if (imageRef) {
                try {
                    elem.imageBase64 = await getBase64(zip, imageRef)
                } catch (err) {
                    console.warn("Image base64 extraction failed", err)
                }
            }
        }

        function _processGraphicFrame(gf) {
            const gfXfrm = gf.querySelector("a\\:xfrm,xfrm")
            const geom = _getGeometry(gfXfrm)
            const chartNode = gf.querySelector("a\\:graphicData > c\\:chart, graphicData > chart")
            if (chartNode) {
                let chartRel = chartNode.getAttribute("r:id") || chartNode.getAttribute("id")
                orderedElements.push({
                    id: `s${slideNumber}-chart${orderedElements.length + 1}`,
                    kind: "chart",
                    type: "chart",
                    placeholderType: null,
                    zIndex: ++orderedZ,
                    geometry: geom,
                    content: null,
                    style: {
                        fillColor: null,
                        outlineColor: null,
                        outlineWidth: null,
                    },
                    chartRelId: chartRel,
                })
                return
            }
            const blip2 = gf.querySelector("a\\:blip,blip")
            if (blip2) {
                const rId2 = blip2.getAttribute("r:embed") || blip2.getAttribute("embed")
                if (rId2 && relsMap[rId2]) {
                    const target2 = relsMap[rId2]
                    const targetPath2 = target2.replace(/\.{2}\//g, "ppt/")
                    const elem = {
                        id: `s${slideNumber}-icon${orderedElements.length + 1}`,
                        kind: "picture",
                        type: "icon",
                        zIndex: ++orderedZ,
                        geometry: geom,
                        imageRef: targetPath2,
                        imageName: targetPath2.split("/").pop(),
                        imageBase64: null,
                    }
                    orderedElements.push(elem)
                    ;(async () => {
                        try {
                            elem.imageBase64 = await getBase64(zip, targetPath2)
                        } catch (err) {
                            console.warn("Icon base64 extraction failed", err)
                        }
                    })()
                }
            }
        }

        async function _traverse(node) {
            const children = node ? Array.from(node.children) : []
            for (const child of children) {
                const localName = child.localName || child.nodeName.split(":").pop()
                if (localName === "sp") {
                    _processShape(child)
                } else if (localName === "pic") {
                    await _processPicture(child)
                } else if (localName === "graphicFrame") {
                    _processGraphicFrame(child)
                } else if (localName === "grpSp") {
                    await _traverse(child)
                }
            }
        }

        const spTreeNode = doc.querySelector("p\\:spTree,spTree")
        if (spTreeNode) {
            await _traverse(spTreeNode)
        }

        // Include placeholders from slide layout (e.g., date, slide number, footer)
        try {
            let layoutRel = null
            for (const relId in relsMap) {
                const tgt = relsMap[relId]
                if (/slideLayout\.xml$/i.test(tgt)) {
                    layoutRel = tgt
                    break
                }
            }
            if (layoutRel) {
                let layoutPath = layoutRel.replace(/\.{2}\//g, "ppt/")
                if (!layoutPath.startsWith("ppt/")) layoutPath = "ppt/" + layoutPath.replace(/^\/+/g, "")
                if (zip.file(layoutPath)) {
                    const layoutXml = await zip.file(layoutPath).async("text")
                    const layoutDoc = parser.parseFromString(layoutXml, "application/xml")
                    const layoutSpTree = layoutDoc.querySelector("p\\:spTree,spTree")
                    if (layoutSpTree) {
                        const layoutChildren = Array.from(layoutSpTree.children)
                        layoutChildren.forEach((child) => {
                            const ln = child.localName || child.nodeName.split(":").pop()
                            if (ln === "sp") {
                                const phNode = child.querySelector("p\\:ph,ph")
                                const phType = phNode ? phNode.getAttribute("type") || null : null
                                if (phType === "dt" || phType === "sldNum" || phType === "ftr" || phType === "hdr" || phType === "pgNum") {
                                    const xfrmNode = child.querySelector("a\\:xfrm,xfrm")
                                    const geom = _getGeometry(xfrmNode)
                                    const spPrNode = child.querySelector("a\\:spPr,spPr")
                                    const baseStyle2 = _parseBaseStyle(spPrNode)
                                    const txBodyNode = child.querySelector("a\\:txBody,txBody")
                                    let content = null
                                    if (txBodyNode) {
                                        const paragraphs2 = txBodyNode.querySelectorAll("a\\:p,p")
                                        const allText2 = []
                                        const paraJson2 = []
                                        paragraphs2.forEach((pNode) => {
                                            const pText2 = []
                                            const runs2 = []
                                            const pPr2 = pNode.querySelector("a\\:pPr,pPr")
                                            let bulletType2 = "none"
                                            let bulletLevel2 = 0
                                            let bulletChar2 = null
                                            let bulletStart2 = null
                                            if (pPr2) {
                                                const buAuto2 = pPr2.querySelector("a\\:buAutoNum,buAutoNum")
                                                const buChar2 = pPr2.querySelector("a\\:buChar,buChar")
                                                const buNone2 = pPr2.querySelector("a\\:buNone,buNone")
                                                if (buAuto2) {
                                                    bulletType2 = "number"
                                                    const startAttr2 = buAuto2.getAttribute("startAt")
                                                    if (startAttr2) bulletStart2 = Number.parseInt(startAttr2, 10)
                                                } else if (buChar2) {
                                                    bulletType2 = "bullet"
                                                    const charAttr2 = buChar2.getAttribute("char") || buChar2.getAttribute("char=") || null
                                                    if (charAttr2) bulletChar2 = charAttr2
                                                } else if (buNone2) {
                                                    bulletType2 = "none"
                                                }
                                                const lvl2 = pPr2.getAttribute("lvl")
                                                if (lvl2) bulletLevel2 = Number.parseInt(lvl2, 10)
                                            }
                                            let align2 = null
                                            if (pPr2) {
                                                const algnNode2 = pPr2.querySelector("a\\:algn,algn")
                                                if (algnNode2 && algnNode2.textContent) align2 = algnNode2.textContent
                                            }
                                            const rNodes2 = pNode.querySelectorAll("a\\:r,r")
                                            rNodes2.forEach((r2) => {
                                                const tNode2 = r2.querySelector("a\\:t,t")
                                                if (!tNode2) return
                                                const text2 = tNode2.textContent
                                                if (!text2.trim()) return
                                                pText2.push(text2)
                                                const rPr2 = r2.querySelector("a\\:rPr,rPr")
                                                const runStyle2 = _getRunStyle(rPr2, baseStyle2)
                                                runs2.push({ text: text2, style: runStyle2 })
                                            })
                                            if (pText2.length) {
                                                allText2.push(pText2.join(""))
                                                paraJson2.push({
                                                    text: pText2.join(""),
                                                    bullet: {
                                                        type: bulletType2,
                                                        level: bulletLevel2,
                                                        char: bulletChar2,
                                                        startAt: bulletStart2,
                                                    },
                                                    align: align2,
                                                    runs: runs2,
                                                })
                                            }
                                        })
                                        if (allText2.length) {
                                            content = {
                                                text: allText2.join("\n"),
                                                paragraphs: paraJson2,
                                            }
                                        }
                                    }
                                    const firstRun = content?.paragraphs?.[0]?.runs?.[0] || {}
                                    const topStyle2 = {
                                        fontFamily: firstRun.style?.fontFamily || null,
                                        fontSize: firstRun.style?.fontSize || null,
                                        bold: firstRun.style?.bold || false,
                                        italic: firstRun.style?.italic || false,
                                        underline: firstRun.style?.underline || false,
                                        align: content?.paragraphs?.[0]?.align || null,
                                        color: firstRun.style?.color || null,
                                        fillColor: baseStyle2.fillColor,
                                        outlineColor: baseStyle2.outlineColor,
                                        outlineWidth: baseStyle2.outlineWidth,
                                    }
                                    const typeValue = phType === "sldNum" ? "slideNumber" : phType === "dt" ? "date" : "footer"
                                    orderedElements.push({
                                        id: `s${slideNumber}-layout${orderedElements.length + 1}`,
                                        kind: "shape",
                                        type: typeValue,
                                        placeholderType: phType,
                                        zIndex: ++orderedZ,
                                        geometry: geom,
                                        content,
                                        style: topStyle2,
                                        shapeType: null,
                                    })
                                }
                            }
                        })
                    }
                }
            }
        } catch (layoutErr) {
            console.warn("Layout extraction error", layoutErr)
        }

        // Replace slideJson.elements with orderedElements
        slideJson.elements = orderedElements

        slidesJson.push(slideJson)
        totalElements += slideJson.elements.length

        // update progress bar
        const percent = Math.round(((i + 1) / slidePaths.length) * 100)
        progressBar.style.width = percent + "%"
        progressLabel.textContent = percent + "%"

        // slide card UI
        const textCount = slideJson.elements.filter((el) => el.type !== "image").length
        const imgCount = slideJson.elements.filter((el) => el.type === "image").length

        const card = document.createElement("div")
        card.className = "slide-card"
        card.innerHTML = `
            <div class="slide-header" onclick="toggleCard(this.parentElement)">
                <div class="slide-info">
                    <div class="slide-title">
                        <span class="slide-number">${slideJson.slideNumber}</span>
                        Slide ${slideJson.slideNumber}
                    </div>
                    <div class="slide-badges">
                        <span class="badge badge-element">${slideJson.elements.length} elements</span>
                        <span class="badge badge-text">${textCount} text</span>
                        <span class="badge badge-image">${imgCount} images</span>
                    </div>
                </div>
                <div class="slide-actions">
                    <button class="btn btn-sm btn-secondary copyBtn" onclick="event.stopPropagation()">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <rect x="9" y="9" width="13" height="13" rx="2" ry="2"/>
                            <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>
                        </svg>
                        Copy
                    </button>
                    <span class="toggle-icon">
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <polyline points="6,9 12,15 18,9"/>
                        </svg>
                    </span>
                </div>
            </div>
            <div class="slide-json">
                <pre>${JSON.stringify(slideJson, null, 2)}</pre>
            </div>
        `
        out.appendChild(card)
    }

    presentationJson = {
        fileName: file.name,
        slideCount: slidesJson.length,
        slideSize: {
            widthEmu: slideW,
            heightEmu: slideH,
            widthCm: emuToCm(slideW),
            heightCm: emuToCm(slideH),
        },
        slides: slidesJson,
    }

    // Save to localStorage
    progressSection.classList.remove("visible")

    fileInfo.textContent = file.name
    statsInfo.textContent = `${slidesJson.length} slides • ${totalElements} elements • ${presentationJson.slideSize.widthCm}cm × ${presentationJson.slideSize.heightCm}cm`
    globalPanel.classList.add("visible")

    // Copy buttons for individual slides
    document.querySelectorAll(".copyBtn").forEach((btn) => {
        btn.onclick = (e) => {
            e.stopPropagation()
            const pre = btn.closest(".slide-card").querySelector(".slide-json pre").innerText
            navigator.clipboard.writeText(pre)
            const originalText = btn.innerHTML
            btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20,6 9,17 4,12"/></svg> Copied!`
            setTimeout(() => (btn.innerHTML = originalText), 1500)
        }
    })

    // Copy all JSON
    document.getElementById("btnCopyAll").onclick = () => {
        if (!presentationJson) return
        navigator.clipboard.writeText(JSON.stringify(presentationJson, null, 2))
        const btn = document.getElementById("btnCopyAll")
        const originalHTML = btn.innerHTML
        btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20,6 9,17 4,12"/></svg> Copied!`
        setTimeout(() => (btn.innerHTML = originalHTML), 1500)
    }

    // Download JSON
    document.getElementById("btnDownloadJson").onclick = () => {
        if (!presentationJson) return
        const blob = new Blob([JSON.stringify(presentationJson, null, 2)], { type: "application/json" })
        const url = URL.createObjectURL(blob)
        const a = document.createElement("a")
        a.href = url
        a.download = (presentationJson.fileName || "presentation").replace(".pptx", "") + ".json"
        document.body.appendChild(a)
        a.click()
        a.remove()
        URL.revokeObjectURL(url)
    }
}
