let jsonData = null
let slides = []
let activeSlideIndex = 0
let zoomLevel = 1

document.getElementById("btnRender").onclick = async () => {
    const txt = document.getElementById("jsonInput").value.trim()

    showLoading(true)

    await new Promise((resolve) => setTimeout(resolve, 10))

    if (txt) {
        try {
            jsonData = await parseJSONAsync(txt)
            normalizeSlides()
            await renderSlidesAsync()
        } catch (e) {
            console.error("JSON parse error:", e)
            alert("Invalid JSON format")
        } finally {
            showLoading(false)
        }
        return
    }

    const file = document.getElementById("fileJson").files[0]
    if (!file) {
        showLoading(false)
        return alert("Please paste JSON or select a file")
    }

    try {
        const t = await file.text()
        if (t.length < 500000) {
            document.getElementById("jsonInput").value = t
        } else {
            document.getElementById("jsonInput").value =
                `[Large file: ${(t.length / 1024 / 1024).toFixed(2)} MB - content hidden for performance]`
            document.getElementById("jsonInput").disabled = true
        }

        jsonData = await parseJSONAsync(t)
        normalizeSlides()
        await renderSlidesAsync()
    } catch (e) {
        console.error("File parse error:", e)
        alert("Invalid JSON format")
    } finally {
        showLoading(false)
    }
}

async function parseJSONAsync(text) {
    return new Promise((resolve, reject) => {
        // Use setTimeout to move parsing off the current execution frame
        setTimeout(() => {
            try {
                const result = JSON.parse(text)
                resolve(result)
            } catch (e) {
                reject(e)
            }
        }, 0)
    })
}

function showLoading(show) {
    const container = document.getElementById("canvasContainer")
    const btn = document.getElementById("btnRender")

    if (show) {
        btn.disabled = true
        btn.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin">
        <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
      </svg>
      Processing...
    `
        container.innerHTML = `
      <div style="text-align: center; color: var(--color-text-secondary); padding: 40px;">
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin" style="margin-bottom: 12px;">
          <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
        </svg>
        <p>Loading slides...</p>
      </div>
    `
    } else {
        btn.disabled = false
        btn.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <polygon points="5,3 19,12 5,21 5,3"/>
      </svg>
      Render Slides
    `
    }
}

// Zoom controls
document.getElementById("zoomIn").onclick = () => {
    zoomLevel = Math.min(zoomLevel + 0.1, 2)
    updateZoom()
}

document.getElementById("zoomOut").onclick = () => {
    zoomLevel = Math.max(zoomLevel - 0.1, 0.5)
    updateZoom()
}

document.getElementById("zoomReset").onclick = () => {
    zoomLevel = 1
    updateZoom()
}

function updateZoom() {
    const canvas = document.querySelector(".slideCanvas")
    if (canvas) {
        canvas.style.transform = `scale(${zoomLevel})`
        canvas.style.transformOrigin = "top left"
    }
}

function normalizeSlides() {
    if (jsonData.slides) slides = jsonData.slides
    else if (jsonData.elements) slides = [jsonData]
    else slides = []
}

async function renderSlidesAsync() {
    if (!slides.length) {
        alert("No slides found in JSON.")
        return
    }

    const slideList = document.getElementById("slideList")
    slideList.innerHTML = ""

    const BATCH_SIZE = 20
    for (let i = 0; i < slides.length; i += BATCH_SIZE) {
        const batch = slides.slice(i, i + BATCH_SIZE)

        batch.forEach((s, j) => {
            const idx = i + j
            const btn = document.createElement("div")
            btn.className = "slide-thumb"
            btn.innerHTML = `
<svg width="18" height="18" viewBox="0 0 24 24" fill="none"
     stroke="currentColor" stroke-width="1.6"
     stroke-linecap="round" stroke-linejoin="round"
     style="vertical-align:-5px">

  <!-- frame -->
  <rect x="2.5" y="4" width="19" height="13" rx="2" ry="2"
        fill="var(--thumb-bg,#ffffff)" stroke="var(--thumb-border,#d0d7e1)"/>

  <!-- title block -->
  <rect x="5" y="6" width="9" height="3" rx="0.5"
        fill="var(--primary,#3b82f6)" opacity="0.22"/>

  <!-- lines -->
  <line x1="5" y1="11" x2="17" y2="11" stroke="#94a3b8" stroke-width="1.2"/>
  <line x1="5" y1="13.5" x2="14" y2="13.5" stroke="#94a3b8" stroke-width="1.2"/>

</svg>

        Slide ${s.slideNumber || idx + 1}
      `
            btn.onclick = () => showSlideAsync(idx)
            slideList.appendChild(btn)
        })

        if (i + BATCH_SIZE < slides.length) {
            await new Promise((resolve) => setTimeout(resolve, 0))
        }
    }

    // Show first slide
    await showSlideAsync(0)
}

async function showSlideAsync(index) {
    activeSlideIndex = index
    const container = document.getElementById("canvasContainer")
    container.innerHTML = ""
    const slide = slides[index]

    const canvas = document.createElement("div")
    canvas.className = "slideCanvas"
    canvas.style.width = "960px"
    canvas.style.height = "540px"
    canvas.style.position = "relative"
    canvas.style.background = "#fff"
    canvas.style.transform = `scale(${zoomLevel})`
    canvas.style.transformOrigin = "top left"

    if (slide.background) {
        if (slide.background.imageBase64) {
            canvas.style.backgroundImage = `url(${slide.background.imageBase64})`
            canvas.style.backgroundRepeat = "no-repeat"
            canvas.style.backgroundSize = "cover"
        } else if (slide.background.fillColor) {
            canvas.style.background = slide.background.fillColor
        }
    }
    container.appendChild(canvas)

    const docWidthCm = slide.size?.widthCm || 33.87
    const docHeightCm = slide.size?.heightCm || 19.05
    const scaleX = 960 / (docWidthCm * 37.8)
    const scaleY = 540 / (docHeightCm * 37.8)

    // Sort elements by zIndex
    const elements = [...(slide.elements || [])].sort((a, b) => (a.zIndex || 0) - (b.zIndex || 0))

    const fragment = document.createDocumentFragment()

    const ELEMENT_BATCH_SIZE = 50
    for (let i = 0; i < elements.length; i += ELEMENT_BATCH_SIZE) {
        const batch = elements.slice(i, i + ELEMENT_BATCH_SIZE)

        batch.forEach((el) => {
            const div = createElementDiv(el, scaleX, scaleY)
            if (div) fragment.appendChild(div)
        })

        if (elements.length > ELEMENT_BATCH_SIZE && i + ELEMENT_BATCH_SIZE < elements.length) {
            await new Promise((resolve) => setTimeout(resolve, 0))
        }
    }

    canvas.appendChild(fragment)

    // Update active state in slide list
    ;[...document.getElementById("slideList").children].forEach((x, i) => x.classList.toggle("active", i === index))
}

function createElementDiv(el, scaleX, scaleY) {
    if (!el.geometry) return null

    const div = document.createElement("div")
    div.className = "slide-element"
    div.style.position = "absolute"
    div.style.left = el.geometry.xCm * 37.8 * scaleX + "px"
    div.style.top = el.geometry.yCm * 37.8 * scaleY + "px"
    div.style.width = el.geometry.wCm * 37.8 * scaleX + "px"
    div.style.height = el.geometry.hCm * 37.8 * scaleY + "px"
    div.style.boxSizing = "border-box"
    // Default overflow - will be adjusted per element type
    div.style.overflow = "visible"

    if (el.geometry.rot && el.geometry.rot !== 0) {
        div.style.transform = `rotate(${el.geometry.rot}deg)`
        div.style.transformOrigin = "top left"
    }
    if (el.zIndex !== undefined && el.zIndex !== null) {
        div.style.zIndex = String(el.zIndex)
    }

    // Chart element
    if (el.kind === "chart") {
        div.style.background = "#f8fafc"
        div.style.border = "2px solid #3b82f6"
        div.style.borderRadius = "8px"
        div.style.display = "flex"
        div.style.flexDirection = "column"
        div.style.alignItems = "center"
        div.style.justifyContent = "center"
        div.style.padding = "8px"
        div.style.fontSize = "12px"
        div.style.color = "#1e40af"
        
        if (el.chartData && el.chartData.type) {
            const chartType = el.chartData.type
            const seriesCount = el.chartData.series ? el.chartData.series.length : 0
            div.innerHTML = `
                <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#3b82f6" stroke-width="2" style="margin-bottom: 4px;">
                    <line x1="12" y1="20" x2="12" y2="10"/>
                    <line x1="18" y1="20" x2="18" y2="4"/>
                    <line x1="6" y1="20" x2="6" y2="16"/>
                </svg>
                <div style="text-align: center; font-weight: 600;">${chartType.charAt(0).toUpperCase() + chartType.slice(1)} Chart</div>
                <div style="font-size: 10px; opacity: 0.7;">${seriesCount} series</div>
            `
        } else {
            div.innerHTML = `
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#3b82f6" stroke-width="2">
                    <line x1="12" y1="20" x2="12" y2="10"/>
                    <line x1="18" y1="20" x2="18" y2="4"/>
                    <line x1="6" y1="20" x2="6" y2="16"/>
                </svg>
                <div style="margin-top: 4px; font-size: 11px;">Chart</div>
            `
        }
        return div
    }

    // Image element
    if (el.type === "image" || el.type === "icon") {
        // Allow images to overflow their container to prevent cropping
        div.style.overflow = "visible"
        let isEmf = false
        if (el.imageName) {
            const ext = el.imageName.split(".").pop().toLowerCase()
            if (ext === "emf" || ext === "wmf") isEmf = true
        } else if (el.imageBase64) {
            isEmf = el.imageBase64.startsWith("data:image/emf") || el.imageBase64.startsWith("data:image/wmf")
        }

        if (isEmf) {
            // Try to display EMF/WMF as base64 data URL if available
            if (el.imageBase64) {
                // Create an object URL or try to display it
                // Note: Browsers can't directly display EMF/WMF, but we can show the base64 data
                div.style.background = "#f3f4f6"
                div.style.display = "flex"
                div.style.flexDirection = "column"
                div.style.alignItems = "center"
                div.style.justifyContent = "center"
                div.style.color = "#6b7280"
                div.style.fontSize = "10px"
                div.style.padding = "8px"
                div.innerHTML = `
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="margin-bottom: 4px; opacity: 0.5;">
                        <rect x="3" y="3" width="18" height="18" rx="2" ry="2"/>
                        <circle cx="8.5" cy="8.5" r="1.5"/>
                        <polyline points="21,15 16,10 5,21"/>
                    </svg>
                    <div style="text-align: center;">
                        <div style="font-weight: 600; margin-bottom: 2px;">${el.imageName ? el.imageName.split('.').pop().toUpperCase() : 'EMF/WMF'}</div>
                        <div style="font-size: 9px; opacity: 0.7;">Vector Image</div>
                    </div>
                `
            } else {
                div.style.background = "#f3f4f6"
                div.style.display = "flex"
                div.style.alignItems = "center"
                div.style.justifyContent = "center"
                div.style.color = "#9ca3af"
                div.style.fontSize = "11px"
                div.textContent = "[EMF/WMF]"
            }
        } else if (el.imageBase64 && el.imageBase64.startsWith("data:image/tiff")) {
            div.style.background = "#f3f4f6"
            div.style.display = "flex"
            div.style.alignItems = "center"
            div.style.justifyContent = "center"
            div.style.color = "#9ca3af"
            div.style.fontSize = "11px"
            div.textContent = "[TIFF]"
        } else {
            const img = document.createElement("img")
            img.src = el.imageBase64 || ""
            img.style.width = "100%"
            img.style.height = "100%"
            img.style.objectFit = "contain"
            img.style.objectPosition = "center"
            img.loading = "lazy"
            img.onerror = () => {
                img.style.display = "none"
                div.style.background = "#f3f4f6"
                div.style.display = "flex"
                div.style.alignItems = "center"
                div.style.justifyContent = "center"
                div.style.color = "#9ca3af"
                div.style.fontSize = "11px"
                div.textContent = "[Image unavailable]"
            }
            div.appendChild(img)
        }
        return div
    }

    // Shape element
    if (el.kind === "shape") {
        // Background
        if (el.style?.fillColor) {
            div.style.background = el.style.fillColor
        } else {
            // Use a light pattern for shapes without fill to indicate transparency
            div.style.background = "repeating-linear-gradient(45deg, #f3f4f6 0 8px, #fff 8px 16px)"
        }
        // Border
        if (el.style?.outlineColor) {
            const widthEmu = el.style.outlineWidth || 0
            // Convert EMU to pixels (1pt = 12700 EMU, 1pt ≈ 1.333px at 96dpi)
            const px = widthEmu ? (widthEmu / 12700) * 1.333 : 1
            div.style.border = `${px.toFixed(1)}px solid ${el.style.outlineColor}`
        } else {
            // No explicit outline: don't draw a border
            div.style.border = "none"
        }
        // Basic shape type adjustments
        if (el.shapeType) {
            const st = el.shapeType.toLowerCase()
            if (st === "ellipse" || st === "oval") div.style.borderRadius = "50%"
            else if (st === "roundrect") div.style.borderRadius = "12px"
        }

        // Shape with text content
        if (el.content && el.content.paragraphs && el.content.paragraphs.length) {
            // Generate default content for date and slide number placeholders if empty
            if ((el.placeholderType === "dt" || el.placeholderType === "sldNum") && 
                (!el.content.text || el.content.text.trim() === "")) {
                if (el.placeholderType === "dt") {
                    const now = new Date()
                    const defaultDate = now.toLocaleDateString("tr-TR", { 
                        year: "numeric", 
                        month: "long", 
                        day: "numeric" 
                    })
                    el.content.text = defaultDate
                    if (el.content.paragraphs[0]) {
                        el.content.paragraphs[0].text = defaultDate
                        if (el.content.paragraphs[0].runs.length === 0) {
                            el.content.paragraphs[0].runs = [{
                                text: defaultDate,
                                style: el.style || {}
                            }]
                        } else {
                            el.content.paragraphs[0].runs[0].text = defaultDate
                        }
                    }
                } else if (el.placeholderType === "sldNum") {
                    const slide = slides[activeSlideIndex]
                    const slideNum = slide?.slideNumber || activeSlideIndex + 1
                    const defaultSlideNum = String(slideNum)
                    el.content.text = defaultSlideNum
                    if (el.content.paragraphs[0]) {
                        el.content.paragraphs[0].text = defaultSlideNum
                        if (el.content.paragraphs[0].runs.length === 0) {
                            el.content.paragraphs[0].runs = [{
                                text: defaultSlideNum,
                                style: el.style || {}
                            }]
                        } else {
                            el.content.paragraphs[0].runs[0].text = defaultSlideNum
                        }
                    }
                }
            }
            // Apply overall text alignment - check shape style first, then paragraph
            let defaultAlign = el.style?.align || null
            if (!defaultAlign && el.content.paragraphs[0]?.align) {
                defaultAlign = el.content.paragraphs[0].align
            }
            // Also check if alignment is in the first paragraph's runs style (some PPTX files store it there)
            if (!defaultAlign && el.content.paragraphs[0]?.runs?.[0]?.style) {
                // Alignment is typically at paragraph level, not run level, but check anyway
            }
            if (defaultAlign) {
                if (defaultAlign === "ctr" || defaultAlign === "center" || defaultAlign === "centered") {
                    div.style.textAlign = "center"
                } else if (defaultAlign === "r" || defaultAlign === "right") {
                    div.style.textAlign = "right"
                } else if (defaultAlign === "l" || defaultAlign === "left") {
                    div.style.textAlign = "left"
                } else if (defaultAlign === "just" || defaultAlign === "justify") {
                    div.style.textAlign = "justify"
                }
            } else {
                // Default to center for title placeholders, left for others
                if (el.placeholderType === "title" || el.placeholderType === "ctrTitle") {
                    div.style.textAlign = "center"
                } else {
                    div.style.textAlign = "left"
                }
            }
            
            // Apply vertical alignment
            if (el.style?.verticalAlign) {
                const vAlign = el.style.verticalAlign
                if (vAlign === "top" || vAlign === "t") div.style.justifyContent = "flex-start"
                else if (vAlign === "middle" || vAlign === "mid" || vAlign === "ctr") div.style.justifyContent = "center"
                else if (vAlign === "bottom" || vAlign === "b") div.style.justifyContent = "flex-end"
            } else {
                // Default vertical alignment for title placeholders
                if (el.placeholderType === "title" || el.placeholderType === "ctrTitle") {
                    div.style.justifyContent = "center"
                }
            }
            
            // Apply text direction (vertical text)
            if (el.style?.textDirection) {
                const textDir = el.style.textDirection
                if (textDir === "vert" || textDir === "vertical") {
                    div.style.writingMode = "vertical-rl"
                    div.style.textOrientation = "upright"
                }
            }
            
            // Set display to flex for vertical alignment
            if (el.style?.verticalAlign || el.placeholderType === "title" || el.placeholderType === "ctrTitle") {
                div.style.display = "flex"
                div.style.flexDirection = "column"
            }
            
            // Apply text wrapping - but don't clip content unnecessarily
            if (el.style?.wrap === "none" || el.style?.wrap === "false") {
                div.style.whiteSpace = "nowrap"
                div.style.overflow = "visible"
                div.style.textOverflow = "ellipsis"
            } else {
                div.style.wordWrap = "break-word"
                div.style.overflowWrap = "break-word"
                div.style.overflow = "visible"
            }
            
            // Ensure text container doesn't clip content
            div.style.padding = "2px"
            
            // Maintain numbering counters per indentation level for numbered lists
            const numberCounters = {}
            el.content.paragraphs.forEach((p, pIndex) => {
                const pEl = document.createElement("div")
                // Slightly tighter line-height for more accurate rendering
                pEl.style.lineHeight = "1.2"
                pEl.style.margin = "0"
                pEl.style.padding = "0"
                
                // Apply paragraph-level alignment if different from default
                if (p.align && p.align !== defaultAlign) {
                    if (p.align === "ctr" || p.align === "center" || p.align === "centered") {
                        pEl.style.textAlign = "center"
                    } else if (p.align === "r" || p.align === "right") {
                        pEl.style.textAlign = "right"
                    } else if (p.align === "l" || p.align === "left") {
                        pEl.style.textAlign = "left"
                    } else if (p.align === "just" || p.align === "justify") {
                        pEl.style.textAlign = "justify"
                    }
                } else if (defaultAlign) {
                    // Apply default alignment if paragraph doesn't have its own
                    if (defaultAlign === "ctr" || defaultAlign === "center" || defaultAlign === "centered") {
                        pEl.style.textAlign = "center"
                    } else if (defaultAlign === "r" || defaultAlign === "right") {
                        pEl.style.textAlign = "right"
                    } else if (defaultAlign === "l" || defaultAlign === "left") {
                        pEl.style.textAlign = "left"
                    } else if (defaultAlign === "just" || defaultAlign === "justify") {
                        pEl.style.textAlign = "justify"
                    }
                }
                
                // Indent according to bullet level
                const level = (p.bullet && typeof p.bullet.level === "number") ? p.bullet.level : 0
                if (level > 0) {
                    // Use em units to preserve scaling relative to font size
                    pEl.style.marginLeft = `${level * 1.2}em`
                }
                // Bullet or number prefix
                if (p.bullet && p.bullet.type && p.bullet.type !== "none") {
                    const bulletSpan = document.createElement("span")
                    bulletSpan.style.display = "inline-block"
                    bulletSpan.style.width = "1.2em"
                    bulletSpan.style.marginRight = "4px"
                    if (p.bullet.type === "bullet") {
                        // Use provided character or fallback to default bullet
                        bulletSpan.textContent = p.bullet.char || "•"
                    } else if (p.bullet.type === "number") {
                        const lvl = level
                        // Initialize counter at startAt or 1
                        if (numberCounters[lvl] === undefined) {
                            const startAt = (typeof p.bullet.startAt === "number" && !isNaN(p.bullet.startAt)) ? p.bullet.startAt : 1
                            numberCounters[lvl] = startAt
                        }
                        const currentNum = numberCounters[lvl]
                        bulletSpan.textContent = `${currentNum}.`
                        numberCounters[lvl] = currentNum + 1
                    }
                    pEl.appendChild(bulletSpan)
                }
                // Runs (text segments)
                p.runs.forEach((run) => {
                    const span = document.createElement("span")
                    // Preserve whitespace and empty text
                    span.textContent = run.text || ""
                    // Font size: run-level overrides shape-level
                    if (run.style?.fontSize) span.style.fontSize = ((run.style.fontSize * 96) / 72) * scaleY + "px"
                    else if (el.style?.fontSize) span.style.fontSize = ((el.style.fontSize * 96) / 72) * scaleY + "px"
                    // Font family
                    if (run.style?.fontFamily) span.style.fontFamily = run.style.fontFamily
                    else if (el.style?.fontFamily) span.style.fontFamily = el.style.fontFamily
                    // Bold/italic/underline
                    const bold = run.style?.bold ?? el.style?.bold
                    if (bold) span.style.fontWeight = "bold"
                    const italic = run.style?.italic ?? el.style?.italic
                    if (italic) span.style.fontStyle = "italic"
                    const underline = run.style?.underline ?? el.style?.underline
                    if (underline) span.style.textDecoration = "underline"
                    // Text color
                    if (run.style?.color) span.style.color = run.style.color
                    else if (el.style?.color) span.style.color = el.style.color
                    // Text background color (highlight)
                    if (run.style?.backgroundColor) {
                        span.style.backgroundColor = run.style.backgroundColor
                        span.style.padding = "2px 4px"
                    }
                    pEl.appendChild(span)
                })
                div.appendChild(pEl)
            })
        }
        return div
    }

    // Text element
    if (el.content && el.content.text) {
        div.textContent = el.content.text
        if (el.style?.fontSize) {
            const fSizePx = (el.style.fontSize * 96) / 72
            div.style.fontSize = fSizePx * scaleY + "px"
        }
        if (el.style?.bold) div.style.fontWeight = "bold"
        if (el.style?.italic) div.style.fontStyle = "italic"
        if (el.style?.underline) div.style.textDecoration = "underline"
        if (el.style?.color) div.style.color = el.style.color
        if (el.style?.align) {
            if (el.style.align === "ctr" || el.style.align === "center" || el.style.align === "centered") {
                div.style.textAlign = "center"
            } else if (el.style.align === "r" || el.style.align === "right") {
                div.style.textAlign = "right"
            } else if (el.style.align === "l" || el.style.align === "left") {
                div.style.textAlign = "left"
            } else if (el.style.align === "just" || el.style.align === "justify") {
                div.style.textAlign = "justify"
            }
        }
        return div
    }

    return div
}

function renderSlides() {
    renderSlidesAsync()
}

function showSlide(index) {
    showSlideAsync(index)
}
document.getElementById("btnPrev").onclick = () => {
    if (activeSlideIndex > 0) {
        showSlideAsync(activeSlideIndex - 1)
    }
}

document.getElementById("btnNext").onclick = () => {
    if (activeSlideIndex < slides.length - 1) {
        showSlideAsync(activeSlideIndex + 1)
    }
}
document.addEventListener("keydown", (e) => {
    if (e.key === "ArrowLeft" && activeSlideIndex > 0) {
        showSlideAsync(activeSlideIndex - 1)
    }
    if (e.key === "ArrowRight" && activeSlideIndex < slides.length - 1) {
        showSlideAsync(activeSlideIndex + 1)
    }
})
