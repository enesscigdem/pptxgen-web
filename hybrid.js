// Hybrid Editor: PDF görsel + JSON metadata
// PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

let pdfDoc = null;
let currentSlideIndex = 0;
let slideMetadata = [];
let presentationJson = null;
const HYBRID_SERVER = 'http://localhost:3001';

// File input
document.getElementById('fileInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (file) {
        await loadPresentationHybrid(file);
    }
});

/**
 * PPTX dosyasını yükle: PDF + Metadata
 */
/**
 * Progress bar güncelle
 */
function updateProgress(percent, text) {
    const progressBar = document.getElementById('progressBar');
    const progressLabel = document.getElementById('progressLabel');
    const progressText = document.getElementById('progressText');
    
    if (progressBar) {
        progressBar.style.width = percent + '%';
    }
    if (progressLabel) {
        progressLabel.textContent = percent + '%';
    }
    if (progressText) {
        progressText.textContent = text || 'İşleniyor...';
    }
}

/**
 * Progress section'ı göster/gizle
 */
function showProgress(show) {
    const progressSection = document.getElementById('progressSection');
    if (progressSection) {
        if (show) {
            progressSection.classList.add('visible');
        } else {
            progressSection.classList.remove('visible');
        }
    }
}

async function loadPresentationHybrid(pptxFile) {
    const fileInfo = document.getElementById('fileInfo');
    const emptyState = document.getElementById('emptyState');
    const pdfCanvas = document.getElementById('pdfCanvas');
    
    fileInfo.textContent = `Yükleniyor: ${pptxFile.name}...`;
    emptyState.style.display = 'none';
    const pdfWrapper = document.getElementById('pdfWrapper');
    if (pdfWrapper) pdfWrapper.style.display = 'none';
    
    // Progress bar'ı göster
    showProgress(true);
    updateProgress(0, 'Dosya hazırlanıyor...');
    
    try {
        // 1. PPTX'i PDF'ye çevir (server-side)
        updateProgress(10, 'PDF\'ye çevriliyor...');
        const pdfBlob = await convertPptxToPdf(pptxFile);
        console.log('PDF converted, size:', pdfBlob.size);
        updateProgress(50, 'PDF oluşturuldu');
        
        // 2. PDF'yi yükle
        updateProgress(60, 'PDF yükleniyor...');
        const pdfUrl = URL.createObjectURL(pdfBlob);
        pdfDoc = await pdfjsLib.getDocument(pdfUrl).promise;
        console.log('PDF loaded, pages:', pdfDoc.numPages);
        updateProgress(70, `PDF yüklendi (${pdfDoc.numPages} slide)`);
        
        // 3. Metadata extract et (client-side, mevcut app.js logic'i)
        updateProgress(75, 'Metadata extract ediliyor...');
        await extractMetadataFromPptx(pptxFile);
        console.log('Metadata extracted, slides:', slideMetadata.length);
        updateProgress(85, `Metadata hazır (${slideMetadata.length} slide)`);
        
        // 4. Slide listesini oluştur
        createSlideList();
        updateProgress(90, 'Slide listesi oluşturuluyor...');
        
        // 5. İlk slide'ı render et (container'ın render olmasını bekle)
        updateProgress(95, 'İlk slide render ediliyor...');
        await new Promise(resolve => {
            if (document.getElementById('pdfContainer').clientWidth > 0) {
                resolve();
            } else {
                // Container henüz render olmamış, biraz bekle
                setTimeout(resolve, 200);
            }
        });
        await renderSlideHybrid(0);
        
        updateProgress(100, 'Tamamlandı!');
        
        // Progress bar'ı gizle
        setTimeout(() => {
            showProgress(false);
        }, 500);
        
        fileInfo.textContent = `${pptxFile.name} • ${pdfDoc.numPages} slide`;
        fileInfo.style.color = 'var(--color-success)';
        
    } catch (error) {
        console.error('Load error:', error);
        fileInfo.textContent = `Hata: ${error.message}`;
        fileInfo.style.color = 'var(--color-accent)';
        emptyState.style.display = 'block';
        const pdfWrapper = document.getElementById('pdfWrapper');
        if (pdfWrapper) pdfWrapper.style.display = 'none';
        
        // Progress bar'ı gizle
        showProgress(false);
    }
}

/**
 * PPTX'i PDF'ye çevir (server kullanarak)
 */
async function convertPptxToPdf(pptxFile) {
    // AbortController ile timeout ekle
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 600000); // 10 dakika timeout
    
    try {
        updateProgress(15, 'Dosya base64\'e çevriliyor...');
        const buffer = await pptxFile.arrayBuffer();
        const uint8 = new Uint8Array(buffer);
        let binary = "";
        // Chunk'lar halinde işle (büyük dosyalar için)
        const chunkSize = 8192;
        for (let i = 0; i < uint8.length; i += chunkSize) {
            const chunk = uint8.slice(i, i + chunkSize);
            binary += String.fromCharCode.apply(null, chunk);
            // Progress güncelle (base64 conversion)
            if (i % (chunkSize * 10) === 0) {
                const progress = 15 + Math.floor((i / uint8.length) * 5);
                updateProgress(progress, 'Dosya hazırlanıyor...');
            }
        }
        updateProgress(20, 'Base64 encoding tamamlandı');
        
        const base64 = btoa(binary);
        updateProgress(25, 'Server\'a gönderiliyor...');
        
        const response = await fetch(`${HYBRID_SERVER}/hybrid/convert-pdf`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                fileName: pptxFile.name,
                fileData: base64
            }),
            signal: controller.signal
        });
        
        clearTimeout(timeoutId);
        
        if (!response.ok) {
            let errorMsg = response.statusText;
            try {
                const errorData = await response.json();
                errorMsg = errorData.error || errorMsg;
            } catch (e) {
                // JSON parse failed, use statusText
            }
            throw new Error(`Server error: ${errorMsg}`);
        }
        
        updateProgress(30, 'PDF response alınıyor...');
        
        // Content-Length'i kontrol et
        const contentLength = response.headers.get('content-length');
        if (contentLength) {
            const totalSize = parseInt(contentLength, 10);
            console.log('Expected PDF size:', totalSize, 'bytes');
        }
        
        // Response'u blob olarak al - streaming ile (büyük dosyalar için)
        try {
            const reader = response.body.getReader();
            const chunks = [];
            let receivedLength = 0;
            const totalLength = contentLength ? parseInt(contentLength, 10) : null;
            
            console.log('Reading response stream, expected size:', totalLength || 'unknown');
            
            while (true) {
                const { done, value } = await reader.read();
                
                if (done) {
                    console.log('Stream reading completed, total received:', receivedLength);
                    break;
                }
                
                if (value) {
                    chunks.push(value);
                    receivedLength += value.length;
                    
                    // Progress güncelle
                    if (totalLength) {
                        const percent = Math.min(45, 30 + Math.floor((receivedLength / totalLength) * 15));
                        updateProgress(percent, `PDF alınıyor... (${(receivedLength / 1024 / 1024).toFixed(2)} MB)`);
                    } else {
                        updateProgress(35, `PDF alınıyor... (${(receivedLength / 1024 / 1024).toFixed(2)} MB)`);
                    }
                }
            }
            
            // Tüm chunk'ları birleştir
            console.log('Combining chunks, count:', chunks.length);
            const allChunks = new Uint8Array(receivedLength);
            let position = 0;
            for (const chunk of chunks) {
                allChunks.set(chunk, position);
                position += chunk.length;
            }
            
            const blob = new Blob([allChunks], { type: 'application/pdf' });
            
            if (!blob || blob.size === 0) {
                throw new Error('PDF response boş geldi');
            }
            
            console.log('PDF blob created, size:', blob.size, 'bytes');
            updateProgress(45, `PDF alındı (${(blob.size / 1024 / 1024).toFixed(2)} MB)`);
            
            return blob;
        } catch (streamError) {
            console.error('Stream reading error:', streamError);
            // Fallback: normal blob() kullan
            console.log('Falling back to response.blob()');
            const blob = await response.blob();
            if (!blob || blob.size === 0) {
                throw new Error('PDF response boş geldi');
            }
            updateProgress(45, `PDF alındı (${(blob.size / 1024 / 1024).toFixed(2)} MB)`);
            return blob;
        }
    } catch (error) {
        clearTimeout(timeoutId);
        if (error.name === 'AbortError') {
            throw new Error('İşlem zaman aşımına uğradı. Dosya çok büyük olabilir, lütfen daha küçük bir dosya deneyin.');
        }
        throw error;
    }
}

/**
 * Metadata extract et (mevcut app.js logic'ini kullan)
 */
async function extractMetadataFromPptx(pptxFile) {
    // Mevcut app.js'deki processFile logic'ini kullan
    // Ama sadece metadata için, görsel render için değil
    
    const arrayBuffer = await pptxFile.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    const parser = new DOMParser();
    
    // Theme colors
    const themeColorMap = {};
    try {
        const presRelsPath = "ppt/_rels/presentation.xml.rels";
        if (zip.file(presRelsPath)) {
            const presRelsXml = await zip.file(presRelsPath).async("text");
            const presRelsDoc = parser.parseFromString(presRelsXml, "application/xml");
            const rels = presRelsDoc.querySelectorAll("Relationship");
            let themeRel = null;
            rels.forEach((r) => {
                const type = r.getAttribute("Type") || "";
                if (type.includes("/theme")) themeRel = r;
            });
            if (themeRel) {
                const target = themeRel.getAttribute("Target");
                let themePath = target.replace("..", "ppt");
                if (!themePath.startsWith("ppt/")) themePath = "ppt/" + themePath.replace(/^\/+/, "");
                if (zip.file(themePath)) {
                    const themeXml = await zip.file(themePath).async("text");
                    const themeDoc = parser.parseFromString(themeXml, "application/xml");
                    const clrScheme = themeDoc.querySelector("a\\:clrScheme,clrScheme");
                    if (clrScheme) {
                        Array.from(clrScheme.children).forEach((node) => {
                            const name = node.localName || node.nodeName.split(":").pop();
                            const srgb = node.querySelector("a\\:srgbClr,srgbClr");
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
    
    // Slide size
    const presXml = await zip.file("ppt/presentation.xml").async("text");
    const presDoc = parser.parseFromString(presXml, "application/xml");
    const sldSz = presDoc.querySelector("p\\:sldSz,sldSz");
    const slideW = +sldSz.getAttribute("cx");
    const slideH = +sldSz.getAttribute("cy");
    
    function emuToCm(emu) {
        const v = Number.parseInt(emu, 10);
        if (isNaN(v)) return null;
        return +(v / 360000).toFixed(2);
    }
    
    // Slide paths
    const slidePaths = Object.keys(zip.files)
        .filter((p) => /ppt\/slides\/slide\d+\.xml$/.test(p))
        .sort((a, b) => Number.parseInt(a.match(/slide(\d+)/)[1]) - Number.parseInt(b.match(/slide(\d+)/)[1]));
    
    slideMetadata = [];
    
    for (let i = 0; i < slidePaths.length; i++) {
        const p = slidePaths[i];
        const xml = await zip.file(p).async("text");
        const doc = parser.parseFromString(xml, "application/xml");
        
        // Slide relationships
        const relsPath = p.replace("slides/", "slides/_rels/") + ".rels";
        const relsMap = {};
        if (zip.file(relsPath)) {
            const relsXml = await zip.file(relsPath).async("text");
            const relsDoc = parser.parseFromString(relsXml, "application/xml");
            const rels = relsDoc.querySelectorAll("Relationship");
            rels.forEach((r) => {
                const id = r.getAttribute("Id");
                const target = r.getAttribute("Target");
                relsMap[id] = target;
            });
        }
        
        const slideNumber = Number.parseInt(p.match(/slide(\d+)/)[1]);
        const slideJson = {
            slideNumber,
            size: {
                widthEmu: slideW,
                heightEmu: slideH,
                widthCm: emuToCm(slideW),
                heightCm: emuToCm(slideH),
            },
            elements: [],
        };
        
        // Extract elements (simplified - full extraction would use app.js logic)
        // For now, extract basic shape and text info
        const shapes = doc.querySelectorAll("p\\:sp,sp");
        shapes.forEach((sp, idx) => {
            const xfrm = sp.querySelector("a\\:xfrm,xfrm");
            if (!xfrm) return;
            
            const off = xfrm.querySelector("a\\:off,off");
            const ext = xfrm.querySelector("a\\:ext,ext");
            if (!off || !ext) return;
            
            const x = +off.getAttribute("x");
            const y = +off.getAttribute("y");
            const w = +ext.getAttribute("cx");
            const h = +ext.getAttribute("cy");
            
            const txBody = sp.querySelector("a\\:txBody,txBody");
            const ph = sp.querySelector("p\\:ph,ph");
            const placeholderType = ph ? ph.getAttribute("type") || null : null;
            
            let text = "";
            if (txBody) {
                const paragraphs = txBody.querySelectorAll("a\\:p,p");
                paragraphs.forEach((pNode) => {
                    const rNodes = pNode.querySelectorAll("a\\:r,r");
                    rNodes.forEach((r) => {
                        const tNode = r.querySelector("a\\:t,t");
                        if (tNode) text += tNode.textContent || "";
                    });
                });
            }
            
            if (text || placeholderType) {
                slideJson.elements.push({
                    id: `s${slideNumber}-el${idx + 1}`,
                    kind: "shape",
                    type: placeholderType === "title" ? "title" : "textbox",
                    placeholderType,
                    geometry: {
                        xEmu: x,
                        yEmu: y,
                        wEmu: w,
                        hEmu: h,
                        xCm: emuToCm(x),
                        yCm: emuToCm(y),
                        wCm: emuToCm(w),
                        hCm: emuToCm(h),
                    },
                    content: {
                        text: text
                    }
                });
            }
        });
        
        slideMetadata.push(slideJson);
    }
    
    presentationJson = {
        fileName: pptxFile.name,
        slideCount: slideMetadata.length,
        slides: slideMetadata
    };
}

/**
 * Slide listesini oluştur
 */
function createSlideList() {
    const slideList = document.getElementById('slideList');
    slideList.innerHTML = '';
    
    for (let i = 0; i < slideMetadata.length; i++) {
        const btn = document.createElement('div');
        btn.className = 'slide-thumb';
        btn.innerHTML = `
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6">
                <rect x="2.5" y="4" width="19" height="13" rx="2" ry="2" fill="var(--thumb-bg,#ffffff)" stroke="var(--thumb-border,#d0d7e1)"/>
                <rect x="5" y="6" width="9" height="3" rx="0.5" fill="var(--primary,#3b82f6)" opacity="0.22"/>
                <line x1="5" y1="11" x2="17" y2="11" stroke="#94a3b8" stroke-width="1.2"/>
                <line x1="5" y1="13.5" x2="14" y2="13.5" stroke="#94a3b8" stroke-width="1.2"/>
            </svg>
            Slide ${i + 1}
        `;
        btn.onclick = () => renderSlideHybrid(i);
        if (i === 0) btn.classList.add('active');
        slideList.appendChild(btn);
    }
}

/**
 * Slide'ı render et: PDF görsel + overlay
 */
async function renderSlideHybrid(slideIndex) {
    if (!pdfDoc || slideIndex < 0 || slideIndex >= pdfDoc.numPages) {
        console.warn('Invalid slide index:', slideIndex);
        return;
    }
    
    try {
        currentSlideIndex = slideIndex;
        const canvas = document.getElementById('pdfCanvas');
        const ctx = canvas.getContext('2d');
        const emptyState = document.getElementById('emptyState');
        const pdfWrapper = document.getElementById('pdfWrapper');
        const pdfContainer = document.getElementById('pdfContainer');
        
        if (!canvas || !pdfWrapper || !pdfContainer) {
            console.error('Required elements not found');
            return;
        }
        
        emptyState.style.display = 'none';
        pdfWrapper.style.display = 'block';
        canvas.style.display = 'block'; // Canvas'ı görünür yap
        
        // PDF page render
        const page = await pdfDoc.getPage(slideIndex + 1);
        console.log('Rendering page:', slideIndex + 1);
        
        // Scale'i container'a göre ayarla - daha büyük görünsün
        let containerWidth = pdfContainer.clientWidth - 40; // padding için
        let containerHeight = pdfContainer.clientHeight - 40;
        
        // Eğer container henüz render olmamışsa, default değerler kullan
        if (containerWidth <= 0) containerWidth = 1200;
        if (containerHeight <= 0) containerHeight = 800;
        
        const viewport = page.getViewport({ scale: 1 });
        // Scale'i artır - daha net görünsün
        const scale = Math.min(
            containerWidth / viewport.width,
            containerHeight / viewport.height,
            3 // Max scale artırıldı (2'den 3'e)
        );
        const scaledViewport = page.getViewport({ scale });
        
        console.log('Viewport:', { width: scaledViewport.width, height: scaledViewport.height, scale });
        
        canvas.width = scaledViewport.width;
        canvas.height = scaledViewport.height;
        
        // Canvas'ı temizle
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        
        await page.render({
            canvasContext: ctx,
            viewport: scaledViewport
        }).promise;
        
        console.log('Page rendered successfully');
        console.log('Canvas dimensions:', canvas.width, 'x', canvas.height);
        console.log('Canvas display:', window.getComputedStyle(canvas).display);
        
        // Canvas'ı tekrar görünür yap (eğer gizlenmişse)
        canvas.style.display = 'block';
        canvas.style.visibility = 'visible';
        
        // Overlay oluştur
        createInteractiveOverlay(slideIndex, scaledViewport);
        
        // Metadata göster
        updateMetadataDisplay(slideIndex);
        
        // Slide list active state
        document.querySelectorAll('.slide-thumb').forEach((btn, i) => {
            btn.classList.toggle('active', i === slideIndex);
        });
    } catch (error) {
        console.error('Render error:', error);
        const emptyState = document.getElementById('emptyState');
        if (emptyState) {
            emptyState.style.display = 'block';
            emptyState.innerHTML = `
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="12" y1="8" x2="12" y2="12"/>
                    <line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                <p>Render hatası: ${error.message}</p>
            `;
        }
    }
}

/**
 * Interactive overlay oluştur
 */
function createInteractiveOverlay(slideIndex, viewport) {
    const overlay = document.getElementById('elementOverlay');
    const canvas = document.getElementById('pdfCanvas');
    overlay.innerHTML = '';
    
    // Overlay'i canvas boyutuna göre ayarla
    overlay.style.width = canvas.width + 'px';
    overlay.style.height = canvas.height + 'px';
    
    const slide = slideMetadata[slideIndex];
    if (!slide || !slide.elements) return;
    
    // PDF viewport scale - canvas boyutuna göre
    const pdfWidthCm = slide.size.widthCm;
    const pdfHeightCm = slide.size.heightCm;
    const scaleX = viewport.width / (pdfWidthCm * 37.8);
    const scaleY = viewport.height / (pdfHeightCm * 37.8);
    
    slide.elements.forEach(el => {
        if (!el.geometry) return;
        
        const div = document.createElement('div');
        div.className = 'element-overlay';
        div.dataset.elementId = el.id;
        
        const x = el.geometry.xCm * 37.8 * scaleX;
        const y = el.geometry.yCm * 37.8 * scaleY;
        const w = el.geometry.wCm * 37.8 * scaleX;
        const h = el.geometry.hCm * 37.8 * scaleY;
        
        div.style.left = x + 'px';
        div.style.top = y + 'px';
        div.style.width = Math.max(w, 5) + 'px'; // Min width
        div.style.height = Math.max(h, 5) + 'px'; // Min height
        
        const text = el.content?.text || el.type || 'Element';
        div.title = text.substring(0, 50) + (text.length > 50 ? '...' : '');
        div.onclick = (e) => {
            e.stopPropagation();
            e.preventDefault();
            selectElement(el);
        };
        
        overlay.appendChild(div);
    });
}

/**
 * Element seç
 */
function selectElement(element) {
    document.querySelectorAll('.element-overlay').forEach(el => {
        el.classList.remove('selected');
    });
    
    const overlay = document.querySelector(`[data-element-id="${element.id}"]`);
    if (overlay) overlay.classList.add('selected');
    
    updateMetadataDisplay(currentSlideIndex, element);
}

/**
 * Metadata göster
 */
function updateMetadataDisplay(slideIndex, selectedElement = null) {
    const content = document.getElementById('metadataContent');
    const slide = slideMetadata[slideIndex];
    
    if (!slide) {
        content.innerHTML = '<p style="color: var(--color-text-secondary);">Metadata yok</p>';
        return;
    }
    
    if (selectedElement) {
        content.innerHTML = `
            <h4 style="margin-bottom: 8px; font-size: 14px;">Seçili Element</h4>
            <pre style="font-size: 11px; line-height: 1.4;">${JSON.stringify(selectedElement, null, 2)}</pre>
        `;
    } else {
        content.innerHTML = `
            <h4 style="margin-bottom: 8px; font-size: 14px;">Slide ${slideIndex + 1}</h4>
            <p style="font-size: 12px; color: var(--color-text-secondary); margin-bottom: 8px;">
                ${slide.elements?.length || 0} element
            </p>
            <details style="margin-top: 8px;">
                <summary style="cursor: pointer; font-size: 12px; color: var(--color-accent);">Tüm metadata'yı göster</summary>
                <pre style="font-size: 10px; line-height: 1.4; margin-top: 8px;">${JSON.stringify(slide, null, 2)}</pre>
            </details>
        `;
    }
}

/**
 * AI düzenleme
 */
document.getElementById('btnAIEdit').addEventListener('click', async () => {
    const instruction = document.getElementById('aiInstruction').value.trim();
    if (!instruction) {
        alert('Lütfen bir talimat girin');
        return;
    }
    
    await applyAIEdit(currentSlideIndex, instruction);
});

async function applyAIEdit(slideIndex, instruction) {
    const btn = document.getElementById('btnAIEdit');
    const originalText = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = `
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin">
            <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
        </svg>
        İşleniyor...
    `;
    
    try {
        // TODO: AI agent API'sine bağlan
        // Şimdilik placeholder
        console.log('AI Edit:', {
            slideIndex,
            instruction,
            metadata: slideMetadata[slideIndex]
        });
        
        // Simüle edilmiş değişiklik
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        alert('AI düzenleme özelliği yakında eklenecek!\n\nŞu an için metadata JSON üzerinden manuel düzenleme yapabilirsiniz.');
        
    } catch (error) {
        console.error('AI edit error:', error);
        alert('Hata: ' + error.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalText;
    }
}

// Navigation
document.getElementById('btnPrev').addEventListener('click', () => {
    if (currentSlideIndex > 0) {
        renderSlideHybrid(currentSlideIndex - 1);
    }
});

document.getElementById('btnNext').addEventListener('click', () => {
    if (pdfDoc && currentSlideIndex < pdfDoc.numPages - 1) {
        renderSlideHybrid(currentSlideIndex + 1);
    }
});

// Keyboard navigation
document.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowLeft' && currentSlideIndex > 0) {
        renderSlideHybrid(currentSlideIndex - 1);
    }
    if (e.key === 'ArrowRight' && pdfDoc && currentSlideIndex < pdfDoc.numPages - 1) {
        renderSlideHybrid(currentSlideIndex + 1);
    }
});

