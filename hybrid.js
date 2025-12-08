// Hybrid Editor: PDF görsel + JSON metadata
// PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

let pdfDoc = null;
let pdfLoadingTask = null; // PDF loading task (cleanup için)
let currentSlideIndex = 0;
let slideMetadata = [];
let presentationJson = null;
let pdfViewport = null; // PDF viewport (overlay güncellemeleri için)
let originalPptxFile = null; // Orijinal PPTX dosyası (yeniden PDF oluşturmak için)
let pptxZip = null; // JSZip instance (XML düzenlemek için)
let currentPdfBlob = null; // Güncellenmiş PDF blob (download için)
let currentPptxBlob = null; // Güncellenmiş PPTX blob (download için)
// Server URL - Otomatik olarak localhost veya production URL'sini seçer
const HYBRID_SERVER = (() => {
    // Eğer window.HYBRID_SERVER_URL manuel olarak ayarlanmışsa onu kullan
    if (window.HYBRID_SERVER_URL) {
        return window.HYBRID_SERVER_URL;
    }
    // Localhost'ta çalışıyorsa localhost kullan
    if (window.location.hostname === 'localhost' || 
        window.location.hostname === '127.0.0.1' ||
        window.location.hostname === '') {
        return 'http://localhost:3001';
    }
    // Production'da Render.com URL'sini kullan
    return 'https://pptx-hybrid-server.onrender.com';
})();

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
        
        // Eski PDF'i temizle
        if (pdfLoadingTask) {
            pdfLoadingTask.destroy();
        }
        
        pdfLoadingTask = pdfjsLib.getDocument(pdfUrl);
        pdfDoc = await pdfLoadingTask.promise;
        console.log('PDF loaded, pages:', pdfDoc.numPages);
        updateProgress(70, `PDF yüklendi (${pdfDoc.numPages} slide)`);
        
        // İlk yüklemede blob'ları sakla (download için)
        currentPdfBlob = pdfBlob;
        // Orijinal PPTX'i blob olarak sakla (download için)
        if (pptxFile instanceof File) {
            const originalBuffer = await pptxFile.arrayBuffer();
            currentPptxBlob = new Blob([originalBuffer], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
        }
        updateDownloadButtons();
        
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
/**
 * OPTİMİZE: Base64 encoding'i daha hızlı yap
 */
function arrayBufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    const chunkSize = 8192;
    for (let i = 0; i < bytes.length; i += chunkSize) {
        const chunk = bytes.subarray(i, i + chunkSize);
        binary += String.fromCharCode.apply(null, chunk);
    }
    return btoa(binary);
}

/**
 * OPTİMİZE: PPTX'i PDF'e çevir (hızlı versiyon, progress göstermez)
 */
async function convertPptxToPdfOptimized(pptxFile) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 600000);
    
    try {
        // Base64 encoding (optimize)
        const buffer = await pptxFile.arrayBuffer();
        const base64 = arrayBufferToBase64(buffer);
        
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
            const errorData = await response.json().catch(() => ({ error: response.statusText }));
            throw new Error(`Server error: ${errorData.error || response.statusText}`);
        }
        
        // Response'u blob olarak al
        const blob = await response.blob();
        if (!blob || blob.size === 0) {
            throw new Error('PDF response boş geldi');
        }
        
        return blob;
    } catch (error) {
        clearTimeout(timeoutId);
        if (error.name === 'AbortError') {
            throw new Error('İşlem zaman aşımına uğradı');
        }
        throw error;
    }
}

async function convertPptxToPdf(pptxFile) {
    // AbortController ile timeout ekle
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 600000); // 10 dakika timeout
    
    try {
        updateProgress(15, 'Dosya base64\'e çevriliyor...');
        const buffer = await pptxFile.arrayBuffer();
        const base64 = arrayBufferToBase64(buffer);
        updateProgress(20, 'Base64 encoding tamamlandı');
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
    
    // Orijinal dosyayı sakla
    originalPptxFile = pptxFile;
    
    const arrayBuffer = await pptxFile.arrayBuffer();
    pptxZip = await JSZip.loadAsync(arrayBuffer);
    const zip = pptxZip;
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
        
        // Viewport'u global olarak sakla (overlay güncellemeleri için)
        pdfViewport = scaledViewport;
        
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
 * Metadata göster ve düzenlenebilir yap
 */
function updateMetadataDisplay(slideIndex, selectedElement = null) {
    const content = document.getElementById('metadataContent');
    const slide = slideMetadata[slideIndex];
    
    if (!slide) {
        content.innerHTML = '<p style="color: var(--color-text-secondary);">Metadata yok</p>';
        return;
    }
    
    if (selectedElement) {
        const elementJson = JSON.stringify(selectedElement, null, 2);
        content.innerHTML = `
            <h4 style="margin-bottom: 8px; font-size: 14px;">Seçili Element</h4>
            <textarea 
                id="metadataEditor" 
                style="width: 100%; min-height: 200px; font-family: var(--font-mono); font-size: 11px; padding: 8px; border: 2px solid var(--color-border); border-radius: 4px; background: rgba(0, 0, 0, 0.02); resize: vertical; box-sizing: border-box;"
            >${elementJson}</textarea>
            <button 
                id="btnSaveMetadata" 
                class="btn btn-primary" 
                style="width: 100%; margin-top: 8px;"
            >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
                    <polyline points="17 21 17 13 7 13 7 21"/>
                    <polyline points="7 3 7 8 15 8"/>
                </svg>
                Değişiklikleri Kaydet
            </button>
        `;
        
        // Kaydet butonuna event listener ekle
        document.getElementById('btnSaveMetadata').addEventListener('click', () => {
            saveMetadataChanges(slideIndex, selectedElement.id);
        });
    } else {
        const slideJson = JSON.stringify(slide, null, 2);
        content.innerHTML = `
            <h4 style="margin-bottom: 8px; font-size: 14px;">Slide ${slideIndex + 1}</h4>
            <p style="font-size: 12px; color: var(--color-text-secondary); margin-bottom: 8px;">
                ${slide.elements?.length || 0} element
            </p>
            <details style="margin-top: 8px;">
                <summary style="cursor: pointer; font-size: 12px; color: var(--color-accent);">Tüm metadata'yı düzenle</summary>
                <textarea 
                    id="metadataEditor" 
                    style="width: 100%; min-height: 300px; font-family: var(--font-mono); font-size: 10px; padding: 8px; border: 2px solid var(--color-border); border-radius: 4px; background: rgba(0, 0, 0, 0.02); resize: vertical; box-sizing: border-box; margin-top: 8px;"
                >${slideJson}</textarea>
                <button 
                    id="btnSaveMetadata" 
                    class="btn btn-primary" 
                    style="width: 100%; margin-top: 8px;"
                >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
                        <polyline points="17 21 17 13 7 13 7 21"/>
                        <polyline points="7 3 7 8 15 8"/>
                    </svg>
                    Slide Metadata'yı Kaydet
                </button>
            </details>
        `;
        
        // Kaydet butonuna event listener ekle
        const saveBtn = document.getElementById('btnSaveMetadata');
        if (saveBtn) {
            saveBtn.addEventListener('click', () => {
                saveSlideMetadata(slideIndex);
            });
        }
    }
}

/**
 * Element metadata değişikliklerini kaydet
 * PPTX XML'ini güncelle, sadece değişen slide için PDF oluştur, canvas'ı güncelle
 */
async function saveMetadataChanges(slideIndex, elementId) {
    try {
        const editor = document.getElementById('metadataEditor');
        const newMetadata = JSON.parse(editor.value);
        
        const slide = slideMetadata[slideIndex];
        if (!slide || !slide.elements) {
            alert('Slide bulunamadı!');
            return;
        }
        
        const elementIndex = slide.elements.findIndex(el => el.id === elementId);
        if (elementIndex === -1) {
            alert('Element bulunamadı!');
            return;
        }
        
        // Metadata'yı güncelle
        slide.elements[elementIndex] = newMetadata;
        
        // Loading overlay göster ve tüm etkileşimleri engelle
        showLoadingOverlay('PPTX güncelleniyor ve PDF yeniden oluşturuluyor...');
        
        try {
            // PPTX XML'ini güncelle ve PDF'i yeniden oluştur
            await updatePptxAndRegeneratePdfSlide(slideIndex, elementId, newMetadata);
            
            // Overlay'i güncelle
            updateOverlayForElement(slideIndex, newMetadata);
            
            // Seçili elementi güncelle
            selectElement(newMetadata);
            
            // Başarı mesajı
            showNotification('Değişiklikler uygulandı ve canvas güncellendi!', 'success');
        } finally {
            // Loading overlay'i kaldır
            hideLoadingOverlay();
        }
        
    } catch (e) {
        alert('Hata: ' + e.message);
        console.error(e);
        showNotification('Hata: ' + e.message, 'error');
        hideLoadingOverlay();
    }
}

// PPTX Blob cache (performans için)
let pptxBlobCache = null;
let pptxBlobCacheTimestamp = 0;

/**
 * PPTX XML'ini güncelle ve sadece değişen slide için PDF oluştur, canvas'ı güncelle
 * OPTİMİZE: Sadece değişen XML'i güncelle, PPTX'i optimize şekilde oluştur
 */
async function updatePptxAndRegeneratePdfSlide(slideIndex, elementId, newMetadata) {
    if (!pptxZip || !originalPptxFile) {
        throw new Error('PPTX dosyası yüklenmemiş!');
    }
    
    try {
        // Slide XML dosyasını bul
        const slidePath = `ppt/slides/slide${slideIndex + 1}.xml`;
        if (!pptxZip.file(slidePath)) {
            throw new Error(`Slide ${slideIndex + 1} XML dosyası bulunamadı!`);
        }
        
        // XML'i parse et
        const slideXml = await pptxZip.file(slidePath).async("text");
        const parser = new DOMParser();
        const slideDoc = parser.parseFromString(slideXml, "application/xml");
        
        // Element'i XML'de bul ve güncelle
        updateElementInXml(slideDoc, newMetadata, elementId);
        
        // XML'i string'e çevir
        const serializer = new XMLSerializer();
        const updatedXml = serializer.serializeToString(slideDoc);
        
        // PPTX'e yaz (sadece bu dosyayı güncelle)
        pptxZip.file(slidePath, updatedXml);
        
        // OPTİMİZE: JSZip'i optimize modda oluştur (sadece değişen dosyaları işle)
        // Stream kullanarak daha hızlı blob oluştur
        const updatedPptxBlob = await pptxZip.generateAsync({ 
            type: "blob",
            compression: "DEFLATE",
            compressionOptions: { level: 1 } // Hızlı compression (level 1-9, 1 en hızlı)
        });
        
        const updatedPptxFile = new File([updatedPptxBlob], originalPptxFile.name, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
        
        // OPTİMİZE: Hızlı PDF conversion (progress göstermez)
        const newPdfBlob = await convertPptxToPdfOptimized(updatedPptxFile);
        
        // Blob'ları sakla (download için)
        currentPptxBlob = updatedPptxBlob;
        currentPdfBlob = newPdfBlob;
        
        // Download butonlarını etkinleştir
        updateDownloadButtons();
        
        // Eski PDF URL'ini temizle (memory leak önleme)
        if (pdfLoadingTask) {
            const oldUrl = pdfLoadingTask._transport?.url;
            if (oldUrl && oldUrl.startsWith('blob:')) {
                URL.revokeObjectURL(oldUrl);
            }
            pdfLoadingTask.destroy();
        }
        
        // Yeni PDF'i yükle (optimize ayarlarla)
        const pdfUrl = URL.createObjectURL(newPdfBlob);
        pdfLoadingTask = pdfjsLib.getDocument({ 
            url: pdfUrl,
            disableAutoFetch: false,
            disableStream: false,
            verbosity: 0 // Log'ları kapat (performans)
        });
        pdfDoc = await pdfLoadingTask.promise;
        
        // Sadece o slide'ı yeniden render et
        await renderSlideHybrid(slideIndex);
        
    } catch (error) {
        console.error('PPTX güncelleme hatası:', error);
        throw error;
    }
}

/**
 * XML'de elementi güncelle
 * Element ID'sine göre bulur: s1-el1 -> slide 1, element index 0
 */
function updateElementInXml(slideDoc, element, elementId) {
    // spTree içindeki tüm shape'leri bul
    const spTree = slideDoc.querySelector("p\\:spTree,spTree");
    if (!spTree) return;
    
    const shapes = spTree.querySelectorAll("p\\:sp,sp");
    
    // Element ID'den index'i çıkar: s1-el1 -> index 0
    let targetIndex = -1;
    if (elementId) {
        const match = elementId.match(/s\d+-el(\d+)/);
        if (match) {
            targetIndex = parseInt(match[1]) - 1; // el1 -> index 0
        }
    }
    
    // Index'e göre bul (daha güvenilir)
    if (targetIndex >= 0 && targetIndex < shapes.length) {
        const shape = shapes[targetIndex];
        updateShapeInXml(shape, element, slideDoc);
        return;
    }
    
    // Fallback: Pozisyon + text kombinasyonu ile eşleştir
    for (let i = 0; i < shapes.length; i++) {
        const shape = shapes[i];
        const xfrm = shape.querySelector("a\\:xfrm,xfrm");
        if (!xfrm) continue;
        
        const off = xfrm.querySelector("a\\:off,off");
        const ext = xfrm.querySelector("a\\:ext,ext");
        if (!off || !ext) continue;
        
        const xEmu = +off.getAttribute("x");
        const yEmu = +off.getAttribute("y");
        
        // Pozisyon toleransı: 1000 EMU (~0.03 cm)
        const tolerance = 1000;
        const posMatch = Math.abs(xEmu - element.geometry.xEmu) < tolerance &&
                        Math.abs(yEmu - element.geometry.yEmu) < tolerance;
        
        if (posMatch) {
            updateShapeInXml(shape, element, slideDoc);
            return;
        }
    }
}

/**
 * Shape'i XML'de güncelle
 */
function updateShapeInXml(shape, element, doc) {
    // Text içeriğini güncelle
    const txBody = shape.querySelector("p\\:txBody,txBody");
    if (txBody && element.content?.text !== undefined) {
        const paragraphs = txBody.querySelectorAll("a\\:p,p");
        if (paragraphs.length > 0) {
            const firstP = paragraphs[0];
            const runs = firstP.querySelectorAll("a\\:r,r");
            
            if (runs.length > 0) {
                // İlk run'un text'ini güncelle
                const textNode = runs[0].querySelector("a\\:t,t");
                if (textNode) {
                    textNode.textContent = element.content.text;
                }
            } else {
                // Run yoksa oluştur
                const run = doc.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:r");
                const text = doc.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:t");
                text.textContent = element.content.text || "";
                run.appendChild(text);
                firstP.appendChild(run);
            }
        }
    }
    
    // Geometry güncelle (pozisyon ve boyut)
    if (element.geometry) {
        const xfrm = shape.querySelector("a\\:xfrm,xfrm");
        if (xfrm) {
            const off = xfrm.querySelector("a\\:off,off");
            const ext = xfrm.querySelector("a\\:ext,ext");
            
            if (off && element.geometry.xEmu !== undefined) {
                off.setAttribute("x", Math.round(element.geometry.xEmu));
            }
            if (off && element.geometry.yEmu !== undefined) {
                off.setAttribute("y", Math.round(element.geometry.yEmu));
            }
            if (ext && element.geometry.wEmu !== undefined) {
                ext.setAttribute("cx", Math.round(element.geometry.wEmu));
            }
            if (ext && element.geometry.hEmu !== undefined) {
                ext.setAttribute("cy", Math.round(element.geometry.hEmu));
            }
        }
    }
    
    // Stil bilgilerini güncelle
    if (element.style) {
        updateElementStyleInXml(shape, element.style, doc);
    }
}

/**
 * XML'de element stilini güncelle
 */
function updateElementStyleInXml(shape, style, doc) {
    const txBody = shape.querySelector("p\\:txBody,txBody");
    if (!txBody) return;
    
    const paragraphs = txBody.querySelectorAll("a\\:p,p");
    paragraphs.forEach(p => {
        const runs = p.querySelectorAll("a\\:r,r");
        runs.forEach(r => {
            let rPr = r.querySelector("a\\:rPr,rPr");
            if (!rPr) {
                rPr = doc.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:rPr");
                r.insertBefore(rPr, r.firstChild);
            }
            
            // Font size
            if (style.fontSize) {
                let sz = rPr.querySelector("a\\:sz,sz");
                if (!sz) {
                    sz = doc.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:sz");
                    rPr.appendChild(sz);
                }
                sz.setAttribute("val", Math.round(style.fontSize * 100)); // PowerPoint uses 100ths of a point
            }
            
            // Font color
            if (style.color) {
                let solidFill = rPr.querySelector("a\\:solidFill,solidFill");
                if (!solidFill) {
                    solidFill = doc.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:solidFill");
                    rPr.appendChild(solidFill);
                }
                
                let srgbClr = solidFill.querySelector("a\\:srgbClr,srgbClr");
                if (!srgbClr) {
                    srgbClr = doc.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:srgbClr");
                    solidFill.appendChild(srgbClr);
                }
                
                // Hex color'ı RGB'ye çevir
                const hex = style.color.replace('#', '');
                srgbClr.setAttribute("val", hex);
            }
            
            // Font weight (bold)
            if (style.fontWeight === 'bold' || style.fontWeight >= 700) {
                let b = rPr.querySelector("a\\:b,b");
                if (!b) {
                    b = doc.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:b");
                    rPr.appendChild(b);
                }
            }
        });
    });
}

/**
 * Belirli bir element için overlay'i güncelle (sadece pozisyon)
 */
function updateOverlayForElement(slideIndex, element) {
    const overlay = document.querySelector(`[data-element-id="${element.id}"]`);
    if (!overlay) return;
    
    const slide = slideMetadata[slideIndex];
    if (!slide || !element.geometry || !pdfViewport) return;
    
    // PDF viewport scale
    const pdfWidthCm = slide.size.widthCm;
    const pdfHeightCm = slide.size.heightCm;
    const scaleX = pdfViewport.width / (pdfWidthCm * 37.8);
    const scaleY = pdfViewport.height / (pdfHeightCm * 37.8);
    
    // Pozisyon ve boyutu hesapla
    const x = element.geometry.xCm * 37.8 * scaleX;
    const y = element.geometry.yCm * 37.8 * scaleY;
    const w = element.geometry.wCm * 37.8 * scaleX;
    const h = element.geometry.hCm * 37.8 * scaleY;
    
    // Overlay pozisyonunu güncelle
    overlay.style.left = x + 'px';
    overlay.style.top = y + 'px';
    overlay.style.width = Math.max(w, 5) + 'px';
    overlay.style.height = Math.max(h, 5) + 'px';
    
    const text = element.content?.text || element.type || 'Element';
    overlay.title = text.substring(0, 50) + (text.length > 50 ? '...' : '');
}

/**
 * Slide metadata değişikliklerini kaydet
 */
function saveSlideMetadata(slideIndex) {
    try {
        const editor = document.getElementById('metadataEditor');
        const newMetadata = JSON.parse(editor.value);
        
        // Slide metadata'yı güncelle
        slideMetadata[slideIndex] = newMetadata;
        
        // Overlay'leri yeniden render et
        renderSlideHybrid(slideIndex);
        
        // Başarı mesajı
        showNotification('Slide metadata kaydedildi ve preview güncellendi!', 'success');
        
    } catch (e) {
        alert('Geçersiz JSON formatı! Lütfen düzeltin.\n\nHata: ' + e.message);
    }
}

/**
 * Bildirim göster
 */
/**
 * Loading overlay göster (ekranın ortasında, tüm etkileşimleri engelle)
 */
function showLoadingOverlay(message = 'İşleniyor...') {
    // Eğer zaten varsa kaldır
    hideLoadingOverlay();
    
    const overlay = document.createElement('div');
    overlay.id = 'loadingOverlay';
    overlay.innerHTML = `
        <div class="loading-content">
            <div class="loading-spinner"></div>
            <p class="loading-message">${message}</p>
        </div>
    `;
    document.body.appendChild(overlay);
    
    // Tüm etkileşimleri engelle
    document.body.style.pointerEvents = 'none';
    overlay.style.pointerEvents = 'all';
}

/**
 * Loading overlay'i kaldır
 */
function hideLoadingOverlay() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.remove();
    }
    // Etkileşimleri tekrar etkinleştir
    document.body.style.pointerEvents = '';
}

function showNotification(message, type = 'info') {
    // Basit bir bildirim sistemi
    const notification = document.createElement('div');
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 12px 20px;
        background: ${type === 'success' ? '#10b981' : '#3b82f6'};
        color: white;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 10000;
        font-size: 14px;
        font-weight: 500;
        animation: slideIn 0.3s ease;
    `;
    notification.textContent = message;
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.style.animation = 'slideOut 0.3s ease';
        setTimeout(() => notification.remove(), 300);
    }, 3000);
}

// CSS animasyonları ve loading overlay stilleri ekle
if (!document.getElementById('notification-styles')) {
    const style = document.createElement('style');
    style.id = 'notification-styles';
    style.textContent = `
        @keyframes slideIn {
            from {
                transform: translateX(100%);
                opacity: 0;
            }
            to {
                transform: translateX(0);
                opacity: 1;
            }
        }
        @keyframes slideOut {
            from {
                transform: translateX(0);
                opacity: 1;
            }
            to {
                transform: translateX(100%);
                opacity: 0;
            }
        }
        
        /* Loading Overlay */
        #loadingOverlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(250, 249, 247, 0.95);
            backdrop-filter: blur(4px);
            z-index: 99999;
            display: flex;
            align-items: center;
            justify-content: center;
            animation: fadeIn 0.2s ease;
        }
        
        .loading-content {
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 20px;
            padding: 40px;
            background: #ffffff;
            border: 2px solid #1e293b;
            border-radius: 12px;
            box-shadow: 6px 6px 0 #1e293b;
            max-width: 400px;
            text-align: center;
        }
        
        .loading-spinner {
            width: 48px;
            height: 48px;
            border: 4px solid #e2e8f0;
            border-top-color: #f97316;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
        }
        
        .loading-message {
            font-family: "Inter", -apple-system, BlinkMacSystemFont, sans-serif;
            font-size: 16px;
            font-weight: 500;
            color: #1e293b;
            margin: 0;
        }
        
        @keyframes fadeIn {
            from {
                opacity: 0;
            }
            to {
                opacity: 1;
            }
        }
        
        @keyframes spin {
            to {
                transform: rotate(360deg);
            }
        }
    `;
    document.head.appendChild(style);
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

/**
 * Download fonksiyonları
 */
function updateDownloadButtons() {
    const btnPptx = document.getElementById('btnDownloadPptx');
    const btnPdf = document.getElementById('btnDownloadPdf');
    
    if (btnPptx) {
        btnPptx.disabled = !currentPptxBlob && !originalPptxFile;
        btnPptx.style.opacity = (currentPptxBlob || originalPptxFile) ? '1' : '0.5';
    }
    
    if (btnPdf) {
        btnPdf.disabled = !currentPdfBlob;
        btnPdf.style.opacity = currentPdfBlob ? '1' : '0.5';
    }
}

function downloadPptx() {
    const blob = currentPptxBlob || originalPptxFile;
    if (!blob) {
        alert('PPTX dosyası bulunamadı!');
        return;
    }
    
    const fileName = originalPptxFile ? originalPptxFile.name : 'presentation.pptx';
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName.replace(/\.pptx?$/i, '_updated.pptx');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    showNotification('PPTX indirildi!', 'success');
}

function downloadPdf() {
    if (!currentPdfBlob) {
        alert('PDF dosyası bulunamadı!');
        return;
    }
    
    const fileName = originalPptxFile ? originalPptxFile.name.replace(/\.pptx?$/i, '.pdf') : 'presentation.pdf';
    const url = URL.createObjectURL(currentPdfBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    showNotification('PDF indirildi!', 'success');
}

// Download butonlarına event listener ekle
const btnPptx = document.getElementById('btnDownloadPptx');
const btnPdf = document.getElementById('btnDownloadPdf');

if (btnPptx) {
    btnPptx.addEventListener('click', downloadPptx);
}

if (btnPdf) {
    btnPdf.addEventListener('click', downloadPdf);
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

