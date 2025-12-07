// Hibrit Renderer: PDF görsel + JSON metadata
// Bu dosya örnek bir implementasyon gösterir

let pdfDoc = null;
let currentSlideIndex = 0;
let slideMetadata = [];

/**
 * PPTX'i PDF'ye çevir ve metadata extract et
 */
async function loadPresentationHybrid(pptxFile) {
    // 1. PPTX'i PDF'ye çevir (LibreOffice server kullan)
    const pdfBlob = await convertPptxToPdf(pptxFile);
    
    // 2. Metadata extract et (mevcut app.js logic'i kullan)
    const metadata = await extractMetadata(pptxFile);
    
    // 3. PDF'yi yükle
    pdfDoc = await pdfjsLib.getDocument(URL.createObjectURL(pdfBlob)).promise;
    
    // 4. Her slide için metadata'yı organize et
    slideMetadata = metadata.slides;
    
    return { pdfDoc, slideMetadata };
}

/**
 * Slide'ı render et: PDF görsel + interactive overlay
 */
async function renderSlideHybrid(slideIndex) {
    const canvas = document.getElementById('slideCanvas');
    const ctx = canvas.getContext('2d');
    
    // 1. PDF'den görseli render et
    const page = await pdfDoc.getPage(slideIndex + 1);
    const viewport = page.getViewport({ scale: 2 });
    
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    
    await page.render({
        canvasContext: ctx,
        viewport: viewport
    }).promise;
    
    // 2. Metadata'dan interactive overlay oluştur
    const metadata = slideMetadata[slideIndex];
    createInteractiveOverlay(metadata.elements, viewport);
}

/**
 * Interactive overlay oluştur - her element için clickable area
 */
function createInteractiveOverlay(elements, viewport) {
    const overlay = document.getElementById('elementOverlay');
    overlay.innerHTML = '';
    
    elements.forEach(el => {
        const div = document.createElement('div');
        div.className = 'element-overlay';
        div.dataset.elementId = el.id;
        
        // PDF viewport'a göre pozisyon hesapla
        const x = (el.geometry.xCm / 25.4) * viewport.scale * 96;
        const y = (el.geometry.yCm / 25.4) * viewport.scale * 96;
        const w = (el.geometry.wCm / 25.4) * viewport.scale * 96;
        const h = (el.geometry.hCm / 25.4) * viewport.scale * 96;
        
        div.style.left = x + 'px';
        div.style.top = y + 'px';
        div.style.width = w + 'px';
        div.style.height = h + 'px';
        
        // Hover'da element bilgilerini göster
        div.title = el.content?.text || el.type;
        div.onclick = () => editElement(el);
        
        overlay.appendChild(div);
    });
}

/**
 * Element düzenleme - AI agent buraya entegre edilecek
 */
function editElement(element) {
    // AI agent burada devreye girer
    // Kullanıcı: "Bu metni değiştir" → AI JSON'u günceller
    console.log('Editing element:', element);
    
    // Örnek: Text element'i düzenle
    if (element.type === 'text' || element.kind === 'shape') {
        const newText = prompt('Yeni metin:', element.content?.text || '');
        if (newText) {
            // JSON'u güncelle
            element.content.text = newText;
            // Overlay'i güncelle (veya PDF'i yeniden render et)
            updateElementInOverlay(element);
        }
    }
}

/**
 * AI Agent için API
 */
async function aiEditSlide(slideIndex, instruction) {
    const slide = slideMetadata[slideIndex];
    
    // AI'a gönder: görsel + metadata + instruction
    const response = await fetch('/api/ai-edit', {
        method: 'POST',
        body: JSON.stringify({
            slideIndex,
            pdfImage: await getSlideAsImage(slideIndex), // PDF'den image extract
            metadata: slide,
            instruction: instruction
        })
    });
    
    const result = await response.json();
    
    // AI'ın yaptığı değişiklikleri uygula
    result.changes.forEach(change => {
        const element = slide.elements.find(el => el.id === change.elementId);
        if (element) {
            Object.assign(element, change.updates);
        }
    });
    
    // Görseli güncelle
    await renderSlideHybrid(slideIndex);
}

/**
 * Helper: PPTX'i PDF'ye çevir
 */
async function convertPptxToPdf(pptxFile) {
    const buffer = await pptxFile.arrayBuffer();
    const base64 = btoa(String.fromCharCode(...new Uint8Array(buffer)));
    
    const response = await fetch('http://localhost:3000/convert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            fileName: pptxFile.name,
            fileData: base64,
            type: 'pptx-to-pdf'
        })
    });
    
    return await response.blob();
}

/**
 * Helper: Metadata extract (mevcut app.js logic'i)
 */
async function extractMetadata(pptxFile) {
    // Mevcut app.js'deki processFile logic'ini kullan
    // Sadece metadata için, görsel render için değil
    // Bu kısım mevcut kodunuzdan alınabilir
    return null; // Placeholder
}

