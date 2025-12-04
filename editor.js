let jsonData = null;
let slides = [];
let activeSlideIndex = 0;

// --- JSON yükleme ---
document.getElementById("btnRender").onclick = () => {
    let txt = document.getElementById("jsonInput").value.trim();

    if (txt) {
        try {
            jsonData = JSON.parse(txt);
            normalizeSlides();
            renderSlides();
        } catch {
            alert("JSON biçimi hatalı");
        }
        return;
    }

    let file = document.getElementById("fileJson").files[0];
    if (!file) return alert("JSON seçin veya yapıştırın");

    file.text().then(t => {
        try {
            jsonData = JSON.parse(t);
            normalizeSlides();
            renderSlides();
        } catch {
            alert("JSON formatı hatalı");
        }
    });
};

// --- JSON format uyarlayıcı ---
function normalizeSlides() {
    if (jsonData.slides) {
        slides = jsonData.slides;
    }
    else if (jsonData.elements) {
        slides = [jsonData];
    }
    else {
        slides = [];
    }
}

// --- Slide listesi UI oluştur ---


function renderSlides() {

    if (!slides.length) {
        alert("JSON içinde slayt bulunamadı.");
        return;
    }

    let slideList = document.getElementById("slideList");
    slideList.innerHTML = "";

    slides.forEach((s, i) => {
        let btn = document.createElement("div");
        btn.className = "thumb p-1 cursor-pointer text-center text-xs";
        btn.innerHTML = `Slayt ${s.slideNumber || i+1}`;
        btn.onclick = () => showSlide(i);
        slideList.appendChild(btn);
    });

    showSlide(0);
}

// --- Slide render ----
function showSlide(index) {

    activeSlideIndex = index;

    let container = document.getElementById("canvasContainer");
    container.innerHTML = "";

    let slide = slides[index];

    let canvas = document.createElement("div");
    canvas.className = "slideCanvas";
    canvas.style.width = "960px";
    canvas.style.height = "540px";
    canvas.style.position = "relative";
    canvas.style.background = "#fff";

    container.appendChild(canvas);

    let docWidthCm  = slide.size?.widthCm  || 33.87;
    let docHeightCm = slide.size?.heightCm || 19.05;

    let scaleX = 960 / (docWidthCm * 37.8);
    let scaleY = 540 / (docHeightCm * 37.8);

    slide.elements.forEach(el => {
        if (!el.geometry) return;
        let div = document.createElement("div");
        div.className = "slide-element";
        div.style.position = "absolute";
        // apply positioning and sizing based on slide dimensions
        div.style.left   = (el.geometry.xCm * 37.8 * scaleX) + "px";
        div.style.top    = (el.geometry.yCm * 37.8 * scaleY) + "px";
        div.style.width  = (el.geometry.wCm * 37.8 * scaleX) + "px";
        div.style.height = (el.geometry.hCm * 37.8 * scaleY) + "px";
        // apply z-index if present
        if (el.zIndex) div.style.zIndex = String(el.zIndex);

        // handle charts
        if (el.kind === "chart") {
            div.style.background = "#dde7ff";
            div.style.border = "2px solid #4f46e5";
            div.style.borderRadius = "50%";
            canvas.appendChild(div);
            return;
        }

        // handle images
        if (el.type === "image") {
            // handle EMF/WMF images by using object tag for better browser support
            let isEmf = false;
            if (el.imageName) {
                let ext = el.imageName.split('.').pop().toLowerCase();
                if (ext === 'emf' || ext === 'wmf') isEmf = true;
            } else if (el.imageBase64) {
                isEmf = el.imageBase64.startsWith('data:image/emf') || el.imageBase64.startsWith('data:image/wmf');
            }
            if (isEmf) {
                let obj = document.createElement('object');
                obj.data = el.imageBase64 || '';
                obj.type = el.imageBase64 && el.imageBase64.startsWith('data:image/wmf') ? 'image/x-wmf' : 'image/x-emf';
                obj.style.width = '100%';
                obj.style.height = '100%';
                div.appendChild(obj);
            } else {
                let img = document.createElement('img');
                img.src = el.imageBase64 || '';
                img.style.width = '100%';
                img.style.height = '100%';
                // keep aspect ratio by covering container area
                img.style.objectFit = 'cover';
                div.appendChild(img);
            }
            canvas.appendChild(div);
            return;
        }

        // handle shapes (text and non-text)
        if (el.kind === 'shape') {
            // background and border for shapes without or with text
            if (el.style?.fillColor) {
                div.style.background = el.style.fillColor;
            }
            if (el.style?.outlineColor) {
                let widthEmu = el.style.outlineWidth || 0;
                // convert outline width EMU to px: 12700 EMU ≈ 1 pt; 1pt ≈ 1.333px
                let px = widthEmu ? (widthEmu / 12700) * 1.333 : 0;
                div.style.border = `${px.toFixed(1)}px solid ${el.style.outlineColor}`;
            }
            // apply borderRadius for roundRect or ellipse
            if (el.shapeType) {
                let st = el.shapeType.toLowerCase();
                if (st === 'ellipse' || st === 'ellipse' || st === 'roundrect' || st === 'oval') {
                    // if ellipse, circle border radius 50%
                    if (st === 'ellipse' || st === 'oval') div.style.borderRadius = '50%';
                    else if (st === 'roundrect') div.style.borderRadius = '12px';
                }
            }
            // if shape has text content
            if (el.content && el.content.paragraphs && el.content.paragraphs.length) {
                // determine text alignment
                let align = el.style?.align || el.content.paragraphs[0]?.align || null;
                if (align) {
                    if (align === 'ctr' || align === 'center') div.style.textAlign = 'center';
                    else if (align === 'r') div.style.textAlign = 'right';
                    else if (align === 'l') div.style.textAlign = 'left';
                }
                // build paragraphs
                el.content.paragraphs.forEach((p, pIndex) => {
                    let pEl = document.createElement('div');
                    pEl.style.lineHeight = '1.2';
                    // bullet handling
                    if (p.bullet && p.bullet.type && p.bullet.type !== 'none') {
                        let bulletSpan = document.createElement('span');
                        bulletSpan.style.display = 'inline-block';
                        bulletSpan.style.width = '1em';
                        bulletSpan.style.marginRight = '4px';
                        if (p.bullet.type === 'bullet') bulletSpan.textContent = '•';
                        else if (p.bullet.type === 'number') bulletSpan.textContent = `${pIndex + 1}.`;
                        pEl.appendChild(bulletSpan);
                    }
                    // runs
                    p.runs.forEach(run => {
                        let span = document.createElement('span');
                        span.textContent = run.text;
                        // font size: use scaleY for vertical scaling
                        let fSizePt = run.style?.fontSize || el.style?.fontSize || 14;
                        // Convert point size to pixels (1pt = 1.333px at 96 DPI) then scale
                        let fSizePx = (fSizePt * 96) / 72;
                        span.style.fontSize = (fSizePx * scaleY) + 'px';
                        span.style.fontFamily = run.style?.fontFamily || el.style?.fontFamily || 'Arial';
                        span.style.fontWeight = run.style?.bold ? 'bold' : (el.style?.bold ? 'bold' : 'normal');
                        span.style.fontStyle = run.style?.italic ? 'italic' : (el.style?.italic ? 'italic' : 'normal');
                        span.style.textDecoration = run.style?.underline || el.style?.underline ? 'underline' : 'none';
                        span.style.color = run.style?.color || el.style?.color || '#000';
                        pEl.appendChild(span);
                    });
                    div.appendChild(pEl);
                });
            }
            canvas.appendChild(div);
            return;
        }

        // fallback: just insert text if present
        if (el.content && el.content.text) {
            div.textContent = el.content.text;
            let fSizePt = el.style?.fontSize || 14;
            // convert pt to px then scale
            let fSizePx = (fSizePt * 96) / 72;
            div.style.fontSize = (fSizePx * scaleY) + 'px';
            div.style.fontWeight = el.style?.bold ? 'bold' : 'normal';
            div.style.fontStyle = el.style?.italic ? 'italic' : 'normal';
            div.style.textDecoration = el.style?.underline ? 'underline' : 'none';
            div.style.color = el.style?.color || '#000';
            if (el.style?.align) {
                if (el.style.align === 'ctr' || el.style.align === 'center') div.style.textAlign = 'center';
                else if (el.style.align === 'r') div.style.textAlign = 'right';
                else if (el.style.align === 'l') div.style.textAlign = 'left';
            }
            canvas.appendChild(div);
            return;
        }
    });

    // Thumbnail highlight
    [...document.getElementById("slideList").children].forEach((x, i) =>
        x.classList.toggle("active-slide", i === index)
    );
}

// --- Base64 resim çözümü ---
function findImageBase64(slide, element) {
    return slide.elements.find(x => x.id === element.id)?.base64 || "";
}
