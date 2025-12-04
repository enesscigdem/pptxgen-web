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
    } else if (jsonData.elements) {
        slides = [jsonData];
    } else {
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
        btn.innerHTML = `Slayt ${s.slideNumber || i + 1}`;
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
    canvas.style.height = "540px";
    canvas.style.position = "relative";
    canvas.style.background = "#fff";

    container.appendChild(canvas);

    let docWidthCm = slide.size?.widthCm || 33.87;
    let docHeightCm = slide.size?.heightCm || 19.05;

    let scaleX = 960 / (docWidthCm * 37.8);
    let scaleY = 540 / (docHeightCm * 37.8);

    slide.elements.sort((a,b) => (a.zIndex || 0) - (b.zIndex || 0));
    slide.elements.forEach(el => {
        if (!el.geometry) return;
        let div = document.createElement("div");
        div.className = "slide-element";
        div.style.position = "absolute";
        // apply positioning and sizing based on slide dimensions
        div.style.left = (el.geometry.xCm * 37.8 * scaleX) + "px";
        div.style.top = (el.geometry.yCm * 37.8 * scaleY) + "px";
        div.style.width = (el.geometry.wCm * 37.8 * scaleX) + "px";
        div.style.height = (el.geometry.hCm * 37.8 * scaleY) + "px";
        // apply z-index if present
        if (el.zIndex !== undefined && el.zIndex !== null) div.style.zIndex = String(el.zIndex);

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
            let imgError = false;
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
                if (el.imageBase64 && el.imageBase64.startsWith("data:image/tiff")) {
                    div.style.background = "#ddd";
                    div.style.display = "flex";
                    div.style.alignItems = "center";
                    div.style.justifyContent = "center";
                    div.textContent = "[TIFF not supported]";
                    canvas.appendChild(div);
                    return;
                }
                img.src = el.imageBase64 || '';
                img.style.width = '100%';
                img.style.height = '100%';
                img.style.objectFit = 'cover';
                img.onerror = function () {
                    imgError = true;
                    img.style.display = 'none';
                    let err = document.createElement('div');
                    err.style.color = '#bbb';
                    err.style.background = '#f8f8f8';
                    err.style.width = '100%';
                    err.style.height = '100%';
                    err.style.display = 'flex';
                    err.style.alignItems = 'center';
                    err.style.justifyContent = 'center';
                    err.textContent = '[img unavailable]';
                    div.appendChild(err);
                };
                div.appendChild(img);
            }
            canvas.appendChild(div);
            return;
        }

        // handle shapes (text and non-text)
        if (el.kind === 'shape') {
            // her shape DOM'a çizilecek!
            // fillColor atanmazsa, test amacıyla debug background ekleyelim
            if (el.style?.fillColor !== undefined && el.style?.fillColor !== null) {
                div.style.background = el.style.fillColor;
            } else {
                // debug/test: hiç renk yoksa %100 görülsün diye gridli arka plan koy
                div.style.background = 'repeating-linear-gradient(45deg, #eee 0 8px,#fff 8px 16px)';
            }
            if (el.style?.outlineColor !== undefined && el.style?.outlineColor !== null) {
                let widthEmu = el.style.outlineWidth || 0;
                let px = widthEmu ? (widthEmu / 12700) * 1.333 : 0;
                div.style.border = `${px.toFixed(1)}px solid ${el.style.outlineColor}`;
            } else {
                div.style.border = '1px dashed #abb'; // debug amaçlı
            }
            if (el.shapeType) {
                let st = el.shapeType.toLowerCase();
                if (st === 'ellipse' || st === 'roundrect' || st === 'oval') {
                    if (st === 'ellipse' || st === 'oval') div.style.borderRadius = '50%';
                    else if (st === 'roundrect') div.style.borderRadius = '12px';
                }
            }
            // shape'in metni varsa çiz...
            if (el.content && el.content.paragraphs && el.content.paragraphs.length) {
                let align = el.style?.align || el.content.paragraphs[0]?.align || null;
                if (align) {
                    if (align === 'ctr' || align === 'center') div.style.textAlign = 'center';
                    else if (align === 'r') div.style.textAlign = 'right';
                    else if (align === 'l') div.style.textAlign = 'left';
                }
                el.content.paragraphs.forEach((p, pIndex) => {
                    let pEl = document.createElement('div');
                    pEl.style.lineHeight = '1.2';
                    if (p.bullet && p.bullet.type && p.bullet.type !== 'none') {
                        let bulletSpan = document.createElement('span');
                        bulletSpan.style.display = 'inline-block';
                        bulletSpan.style.width = '1em';
                        bulletSpan.style.marginRight = '4px';
                        if (p.bullet.type === 'bullet') bulletSpan.textContent = '•';
                        else if (p.bullet.type === 'number') bulletSpan.textContent = `${pIndex + 1}.`;
                        pEl.appendChild(bulletSpan);
                    }
                    p.runs.forEach(run => {
                        let span = document.createElement('span');
                        span.textContent = run.text;
                        if (run.style?.fontSize !== undefined && run.style?.fontSize !== null)
                            span.style.fontSize = (((run.style.fontSize * 96) / 72) * scaleY) + 'px';
                        else if (el.style?.fontSize !== undefined && el.style?.fontSize !== null)
                            span.style.fontSize = (((el.style.fontSize * 96) / 72) * scaleY) + 'px';
                        if (run.style?.fontFamily !== undefined && run.style?.fontFamily !== null)
                            span.style.fontFamily = run.style.fontFamily;
                        else if (el.style?.fontFamily !== undefined && el.style?.fontFamily !== null)
                            span.style.fontFamily = el.style.fontFamily;
                        if (run.style?.bold !== undefined && run.style?.bold !== null)
                            span.style.fontWeight = run.style.bold ? 'bold' : 'normal';
                        else if (el.style?.bold !== undefined && el.style?.bold !== null)
                            span.style.fontWeight = el.style.bold ? 'bold' : 'normal';
                        if (run.style?.italic !== undefined && run.style?.italic !== null)
                            span.style.fontStyle = run.style.italic ? 'italic' : 'normal';
                        else if (el.style?.italic !== undefined && el.style?.italic !== null)
                            span.style.fontStyle = el.style.italic ? 'italic' : 'normal';
                        if (run.style?.underline !== undefined && run.style?.underline !== null)
                            span.style.textDecoration = run.style.underline ? 'underline' : 'none';
                        else if (el.style?.underline !== undefined && el.style?.underline !== null)
                            span.style.textDecoration = el.style.underline ? 'underline' : 'none';
                        if (run.style?.color !== undefined && run.style?.color !== null)
                            span.style.color = run.style.color;
                        else if (el.style?.color !== undefined && el.style?.color !== null)
                            span.style.color = el.style.color;
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
            if (el.style?.fontSize !== undefined && el.style?.fontSize !== null) {
                let fSizePx = (el.style.fontSize * 96) / 72;
                div.style.fontSize = (fSizePx * scaleY) + 'px';
            }
            if (el.style?.bold !== undefined && el.style?.bold !== null)
                div.style.fontWeight = el.style.bold ? 'bold' : 'normal';
            if (el.style?.italic !== undefined && el.style?.italic !== null)
                div.style.fontStyle = el.style.italic ? 'italic' : 'normal';
            if (el.style?.underline !== undefined && el.style?.underline !== null)
                div.style.textDecoration = el.style.underline ? 'underline' : 'none';
            if (el.style?.color !== undefined && el.style?.color !== null)
                div.style.color = el.style.color;
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
