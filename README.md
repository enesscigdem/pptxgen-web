# PPTX Analyzer - Hybrid Editor

## 📋 Genel Bakış

Bu proje, PowerPoint (PPTX) dosyalarını analiz edip, görselleştirip ve AI ile düzenleyebileceğiniz bir web uygulamasıdır.

## 🎯 Nasıl Çalışıyor?

### 1. PPTX → PDF Dönüşümü

**Server-Side (Node.js + LibreOffice):**
- Kullanıcı PPTX dosyasını yükler
- Dosya browser'da base64'e çevrilir
- Server'a gönderilir (`hybrid_server.js`)
- LibreOffice kullanılarak PPTX → PDF'ye çevrilir
- PDF browser'a geri gönderilir

**Neden PDF?**
- LibreOffice ile **pixel-perfect** görüntü elde ediyoruz
- Orijinal slaytların %100 aynısını gösteriyoruz
- HTML/CSS render'dan çok daha doğru

### 2. PDF Görselleştirme

**Client-Side (PDF.js):**
- PDF browser'da PDF.js kütüphanesi ile render edilir
- Her slide bir canvas üzerinde gösterilir
- Zoom, navigation gibi özellikler var

### 3. Metadata Extraction

**Client-Side (JSZip + DOMParser):**
- PPTX dosyası browser'da açılır (JSZip ile)
- XML dosyaları parse edilir (DOMParser ile)
- Her element için şu bilgiler extract edilir:
  - **Pozisyon:** x, y, width, height (EMU ve cm cinsinden)
  - **İçerik:** Text, images, charts
  - **Stil:** Font, renk, alignment, bold, italic, vb.
  - **Tip:** Shape, textbox, title, image, chart, vb.

### 4. Interactive Overlay

**Nasıl Çalışıyor?**
- PDF canvas'ının üzerine **invisible overlay** div'leri eklenir
- Her element için bir overlay div oluşturulur
- Overlay'ler element'in pozisyonuna göre yerleştirilir
- Kullanıcı overlay'e tıkladığında:
  - Element seçilir (orange border)
  - Metadata gösterilir (sol sidebar'da)
  - AI düzenleme için hazır hale gelir

**Pozisyon Hesaplama:**
```
PDF viewport scale = canvas width / (slide width in cm * 37.8)
Element X (px) = element X (cm) * 37.8 * scale
Element Y (px) = element Y (cm) * 37.8 * scale
```

### 5. Element Metadata

**Metadata Nereden Geliyor?**
- PPTX dosyasının XML'lerinden extract ediliyor
- Her slide için `slideX.xml` dosyası parse ediliyor
- Element'ler `spTree` içinde sırayla bulunuyor
- Her element'in:
  - `a:xfrm` → pozisyon ve boyut
  - `a:txBody` → text içeriği
  - `a:spPr` → stil (renk, border, vb.)
  - `a:rPr` → text stil (font, size, bold, vb.)

**Seçilince Ne Oluyor?**
- `selectElement()` fonksiyonu çağrılıyor
- Element'in JSON metadata'sı gösteriliyor
- Overlay'de orange border görünüyor

## 🛠️ Kullanılan Teknolojiler

### Frontend
- **HTML/CSS/JavaScript** - Temel web teknolojileri
- **PDF.js** - PDF render için (Mozilla'nın kütüphanesi)
- **JSZip** - PPTX dosyalarını açmak için
- **DOMParser** - XML parse etmek için

### Backend
- **Node.js** - Server runtime
- **LibreOffice** - PPTX → PDF conversion
- **HTTP Server** - Basit Node.js HTTP server

### Formatlar
- **PPTX** - PowerPoint dosya formatı (ZIP + XML)
- **PDF** - Görsel render için
- **JSON** - Metadata formatı

## 📁 Dosya Yapısı

```
pptxgen-web/
├── hybrid.html          # Ana sayfa (Hybrid Editor)
├── hybrid.js            # Client-side logic
├── hybrid_server.js     # Server (PDF conversion)
├── app.js               # Metadata extraction logic
├── editor.html          # Eski renderer (JSON → HTML)
├── editor.js            # Eski renderer logic
├── index.html           # JSON Extractor
└── styles.css           # Tüm sayfalar için CSS
```

## 🔄 İş Akışı

```
1. Kullanıcı PPTX yükler
   ↓
2. Browser: PPTX → base64
   ↓
3. Server: base64 → PPTX dosyası → LibreOffice → PDF
   ↓
4. Browser: PDF → PDF.js ile render
   ↓
5. Browser: PPTX → JSZip → XML parse → Metadata extract
   ↓
6. Browser: Metadata + PDF → Overlay oluştur
   ↓
7. Kullanıcı element'e tıklar → Metadata gösterilir
```

## 🤖 AI Agent Entegrasyonu

### Mevcut Durum

**Hazır Olan:**
- ✅ PDF görseli (AI görebilir)
- ✅ JSON metadata (AI anlayabilir)
- ✅ Element seçimi (hangi element düzenlenecek)
- ✅ UI güncelleme mekanizması

**Eksik Olan:**
- ❌ AI API entegrasyonu
- ❌ Metadata güncelleme
- ❌ PDF yeniden render
- ❌ Değişiklikleri kaydetme

### AI Agent Nasıl Çalışacak?

#### Senaryo: "İlk 3 slayttaki başlıkların rengi mavi ve 48px olsun"

**1. Kullanıcı Talimatı:**
```
"İlk 3 slayttaki başlıkların rengi mavi ve 48px olsun"
```

**2. AI'a Gönderilecek Veri:**
```json
{
  "instruction": "İlk 3 slayttaki başlıkların rengi mavi ve 48px olsun",
  "slides": [
    {
      "slideNumber": 1,
      "pdfImage": "base64...",  // PDF'den extract edilmiş görsel
      "elements": [
        {
          "id": "s1-el1",
          "type": "title",
          "content": { "text": "Başlık 1" },
          "style": { "color": "#000000", "fontSize": 24 },
          "geometry": { "x": 100, "y": 50, "width": 200, "height": 30 }
        }
      ]
    }
  ]
}
```

**3. AI'ın Yapacağı İşlem:**
```javascript
// AI şunu anlayacak:
// - "İlk 3 slide" → slide 1, 2, 3
// - "Başlıklar" → type === "title" olan elementler
// - "Mavi renk" → color: "#0000FF"
// - "48px" → fontSize: 48

// AI'ın döndüreceği:
{
  "changes": [
    {
      "slideNumber": 1,
      "elementId": "s1-el1",
      "updates": {
        "style": {
          "color": "#0000FF",
          "fontSize": 48
        }
      }
    },
    {
      "slideNumber": 2,
      "elementId": "s2-el1",
      "updates": {
        "style": {
          "color": "#0000FF",
          "fontSize": 48
        }
      }
    }
    // ...
  ]
}
```

**4. UI Güncelleme:**
```javascript
// 1. Metadata'yı güncelle
changes.forEach(change => {
  const slide = slideMetadata[change.slideNumber - 1];
  const element = slide.elements.find(el => el.id === change.elementId);
  if (element) {
    Object.assign(element.style, change.updates.style);
  }
});

// 2. Overlay'i güncelle (renk değişikliği görsel olarak)
updateOverlayStyles();

// 3. PDF'yi yeniden render et (veya overlay'de stil değişikliği göster)
// Not: PDF'yi değiştiremeyiz, ama overlay'de stil gösterebiliriz
```

### Implementasyon Adımları

#### Adım 1: AI API Entegrasyonu

```javascript
// hybrid.js içinde
async function applyAIEdit(slideIndex, instruction) {
    // 1. Seçili slide'ın PDF görselini al
    const pdfImage = await getSlideAsImage(slideIndex);
    
    // 2. Metadata'yı hazırla
    const slideData = {
        slideIndex,
        pdfImage,
        metadata: slideMetadata[slideIndex],
        instruction
    };
    
    // 3. AI API'ye gönder
    const response = await fetch('/api/ai-edit', {
        method: 'POST',
        body: JSON.stringify(slideData)
    });
    
    // 4. AI'ın döndürdüğü değişiklikleri uygula
    const result = await response.json();
    applyChanges(result.changes);
}
```

#### Adım 2: PDF'den Image Extract

```javascript
// PDF'den belirli bir sayfayı image'e çevir
async function getSlideAsImage(slideIndex) {
    const page = await pdfDoc.getPage(slideIndex + 1);
    const viewport = page.getViewport({ scale: 2 });
    
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    
    await page.render({
        canvasContext: canvas.getContext('2d'),
        viewport: viewport
    }).promise;
    
    // Canvas'ı base64 image'e çevir
    return canvas.toDataURL('image/png');
}
```

#### Adım 3: Değişiklikleri Uygula

```javascript
function applyChanges(changes) {
    changes.forEach(change => {
        const slide = slideMetadata[change.slideIndex];
        const element = slide.elements.find(el => el.id === change.elementId);
        
        if (element) {
            // Metadata'yı güncelle
            if (change.updates.style) {
                Object.assign(element.style, change.updates.style);
            }
            if (change.updates.content) {
                Object.assign(element.content, change.updates.content);
            }
            
            // Overlay'i güncelle
            updateElementOverlay(element);
        }
    });
    
    // Metadata panelini güncelle
    updateMetadataDisplay(currentSlideIndex);
}
```

#### Adım 4: Overlay'de Stil Göster

```javascript
function updateElementOverlay(element) {
    const overlay = document.querySelector(`[data-element-id="${element.id}"]`);
    if (!overlay) return;
    
    // Text element'ler için stil göster
    if (element.kind === 'shape' && element.content?.text) {
        // Overlay içine bir div ekle ve stil uygula
        const textDiv = overlay.querySelector('.element-text-preview');
        if (textDiv) {
            if (element.style.color) {
                textDiv.style.color = element.style.color;
            }
            if (element.style.fontSize) {
                textDiv.style.fontSize = element.style.fontSize + 'px';
            }
        }
    }
}
```

### AI API Seçenekleri

**1. OpenAI GPT-4 Vision:**
- PDF görselini görebilir
- JSON metadata'yı anlayabilir
- Talimatları işleyebilir

**2. Claude (Anthropic):**
- Vision desteği var
- JSON işleme güçlü

**3. Custom AI Model:**
- Kendi modelinizi eğitebilirsiniz
- PPTX düzenleme için özelleştirilebilir

### Örnek AI Prompt

```
Sen bir PowerPoint düzenleme asistanısın. Kullanıcının talimatını anlayıp, 
JSON metadata'yı güncelle.

Kullanıcı Talimatı: "İlk 3 slayttaki başlıkların rengi mavi ve 48px olsun"

Mevcut Metadata:
{
  "slides": [
    {
      "slideNumber": 1,
      "elements": [
        {
          "id": "s1-el1",
          "type": "title",
          "style": { "color": "#000000", "fontSize": 24 }
        }
      ]
    }
  ]
}

Görevin:
1. İlk 3 slide'ı bul (slideNumber: 1, 2, 3)
2. Her slide'da type === "title" olan elementleri bul
3. Bu elementlerin style.color = "#0000FF" yap
4. Bu elementlerin style.fontSize = 48 yap

Döndür:
{
  "changes": [
    {
      "slideIndex": 0,
      "elementId": "s1-el1",
      "updates": {
        "style": {
          "color": "#0000FF",
          "fontSize": 48
        }
      }
    }
  ]
}
```

## 🚀 Kullanım

### 1. Server'ı Başlat

```bash
node hybrid_server.js
```

Server `http://localhost:3001` portunda çalışacak.

### 2. Tarayıcıda Aç

`hybrid.html` dosyasını tarayıcıda açın.

### 3. PPTX Yükle

Sol sidebar'dan PPTX dosyanızı seçin.

### 4. Element Seç

Sağ taraftaki slide'da element'lere tıklayın, metadata'yı görün.

## 📝 Notlar

- **LibreOffice gerekli:** Server'da LibreOffice kurulu olmalı
- **Büyük dosyalar:** 200MB'a kadar destekleniyor
- **Timeout:** 10 dakika (büyük dosyalar için)

## 🔮 Gelecek Geliştirmeler

- [ ] AI agent entegrasyonu
- [ ] Değişiklikleri PPTX'e geri kaydetme
- [ ] Real-time collaboration
- [ ] Export options (PDF, PNG, vb.)

