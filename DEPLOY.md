# 🚀 Ücretsiz Web Deployment Rehberi

Bu projeyi tamamen ücretsiz bir şekilde web'e yüklemek için iki seçenek var:

## 📋 Seçenek 1: Render.com (Önerilen - En Kolay)

### Avantajlar:
- ✅ Tamamen ücretsiz (Free tier)
- ✅ LibreOffice desteği (Docker ile)
- ✅ Kolay deploy
- ✅ Otomatik HTTPS
- ✅ Environment variables desteği

### Adımlar:

#### 1. GitHub'a Push Et
```bash
git add .
git commit -m "Deploy ready"
git push origin main
```

#### 2. Render.com'da Hesap Oluştur
1. https://render.com adresine git
2. "Get Started for Free" ile GitHub hesabınla giriş yap

#### 3. Hybrid Server Deploy Et
1. Render Dashboard'da "New +" → "Web Service"
2. GitHub repo'nu bağla
3. Ayarlar:
   - **Name:** `pptx-hybrid-server`
   - **Environment:** `Node`
   - **Build Command:** `echo "No build needed"`
   - **Start Command:** `node hybrid_server.js`
   - **Plan:** `Free`
   - **Environment Variables:**
     - `PORT` = `3001` (Render otomatik atar, ama ekleyebilirsin)

#### 4. Convert Server Deploy Et
1. Yine "New +" → "Web Service"
2. Aynı repo'yu seç
3. Ayarlar:
   - **Name:** `pptx-convert-server`
   - **Environment:** `Node`
   - **Build Command:** `echo "No build needed"`
   - **Start Command:** `node convert_server.js`
   - **Plan:** `Free`
   - **Environment Variables:**
     - `PORT` = `3002`

#### 5. LibreOffice Kurulumu
Render.com'da LibreOffice kurmak için **Dockerfile** kullanmalısın:

`Dockerfile` oluştur (her iki server için):
```dockerfile
FROM node:18-slim

# LibreOffice kurulumu
RUN apt-get update && \
    apt-get install -y libreoffice && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . .

EXPOSE 3001
CMD ["node", "hybrid_server.js"]
```

**VEYA** Render'ın otomatik kurulumunu kullan:
- Render.com'da "Environment" sekmesinde:
  - `INSTALL_LIBREOFFICE=true` ekle
  - Build command: `apt-get update && apt-get install -y libreoffice && node hybrid_server.js`

#### 6. Frontend'i GitHub Pages'e Deploy Et
1. GitHub repo'nda Settings → Pages
2. Source: `main` branch, `/` folder
3. `hybrid.js` dosyasında server URL'lerini güncelle:

```javascript
// hybrid.js içinde
const HYBRID_SERVER = 'https://pptx-hybrid-server.onrender.com';
const CONVERT_SERVER = 'https://pptx-convert-server.onrender.com';
```

#### 7. CORS Ayarları
Render.com'da otomatik HTTPS var, ama CORS ayarlarını kontrol et:
- `hybrid_server.js` ve `convert_server.js`'de zaten CORS headers var ✅

---

## 📋 Seçenek 2: Railway.app (Alternatif)

### Avantajlar:
- ✅ Ücretsiz tier ($5 kredi/ay)
- ✅ LibreOffice desteği
- ✅ Kolay deploy

### Adımlar:

1. https://railway.app → GitHub ile giriş
2. "New Project" → "Deploy from GitHub repo"
3. Her iki server için ayrı service oluştur
4. Environment variables ekle
5. LibreOffice için Dockerfile kullan (yukarıdaki gibi)

---

## 📋 Seçenek 3: Fly.io (Alternatif)

### Avantajlar:
- ✅ Ücretsiz tier (3 VM)
- ✅ LibreOffice desteği
- ✅ Global CDN

### Adımlar:

1. `flyctl` kurulumu
2. `fly launch` komutu ile deploy
3. LibreOffice için Dockerfile ekle

---

## 🔧 Frontend URL Güncelleme

Deploy sonrası `hybrid.js` ve `convert.js` dosyalarında server URL'lerini güncelle:

```javascript
// hybrid.js
const HYBRID_SERVER = 'https://your-hybrid-server.onrender.com';

// convert.js  
const CONVERT_SERVER = 'https://your-convert-server.onrender.com';
```

**VEYA** Environment variable kullan (daha iyi):

```javascript
// hybrid.js
const HYBRID_SERVER = window.HYBRID_SERVER_URL || 'http://localhost:3001';
```

HTML'de:
```html
<script>
  window.HYBRID_SERVER_URL = 'https://your-server.onrender.com';
</script>
```

---

## 🐳 Dockerfile Örneği (LibreOffice ile)

`Dockerfile`:
```dockerfile
FROM node:18-slim

# LibreOffice ve gerekli kütüphaneleri kur
RUN apt-get update && \
    apt-get install -y \
    libreoffice \
    libreoffice-writer \
    libreoffice-impress \
    && apt-get clean && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package.json .
COPY hybrid_server.js .
COPY convert_server.js .

EXPOSE 3001
CMD ["node", "hybrid_server.js"]
```

---

## ⚠️ Önemli Notlar

1. **LibreOffice Kurulumu:** Render.com'da LibreOffice kurmak için Dockerfile veya build script gerekli
2. **Free Tier Limitleri:**
   - Render: 750 saat/ay (yeterli)
   - Railway: $5 kredi/ay
   - Fly.io: 3 VM (yeterli)
3. **Cold Start:** Free tier'da ilk istek yavaş olabilir (server uyuyor)
4. **Timeout:** Render.com'da 30 saniye timeout var, büyük dosyalar için upgrade gerekebilir

---

## 🎯 Hızlı Başlangıç (Render.com)

1. GitHub'a push et
2. Render.com'da 2 web service oluştur
3. Dockerfile ekle (LibreOffice için)
4. Frontend'de URL'leri güncelle
5. GitHub Pages'de deploy et

**Toplam Süre:** ~15 dakika

---

## 📞 Sorun Giderme

### LibreOffice Bulunamıyor
- Dockerfile'da LibreOffice kurulumunu kontrol et
- Build logs'u kontrol et

### CORS Hatası
- Server'larda CORS headers kontrol et
- Frontend URL'ini kontrol et

### Timeout Hatası
- Büyük dosyalar için timeout artır
- Render.com'da upgrade gerekebilir

