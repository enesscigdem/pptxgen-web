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

# Port'u environment variable'dan al
EXPOSE 3001

# Default olarak hybrid_server çalıştır
# Render.com'da startCommand ile override edilebilir
CMD ["node", "hybrid_server.js"]

