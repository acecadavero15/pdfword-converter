FROM node:20-bookworm-slim
# LibreOffice plus fonts that match Arial, Times New Roman, Calibri and Cambria,
# so line breaks and page counts stay the same as in Word
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer fonts-liberation fonts-dejavu-core fonts-crosextra-carlito fonts-crosextra-caladea \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY package.json .
RUN npm install --omit=dev
COPY server.js .
ENV PORT=10000
EXPOSE 10000
CMD ["node", "server.js"]
