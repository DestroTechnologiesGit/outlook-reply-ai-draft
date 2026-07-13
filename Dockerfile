FROM node:20-alpine

WORKDIR /app

# Install production deps first so Docker layer-caches them across code changes.
COPY server/package.json server/package-lock.json server/
RUN cd server && npm ci --omit=dev

COPY server/ server/
COPY src/ src/

ENV PORT=3000
EXPOSE 3000

CMD ["node", "server/server.js"]
