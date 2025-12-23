# =============================================
# DOCKERFILE - CONSTRUCTOR DE DOCUMENTOS SUMATE
# =============================================

# Usar Node.js 20 Alpine para compatibilidad con ExcelJS
FROM node:20-alpine

# Establecer directorio de trabajo
WORKDIR /app

# Instalar dependencias del sistema necesarias para ExcelJS
RUN apk add --no-cache \
    tini \
    python3 \
    make \
    g++ \
    && rm -rf /var/cache/apk/*

# Copiar package files
COPY package*.json ./

# Instalar dependencias de producción
RUN npm install --only=production && npm cache clean --force

# Copiar código fuente
COPY . .

# Crear usuario no-root
RUN addgroup -g 1001 -S nodejs && \
    adduser -S nextjs -u 1001

# Cambiar ownership de la app
RUN chown -R nextjs:nodejs /app
USER nextjs

# Exponer puerto
EXPOSE 3001

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
  CMD node -e "require('http').get('http://localhost:3001/health', (res) => { process.exit(res.statusCode === 200 ? 0 : 1); })\" || exit 1

# Usar tini como PID 1 para manejo correcto de señales
ENTRYPOINT ["/sbin/tini", "--"]

# Comando por defecto
CMD ["node", "server.js"]