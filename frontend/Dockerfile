# Frontend - Node.js + Vite
FROM node:18-alpine AS build

WORKDIR /app

ARG VITE_API_URL=http://localhost:8000
ENV VITE_API_URL=$VITE_API_URL

# Copiar package.json e instalar dependencias
COPY package.json package-lock.json ./
RUN npm ci

# Copiar código fuente
COPY . .

# Build
RUN npm run build

# Etapa de producción: servir con nginx
FROM nginx:alpine

# Copiar archivo nginx.conf personalizado
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Copiar build al directorio de nginx
COPY --from=build /app/dist /usr/share/nginx/html

EXPOSE 80

CMD ["nginx", "-g", "daemon off;"]
