# Docker Deployment

## Requisitos

- Docker >= 20.10
- Docker Compose >= 1.29

## Ejecución local con Docker Compose

```bash
# Construir e iniciar los servicios
docker-compose up --build

# O para correr en segundo plano
docker-compose up -d --build
```

El frontend estará disponible en `http://localhost:80` y el backend en `http://localhost:8000`.

## Estructura

- **Backend**: FastAPI en Python, expuesto en puerto 8000
- **Frontend**: React + Vite, servido por Nginx en puerto 80
- **Network**: Bridge network `app-network` para comunicación entre servicios

## Variables de entorno

### Backend (integrador.py)

- `PORT`: Puerto en el que escucha (default: 8000)
- `ALLOWED_ORIGINS`: Origins permitidos para CORS (default: localhost + Netlify)
- `PYTHONUNBUFFERED`: 1 para logs en tiempo real

### Frontend

- `VITE_API_URL`: URL del backend (default: http://backend:8000)

## Comandos útiles

```bash
# Ver logs
docker-compose logs -f

# Logs de un servicio específico
docker-compose logs -f backend
docker-compose logs -f frontend

# Detener servicios
docker-compose down

# Reconstruir imágenes
docker-compose build

# Eliminar volúmenes (resetea caché)
docker-compose down -v
```

## Build para producción

```bash
# Build solo del frontend (build Vite + nginx)
docker build -f frontend/Dockerfile -t conciliador-frontend:latest ./frontend

# Build solo del backend
docker build -f backend/Dockerfile -t conciliador-backend:latest ./backend
```

## Despliegue en Render/Heroku

Para Render.yaml o Procfile, usar:

```bash
# Backend
gunicorn main:app --bind 0.0.0.0:8000

# O directo con uvicorn
uvicorn main:app --host 0.0.0.0 --port 8000
```

Para el frontend, usar Node.js buildpack + script de build en package.json.

## Troubleshooting

- **502 Bad Gateway**: Backend no responde. Verificar logs: `docker-compose logs backend`
- **Frontend no carga**: CORS issues. Verificar `ALLOWED_ORIGINS` en backend.
- **API no funciona**: Verificar que `VITE_API_URL` apunta a backend correcto.

