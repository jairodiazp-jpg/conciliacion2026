# Logo CALYPSO - Documentación de Diseño

## Descripción del Proyecto

**CALYPSO** es el nombre diseñado para el sistema de conciliación contable. El logo refleja profesionalismo, modernidad y la naturaleza fluida del procesamiento de datos contables.

## Archivos del Logo

### 1. Logo Principal Horizontal
- **Archivo**: `calypso-logo.svg`
- **Dimensiones**: 400x120px
- **Uso**: Encabezados, sidebar, documentación, presentaciones
- **Características**:
  - Icono cuadrado con esquinas redondeadas (90x90px)
  - Texto "CALYPSO" en tipografía Space Grotesk Bold
  - Línea decorativa inferior
  - Tagline "Conciliador" en tamaño reducido

### 2. Favicon/Icono
- **Archivo**: `calypso-favicon.svg`
- **Dimensiones**: 100x100px
- **Uso**: Favicon del navegador, iconos de aplicaciones, avatares
- **Características**:
  - Versión simplificada del icono
  - Mantiene la identidad visual en espacios pequeños

## Paleta de Colores

### Rojo Principal
- **Rojo vibrante**: `#DC2626` ( Tailwind Red-600 )
- **Rojo oscuro**: `#991B1B` ( Tailwind Red-800 )
- **Rojo brillante**: `#EF4444` ( Tailwind Red-500 )

### Gradientes
El logo utiliza gradientes rojos que van desde tonos más brillantes hasta más oscuros, creando profundidad y dinamismo.

## Elementos de Diseño

### Icono
- **Forma**: Rectángulo con esquinas redondeadas (rx="22")
- **Patrón**: Tres líneas onduladas que representan:
  - Flujo de datos contables
  - Movimiento y procesamiento
  - Conexión entre diferentes elementos
- **Efectos**: Sombra suave y brillo superior para dar sensación de volumen

### Tipografía
- **Fuente principal**: Space Grotesk (Bold, 700)
- **Fuente secundaria**: Space Grotesk (Regular, 400)
- **Letter spacing**: 2px para "CALYPSO", 4px para "Conciliador"

## Cómo Usar en el Proyecto

### En el Frontend (React/Vite)

1. **Importar en componentes**:
```jsx
import logo from './public/calypso-logo.svg';
// o
const logoUrl = '/calypso-logo.svg';
```

2. **Usar en el sidebar** (reemplazando el brand-mark actual):
```jsx
<div className="brand-box">
  <img src="/calypso-logo.svg" alt="CALYPSO" className="brand-logo" />
  <div>
    <h1 className="brand-name">CALYPSO</h1>
    <p className="brand-subtitle">Conciliador Contable</p>
  </div>
</div>
```

3. **Como favicon** en `frontend/index.html`:
```html
<link rel="icon" type="image/svg+xml" href="/calypso-favicon.svg" />
```

### Estilos CSS Sugeridos

```css
.brand-logo {
  width: 48px;
  height: 48px;
  border-radius: 12px;
}

/* Para el logo completo en header */
.logo-horizontal {
  max-width: 300px;
  height: auto;
}
```

## Significado del Nombre

**CALYPSO** evoca:
- **Calipso**: Ninfa de la mitología griega conocida por su capacidad de retener y transformar
- **Calypso**: Género musical caribeño que representa armonía y ritmo
- **Sonoridad**: Suena moderno, profesional y fácil de recordar

El nombre refleja la esencia del sistema: retener, procesar y transformar datos contables de manera armoniosa y eficiente.

## Variantes y Consideraciones

### Para Fondos Oscuros
El logo está optimizado para fondos claros. Para fondos oscuros, se recomienda:
- Usar la versión con fondo blanco
- O crear una versión con texto blanco y icono claro

### Tamaños Mínimos
- **Logo horizontal**: No reducir de 150px de ancho
- **Favicon**: No reducir de 32x32px

### Formatos Alternativos
Los archivos SVG son escalables y mantienen calidad en cualquier tamaño. Para usos específicos:
- **Web**: SVG (recomendado)
- **Impresión**: SVG o exportar a PDF/AI
- **Redes sociales**: PNG con fondo transparente (exportar desde SVG)

## Scripts de Generación

El logo fue generado mediante scripts Python incluidos en el directorio:
- `generate_logo.py`: Genera el logo principal
- `generate_favicon.py`: Genera el favicon

Para regenerar los logos, ejecutar:
```bash
cd frontend/public
python generate_logo.py
python generate_favicon.py
```

## Créditos

Diseño creado el 14 de mayo de 2026 para el proyecto Conciliador App.