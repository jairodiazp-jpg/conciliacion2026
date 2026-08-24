#!/usr/bin/env python3
"""Genera el favicon de CALYPSO en formato SVG (versión cuadrada del icono)"""

svg_content = '''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" width="100" height="100">
  <defs>
    <linearGradient id="redGradient" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" style="stop-color:#DC2626;stop-opacity:1" />
      <stop offset="100%" style="stop-color:#991B1B;stop-opacity:1" />
    </linearGradient>
    <filter id="iconShadow" x="-50%" y="-50%" width="200%" height="200%">
      <feDropShadow dx="0" dy="4" stdDeviation="6" flood-color="#DC2626" flood-opacity="0.4"/>
    </filter>
  </defs>
  <g filter="url(#iconShadow)">
    <rect x="0" y="0" width="100" height="100" rx="22" fill="url(#redGradient)" />
    <path d="M 20 45 Q 35 25, 50 45 T 80 45" stroke="white" stroke-width="4" fill="none" stroke-linecap="round" opacity="0.9"/>
    <path d="M 15 55 Q 30 35, 45 55 T 75 55" stroke="white" stroke-width="3" fill="none" stroke-linecap="round" opacity="0.7"/>
    <path d="M 25 65 Q 40 50, 55 65" stroke="white" stroke-width="2.5" fill="none" stroke-linecap="round" opacity="0.5"/>
    <ellipse cx="50" cy="25" rx="25" ry="12" fill="white" opacity="0.15"/>
  </g>
</svg>'''

with open('calypso-favicon.svg', 'w', encoding='utf-8') as f:
    f.write(svg_content)

print('Favicon CALYPSO creado exitosamente en: calypso-favicon.svg')