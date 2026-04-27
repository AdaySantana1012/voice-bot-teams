# Voice Bot para Microsoft Teams

App de Teams que graba audio con el micrófono del móvil, transcribe con Web Speech API y envía el texto al bot de Copilot Studio via Direct Line.

---

## Estructura

```
teams-voice-app/
├── index.html        ← La tab principal (grabación + chat)
├── config.html       ← Página de configuración (para tabs de canal)
├── manifest.json     ← Manifiesto de la app de Teams
├── icon-color.png    ← Icono 192x192 (debes añadirlo)
├── icon-outline.png  ← Icono 32x32 en blanco (debes añadirlo)
└── README.md
```

---

## Pasos para desplegar

### 1. Sube los archivos a un servidor HTTPS

Teams requiere HTTPS. Opciones gratuitas/rápidas:

- **GitHub Pages**: Sube el repo y activa Pages → `https://tuuser.github.io/voice-bot/`
- **Azure Static Web Apps**: Gratis, fácil desde el portal de Azure
- **Netlify / Vercel**: Arrastrar la carpeta y ya tienes URL HTTPS

### 2. Actualiza `manifest.json`

Reemplaza `TU-URL-AQUI.com` con tu dominio real en:
- `developer.websiteUrl`
- `developer.privacyUrl`
- `developer.termsOfUseUrl`
- `staticTabs[0].contentUrl`
- `configurableTabs[0].configurationUrl`
- `validDomains`

### 3. Añade los iconos

- `icon-color.png`: 192×192 px, fondo de color
- `icon-outline.png`: 32×32 px, icono blanco sobre transparente

Puedes generarlos en https://www.canva.com o cualquier editor.

### 4. Empaqueta el .zip

El zip debe contener directamente (sin carpeta raíz):
```
manifest.json
icon-color.png
icon-outline.png
```

```bash
zip -j voicebot.zip manifest.json icon-color.png icon-outline.png
```

### 5. Sube a Teams

**Opción A – Solo para ti (más rápido):**
1. Teams → Aplicaciones → Administrar tus aplicaciones
2. → Cargar una aplicación → Cargar una aplicación personalizada
3. Selecciona el `voicebot.zip`

**Opción B – Para toda la organización:**
1. Portal de administración de Teams → Aplicaciones de Teams → Administrar aplicaciones
2. → Cargar → sube el zip

### 6. Configura el Direct Line

Al abrir la app en Teams:
1. Pega tu **Direct Line Secret Key** de Copilot Studio
2. (Copilot Studio → Configuración → Canales → Direct Line)
3. Pulsa "Guardar y conectar"

---

## Cómo funciona

```
Móvil (Teams) 
  → getUserMedia() pide micrófono del dispositivo
  → Web Speech API transcribe el audio en tiempo real
  → Texto → Direct Line API → Copilot Studio Bot
  → Respuesta del bot → mostrada en el chat
```

## Notas

- **Web Speech API** funciona en Chrome y en el WebView de Teams (Android/iOS).
- En iOS puede requerir que el usuario toque primero para activar el audio.
- Si necesitas más precisión en la transcripción, sustituye Web Speech API por **Azure Speech to Text** (añade tu clave en el código).
