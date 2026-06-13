# Bot de Gastos — Railway + Google Drive + Claude

## Resumen
- Railway: servidor 24/7 gratis ($5 crédito/mes renovable)
- Google Drive: fotos permanentes gratis (15GB)
- Claude: lee las boletas automáticamente ($5 = ~1500 boletas)

---

## PASO 1 — Crear bot en Telegram (2 min)
1. Abre Telegram → busca @BotFather
2. Escribe `/newbot` → sigue los pasos
3. Guarda el TOKEN que te entrega

---

## PASO 2 — Google Drive API (10 min)

1. Ve a https://console.cloud.google.com
2. Crea un proyecto nuevo (botón arriba a la izquierda)
3. Busca "Google Drive API" → habilitar
4. Ve a "Credenciales" → "Crear credenciales" → "Cuenta de servicio"
5. Ponle nombre: `bot-gastos` → Crear
6. Haz clic en la cuenta creada → pestaña "Claves"
7. "Agregar clave" → "Crear clave nueva" → JSON → Descargar
8. Abre el JSON descargado con el Bloc de notas → copia TODO el contenido

9. Crea una carpeta en Google Drive llamada "Gastos Bot"
10. Haz clic derecho → Compartir → pega el email de la cuenta de servicio
    (está en el JSON: campo "client_email", ej: bot-gastos@proyecto.iam.gserviceaccount.com)
11. Dale permiso de Editor → Compartir
12. Copia el ID de la carpeta desde la URL:
    https://drive.google.com/drive/folders/[ESTE_ES_EL_ID]

---

## PASO 3 — Subir a GitHub (3 min)

1. Ve a github.com → crea cuenta gratuita
2. Nuevo repositorio → nombre: `bot-gastos`
3. Sube todos los archivos de esta carpeta

---

## PASO 4 — Desplegar en Railway (5 min)

1. Ve a railway.app → "Start a New Project"
2. "Deploy from GitHub repo" → conecta tu repo
3. Ve a "Variables" y agrega estas 4:

   TELEGRAM_TOKEN
   → el token de BotFather

   ANTHROPIC_API_KEY
   → tu key de console.anthropic.com

   GOOGLE_SERVICE_ACCOUNT_JSON
   → pega TODO el contenido del archivo JSON descargado
     (en una sola línea, sin saltos de línea)

   GOOGLE_DRIVE_FOLDER_ID
   → el ID de la carpeta de Drive del paso 2

4. Railway despliega automáticamente → el bot queda activo 24/7

---

## Cómo usar el bot

| Acción | Qué hacer |
|--------|-----------|
| Registrar gasto | Enviar foto de boleta |
| Recibir Excel | Escribir /planilla |
| Cancelar | Tocar Cancelar |

## Flujo completo
```
Tú  → foto boleta
Bot → Fecha: 12/06/2026 | Monto: $9.350 | Comercio: Cafe Helado
      Cual es el MOTIVO?
Tú  → Alimentacion
Bot → Cual es el DESTINO?
Tú  → Administrativo
Bot → Detalle? (o Omitir)
Tú  → Omitir
Bot → Resumen... Confirmar?
Tú  → Confirmar
Bot → Gasto #0001 guardado!
      Foto en Drive: https://drive.google.com/...
```

## Qué se guarda
- Excel con todos los gastos acumulados (en Railway y en Drive)
- Fotos de las boletas en Google Drive como GASTO_0001.jpg, GASTO_0002.jpg...
- Resumen por Motivo y por Destino en el Excel
