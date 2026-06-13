# 🤖 Bot de Gastos Telegram — Instalación en 10 minutos

## Lo que necesitas
- Telegram en el celular
- Cuenta gratis en render.com
- API key de Anthropic (console.anthropic.com)

---

## PASO 1 — Crear tu bot en Telegram (2 minutos)

1. Abre Telegram y busca **@BotFather**
2. Escríbele: `/newbot`
3. Ponle un nombre: ej. `Bot Gastos Empresa`
4. Ponle un usuario: ej. `gastos_empresa_bot`
5. BotFather te entregará un **TOKEN** así:
   ```
   123456789:AAxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
   ```
   ⚠️ Guárdalo, lo necesitas en el Paso 3.

---

## PASO 2 — Subir el código a GitHub (3 minutos)

1. Ve a **github.com** y crea cuenta gratuita
2. Haz clic en **New repository** → nombre: `bot-gastos`
3. Sube todos los archivos de esta carpeta

---

## PASO 3 — Desplegar en Render (5 minutos)

1. Ve a **render.com** → crea cuenta gratuita
2. Haz clic en **New → Web Service**
3. Conecta tu repositorio de GitHub
4. Render detecta el `render.yaml` automáticamente
5. Agrega estas variables de entorno:
   ```
   TELEGRAM_TOKEN     = (el token de BotFather)
   ANTHROPIC_API_KEY  = (de console.anthropic.com)
   ```
6. Haz clic en **Deploy** y listo ✅

---

## PASO 4 — Usar el bot

Abre Telegram, busca tu bot por el nombre que le pusiste y escríbele:

```
/start
```

Luego envía una foto de boleta y el bot hace todo solo.

---

## Cómo usar el bot

| Acción | Qué hacer |
|--------|-----------|
| Registrar gasto | Enviar foto de boleta |
| Recibir Excel | Escribir `/planilla` |
| Cancelar | Tocar ❌ Cancelar |

### Flujo completo:
```
Tú  →  📷 [foto boleta]
Bot →  ✅ Fecha: 12/06/2026 | Monto: CLP $15.990 | Comercio: Jumbo
       ¿Motivo?

Tú  →  Alimentación
Bot →  ¿Destino?

Tú  →  Administrativo
Bot →  ¿Detalle? (o Omitir)

Tú  →  Omitir
Bot →  📋 Resumen... [✅ Confirmar] [❌ Cancelar]

Tú  →  ✅ Confirmar
Bot →  ✅ Gasto #0001 guardado. Planilla con 1 registro.

Tú  →  /planilla
Bot →  📎 [Registro_Gastos.xlsx]
```

---

## ⚠️ Nota sobre Render gratis
El servidor "duerme" tras 15 min sin uso.
El primer mensaje del día puede tardar ~30 segundos en responder.
Para evitarlo puedes usar **Railway** (también gratis con $5 de crédito mensual).
