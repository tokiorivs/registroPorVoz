// Pegar en Extensiones > Apps Script de tu hoja de Google Sheets.
// Voz: en Configuración del proyecto > Propiedades del script, crea:
//   FISH_API_KEY  (obligatoria)  tu API key de Fish Audio
//   FISH_VOICE_ID (opcional)     ID de la voz elegida en fish.audio
//   FISH_MODEL    (opcional)     por defecto s2.1-pro; usa s2.1-pro-free en el plan gratuito
//   APP_PASSWORD  (obligatoria)  contraseña de acceso; sin ella el script rechaza todo
//   GEMINI_API_KEY (obligatoria) API key de Google AI Studio, para limpiar lo que dice el participante
//   GEMINI_MODEL  (opcional)     por defecto gemini-3.5-flash-lite

// Extrae solo el dato pedido de lo que dijo el participante, con Gemini.
// "mi medio de comunicación es el diario correo" -> "Diario Correo"
function extract(field, hint, text) {
  text = String(text || "").trim();
  field = String(field || "").slice(0, 60);
  hint = String(hint || "").slice(0, 160);
  if (!text || text.length > 300) return { ok: false, error: "texto inválido" };

  const props = PropertiesService.getScriptProperties();
  const system =
    "Extraes un dato de lo que una persona dijo en voz alta (transcripción en español que puede tener errores de reconocimiento). " +
    "Campo pedido: " + field + " (" + hint + "). " +
    "Devuelve SOLO ese dato, sin frases de relleno como 'mi nombre es' o 'trabajo en'. " +
    "Corrige mayúsculas y errores evidentes de transcripción (por ejemplo 'diario correo' -> 'Diario Correo'). " +
    "Acepta cualquier nombre aunque no lo conozcas o no parezca de un medio: 'las palmeritas' es válido y se devuelve como 'Las Palmeritas'. " +
    "Devuelve una cadena vacía SOLO si la persona no dio ningún nombre (dice que no sabe, o no responde a lo preguntado). " +
    "No inventes nada. " +
    'Responde en JSON con esta forma: {"value": "..."}';
  const model = props.getProperty("GEMINI_MODEL") || "gemini-3.5-flash-lite";

  const res = UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent", {
      method: "post",
      contentType: "application/json",
      headers: { "x-goog-api-key": props.getProperty("GEMINI_API_KEY") },
      payload: JSON.stringify({
        systemInstruction: { parts: [{ text: system }] },
        contents: [{ role: "user", parts: [{ text: text }] }],
        generationConfig: { temperature: 0, responseMimeType: "application/json" }
      }),
      muteHttpExceptions: true
    });
  if (res.getResponseCode() !== 200) {
    return { ok: false, error: "Gemini " + res.getResponseCode() + ": " + res.getContentText().slice(0, 200) };
  }
  try {
    const parts = JSON.parse(res.getContentText()).candidates[0].content.parts;
    const out = JSON.parse(parts.map(function (p) { return p.text || ""; }).join(""));
    return { ok: true, value: String(out.value || "").trim().slice(0, 120) };
  } catch (err) {
    return { ok: false, error: "respuesta inesperada" };
  }
}

// Convierte texto en audio mp3 (base64) con Fish Audio. La API key nunca sale de aquí.
function tts(text) {
  text = String(text || "").trim();
  if (!text || text.length > 200) return { ok: false, error: "texto inválido" };

  const cache = CacheService.getScriptCache();
  const key = "tts_" + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text, Utilities.Charset.UTF_8));
  const hit = cache.get(key);
  if (hit) return { ok: true, audio: hit };

  const props = PropertiesService.getScriptProperties();
  const res = UrlFetchApp.fetch("https://api.fish.audio/v1/tts", {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + props.getProperty("FISH_API_KEY"),
      model: props.getProperty("FISH_MODEL") || "s2.1-pro"
    },
    payload: JSON.stringify({
      text: text,
      reference_id: props.getProperty("FISH_VOICE_ID") || undefined,
      format: "mp3",
      mp3_bitrate: 64,
      latency: "low",
      language: "es-MX"
    }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    return { ok: false, error: "Fish " + res.getResponseCode() + ": " + res.getContentText().slice(0, 200) };
  }
  const audio = Utilities.base64Encode(res.getContent());
  try { cache.put(key, audio, 21600); } catch (_) {}   // 6 h; las frases repetidas no gastan créditos
  return { ok: true, audio: audio };
}
// Cada envío del formulario añade una fila: Fecha | Nombre | Medio de comunicación

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const expected = PropertiesService.getScriptProperties().getProperty("APP_PASSWORD");
    if (!expected || String(data.pw || "") !== expected) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "no autorizado" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (data.action === "auth") {
      return ContentService.createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (data.action === "tts") {
      return ContentService.createTextOutput(JSON.stringify(tts(data.text)))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (data.action === "extract") {
      return ContentService.createTextOutput(JSON.stringify(extract(data.field, data.hint, data.text)))
        .setMimeType(ContentService.MimeType.JSON);
    }
    // Nunca se escribe una fila vacía ni por una acción desconocida
    if (data.action || (!data.campo1 && !data.campo2)) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "petición no válida" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Fecha", "Nombre", "Medio de comunicación"]);
    }
    sheet.appendRow([new Date(), data.campo1 || "", data.campo2 || ""]);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
