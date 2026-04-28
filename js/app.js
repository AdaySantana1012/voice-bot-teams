const APP_ID = "4a693834-2aa8-478e-a086-98d7de9ade7e";
const APP_DOMAIN = "voicebot.compracamper.com";
const RESOURCE = `api://${APP_DOMAIN}/${APP_ID}`;
const PROXY_BASE = "https://voicebot-proxy.azurewebsites.net/api";
const TOKEN_URL = `${PROXY_BASE}/token`;
const TRANSCRIBE_URL = `${PROXY_BASE}/transcribe`;
const DIRECTLINE_BASE =
  "https://europe.directline.botframework.com/v3/directline";
const MAX_FILE_BYTES = 4 * 1024 * 1024;

let ssoToken = null,
  dlToken = null,
  conversationId = null,
  watermark = null,
  userId = null,
  dlUserId = null,
  userName = "Usuario",
  polling = false;
let mediaRecorder = null,
  audioChunks = [],
  recordingStream = null,
  recordStart = 0,
  recordTimer = null;
let attachedFile = null;

async function init() {
  try {
    await microsoftTeams.app.initialize();
    setStatus("thinking", "Obteniendo credenciales...");
    ssoToken = await new Promise((r, j) => {
      microsoftTeams.authentication.getAuthToken({
        resources: [RESOURCE],
        successCallback: r,
        failureCallback: (e) =>
          j(new Error(typeof e === "string" ? e : JSON.stringify(e))),
      });
    });
    document.getElementById("dbSSO").textContent =
      ssoToken.substring(0, 20) + "...";
    try {
      const p = JSON.parse(atob(ssoToken.split(".")[1]));
      userId = p.oid || p.sub;
      dlUserId = `dl_${userId}`;
      userName = p.name || p.preferred_username || "Usuario";
      document.getElementById("dbUser").textContent =
        p.preferred_username || p.upn || userId || "unknown";
      document.getElementById("dbDlUser").textContent = dlUserId;
    } catch (_) {}
    setStatus("thinking", "Conectando con el asistente...");
    const tokenRes = await fetch(TOKEN_URL, {
      method: "GET",
      headers: { Authorization: `Bearer ${ssoToken}` },
    });
    if (!tokenRes.ok) {
      const e = await tokenRes.text();
      setStatus("error", `Error (${tokenRes.status})`);
      addBotMessage(`No se pudo conectar. ${e}`);
      return;
    }
    const td = await tokenRes.json();
    dlToken = td.token;
    document.getElementById("dbToken").textContent =
      dlToken.substring(0, 20) + "...";
      const convRes = await fetch(`${DIRECTLINE_BASE}/conversations`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${dlToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          user: { id: dlUserId, name: userName },
          tokenExchangeRequest: {
            uri: `api://${APP_DOMAIN}/${APP_ID}/access_as_user`,
            token: ssoToken
          }
        })
      });
    if (!convRes.ok) {
      const e = await convRes.text();
      setStatus("error", "Error conversación");
      addBotMessage(e);
      return;
    }
    const cd = await convRes.json();
    conversationId = cd.conversationId;
    document.getElementById("dbConv").textContent = conversationId;
    setStatus("ok", "Conectado");
    document.getElementById("sendBtn").disabled = false;
    document.getElementById("micBtn").disabled = false;
    document.getElementById("attachBtn").disabled = false;
    startContinuousPolling();
    await sendConversationUpdate();
  } catch (e) {
    setStatus("error", "Error de inicialización");
    addBotMessage(e.message);
  }
}

async function sendConversationUpdate() {
  const h = {
    Authorization: `Bearer ${dlToken}`,
    "Content-Type": "application/json",
  };
  const u = `${DIRECTLINE_BASE}/conversations/${conversationId}/activities`;
  try {
    await fetch(u, {
      method: "POST",
      headers: h,
      body: JSON.stringify({
        type: "conversationUpdate",
        from: { id: dlUserId, name: userName },
        membersAdded: [{ id: dlUserId, name: userName }],
      }),
    });
  } catch (e) {}
  try {
    await fetch(u, {
      method: "POST",
      headers: h,
      body: JSON.stringify({
        type: "event",
        name: "startConversation",
        from: { id: dlUserId, name: userName },
      }),
    });
  } catch (e) {}
}

function onFilePicked(ev) {
  const f = ev.target.files && ev.target.files[0];
  ev.target.value = "";
  if (!f) return;
  if (f.type !== "application/pdf") {
    addBotMessage("Solo se admiten archivos PDF.");
    return;
  }
  if (f.size > MAX_FILE_BYTES) {
    addBotMessage(
      `Archivo demasiado grande. Máximo ${(
        MAX_FILE_BYTES /
        1024 /
        1024
      ).toFixed(0)}MB.`
    );
    return;
  }
  attachedFile = {
    file: f,
    name: f.name,
    size: f.size,
    mimeType: f.type,
  };
  showFilePreview();
}
function showFilePreview() {
  if (!attachedFile) return;
  document.getElementById("filePreviewName").textContent =
    attachedFile.name;
  document.getElementById("filePreviewSize").textContent = formatBytes(
    attachedFile.size
  );
  document.getElementById("filePreview").classList.add("active");
}
function clearAttachedFile() {
  attachedFile = null;
  document.getElementById("filePreview").classList.remove("active");
}
function formatBytes(b) {
  if (b < 1024) return b + " B";
  if (b < 1048576) return (b / 1024).toFixed(1) + " KB";
  return (b / 1048576).toFixed(2) + " MB";
}

async function sendTextMessage() {
  const i = document.getElementById("textInput");
  const t = i.value.trim();
  if (!conversationId) return;
  if (!t && !attachedFile) return;
  i.value = "";
  if (attachedFile)
    addUserFileMessage(attachedFile.name, attachedFile.size);
  if (t) addUserTextMessage(t);
  if (attachedFile) await sendActivityWithUpload(t, attachedFile);
  else await sendActivityTextOnly(t);
  clearAttachedFile();
}

async function sendActivityTextOnly(text) {
  try {
    const r = await fetch(
      `${DIRECTLINE_BASE}/conversations/${conversationId}/activities`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${dlToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          type: "message",
          from: { id: dlUserId, name: userName },
          text: text || "",
        }),
      }
    );
    if (!r.ok) {
      addBotMessage(`Error: ${r.status}`);
      return;
    }
    setStatus("thinking", "Procesando...");
    showThinking();
    disableLastSuggestedActions();
  } catch (e) {
    addBotMessage("Error: " + e.message);
  }
}

async function sendActivityWithUpload(text, fileObj) {
  try {
    const f = new FormData();
    const a = {
      type: "message",
      from: { id: dlUserId, name: userName },
      text: text || "",
    };
    f.append(
      "activity",
      new Blob([JSON.stringify(a)], {
        type: "application/vnd.microsoft.activity",
      })
    );
    f.append("file", fileObj.file, fileObj.name);
    const u = `${DIRECTLINE_BASE}/conversations/${conversationId}/upload?userId=${encodeURIComponent(
      dlUserId
    )}`;
    const r = await fetch(u, {
      method: "POST",
      headers: { Authorization: `Bearer ${dlToken}` },
      body: f,
    });
    if (!r.ok) {
      addBotMessage(`Error upload: ${r.status}`);
      return;
    }
    setStatus("thinking", "Procesando...");
    showThinking();
    disableLastSuggestedActions();
  } catch (e) {
    addBotMessage("Error: " + e.message);
  }
}

async function sendActionAsUser(display, value) {
  addUserTextMessage(display);
  try {
    const r = await fetch(
      `${DIRECTLINE_BASE}/conversations/${conversationId}/activities`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${dlToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          type: "message",
          from: { id: dlUserId, name: userName },
          text: value || display,
        }),
      }
    );
    if (!r.ok) return;
    setStatus("thinking", "Procesando...");
    showThinking();
    disableLastSuggestedActions();
  } catch (e) {
    addBotMessage("Error: " + e.message);
  }
}

async function sendCardSubmit(data) {
  try {
    await fetch(
      `${DIRECTLINE_BASE}/conversations/${conversationId}/activities`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${dlToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          type: "message",
          from: { id: dlUserId, name: userName },
          value: data,
          text: "",
        }),
      }
    );
    setStatus("thinking", "Procesando...");
    showThinking();
  } catch (e) {
    addBotMessage("Error: " + e.message);
  }
}

async function toggleRecording() {
  if (!conversationId) return;
  if (mediaRecorder && mediaRecorder.state === "recording")
    stopRecording();
  else await startRecording();
}
async function startRecording() {
  try {
    recordingStream = await navigator.mediaDevices.getUserMedia({
      audio: true,
    });
    const c = [
      "audio/webm;codecs=opus",
      "audio/webm",
      "audio/mp4",
      "audio/ogg;codecs=opus",
    ];
    const m = c.find((t) => MediaRecorder.isTypeSupported(t)) || "";
    mediaRecorder = new MediaRecorder(
      recordingStream,
      m ? { mimeType: m } : undefined
    );
    audioChunks = [];
    mediaRecorder.addEventListener("dataavailable", (e) => {
      if (e.data.size > 0) audioChunks.push(e.data);
    });
    mediaRecorder.addEventListener("stop", async () => {
      const b = new Blob(audioChunks, {
        type: mediaRecorder.mimeType || "audio/webm",
      });
      recordingStream.getTracks().forEach((t) => t.stop());
      recordingStream = null;
      await handleAudioBlob(b);
    });
    mediaRecorder.start();
    recordStart = Date.now();
    document.getElementById("micBtn").classList.add("recording");
    document.getElementById("recIndicator").classList.add("active");
    document.getElementById("textInput").disabled = true;
    document.getElementById("sendBtn").disabled = true;
    document.getElementById("attachBtn").disabled = true;
    recordTimer = setInterval(() => {
      const s = Math.floor((Date.now() - recordStart) / 1000);
      document.getElementById("recTime").textContent = `${String(
        Math.floor(s / 60)
      ).padStart(2, "0")}:${String(s % 60).padStart(2, "0")}`;
    }, 250);
  } catch (e) {
    addBotMessage("Micrófono no accesible: " + e.message);
  }
}
function stopRecording() {
  if (mediaRecorder && mediaRecorder.state === "recording")
    mediaRecorder.stop();
  clearInterval(recordTimer);
  document.getElementById("micBtn").classList.remove("recording");
  document.getElementById("recIndicator").classList.remove("active");
  document.getElementById("textInput").disabled = false;
  document.getElementById("sendBtn").disabled = false;
  document.getElementById("attachBtn").disabled = false;
  document.getElementById("recTime").textContent = "00:00";
}
async function handleAudioBlob(blob) {
  const au = URL.createObjectURL(blob);
  const tr = addUserAudioMessage(au);
  try {
    const f = new FormData();
    f.append("audio", blob, "recording.webm");
    const r = await fetch(TRANSCRIBE_URL, {
      method: "POST",
      headers: { Authorization: `Bearer ${ssoToken}` },
      body: f,
    });
    if (!r.ok) {
      tr.classList.remove("loading");
      tr.textContent = "No se pudo transcribir.";
      return;
    }
    const d = await r.json();
    const t = (d.text || "").trim();
    tr.classList.remove("loading");
    tr.textContent = t || "(sin voz)";
    if (t) await sendActivityTextOnly(t);
  } catch (e) {
    tr.classList.remove("loading");
    tr.textContent = "Error: " + e.message;
  }
}

async function startContinuousPolling() {
  if (polling) return;
  polling = true;
  while (polling) {
    try {
      const u =
        `${DIRECTLINE_BASE}/conversations/${conversationId}/activities` +
        (watermark ? `?watermark=${watermark}` : "");
      const r = await fetch(u, {
        headers: { Authorization: `Bearer ${dlToken}` },
      });
      if (r.ok) {
        const d = await r.json();
        watermark = d.watermark;
        const msgs = (d.activities || []).filter(
          (a) =>
            a.type === "message" &&
            a.from &&
            a.from.id !== dlUserId &&
            a.from.id !== userId &&
            a.from.id !== "user1"
        );
        if (msgs.length > 0) {
          hideThinking();
          msgs.forEach((act) => renderBotActivity(act));
          setStatus("ok", "Conectado");
        }
      }
    } catch (e) {}
    await sleep(1500);
  }
}

function renderBotActivity(act) {
  if (act.text && act.text.trim()) addBotMessage(act.text);
  if (Array.isArray(act.attachments))
    act.attachments.forEach((att) => renderAttachment(att));
  if (
    act.suggestedActions &&
    Array.isArray(act.suggestedActions.actions) &&
    act.suggestedActions.actions.length > 0
  )
    renderSuggestedActions(act.suggestedActions.actions);
}
function renderAttachment(att) {
  const ct = att.contentType || "";
  if (ct === "application/vnd.microsoft.card.adaptive") {
    renderFullAdaptiveCard(att.content);
    return;
  }
  if (ct === "application/vnd.microsoft.card.hero") {
    renderHeroCard(att.content);
    return;
  }
}
function renderFullAdaptiveCard(cardJson) {
  const chat = document.getElementById("chat");
  const wrap = document.createElement("div");
  wrap.className = "msg bot";
  const l = document.createElement("div");
  l.className = "msg-label";
  l.textContent = "Asistente";
  wrap.appendChild(l);
  const host = document.createElement("div");
  host.className = "adaptive-card-host";
  try {
    const ac = new AdaptiveCards.AdaptiveCard();
    ac.parse(cardJson);
    ac.onExecuteAction = (action) => {
      if (action instanceof AdaptiveCards.OpenUrlAction)
        window.open(action.url, "_blank", "noopener");
      else if (action instanceof AdaptiveCards.SubmitAction) {
        const inputValues = ac.getAllInputs().reduce((acc, input) => {
          acc[input.id] = input.value;
          return acc;
        }, {});
      
        sendCardSubmit({
          ...(action.data || {}),
          ...inputValues,
        });
      }
    };
    const rendered = ac.render();
    host.appendChild(rendered);
  } catch (e) {
    host.textContent = "No se pudo mostrar la tarjeta.";
    console.error("AC render error:", e);
  }
  wrap.appendChild(host);
  chat.appendChild(wrap);
  chat.scrollTop = chat.scrollHeight;
}
function renderHeroCard(content) {
  const card = { body: [], actions: [] };
  if (content.title)
    card.body.push({
      type: "TextBlock",
      text: content.title,
      weight: "Bolder",
      size: "Medium",
    });
  if (content.subtitle)
    card.body.push({ type: "TextBlock", text: content.subtitle });
  if (content.text)
    card.body.push({ type: "TextBlock", text: content.text });
  card.actions = (content.buttons || []).map((b) =>
    b.type === "openUrl"
      ? { type: "Action.OpenUrl", title: b.title, url: b.value }
      : {
          type: "Action.Submit",
          title: b.title,
          data: { __action: b.value },
        }
  );
  renderFullAdaptiveCard(card);
}
function renderSuggestedActions(actions) {
  const chat = document.getElementById("chat");
  const wrap = document.createElement("div");
  wrap.className = "msg bot suggested-actions-wrap";
  const row = document.createElement("div");
  row.className = "suggested-actions";
  actions.forEach((a) => {
    const btn = document.createElement("button");
    btn.className = "suggested-action";
    btn.textContent = a.title || a.value || "...";
    btn.onclick = () => {
      const d = a.title || String(a.value);
      const v =
        a.type === "imBack" || a.type === "messageBack" || !a.type
          ? a.value || a.title
          : a.title || a.value;
      sendActionAsUser(d, v);
    };
    row.appendChild(btn);
  });
  wrap.appendChild(row);
  chat.appendChild(wrap);
  chat.scrollTop = chat.scrollHeight;
}
function disableLastSuggestedActions() {
  document
    .querySelectorAll(".suggested-actions .suggested-action")
    .forEach((b) => b.classList.add("disabled"));
}

function setStatus(type, text) {
  document.getElementById("statusBar").className = "status-bar " + type;
  document.getElementById("statusText").textContent = text;
}
function addUserTextMessage(text) {
  const c = document.getElementById("chat");
  const m = document.createElement("div");
  m.className = "msg user";
  m.innerHTML = `<div class="msg-label">Tú</div><div class="msg-bubble"></div>`;
  m.querySelector(".msg-bubble").textContent = text;
  c.appendChild(m);
  c.scrollTop = c.scrollHeight;
}
function addUserAudioMessage(url) {
  const c = document.getElementById("chat");
  const m = document.createElement("div");
  m.className = "msg user";
  m.innerHTML = `<div class="msg-label">Tú</div><div class="msg-bubble audio-bubble"><audio controls src="${url}"></audio><div class="audio-transcript loading">Transcribiendo...</div></div>`;
  c.appendChild(m);
  c.scrollTop = c.scrollHeight;
  return m.querySelector(".audio-transcript");
}
function addUserFileMessage(name, size) {
  const c = document.getElementById("chat");
  const m = document.createElement("div");
  m.className = "msg user";
  m.innerHTML = `<div class="msg-label">Tú</div><div class="msg-bubble file-bubble"><div class="file-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg></div><div class="file-info"><div class="file-name"></div><div class="file-size"></div></div></div>`;
  m.querySelector(".file-name").textContent = name;
  m.querySelector(".file-size").textContent = formatBytes(size);
  c.appendChild(m);
  c.scrollTop = c.scrollHeight;
}

// Bot message con soporte Markdown
function addBotMessage(text) {
  const c = document.getElementById("chat");
  const m = document.createElement("div");
  m.className = "msg bot";
  const bubble = document.createElement("div");
  bubble.className = "msg-bubble";
  // renderMarkdown viene de voicebot-extensions.js (usa markdown-it si disponible)
  if (typeof renderMarkdown === "function") {
    bubble.innerHTML = renderMarkdown(text || "...");
  } else {
    bubble.textContent = text || "...";
  }
  m.innerHTML = `<div class="msg-label">Asistente</div>`;
  m.appendChild(bubble);
  c.appendChild(m);
  c.scrollTop = c.scrollHeight;
}
function showThinking() {
  if (document.getElementById("thinking")) return;
  const c = document.getElementById("chat");
  const t = document.createElement("div");
  t.className = "msg bot";
  t.id = "thinking";
  t.innerHTML = `<div class="msg-label">Asistente</div><div class="msg-bubble thinking-bubble"><div class="dot"></div><div class="dot"></div><div class="dot"></div></div>`;
  c.appendChild(t);
  c.scrollTop = c.scrollHeight;
}
function hideThinking() {
  const t = document.getElementById("thinking");
  if (t) t.remove();
}
function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}
init();