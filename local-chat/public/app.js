(function () {
  "use strict";

  var STORAGE_KEY = "local-nvidia-chat.settings.v1";
  var DEFAULTS = {
    model: "nvidia/nemotron-3.5-super-vl-120b-a12b",
    systemPrompt: "You are a helpful assistant running locally. Answer plainly and concisely.",
    temperature: 1,
    topP: 0.95,
    maxTokens: 4096,
  };

  var messagesEl = document.getElementById("messages");
  var emptyEl = document.getElementById("empty");
  var bannerEl = document.getElementById("banner");
  var formEl = document.getElementById("composer");
  var draftEl = document.getElementById("draft");
  var sendEl = document.getElementById("send");
  var stopEl = document.getElementById("stop");
  var fileEl = document.getElementById("file");
  var attachEl = document.getElementById("attach");
  var attachmentEl = document.getElementById("attachment");
  var attachmentImageEl = document.getElementById("attachment-image");
  var removeImageEl = document.getElementById("remove-image");
  var clearEl = document.getElementById("clear");
  var settingsEl = document.getElementById("settings");
  var settingsToggleEl = document.getElementById("settings-toggle");
  var modelBadgeEl = document.getElementById("model-badge");
  var resetEl = document.getElementById("reset-settings");
  var statusEl = document.getElementById("status");

  var fields = {
    model: document.getElementById("setting-model"),
    systemPrompt: document.getElementById("setting-system"),
    maxTokens: document.getElementById("setting-max-tokens"),
    temperature: document.getElementById("setting-temperature"),
    topP: document.getElementById("setting-top-p"),
  };

  var settings = loadSettings();
  var messages = [];
  var nodes = new Map();
  var pendingImage = "";
  var busy = false;
  var controller = null;
  var waitStarted = 0;
  var waitTimer = null;

  function uid() {
    return String(Date.now()) + "-" + Math.random().toString(36).slice(2);
  }

  function numberOr(value, fallback) {
    var parsed = Number(value);
    return isFinite(parsed) ? parsed : fallback;
  }

  function loadSettings() {
    try {
      var raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) return Object.assign({}, DEFAULTS);
      var parsed = JSON.parse(raw) || {};
      var storedModel = typeof parsed.model === "string" && parsed.model.trim() ? parsed.model : "";
      // The previous default never answers on this account; migrate it once so
      // an old saved setting does not keep the UI broken.
      if (!storedModel || /^deepseek-ai\/deepseek-v4\.1-flash$/i.test(storedModel)) {
        storedModel = DEFAULTS.model;
      }
      return {
        model: storedModel,
        systemPrompt: typeof parsed.systemPrompt === "string" ? parsed.systemPrompt : DEFAULTS.systemPrompt,
        temperature: numberOr(parsed.temperature, DEFAULTS.temperature),
        topP: numberOr(parsed.topP, DEFAULTS.topP),
        maxTokens: numberOr(parsed.maxTokens, DEFAULTS.maxTokens),
      };
    } catch (_error) {
      return Object.assign({}, DEFAULTS);
    }
  }

  function saveSettings() {
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
    } catch (_error) {
      /* storage can be blocked; settings stay in memory */
    }
  }

  function fillSettingsInputs() {
    fields.model.value = settings.model;
    fields.systemPrompt.value = settings.systemPrompt;
    fields.maxTokens.value = String(settings.maxTokens);
    fields.temperature.value = String(settings.temperature);
    fields.topP.value = String(settings.topP);
    modelBadgeEl.textContent = settings.model;
  }

  function readSettingsInputs() {
    settings.model = fields.model.value.trim() || DEFAULTS.model;
    settings.systemPrompt = fields.systemPrompt.value;
    settings.maxTokens = Math.max(1, Math.round(numberOr(fields.maxTokens.value, DEFAULTS.maxTokens)));
    settings.temperature = numberOr(fields.temperature.value, DEFAULTS.temperature);
    settings.topP = numberOr(fields.topP.value, DEFAULTS.topP);
    modelBadgeEl.textContent = settings.model;
    saveSettings();
  }

  function setBanner(text, kind) {
    if (!text) {
      bannerEl.hidden = true;
      bannerEl.textContent = "";
      return;
    }
    bannerEl.hidden = false;
    bannerEl.textContent = text;
    bannerEl.className = kind === "error" ? "banner" : "banner info";
  }

  function scrollToEnd() {
    messagesEl.scrollTop = messagesEl.scrollHeight;
  }

  function updateWaitStatus() {
    if (!busy) return;
    var seconds = Math.max(0, Math.round((Date.now() - waitStarted) / 1000));
    statusEl.textContent = "Request sent to NVIDIA through this server \u00b7 waiting " + seconds + "s";
  }

  function setBusy(next) {
    busy = next;
    sendEl.disabled = next;
    draftEl.disabled = next;
    attachEl.disabled = next;
    stopEl.hidden = !next;
    if (next) {
      waitStarted = Date.now();
      updateWaitStatus();
      waitTimer = window.setInterval(updateWaitStatus, 1000);
    } else {
      if (waitTimer) {
        window.clearInterval(waitTimer);
        waitTimer = null;
      }
      statusEl.textContent = "";
      draftEl.focus();
    }
  }

  function addMessageNode(message) {
    emptyEl.hidden = true;
    var root = document.createElement("article");
    root.className = "message message-" + message.role;

    var meta = document.createElement("div");
    meta.className = "message-meta";
    meta.textContent = message.role === "user" ? "You" : "Assistant";
    root.appendChild(meta);

    if (message.image) {
      var img = document.createElement("img");
      img.className = "message-image";
      img.src = message.image;
      img.alt = "Attached image";
      root.appendChild(img);
    }

    var text = document.createElement("div");
    text.className = "message-text";
    text.textContent = message.text || "";
    root.appendChild(text);

    messagesEl.appendChild(root);
    nodes.set(message.id, { root: root, text: text });
    scrollToEnd();
  }

  function updateMessageText(id, value) {
    var node = nodes.get(id);
    if (!node) return;
    node.text.textContent = value;
    scrollToEnd();
  }

  function updateMessageReasoning(id, value) {
    var node = nodes.get(id);
    if (!node) return;
    if (!value) {
      if (node.reasoningBox) node.reasoningBox.hidden = true;
      return;
    }
    if (!node.reasoningBox) {
      var details = document.createElement("details");
      details.className = "message-reasoning";
      details.open = true;
      var summary = document.createElement("summary");
      summary.textContent = "Reasoning";
      var body = document.createElement("div");
      body.className = "reasoning-text";
      details.appendChild(summary);
      details.appendChild(body);
      node.root.insertBefore(details, node.text);
      node.reasoningBox = details;
      node.reasoningText = body;
      node.reasoningSummary = summary;
    }
    node.reasoningBox.hidden = false;
    node.reasoningText.textContent = value;
    node.reasoningSummary.textContent = "Reasoning (" + value.length + " chars)";
    scrollToEnd();
  }

  function collapseReasoning(id) {
    var node = nodes.get(id);
    if (node && node.reasoningBox) node.reasoningBox.open = false;
  }

  function removeMessageNode(id) {
    var node = nodes.get(id);
    if (!node) return;
    node.root.remove();
    nodes.delete(id);
  }

  function toApiMessages(history, systemPrompt) {
    var result = [];
    if (systemPrompt.trim()) result.push({ role: "system", content: systemPrompt.trim() });
    history.forEach(function (message) {
      if (message.role === "user" && message.image) {
        result.push({
          role: "user",
          content: [
            { type: "text", text: message.text || "Describe this image." },
            { type: "image_url", image_url: { url: message.image } },
          ],
        });
      } else {
        result.push({ role: message.role, content: message.text });
      }
    });
    return result;
  }

  /**
   * Reasoning models (Nemotron, DeepSeek) stream chain-of-thought separately as
   * `delta.reasoning_content`, then the answer as `delta.content`.
   */
  function deltaFromEvent(data) {
    var result = { text: "", reasoning: "" };
    try {
      var parsed = JSON.parse(data);
      var choice = parsed && parsed.choices && parsed.choices[0];
      if (!choice) return result;
      if (choice.delta) {
        if (typeof choice.delta.content === "string") result.text = choice.delta.content;
        if (typeof choice.delta.reasoning_content === "string") result.reasoning = choice.delta.reasoning_content;
      }
      if (!result.text && choice.message && typeof choice.message.content === "string") result.text = choice.message.content;
      if (!result.reasoning && choice.message && typeof choice.message.reasoning_content === "string") {
        result.reasoning = choice.message.reasoning_content;
      }
      if (!result.text && typeof choice.text === "string") result.text = choice.text;
    } catch (_error) {
      /* keep-alive or partial chunk */
    }
    return result;
  }

  function textFromJson(json) {
    var choice = json && json.choices && json.choices[0];
    if (!choice) return "";
    if (choice.message && typeof choice.message.content === "string") return choice.message.content;
    if (typeof choice.text === "string") return choice.text;
    return "";
  }

  function readEventStream(response, onDelta) {
    if (!response.body) {
      return Promise.reject(new Error("The server returned no response body."));
    }
    var reader = response.body.getReader();
    var decoder = new TextDecoder();
    var buffer = "";
    var text = "";
    var reasoning = "";

    function consume(block) {
      block.split("\n").forEach(function (line) {
        var trimmed = line.trim();
        if (trimmed.indexOf("data:") !== 0) return;
        var data = trimmed.slice(5).trim();
        if (!data || data === "[DONE]") return;
        var delta = deltaFromEvent(data);
        if (delta.text || delta.reasoning) {
          text += delta.text;
          reasoning += delta.reasoning;
          onDelta(text, reasoning);
        }
      });
    }

    function step() {
      return reader.read().then(function (result) {
        if (result.done) {
          buffer += decoder.decode();
          if (buffer.trim()) consume(buffer);
          return text;
        }
        buffer += decoder.decode(result.value, { stream: true });
        var blocks = buffer.split("\n\n");
        buffer = blocks.pop() || "";
        blocks.forEach(consume);
        return step();
      });
    }

    return step();
  }

  function send() {
    var text = draftEl.value.trim();
    if (busy) return;
    if (!text && !pendingImage) return;

    var userMessage = { id: uid(), role: "user", text: text, image: pendingImage || undefined };
    messages.push(userMessage);
    addMessageNode(userMessage);

    var assistantMessage = { id: uid(), role: "assistant", text: "" };
    messages.push(assistantMessage);
    addMessageNode(assistantMessage);

    var history = messages.slice(0, -1);

    draftEl.value = "";
    pendingImage = "";
    attachmentEl.hidden = true;
    attachmentImageEl.removeAttribute("src");
    setBanner("");
    setBusy(true);

    controller = new AbortController();
    var received = "";
    var receivedReasoning = "";

    fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Accept": "text/event-stream" },
      body: JSON.stringify({
        model: settings.model,
        messages: toApiMessages(history, settings.systemPrompt),
        temperature: settings.temperature,
        top_p: settings.topP,
        max_tokens: settings.maxTokens,
        stream: true,
      }),
      signal: controller.signal,
    })
      .then(function (response) {
        if (!response.ok) {
          return response.text().then(function (detail) {
            var message = detail;
            try {
              var parsed = JSON.parse(detail);
              if (parsed && typeof parsed.error === "string") message = parsed.error;
              else if (parsed && parsed.error && typeof parsed.error.message === "string") message = parsed.error.message;
            } catch (_error) {
              /* not JSON: show the raw body */
            }
            throw new Error("HTTP " + response.status + (message ? ": " + message : ""));
          });
        }
        var type = response.headers.get("content-type") || "";
        if (type.indexOf("text/event-stream") >= 0) {
          return readEventStream(response, function (partial, reasoning) {
            if (!received && !partial) statusEl.textContent = "Model is reasoning\u2026";
            if (partial && !received) {
              statusEl.textContent = "Streaming reply\u2026";
              collapseReasoning(assistantMessage.id);
            }
            received = partial;
            receivedReasoning = reasoning;
            assistantMessage.text = partial;
            assistantMessage.reasoning = reasoning;
            if (partial) updateMessageText(assistantMessage.id, partial);
            if (reasoning) updateMessageReasoning(assistantMessage.id, reasoning);
          }).then(function (full) {
            if (!full && !receivedReasoning) updateMessageText(assistantMessage.id, "(empty response)");
          });
        }
        return response.json().then(function (json) {
          var content = textFromJson(json) || "(empty response)";
          received = content;
          assistantMessage.text = content;
          updateMessageText(assistantMessage.id, content);
        });
      })
      .catch(function (error) {
        var aborted = controller && controller.signal.aborted;
        if (aborted) {
          if (!received && !receivedReasoning) updateMessageText(assistantMessage.id, "(stopped)");
          return;
        }
        var detail = error instanceof Error ? error.message : String(error);
        setBanner(detail === "Failed to fetch"
          ? "Could not reach the local server. Is it still running?"
          : detail, "error");
        if (!received && !receivedReasoning) {
          removeMessageNode(assistantMessage.id);
          messages = messages.filter(function (message) { return message.id !== assistantMessage.id; });
        }
      })
      .then(function () {
        controller = null;
        setBusy(false);
      });
  }

  function clearChat() {
    if (controller) controller.abort();
    messages = [];
    nodes.forEach(function (node) { node.root.remove(); });
    nodes.clear();
    emptyEl.hidden = false;
    setBanner("");
  }

  function chooseImage(event) {
    var file = event.target.files && event.target.files[0];
    event.target.value = "";
    if (!file) return;
    if (file.type.indexOf("image/") !== 0) {
      setBanner("Choose an image file.", "error");
      return;
    }
    if (file.size > 5 * 1024 * 1024) {
      setBanner("Choose an image under 5 MB.", "error");
      return;
    }
    var reader = new FileReader();
    reader.onload = function () {
      pendingImage = String(reader.result || "");
      attachmentImageEl.src = pendingImage;
      attachmentEl.hidden = false;
    };
    reader.onerror = function () { setBanner("Could not read that image.", "error"); };
    reader.readAsDataURL(file);
  }

  formEl.addEventListener("submit", function (event) {
    event.preventDefault();
    send();
  });

  stopEl.addEventListener("click", function () {
    if (controller) controller.abort();
  });

  clearEl.addEventListener("click", clearChat);

  attachEl.addEventListener("click", function () { fileEl.click(); });
  fileEl.addEventListener("change", chooseImage);
  removeImageEl.addEventListener("click", function () {
    pendingImage = "";
    attachmentEl.hidden = true;
    attachmentImageEl.removeAttribute("src");
  });

  draftEl.addEventListener("keydown", function (event) {
    if (event.key !== "Enter" || event.shiftKey || event.isComposing) return;
    event.preventDefault();
    if (typeof formEl.requestSubmit === "function") formEl.requestSubmit();
    else send();
  });

  settingsToggleEl.addEventListener("click", function () {
    var open = settingsEl.hidden;
    settingsEl.hidden = !open;
    settingsToggleEl.setAttribute("aria-expanded", open ? "true" : "false");
    settingsToggleEl.textContent = open ? "Hide settings" : "Settings";
  });

  Object.keys(fields).forEach(function (key) {
    fields[key].addEventListener("input", readSettingsInputs);
  });

  resetEl.addEventListener("click", function () {
    settings = Object.assign({}, DEFAULTS);
    saveSettings();
    fillSettingsInputs();
  });

  fillSettingsInputs();

  fetch("/api/health")
    .then(function (response) { return response.json(); })
    .then(function (data) {
      if (!data.keyConfigured) {
        setBanner("NVIDIA_API_KEY is missing in local-chat/.env. Add it and restart the server.", "error");
      }
    })
    .catch(function () {
      setBanner("Could not reach the local server.", "error");
    });
})();
