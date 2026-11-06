import * as React from "react";
import {
  DefaultButton,
  IconButton,
  Label,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Stack,
  TextField,
} from "@fluentui/react";

/**
 * Browser chat client for an OpenAI-compatible NVIDIA endpoint.
 *
 * The browser never holds the API key. SharePoint is a static, client-side
 * host, so the request goes through a small proxy (see tools/chat-proxy) that
 * adds the key server-side and returns CORS headers. The proxy URL is stored
 * in this browser only.
 */

const STORAGE_KEY = "sharepoint-api-chatbot.v1";
const DEFAULT_MODEL = "nvidia/nemotron-3.5-super-vl-120b-a12b";
const DEFAULT_SYSTEM_PROMPT = "You are a helpful assistant embedded in a SharePoint page. Answer plainly and concisely.";

interface ChatSettings {
  proxyUrl: string;
  model: string;
  systemPrompt: string;
  temperature: number;
  topP: number;
  maxTokens: number;
}

interface ChatMessage {
  id: string;
  role: "user" | "assistant";
  text: string;
  image?: string;
}

const DEFAULT_SETTINGS: ChatSettings = {
  proxyUrl: "",
  model: DEFAULT_MODEL,
  systemPrompt: DEFAULT_SYSTEM_PROMPT,
  temperature: 1,
  topP: 0.95,
  maxTokens: 4096,
};

function loadSettings(): ChatSettings {
  try {
    if (typeof window === "undefined" || !window.localStorage) return DEFAULT_SETTINGS;
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return DEFAULT_SETTINGS;
    const parsed = JSON.parse(raw) as Partial<ChatSettings>;
    const savedModel = typeof parsed.model === "string" && parsed.model.trim() ? parsed.model : DEFAULT_MODEL;
    // The old default never answers on this account; migrate it once.
    const migratedModel = /^deepseek-ai\/deepseek-v4\.1-flash$/i.test(savedModel) ? DEFAULT_MODEL : savedModel;
    return {
      proxyUrl: typeof parsed.proxyUrl === "string" ? parsed.proxyUrl : "",
      model: migratedModel,
      systemPrompt: typeof parsed.systemPrompt === "string" ? parsed.systemPrompt : DEFAULT_SYSTEM_PROMPT,
      temperature: typeof parsed.temperature === "number" ? parsed.temperature : 1,
      topP: typeof parsed.topP === "number" ? parsed.topP : 0.95,
      maxTokens: typeof parsed.maxTokens === "number" ? parsed.maxTokens : 4096,
    };
  } catch {
    return DEFAULT_SETTINGS;
  }
}

/** Accepts a base URL, a /v1 URL or the full completions URL. */
function chatEndpoint(proxyUrl: string): string {
  const base = proxyUrl.trim().replace(/\/+$/, "");
  if (!base) return "";
  if (/\/chat\/completions$/i.test(base)) return base;
  if (/\/v1$/i.test(base)) return base + "/chat/completions";
  return base + "/v1/chat/completions";
}

function proxyHost(proxyUrl: string): string {
  try {
    return new URL(proxyUrl).host;
  } catch {
    return proxyUrl;
  }
}

function newId(): string {
  return String(Date.now()) + "-" + Math.random().toString(36).slice(2);
}

/** Mirrors the sample: text turns as strings, image turns as content parts. */
function toApiMessages(messages: ChatMessage[], systemPrompt: string): unknown[] {
  const result: unknown[] = [];
  if (systemPrompt.trim()) result.push({ role: "system", content: systemPrompt.trim() });
  messages.forEach((message) => {
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

function deltaFromEvent(data: string): string {
  try {
    const parsed = JSON.parse(data) as {
      choices?: Array<{ delta?: { content?: string }; message?: { content?: string }; text?: string }>;
    };
    const choice = parsed.choices && parsed.choices[0];
    if (!choice) return "";
    if (choice.delta && typeof choice.delta.content === "string") return choice.delta.content;
    if (choice.message && typeof choice.message.content === "string") return choice.message.content;
    if (typeof choice.text === "string") return choice.text;
  } catch {
    /* keep-alive or partial chunk */
  }
  return "";
}

async function readEventStream(
  response: Response,
  onText: (fullText: string) => void
): Promise<string> {
  if (!response.body) throw new Error("The proxy returned no response body.");
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let full = "";
  const consume = (block: string): void => {
    const lines = block.split("\n");
    lines.forEach((line) => {
      const trimmed = line.trim();
      if (trimmed.indexOf("data:") !== 0) return;
      const data = trimmed.slice(5).trim();
      if (!data || data === "[DONE]") return;
      const delta = deltaFromEvent(data);
      if (delta) {
        full += delta;
        onText(full);
      }
    });
  };
  for (;;) {
    const step = await reader.read();
    if (step.done) break;
    buffer += decoder.decode(step.value, { stream: true });
    const blocks = buffer.split("\n\n");
    buffer = blocks.pop() || "";
    blocks.forEach(consume);
  }
  buffer += decoder.decode();
  if (buffer.trim()) consume(buffer);
  return full;
}

interface ChatCompletionChoice {
  message?: { content?: string };
  text?: string;
}

function messageFromJson(json: unknown): string {
  const parsed = (json || {}) as { choices?: ChatCompletionChoice[] };
  const choice = parsed.choices && parsed.choices[0];
  if (!choice) return "";
  if (choice.message && typeof choice.message.content === "string") return choice.message.content;
  if (typeof choice.text === "string") return choice.text;
  return "";
}

export default function ChatbotPanel(): React.ReactElement {
  const [settings, setSettings] = React.useState<ChatSettings>(loadSettings);
  const [settingsOpen, setSettingsOpen] = React.useState<boolean>(false);
  const [messages, setMessages] = React.useState<ChatMessage[]>([]);
  const [draft, setDraft] = React.useState<string>("");
  const [image, setImage] = React.useState<string>("");
  const [busy, setBusy] = React.useState<boolean>(false);
  const [notice, setNotice] = React.useState<string>("");
  const [noticeType, setNoticeType] = React.useState<MessageBarType>(MessageBarType.info);
  const abortRef = React.useRef<AbortController | null>(null);
  const fileRef = React.useRef<HTMLInputElement | null>(null);
  const endRef = React.useRef<HTMLDivElement | null>(null);

  const endpoint = chatEndpoint(settings.proxyUrl);

  React.useEffect(() => {
    try {
      if (window.localStorage) window.localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
    } catch {
      /* private mode or blocked storage: settings stay in memory */
    }
  }, [settings]);

  React.useEffect(() => {
    if (endRef.current) endRef.current.scrollIntoView({ block: "nearest" });
  }, [messages, busy]);

  const setSetting = (key: keyof ChatSettings, value: string | number): void => {
    setSettings((current) => Object.assign({}, current, { [key]: value }));
  };

  const showNotice = (text: string, type: MessageBarType): void => {
    setNotice(text);
    setNoticeType(type);
  };

  const appendAssistantDelta = (id: string, text: string): void => {
    setMessages((current) => current.map((message) => message.id === id
      ? Object.assign({}, message, { text })
      : message));
  };

  const stop = (): void => {
    if (abortRef.current) abortRef.current.abort();
  };

  const send = async (): Promise<void> => {
    const text = draft.trim();
    if (busy) return;
    if (!text && !image) return;
    if (!endpoint) {
      showNotice("Add the proxy URL in Settings first. A SharePoint page cannot call NVIDIA directly.", MessageBarType.warning);
      return;
    }
    const userMessage: ChatMessage = { id: newId(), role: "user", text, image: image || undefined };
    const history = messages.concat(userMessage);
    const assistantId = newId();
    setMessages(history.concat({ id: assistantId, role: "assistant", text: "" }));
    setDraft("");
    setImage("");
    setNotice("");
    setBusy(true);
    const controller = new AbortController();
    abortRef.current = controller;
    try {
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "text/event-stream" },
        body: JSON.stringify({
          model: settings.model,
          messages: toApiMessages(history, settings.systemPrompt),
          temperature: settings.temperature,
          top_p: settings.topP,
          max_tokens: settings.maxTokens,
          stream: true,
        }),
        signal: controller.signal,
      });
      if (!response.ok) {
        const detail = (await response.text()).slice(0, 800);
        throw new Error("Proxy returned HTTP " + response.status + " " + response.statusText + (detail ? ": " + detail : ""));
      }
      const contentType = response.headers.get("content-type") || "";
      if (contentType.indexOf("text/event-stream") >= 0) {
        const full = await readEventStream(response, (partial) => appendAssistantDelta(assistantId, partial));
        if (!full) appendAssistantDelta(assistantId, "(empty response)");
      } else {
        const json = await response.json();
        appendAssistantDelta(assistantId, messageFromJson(json) || "(empty response)");
      }
    } catch (error) {
      if (controller.signal.aborted) {
        setMessages((current) => current.map((message) => message.id === assistantId && !message.text
          ? Object.assign({}, message, { text: "(stopped)" })
          : message).filter((message) => message.id !== assistantId || !!message.text));
      } else {
        const detail = error instanceof Error ? error.message : String(error);
        const friendly = detail === "Failed to fetch"
          ? "Could not reach the proxy. Check the URL and confirm it allows CORS from " + window.location.origin + "."
          : detail;
        showNotice(friendly, MessageBarType.error);
        setMessages((current) => current.filter((message) => message.id !== assistantId || !!message.text));
      }
    } finally {
      abortRef.current = null;
      setBusy(false);
    }
  };

  const chooseImage = (event: React.ChangeEvent<HTMLInputElement>): void => {    const file = event.target.files && event.target.files[0];
    event.target.value = "";
    if (!file) return;
    if (file.size > 5 * 1024 * 1024) {
      showNotice("Choose an image smaller than 5 MB.", MessageBarType.error);
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === "string") setImage(reader.result);
    };
    reader.onerror = () => showNotice("Could not read that image.", MessageBarType.error);
    reader.readAsDataURL(file);
  };

  const clearChat = (): void => {
    stop();
    setMessages([]);
    setDraft("");
    setImage("");
    setNotice("");
  };

  const sendFromUi = (): void => {
    send().catch((error) => showNotice(error instanceof Error ? error.message : String(error), MessageBarType.error));
  };

  const onComposerKeyDown = (event: React.KeyboardEvent<HTMLDivElement>): void => {
    if (event.key !== "Enter" || event.shiftKey) return;
    event.preventDefault();
    sendFromUi();
  };

  return (
    <section className="chatbotPanel" aria-label="NVIDIA chatbot">
      <Stack horizontal wrap verticalAlign="center" tokens={{ childrenGap: 8 }}>
        <Label className="chatbotPanel__title">NVIDIA chatbot</Label>
        <span className="chatbotPanel__meta">{endpoint ? settings.model + " · via " + proxyHost(settings.proxyUrl) : "proxy not configured"}</span>
        <DefaultButton text={settingsOpen ? "Hide settings" : "Settings"} onClick={() => setSettingsOpen((open) => !open)} />
        <DefaultButton text="Clear chat" onClick={clearChat} disabled={busy || messages.length === 0} />
      </Stack>

      {settingsOpen && <Stack tokens={{ childrenGap: 8 }} className="chatbotPanel__settings">
        <TextField
          label="Proxy URL"
          placeholder="https://your-proxy.example.workers.dev"
          description="Required. The proxy adds the NVIDIA key server-side and sends CORS headers. See tools/chat-proxy."
          value={settings.proxyUrl}
          onChange={(_event, value) => setSetting("proxyUrl", value || "")}
        />
        <TextField
          label="Model"
          value={settings.model}
          onChange={(_event, value) => setSetting("model", value || DEFAULT_MODEL)}
        />
        <TextField
          label="System prompt"
          multiline
          autoAdjustHeight
          value={settings.systemPrompt}
          onChange={(_event, value) => setSetting("systemPrompt", value || "")}
        />
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <TextField
            label="Temperature"
            type="number"
            value={String(settings.temperature)}
            onChange={(_event, value) => {
              const parsed = Number(value);
              if (!isNaN(parsed)) setSetting("temperature", parsed);
            }}
          />
          <TextField
            label="Top P"
            type="number"
            value={String(settings.topP)}
            onChange={(_event, value) => {
              const parsed = Number(value);
              if (!isNaN(parsed)) setSetting("topP", parsed);
            }}
          />
          <TextField
            label="Max tokens"
            type="number"
            value={String(settings.maxTokens)}
            onChange={(_event, value) => {
              const parsed = Number(value);
              if (!isNaN(parsed)) setSetting("maxTokens", parsed);
            }}
          />
        </Stack>
        <span className="chatbotPanel__meta">Saved in this browser only. The API key is never stored here.</span>
      </Stack>}

      {!endpoint && <MessageBar messageBarType={MessageBarType.warning}>
        Set a proxy URL in Settings to start. NVIDIA blocks browser requests (no CORS headers) and an API key in SharePoint code would be public.
      </MessageBar>}

      {notice && <MessageBar messageBarType={noticeType}>{notice}</MessageBar>}

      <div className="chatbotPanel__messages" role="log" aria-live="polite">
        {messages.length === 0 && <p className="chatbotPanel__empty">Ask a question, or attach an image and ask about it.</p>}
        {messages.map((message) => (
          <div key={message.id} className={"chatbotPanel__message chatbotPanel__message--" + message.role}>
            <span className="chatbotPanel__role">{message.role === "user" ? "You" : "Assistant"}</span>
            {message.image && <img className="chatbotPanel__image" src={message.image} alt="Attached" />}
            <div className="chatbotPanel__text">{message.text}{message.role === "assistant" && busy && !message.text && <Spinner size={SpinnerSize.small} label="Thinking" />}</div>
          </div>
        ))}
        <div ref={endRef} />
      </div>

      <div className="chatbotPanel__composer">
        {image && <div className="chatbotPanel__attachment">
          <img className="chatbotPanel__image" src={image} alt="Attached" />
          <IconButton
            iconProps={{ iconName: "Cancel" }}
            title="Remove image"
            ariaLabel="Remove image"
            onClick={() => setImage("")}
            disabled={busy}
          />
        </div>}
        <div onKeyDown={onComposerKeyDown}>
          <TextField
            multiline
            autoAdjustHeight
            rows={2}
            placeholder="Ask the model. Enter sends, Shift+Enter adds a line."
            value={draft}
            onChange={(_event, value) => setDraft(value || "")}
            disabled={busy}
          />
        </div>
        <Stack horizontal wrap verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <DefaultButton
            text="Attach image"
            onClick={() => { if (fileRef.current) fileRef.current.click(); }}
            disabled={busy}
          />
          <input ref={fileRef} type="file" accept="image/*" className="chatbotPanel__file" onChange={chooseImage} />
          {busy
            ? <DefaultButton text="Stop" onClick={stop} />
            : <PrimaryButton text="Send" onClick={sendFromUi} disabled={!draft.trim() && !image} />}
          {busy && <Spinner size={SpinnerSize.small} label="Waiting for the model" />}
        </Stack>
      </div>
    </section>
  );
}
