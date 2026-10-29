import { useEffect, useRef, useState } from "react";
import { getJson, imageData, toolPath, toolSend } from "../toolsApi";
import { useImagePaste } from "./translate/useImagePaste";
import "../tools.css";

interface SourceImage {
  mediaType: string;
  base64: string;
}
interface Settings {
  available: boolean;
  model: string;
  translateEffort: string;
  isCli: boolean;
}
interface Translation {
  sourceText: string;
  translation: string;
  usage: string;
}
const languages = [
  "Korean",
  "English",
  "Japanese",
  "Chinese (Simplified)",
  "Chinese (Traditional)",
  "Spanish",
  "French",
  "German",
  "Vietnamese",
];

export default function Translate() {
  const [text, setText] = useState("");
  const [target, setTarget] = useState("Korean");
  const [images, setImages] = useState<SourceImage[]>([]);
  const [settings, setSettings] = useState<Settings>();
  const [result, setResult] = useState<Translation>();
  const [resultLanguage, setResultLanguage] = useState("");
  const [hasResultImages, setHasResultImages] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");
  const request = useRef<AbortController | null>(null);
  const paste = useImagePaste(addPastedImage, setError);

  useEffect(() => {
    getJson<Settings>(toolPath("translate"))
      .then(setSettings)
      .catch((e: Error) => setError(e.message));
    return () => request.current?.abort();
  }, []);

  function addPastedImage(mediaType: string, base64: string) {
    setImages((current) => [...current, { mediaType, base64 }]);
  }

  async function addFiles(files: FileList | null) {
    try {
      for (const file of Array.from(files ?? [])) {
        if (file.type.startsWith("image/"))
          addPastedImage(file.type, await imageData(file));
      }
    } catch (e) {
      setError((e as Error).message);
    }
  }

  async function translate() {
    setBusy(true);
    setError("");
    setStatus("");
    setResult(undefined);
    request.current = new AbortController();
    try {
      const value = await toolSend<Translation>(
        "translate",
        { text, targetLanguage: target, images },
        "POST",
        request.current.signal,
      );
      setResult(value);
      setResultLanguage(target);
      setHasResultImages(images.length > 0);
    } catch (e) {
      setError(
        (e as Error).name === "AbortError"
          ? "Translation cancelled."
          : (e as Error).message,
      );
    } finally {
      setBusy(false);
      request.current = null;
    }
  }

  async function copyResult() {
    try {
      await navigator.clipboard.writeText(result?.translation ?? "");
      setStatus("Translation copied.");
    } catch (e) {
      setError((e as Error).message);
    }
  }

  return (
    <div className="tools-page">
      <header className="page-head">
        <h1>Translate</h1>
      </header>
      {error && (
        <div className="alert" role="alert">
          {error}
        </div>
      )}
      {settings && !settings.available && (
        <div className="alert">
          {settings.isCli
            ? "CLI is disabled. Enable CLI in Settings."
            : "Translation API is not configured. Save an API key in Settings."}
        </div>
      )}
      <div className="translate-layout">
        <section className="card">
          <h2>Source</h2>
          <label className="tool-field">
            Source text
            <textarea
              className="inp tool-textarea"
              rows={10}
              value={text}
              onChange={(e) => setText(e.target.value)}
              placeholder="Enter text to translate, paste an image, or both…"
            />
          </label>
          <div className="form-row">
            <button className="btn sm" onClick={paste.openCapture}>
              Capture / paste image
            </button>
            <button
              className="btn sm"
              onClick={paste.cancelCapture}
              disabled={!paste.capture}
            >
              Cancel capture
            </button>
            <label>
              Add images
              <input
                type="file"
                accept="image/*"
                multiple
                onChange={(e) => void addFiles(e.target.files)}
              />
            </label>
          </div>
          {paste.capture && (
            <p className="muted small">
              Take a screenshot with your system capture shortcut, then press
              Ctrl+V anywhere on this page. Escape cancels capture.
            </p>
          )}
          <div className="tool-thumbnails">
            {images.map((image, index) => (
              <figure
                key={index}
                title={`${image.mediaType} · approximately ${Math.round((image.base64.length * 0.75) / 1024)} KB`}
              >
                <img
                  src={`data:${image.mediaType};base64,${image.base64}`}
                  alt={`Source ${index + 1}`}
                />
                <button
                  className="btn sm"
                  onClick={() =>
                    setImages((current) =>
                      current.filter((_, i) => i !== index),
                    )
                  }
                >
                  Remove
                </button>
              </figure>
            ))}
          </div>
          {images.length > 0 && (
            <button className="btn sm" onClick={() => setImages([])}>
              Clear all ({images.length})
            </button>
          )}
          <div className="form-row">
            <label>
              Target language
              <select
                className="inp"
                value={target}
                onChange={(e) => setTarget(e.target.value)}
              >
                {languages.map((language) => (
                  <option key={language}>{language}</option>
                ))}
              </select>
            </label>
            <button
              className="btn sm primary"
              disabled={
                busy ||
                !settings?.available ||
                (!text.trim() && images.length === 0)
              }
              onClick={() => void translate()}
            >
              {busy
                ? "Translating…"
                : `Translate (${settings?.model ?? ""}${settings?.isCli ? " CLI" : ` · ${settings?.translateEffort ?? ""}`})`}
            </button>
          </div>
        </section>
        <section className="card translate-output" aria-busy={busy}>
          <h2>Translation Result {result && `— ${resultLanguage}`}</h2>
          {busy ? (
            <p className="muted">Translating…</p>
          ) : result ? (
            <>
              {hasResultImages && result.sourceText && (
                <>
                  <h3>AI OCR source text</h3>
                  <pre className="translate-source">{result.sourceText}</pre>
                </>
              )}
              <p className="tag ok">{result.usage}</p>
              <pre className="translate-result">{result.translation}</pre>
              <button className="btn sm" onClick={() => void copyResult()}>
                Copy
              </button>
            </>
          ) : (
            <p className="muted">Translation will appear here.</p>
          )}
          <p role="status">{status}</p>
        </section>
      </div>
    </div>
  );
}
