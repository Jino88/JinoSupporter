import { useEffect, useRef, useState } from "react";
import { loadLegacyTools } from "../../toolsApi";

interface PasteHandler {
  init(ref: {
    invokeMethodAsync: (
      method: string,
      mediaType: string,
      dataUrl: string,
    ) => Promise<void>;
  }): void;
  openCapture(): void;
  cancelCapture(): void;
  cancelPasteCapture(): void;
}

/** Bridge the original white-background image capture handler to React, with route cleanup. */
export function useImagePaste(
  onImage: (mediaType: string, base64: string) => void,
  onError: (message: string) => void,
) {
  const receiver = useRef(onImage);
  receiver.current = onImage;
  const [capture, setCapture] = useState(true);
  const handler = useRef<PasteHandler | undefined>(undefined);

  useEffect(() => {
    let disposed = false;
    loadLegacyTools()
      .then(() => {
        if (disposed) return;
        const paste = (window as unknown as { pasteImageHandler: PasteHandler })
          .pasteImageHandler;
        handler.current = paste;
        paste.init({
          invokeMethodAsync: async (_method, mediaType, dataUrl) => {
            if (disposed) return;
            receiver.current(
              mediaType,
              dataUrl.slice(dataUrl.indexOf(",") + 1),
            );
            paste.openCapture();
          },
        });
        paste.openCapture();
      })
      .catch((error: Error) => {
        if (!disposed) onError(error.message);
      });
    return () => {
      disposed = true;
      handler.current?.cancelPasteCapture();
      handler.current = undefined;
    };
  }, []);

  function openCapture() {
    handler.current?.openCapture();
    setCapture(true);
  }
  function cancelCapture() {
    handler.current?.cancelCapture();
    setCapture(false);
  }
  useEffect(() => {
    function onKey(event: KeyboardEvent) {
      if (event.key === "Escape") cancelCapture();
    }
    document.addEventListener("keydown", onKey);
    return () => document.removeEventListener("keydown", onKey);
  }, []);
  return { capture, openCapture, cancelCapture };
}
