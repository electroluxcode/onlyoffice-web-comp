import { STATIC_RESOURCE } from "../const";

let initializePromise: Promise<void> | null = null;

function preloadEditorFrame() {
  if (document.querySelector(`iframe[data-onlyoffice-preload="${STATIC_RESOURCE.onlyoffice.preloadHtml}"]`)) {
    return;
  }

  const iframe = document.createElement("iframe");
  iframe.src = STATIC_RESOURCE.onlyoffice.preloadUrl;
  iframe.dataset.onlyofficePreload = STATIC_RESOURCE.onlyoffice.preloadHtml;
  iframe.className = "w-0 h-0 hidden absolute -z-10";
  document.body.appendChild(iframe);
}

export async function initializeOnlyOffice() {
  if (typeof window === "undefined") return;

  if (initializePromise) {
    return initializePromise;
  }

  initializePromise = new Promise<void>((resolve, reject) => {
    preloadEditorFrame();

    if (window.DocsAPI?.DocEditor) {
      resolve();
      return;
    }

    const apiUrl = STATIC_RESOURCE.onlyoffice.apiUrl;
    let script = document.querySelector<HTMLScriptElement>(
      `script[src="${apiUrl}"]`,
    );

    if (!script) {
      script = document.createElement("script");
      script.src = apiUrl;
      document.head.appendChild(script);
    }

    script.addEventListener("load", () => resolve(), { once: true });
    script.addEventListener(
      "error",
      () => {
        initializePromise = null;
        reject(new Error("Failed to load OnlyOffice DocsAPI script"));
      },
      { once: true },
    );
  });

  return initializePromise;
}
