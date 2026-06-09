import { createFetchProxy } from "./fetch";
import { createXHRProxy } from "./xhr";
import type { EditorServer } from "./server";
import type { MockSocket, MockSocketOptions } from "./socket";

export function shouldBypassOnlyOfficeProxy(url: string, baseUrl: string) {
  const pathname = new URL(url, baseUrl).pathname;

  return (
    pathname.includes("/sdkjs/common/AllFonts.js") ||
    pathname.includes("/sdkjs/common/libfont/") ||
    pathname.includes("/fonts/")
  );
}

export type ScopedIoFactory = (
  url?: string,
  options?: MockSocketOptions,
) => MockSocket;

export type OnlyOfficeProxyWindow = Window & {
  __ONLYOFFICE_PROXIES_INSTALLED__?: boolean;
  XMLHttpRequest: typeof XMLHttpRequest;
  Worker: typeof Worker;
};

export function installOnlyOfficeProxies(
  win: OnlyOfficeProxyWindow,
  server: EditorServer,
  createIo: ScopedIoFactory,
) {
  if (win.__ONLYOFFICE_PROXIES_INSTALLED__) {
    return;
  }

  const xhr = createXHRProxy(win.XMLHttpRequest, {
    baseUrl: win.location.href,
    shouldBypass: (url) => shouldBypassOnlyOfficeProxy(url, win.location.href),
  });
  const fetchProxy = createFetchProxy(win);
  const WorkerCtor = win.Worker;

  xhr.use((request) => server.handleRequest(request));
  fetchProxy.use((request) => server.handleRequest(request));

  Object.assign(win, {
    io: createIo,
    XMLHttpRequest: xhr,
    fetch: fetchProxy,
    Worker: function Worker(url: string, options?: WorkerOptions) {
      const u = new URL(url, win.location.origin);
      return new WorkerCtor(
        u.href.replace(u.origin, win.location.origin),
        options,
      );
    },
  });
  win.__ONLYOFFICE_PROXIES_INSTALLED__ = true;
}

export const REPORTER_HTML = "index.reporter.html";

export type ReporterBridge = {
  install: (target: Window) => void;
};

export type ReporterHookWindow = Window & {
  open: typeof window.open;
  __ONLYOFFICE_REPORTER_HOOK__?: boolean;
  __ONLYOFFICE_REPORTER_BRIDGE__?: ReporterBridge;
};

export function installReporterWindowHook(
  win: ReporterHookWindow,
  installProxies: (target: Window) => void,
) {
  if (win.__ONLYOFFICE_REPORTER_HOOK__) {
    return;
  }

  win.__ONLYOFFICE_REPORTER_BRIDGE__ = { install: installProxies };
  win.__ONLYOFFICE_REPORTER_HOOK__ = true;

  const nativeOpen = win.open.bind(win);
  win.open = function openReporter(
    url?: string | URL,
    target?: string,
    features?: string,
  ) {
    const popup = nativeOpen(url, target, features);
    const href = typeof url === "string" ? url : url?.toString() ?? "";

    if (popup && href.includes(REPORTER_HTML)) {
      watchReporterWindow(popup, installProxies);
    }

    return popup;
  };
}

function watchReporterWindow(
  popup: Window,
  installProxies: (target: Window) => void,
) {
  const tryInstall = () => {
    if (popup.closed) {
      return true;
    }

    try {
      if (popup.location.href.includes(REPORTER_HTML)) {
        installProxies(popup);
        return true;
      }
    } catch {
      // Navigation in progress; keep polling.
    }

    return false;
  };

  if (tryInstall()) {
    return;
  }

  const interval = window.setInterval(() => {
    if (tryInstall()) {
      window.clearInterval(interval);
    }
  }, 1);

  popup.addEventListener(
    "load",
    () => {
      tryInstall();
      window.clearInterval(interval);
    },
    { once: true },
  );
}
