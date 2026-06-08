"use client";

import { memo, useEffect, useRef, useState } from "react";
import {
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_ID,
  ONLYOFFICE_LANG_KEY,
  OnlyOfficeManager,
  editorManagerFactory,
  type FileType,
} from "@/components/onlyoffice-web-comp";

type OfficePreviewPageProps = {
  title: string;
  badge: string;
  badgeClassName: string;
  defaultFileName: string;
  fileType: FileType;
  accept: string;
  newButtonLabel: string;
  /** public 目录下的默认文件路径，如 /test.xlsx */
  initialFileUrl?: string;
};

function LoadingOverlay() {
  return (
    <div className="absolute inset-0 z-20 flex items-center justify-center bg-white/70 backdrop-blur-sm">
      <div className="rounded-lg border border-gray-200 bg-white px-4 py-3 text-sm text-gray-700 shadow">
        加载中...
      </div>
    </div>
  );
}

const OnlyOfficeHost = memo(function OnlyOfficeHost() {
  return (
    <div
      className={`${ONLYOFFICE_CONTAINER_CONFIG.PARENT_CLASS_NAME} absolute inset-0`}
    >
      <div id={ONLYOFFICE_ID} className="absolute inset-0" />
    </div>
  );
});

async function fetchPublicFile(url: string, fileName: string) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}`);
  }
  const blob = await response.blob();
  return new File([blob], fileName, { type: blob.type });
}

export function OfficePreviewPage({
  title,
  badge,
  badgeClassName,
  defaultFileName,
  fileType,
  accept,
  newButtonLabel,
  initialFileUrl,
}: OfficePreviewPageProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const managerRef = useRef<OnlyOfficeManager | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [readOnly, setReadOnly] = useState(false);
  const [currentLang, setCurrentLangState] = useState(
    ONLYOFFICE_LANG_KEY.ZH as string,
  );
  const [editorReady, setEditorReady] = useState(false);

  useEffect(() => {
    let unsubscribeLoading: (() => void) | undefined;
    let disposed = false;
    let ownedManager: OnlyOfficeManager | null = null;
    const containerId = ONLYOFFICE_ID;

    const init = async () => {
      editorManagerFactory.destroy(containerId);
      const loadSession = editorManagerFactory.beginLoadSession(containerId);

      let manager: OnlyOfficeManager;

      if (initialFileUrl) {
        const file = await fetchPublicFile(initialFileUrl, defaultFileName);
        if (disposed) return;

        manager = await OnlyOfficeManager.createWithFile(
          {
            containerId,
            fileType,
            defaultFileName,
            readOnly,
            loadSession,
          },
          file,
        );
      } else {
        manager = await OnlyOfficeManager.create({
          containerId,
          fileType,
          defaultFileName,
          readOnly,
          loadSession,
        });
      }

      if (
        disposed ||
        !editorManagerFactory.isLoadSessionActive(containerId, loadSession)
      ) {
        return;
      }

      ownedManager = manager;
      managerRef.current = manager;
      setCurrentLangState(manager.getLanguage());
      setEditorReady(true);
      unsubscribeLoading = manager.onLoadingChange(({ loading: next }) => {
        setLoading(next);
      });
    };

    init().catch((err) => {
      if (disposed) return;
      setError("无法加载编辑器组件");
      console.error("Failed to initialize OnlyOffice:", err);
    });

    return () => {
      disposed = true;
      unsubscribeLoading?.();
      ownedManager?.destroy();
      editorManagerFactory.destroy(containerId);
      managerRef.current = null;
      setEditorReady(false);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const runAction = async (action: () => Promise<void>, message: string) => {
    try {
      setError(null);
      await action();
    } catch (err) {
      setError(message);
      console.error(message, err);
    }
  };

  const handleOpenDocument = (
    fileName: string,
    file?: File,
    nextReadOnly = readOnly,
  ) =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }
      await manager.openDocument({ fileName, file, readOnly: nextReadOnly });
      setReadOnly(nextReadOnly);
    }, "操作失败");

  const handleLanguageSwitch = () =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }
      const nextLang = await manager.toggleLanguage();
      setCurrentLangState(nextLang);
    }, "切换语言失败");

  const handleExport = () =>
    runAction(async () => {
      await managerRef.current?.downloadExport();
    }, "导出失败");

  const handleToggleReadOnly = () =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }
      manager.toggleReadOnly();
      setReadOnly(manager.getReadOnly());
    }, "切换模式失败");

  return (
    <div className="flex h-screen flex-col bg-white">
      <div className="border-b border-gray-200 bg-gradient-to-r from-white to-gray-50 shadow-sm">
        <div className="mx-auto flex max-w-7xl flex-wrap items-center gap-4 px-5 py-4">
          <div className="mr-auto flex items-center gap-3">
            <div
              className={`flex h-8 w-8 items-center justify-center rounded-lg font-bold text-white ${badgeClassName}`}
            >
              {badge}
            </div>
            <h1 className="text-lg font-semibold text-gray-900">{title}</h1>
          </div>

          <div className="flex flex-wrap gap-3">
            <button
              type="button"
              onClick={handleLanguageSwitch}
              className="rounded-md bg-gray-100 px-3 py-2 text-sm font-medium text-gray-700 transition-colors hover:bg-gray-200"
            >
              {currentLang === ONLYOFFICE_LANG_KEY.ZH ? "点击切换 EN" : "点击切换中文"}
            </button>
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="rounded-md bg-blue-500 px-4 py-2 text-white transition-colors hover:bg-blue-600"
            >
              上传文档
            </button>
            <button
              type="button"
              onClick={() => handleOpenDocument(defaultFileName)}
              className="rounded-md border border-gray-300 bg-white px-4 py-2 transition-colors hover:bg-gray-50"
            >
              {newButtonLabel}
            </button>
            {editorReady && (
              <>
                <button
                  type="button"
                  onClick={handleExport}
                  className="rounded-md border border-gray-300 bg-white px-4 py-2 transition-colors hover:bg-gray-50"
                >
                  导出
                </button>
                <button
                  type="button"
                  onClick={handleToggleReadOnly}
                  className={`rounded-md px-4 py-2 transition-colors ${
                    readOnly
                      ? "bg-yellow-500 text-white hover:bg-yellow-600"
                      : "border border-gray-300 bg-white hover:bg-gray-50"
                  }`}
                >
                  {readOnly ? "只读模式" : "编辑模式"}
                </button>
              </>
            )}
          </div>
        </div>
      </div>

      {error && (
        <div className="mx-4 mt-4 rounded border-l-4 border-red-500 bg-red-50 p-4 text-red-700">
          <p className="font-medium">错误：{error}</p>
        </div>
      )}

      <div className="relative flex-1">
        <OnlyOfficeHost />
        {loading && <LoadingOverlay />}
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept={accept}
        className="hidden"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) {
            handleOpenDocument(file.name, file);
            if (fileInputRef.current) fileInputRef.current.value = "";
          }
        }}
      />
    </div>
  );
}
