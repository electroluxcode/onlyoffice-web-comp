"use client";

import { memo, useEffect, useRef, useState } from "react";
import {
  FILE_TYPE,
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_EVENT_KEYS,
  OnlyOfficeManager,
  onlyOfficeManagerFactory,
  onlyofficeEventbus,
  type FileType,
} from "@/components/onlyoffice-web-comp";

type PanelKey = "word" | "excel" | "ppt";

type PanelConfig = {
  key: PanelKey;
  title: string;
  badge: string;
  color: string;
  containerId: string;
  defaultFileName: string;
  fileType: FileType;
  accept: string;
};

const PANELS: PanelConfig[] = [
  {
    key: "word",
    title: "Word",
    badge: "W",
    color: "from-blue-500 to-blue-700",
    containerId: "multi-word-editor",
    defaultFileName: "New_Document.docx",
    fileType: FILE_TYPE.DOCX,
    accept: ".docx,.doc,.odt,.rtf,.txt",
  },
  {
    key: "excel",
    title: "Excel",
    badge: "E",
    color: "from-green-500 to-green-700",
    containerId: "multi-excel-editor",
    defaultFileName: "New_Spreadsheet.xlsx",
    fileType: FILE_TYPE.XLSX,
    accept: ".xlsx,.xls,.ods,.csv",
  },
  {
    key: "ppt",
    title: "PowerPoint",
    badge: "P",
    color: "from-orange-500 to-orange-700",
    containerId: "multi-ppt-editor",
    defaultFileName: "New_Presentation.pptx",
    fileType: FILE_TYPE.PPTX,
    accept: ".pptx,.ppt,.odp",
  },
];

function getPanel(key: PanelKey) {
  const panel = PANELS.find((item) => item.key === key);
  if (!panel) throw new Error(`Unknown panel ${key}`);
  return panel;
}

const OnlyOfficePanelHost = memo(function OnlyOfficePanelHost({
  containerId,
}: {
  containerId: string;
}) {
  return (
    <div
      className={`${ONLYOFFICE_CONTAINER_CONFIG.PARENT_CLASS_NAME} absolute inset-0`}
      data-onlyoffice-container-id={containerId}
    >
      <div id={containerId} className="absolute inset-0" />
    </div>
  );
});

export function ProgressiveMultiPage() {
  const initializedRef = useRef(false);
  const inputRefs = {
    word: useRef<HTMLInputElement>(null),
    excel: useRef<HTMLInputElement>(null),
    ppt: useRef<HTMLInputElement>(null),
  };
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [managers, setManagers] = useState<
    Partial<Record<PanelKey, OnlyOfficeManager>>
  >({});
  const [documents, setDocuments] = useState<
    Record<PanelKey, { fileName: string; file?: File }>
  >({
    word: { fileName: getPanel("word").defaultFileName },
    excel: { fileName: getPanel("excel").defaultFileName },
    ppt: { fileName: getPanel("ppt").defaultFileName },
  });
  const [readOnlyStates, setReadOnlyStates] = useState<
    Record<PanelKey, boolean>
  >({
    word: false,
    excel: false,
    ppt: false,
  });

  const openPanel = async (
    key: PanelKey,
    nextDocument = documents[key],
    nextReadOnly = readOnlyStates[key],
  ) => {
    const panel = getPanel(key);
    setError(null);

    try {
      const manager = await onlyOfficeManagerFactory.open(
        {
          containerId: panel.containerId,
          fileType: panel.fileType,
          defaultFileName: panel.defaultFileName,
          readOnly: nextReadOnly,
        },
        {
          fileName: nextDocument.fileName,
          file: nextDocument.file,
          isNew: !nextDocument.file,
          readOnly: nextReadOnly,
        },
      );

      setManagers((prev) => ({
        ...prev,
        [key]: manager,
      }));
    } catch (err) {
      setError(err instanceof Error ? err.message : "打开编辑器失败");
      console.error("Failed to open editor:", err);
    }
  };

  const uploadPanelFile = async (key: PanelKey, file: File) => {
    const document = { fileName: file.name, file };
    setDocuments((prev) => ({
      ...prev,
      [key]: document,
    }));
    await openPanel(key, document);
  };

  const newPanelDocument = async (key: PanelKey) => {
    const panel = getPanel(key);
    const document = { fileName: panel.defaultFileName, file: undefined };
    setDocuments((prev) => ({
      ...prev,
      [key]: document,
    }));
    await openPanel(key, document);
  };

  const toggleReadOnly = async (key: PanelKey) => {
    const nextReadOnly = !readOnlyStates[key];
    setReadOnlyStates((prev) => ({
      ...prev,
      [key]: nextReadOnly,
    }));

    try {
      const manager = managers[key];
      if (manager?.isReady()) {
        manager.setReadOnly(nextReadOnly);
      } else {
        await openPanel(key, documents[key], nextReadOnly);
      }
    } catch (err) {
      setError("切换模式失败");
      console.error("Failed to toggle read-only:", err);
    }
  };

  const exportPanel = async (key: PanelKey) => {
    const manager = managers[key];

    if (!manager) {
      setError("编辑器未初始化");
      return;
    }

    try {
      await manager.downloadExport();
    } catch (err) {
      setError("导出失败");
      console.error("Export failed:", err);
    }
  };

  useEffect(() => {
    const init = async () => {
      if (initializedRef.current) return;
      initializedRef.current = true;

      for (const panel of PANELS) {
        await openPanel(
          panel.key,
          documents[panel.key],
          readOnlyStates[panel.key],
        );
      }
    };

    init().catch((err) => {
      setError("无法加载编辑器组件");
      console.error("Failed to initialize multi editors:", err);
    });

    const handleLoadingChange = (data: { loading: boolean }) => {
      setLoading(data.loading);
    };
    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, handleLoadingChange);

    return () => {
      onlyofficeEventbus.off(
        ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE,
        handleLoadingChange,
      );
      onlyOfficeManagerFactory.destroyAll();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div className="flex min-h-screen flex-col bg-gray-50">
      <div className="border-b border-gray-200 bg-white px-5 py-4 shadow-sm">
        <h1 className="text-lg font-semibold text-gray-900">真正多实例实验</h1>
        <p className="mt-1 text-sm text-gray-500">
          当前页面同时创建 Word / Excel / PowerPoint 三个独立 OnlyOffice 实例。
        </p>
      </div>

      {error && (
        <div className="mx-4 mt-4 rounded border-l-4 border-red-500 bg-red-50 p-4 text-red-700">
          {error}
        </div>
      )}

      <div className="grid flex-1 grid-cols-1 gap-4 p-4 lg:grid-cols-3">
        {PANELS.map((panel) => (
          <section
            key={panel.key}
            className="flex min-h-[560px] flex-col overflow-hidden rounded-lg border border-gray-200 bg-white shadow-sm"
          >
            <div className="flex flex-wrap items-center gap-2 border-b border-gray-200 bg-gray-50 px-3 py-3">
              <div
                className={`flex h-7 w-7 items-center justify-center rounded bg-gradient-to-br text-sm font-bold text-white ${panel.color}`}
              >
                {panel.badge}
              </div>
              <div className="mr-auto min-w-0">
                <h2 className="text-sm font-semibold text-gray-900">
                  {panel.title}
                </h2>
                <p className="truncate text-xs text-gray-500">
                  {documents[panel.key].fileName}
                </p>
              </div>
              <button
                type="button"
                onClick={() => inputRefs[panel.key].current?.click()}
                className="rounded bg-blue-500 px-3 py-1.5 text-xs text-white hover:bg-blue-600"
              >
                上传
              </button>
              <button
                type="button"
                onClick={() => newPanelDocument(panel.key)}
                className="rounded border border-gray-300 bg-white px-3 py-1.5 text-xs hover:bg-gray-50"
              >
                新建
              </button>
              <button
                type="button"
                onClick={() => exportPanel(panel.key)}
                className="rounded border border-gray-300 bg-white px-3 py-1.5 text-xs hover:bg-gray-50"
              >
                导出
              </button>
              <button
                type="button"
                onClick={() => toggleReadOnly(panel.key)}
                className={`rounded px-3 py-1.5 text-xs ${
                  readOnlyStates[panel.key]
                    ? "bg-yellow-500 text-white hover:bg-yellow-600"
                    : "border border-gray-300 bg-white hover:bg-gray-50"
                }`}
              >
                {readOnlyStates[panel.key] ? "只读" : "编辑"}
              </button>
            </div>

            <div className="relative flex-1">
              <OnlyOfficePanelHost containerId={panel.containerId} />
            </div>

            <input
              ref={inputRefs[panel.key]}
              type="file"
              accept={panel.accept}
              className="hidden"
              onChange={(event) => {
                const file = event.target.files?.[0];
                if (file) {
                  uploadPanelFile(panel.key, file);
                  const input = inputRefs[panel.key].current;
                  if (input) input.value = "";
                }
              }}
            />
          </section>
        ))}
      </div>

      {loading && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-white/40">
          <div className="rounded bg-white px-4 py-3 text-sm shadow">
            加载中...
          </div>
        </div>
      )}
    </div>
  );
}
