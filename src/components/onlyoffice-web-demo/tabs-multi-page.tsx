"use client";

import { memo, useEffect, useRef, useState } from "react";
import { nanoid } from "nanoid";
import {
  FILE_TYPE,
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_EVENT_KEYS,
  onlyOfficeManagerFactory,
  onlyofficeEventbus,
  type FileType,
} from "@/components/onlyoffice-web-comp";

type DocKind = "word" | "excel" | "ppt";

type DocPreset = {
  label: string;
  badge: string;
  badgeClassName: string;
  tabAccent: string;
  tabIdleBg: string;
  tabIdleText: string;
  tabAddBtn: string;
  fileType: FileType;
  defaultFileName: string;
  accept: string;
};

type TabItem = {
  id: string;
  label: string;
  containerId: string;
  fileName: string;
  readOnly: boolean;
  docKind: DocKind;
};

const DOC_PRESETS: Record<DocKind, DocPreset> = {
  word: {
    label: "Word",
    badge: "W",
    badgeClassName: "bg-blue-100 text-blue-600",
    tabAccent: "border-t-blue-300",
    tabIdleBg: "hover:bg-white/60",
    tabIdleText: "text-gray-500",
    tabAddBtn:
      "border border-blue-100 bg-blue-50/60 text-blue-600/80 hover:bg-blue-50 hover:border-blue-200",
    fileType: FILE_TYPE.DOCX,
    defaultFileName: "New_Document.docx",
    accept: ".docx,.doc,.odt,.rtf,.txt",
  },
  excel: {
    label: "Excel",
    badge: "E",
    badgeClassName: "bg-emerald-50 text-emerald-600",
    tabAccent: "border-t-emerald-300",
    tabIdleBg: "hover:bg-white/60",
    tabIdleText: "text-gray-500",
    tabAddBtn:
      "border border-emerald-100 bg-emerald-50/60 text-emerald-600/80 hover:bg-emerald-50 hover:border-emerald-200",
    fileType: FILE_TYPE.XLSX,
    defaultFileName: "New_Spreadsheet.xlsx",
    accept: ".xlsx,.xls,.ods,.csv",
  },
  ppt: {
    label: "PPT",
    badge: "P",
    badgeClassName: "bg-orange-50 text-orange-600",
    tabAccent: "border-t-orange-300",
    tabIdleBg: "hover:bg-white/60",
    tabIdleText: "text-gray-500",
    tabAddBtn:
      "border border-orange-100 bg-orange-50/60 text-orange-600/80 hover:bg-orange-50 hover:border-orange-200",
    fileType: FILE_TYPE.PPTX,
    defaultFileName: "New_Presentation.pptx",
    accept: ".pptx,.ppt,.odp",
  },
};

const OnlyOfficeTabHost = memo(function OnlyOfficeTabHost({
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

function getPreset(docKind: DocKind) {
  return DOC_PRESETS[docKind];
}

function isNewDocument(tab: TabItem) {
  return tab.fileName === getPreset(tab.docKind).defaultFileName;
}

function createTab(index: number, docKind: DocKind): TabItem {
  const id = nanoid(6);
  const preset = getPreset(docKind);
  return {
    id,
    label: `${preset.label} ${index}`,
    containerId: `tab-editor-${id}`,
    fileName: preset.defaultFileName,
    readOnly: false,
    docKind,
  };
}

function createInitialTabState() {
  const initialTab = createTab(1, "word");
  return { tabs: [initialTab], activeId: initialTab.id };
}

export function TabsMultiPage() {
  const [tabs, setTabs] = useState<TabItem[]>([]);
  const [activeId, setActiveId] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const initializedRef = useRef(new Set<string>());
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const { tabs: initialTabs, activeId: initialActiveId } = createInitialTabState();
    setTabs(initialTabs);
    setActiveId(initialActiveId);
  }, []);

  const activeTab = tabs.find((tab) => tab.id === activeId);
  const activePreset = activeTab ? getPreset(activeTab.docKind) : null;

  const updateTab = (tabId: string, patch: Partial<TabItem>) => {
    setTabs((prev) =>
      prev.map((tab) => (tab.id === tabId ? { ...tab, ...patch } : tab)),
    );
  };

  const runAction = async (action: () => Promise<void>, message: string) => {
    try {
      setError(null);
      await action();
    } catch (err) {
      setError(message);
      console.error(message, err);
    }
  };

  const openTabEditor = async (tab: TabItem) => {
    const preset = getPreset(tab.docKind);

    await onlyOfficeManagerFactory.open(
      {
        containerId: tab.containerId,
        fileType: preset.fileType,
        defaultFileName: preset.defaultFileName,
        readOnly: tab.readOnly,
      },
      {
        fileName: tab.fileName,
        isNew: isNewDocument(tab),
        readOnly: tab.readOnly,
      },
    );

    initializedRef.current.add(tab.id);
  };

  useEffect(() => {
    if (!activeId) return;

    const tab = tabs.find((item) => item.id === activeId);
    if (!tab || initializedRef.current.has(tab.id)) return;

    let cancelled = false;

    openTabEditor(tab)
      .then(() => {
        if (cancelled) {
          onlyOfficeManagerFactory.destroy(tab.containerId);
          initializedRef.current.delete(tab.id);
        }
      })
      .catch((err) => {
        if (cancelled) return;
        setError("无法加载编辑器");
        console.error("Failed to open tab editor:", err);
      });

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeId, tabs]);

  useEffect(() => {
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
      initializedRef.current.clear();
    };
  }, []);

  const addTab = (docKind: DocKind) => {
    const nextTab = createTab(tabs.length + 1, docKind);
    setTabs((prev) => [...prev, nextTab]);
    setActiveId(nextTab.id);
  };

  const closeTab = (tabId: string) => {
    if (tabs.length <= 1) return;

    const tab = tabs.find((item) => item.id === tabId);
    if (tab) {
      onlyOfficeManagerFactory.destroy(tab.containerId);
      initializedRef.current.delete(tab.id);
    }

    setTabs((prev) => {
      const next = prev.filter((item) => item.id !== tabId);
      if (activeId === tabId) {
        setActiveId(next[0]?.id ?? "");
      }
      return next;
    });
  };

  const ensureActiveManager = async () => {
    if (!activeTab) throw new Error("No active tab");

    if (!initializedRef.current.has(activeTab.id)) {
      await openTabEditor(activeTab);
    }

    const manager = onlyOfficeManagerFactory.get(activeTab.containerId);
    if (!manager) throw new Error("Editor is not initialized");
    return manager;
  };

  const uploadFile = (file: File) =>
    runAction(async () => {
      if (!activeTab) return;

      const preset = getPreset(activeTab.docKind);

      await onlyOfficeManagerFactory.open(
        {
          containerId: activeTab.containerId,
          fileType: preset.fileType,
          defaultFileName: preset.defaultFileName,
          readOnly: activeTab.readOnly,
        },
        {
          fileName: file.name,
          file,
          readOnly: activeTab.readOnly,
        },
      );

      initializedRef.current.add(activeTab.id);
      updateTab(activeTab.id, { fileName: file.name });
    }, "上传失败");

  const newDocument = () =>
    runAction(async () => {
      if (!activeTab) return;

      const preset = getPreset(activeTab.docKind);

      await onlyOfficeManagerFactory.open(
        {
          containerId: activeTab.containerId,
          fileType: preset.fileType,
          defaultFileName: preset.defaultFileName,
          readOnly: activeTab.readOnly,
        },
        {
          fileName: preset.defaultFileName,
          isNew: true,
          readOnly: activeTab.readOnly,
        },
      );

      initializedRef.current.add(activeTab.id);
      updateTab(activeTab.id, { fileName: preset.defaultFileName });
    }, "新建失败");

  const exportDocument = () =>
    runAction(async () => {
      const manager = await ensureActiveManager();
      await manager.downloadExport();
    }, "导出失败");

  const toggleReadOnly = () =>
    runAction(async () => {
      if (!activeTab) return;

      const preset = getPreset(activeTab.docKind);
      const manager = await ensureActiveManager();
      const nextReadOnly = !activeTab.readOnly;

      if (manager.isReady()) {
        await manager.setReadOnly(nextReadOnly);
      } else {
        await onlyOfficeManagerFactory.open(
          {
            containerId: activeTab.containerId,
            fileType: preset.fileType,
            defaultFileName: preset.defaultFileName,
            readOnly: nextReadOnly,
          },
          {
            fileName: activeTab.fileName,
            isNew: isNewDocument(activeTab),
            readOnly: nextReadOnly,
          },
        );
      }

      updateTab(activeTab.id, { readOnly: nextReadOnly });
    }, "切换模式失败");

  return (
    <div className="flex h-screen flex-col bg-gray-50">
      <div className="border-b border-gray-200 bg-gradient-to-r from-white to-gray-50 shadow-sm">
        <div className="flex flex-wrap items-center gap-4 px-5 py-4">
          <div className="mr-auto flex min-w-0 items-center gap-3">
            <div
              className={`flex h-7 w-7 shrink-0 items-center justify-center rounded-md text-xs font-semibold ${
                activePreset?.badgeClassName ?? "bg-blue-100 text-blue-600"
              }`}
            >
              {activePreset?.badge ?? "T"}
            </div>
            <div className="min-w-0">
              <h1 className="text-lg font-semibold text-gray-900">
                多实例 Tab 演示
              </h1>
              <p className="truncate text-xs text-gray-500">
                {activeTab
                  ? `${activePreset?.label} · ${activeTab.fileName}`
                  : "切换标签页，实例状态会保留"}
              </p>
            </div>
          </div>

          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="rounded-md bg-blue-500 px-4 py-2 text-sm text-white transition-colors hover:bg-blue-600"
            >
              上传文档
            </button>
            <button
              type="button"
              onClick={newDocument}
              className="rounded-md border border-gray-300 bg-white px-4 py-2 text-sm transition-colors hover:bg-gray-50"
            >
              新建{activePreset?.label ?? "文档"}
            </button>
            <button
              type="button"
              onClick={exportDocument}
              className="rounded-md border border-gray-300 bg-white px-4 py-2 text-sm transition-colors hover:bg-gray-50"
            >
              导出
            </button>
            <button
              type="button"
              onClick={toggleReadOnly}
              className={`rounded-md px-4 py-2 text-sm transition-colors ${
                activeTab?.readOnly
                  ? "bg-yellow-500 text-white hover:bg-yellow-600"
                  : "border border-gray-300 bg-white hover:bg-gray-50"
              }`}
            >
              {activeTab?.readOnly ? "只读模式" : "编辑模式"}
            </button>
          </div>
        </div>

        <div className="border-t border-gray-100 bg-gray-50/80 px-3 py-1">
          <div className="flex items-end gap-0.5 overflow-x-auto">
            {tabs.map((tab) => {
              const preset = getPreset(tab.docKind);
              const isActive = activeId === tab.id;
              return (
                <div
                  key={tab.id}
                  className={`group relative flex max-w-[200px] min-w-[96px] shrink-0 items-stretch rounded-t-md border border-b-0 transition-colors ${
                    isActive
                      ? `z-10 -mb-px border-gray-200 bg-white ${preset.tabAccent} border-t-2`
                      : `border-transparent ${preset.tabIdleBg}`
                  }`}
                >
                  <button
                    type="button"
                    onClick={() => setActiveId(tab.id)}
                    className={`flex min-w-0 flex-1 items-center gap-1.5 px-2.5 py-1.5 text-xs transition-colors ${
                      isActive ? "text-gray-800" : preset.tabIdleText
                    }`}
                    title={tab.fileName}
                  >
                    <span
                      className={`flex h-4 w-4 shrink-0 items-center justify-center rounded text-[9px] font-semibold ${preset.badgeClassName}`}
                    >
                      {preset.badge}
                    </span>
                    <span className="truncate">{tab.label}</span>
                  </button>
                  {tabs.length > 1 && (
                    <button
                      type="button"
                      onClick={() => closeTab(tab.id)}
                      className={`mr-1 self-center rounded px-1 text-[10px] transition-colors ${
                        isActive
                          ? "text-gray-400 hover:text-gray-600"
                          : "text-gray-300 opacity-0 hover:text-gray-500 group-hover:opacity-100"
                      }`}
                      aria-label={`关闭 ${tab.label}`}
                    >
                      ×
                    </button>
                  )}
                </div>
              );
            })}

            <div className="mb-px flex shrink-0 items-center gap-1.5 pl-2">
              {(Object.keys(DOC_PRESETS) as DocKind[]).map((kind) => {
                const preset = getPreset(kind);
                return (
                  <button
                    key={kind}
                    type="button"
                    onClick={() => addTab(kind)}
                    className={`rounded px-2 py-1 text-xs transition-colors ${preset.tabAddBtn}`}
                    title={`新建 ${preset.label} 标签页`}
                  >
                    + {preset.label}
                  </button>
                );
              })}
            </div>
          </div>
        </div>
      </div>

      {error && (
        <div className="mx-4 mt-4 rounded border-l-4 border-red-500 bg-red-50 p-3 text-sm text-red-700">
          {error}
        </div>
      )}

      <div className="relative flex-1 bg-white">
        {tabs.length === 0 && (
          <div className="absolute inset-0 flex items-center justify-center text-sm text-gray-400">
            加载中...
          </div>
        )}
        {tabs.map((tab) => (
          <div
            key={tab.id}
            className={`absolute inset-0 ${
              activeId === tab.id ? "visible z-10" : "invisible z-0"
            }`}
          >
            <OnlyOfficeTabHost containerId={tab.containerId} />
          </div>
        ))}

        {loading && (
          <div className="absolute inset-0 z-20 flex items-center justify-center bg-white/70 backdrop-blur-sm">
            <div className="rounded-lg border border-gray-200 bg-white px-4 py-3 text-sm text-gray-700 shadow">
              加载中...
            </div>
          </div>
        )}
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept={activePreset?.accept ?? ".docx,.doc,.odt,.rtf,.txt"}
        className="hidden"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) {
            uploadFile(file);
            if (fileInputRef.current) fileInputRef.current.value = "";
          }
        }}
      />
    </div>
  );
}
