import { EventEmitter, type EventListener } from "../util/event-emitter";
import { ONLYOFFICE_EVENT_KEYS } from "../const";

export type DocumentReadyData = {
  fileName: string;
  fileType: string;
  instanceId?: string;
};

export type SaveDocumentData = {
  fileName: string;
  fileType: string;
  binData: Uint8Array;
  instanceId: string;
  media?: Record<string, Uint8Array>;
};

export type OnSaveData = {
  fileName: string;
  instanceId: string;
};

export type LoadingChangeData = {
  loading: boolean;
};

type OnlyOfficeEventPayloads = {
  [ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY]: DocumentReadyData;
  [ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT]: SaveDocumentData;
  [ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE]: LoadingChangeData;
  [ONLYOFFICE_EVENT_KEYS.ONSAVE]: OnSaveData;
};

class OnlyOfficeEventBus {
  private emitter = new EventEmitter();

  on<Key extends keyof OnlyOfficeEventPayloads>(
    key: Key,
    handler: (data: OnlyOfficeEventPayloads[Key]) => void,
  ) {
    this.emitter.on(key, handler as EventListener);
  }

  off<Key extends keyof OnlyOfficeEventPayloads>(
    key: Key,
    handler: (data: OnlyOfficeEventPayloads[Key]) => void,
  ) {
    this.emitter.off(key, handler as EventListener);
  }

  emit<Key extends keyof OnlyOfficeEventPayloads>(
    key: Key,
    data: OnlyOfficeEventPayloads[Key],
  ) {
    this.emitter.emit(key, data);
  }

  waitFor<Key extends keyof OnlyOfficeEventPayloads>(
    key: Key,
    timeout = 30000,
  ): Promise<OnlyOfficeEventPayloads[Key]> {
    return new Promise((resolve, reject) => {
      const timer = window.setTimeout(() => {
        this.off(key, handler);
        reject(new Error(`Timed out waiting for ${key}`));
      }, timeout);

      const handler = (data: OnlyOfficeEventPayloads[Key]) => {
        window.clearTimeout(timer);
        this.off(key, handler);
        resolve(data);
      };

      this.on(key, handler);
    });
  }
}

export const onlyofficeEventbus = new OnlyOfficeEventBus();
