type EventListener = (...args: any[]) => void;

export type { EventListener };

/**
 * 轻量 EventEmitter，供 EventBus 与 MockSocket 使用。
 * 仅实现项目内用到的 on / off / once / emit / removeAllListeners。
 */
export class EventEmitter {
  private listeners = new Map<string | symbol, EventListener[]>();

  on(event: string | symbol, listener: EventListener): this {
    const list = this.listeners.get(event) ?? [];
    list.push(listener);
    this.listeners.set(event, list);
    return this;
  }

  once(event: string | symbol, listener: EventListener): this {
    const wrapper: EventListener = (...args) => {
      this.off(event, wrapper);
      listener(...args);
    };
    return this.on(event, wrapper);
  }

  off(event: string | symbol, listener?: EventListener): this {
    if (!listener) {
      this.listeners.delete(event);
      return this;
    }

    const list = this.listeners.get(event);
    if (!list) {
      return this;
    }

    const next = list.filter((fn) => fn !== listener);
    if (next.length) {
      this.listeners.set(event, next);
    } else {
      this.listeners.delete(event);
    }
    return this;
  }

  removeAllListeners(event?: string | symbol): this {
    if (event === undefined) {
      this.listeners.clear();
    } else {
      this.listeners.delete(event);
    }
    return this;
  }

  emit(event: string | symbol, ...args: any[]): boolean {
    const list = this.listeners.get(event);
    if (!list?.length) {
      return false;
    }

    for (const listener of [...list]) {
      listener(...args);
    }
    return true;
  }
}
