interface DocEditorConstructor {
  new (id?: string, config: any): DocEditor;
  static defaultConfig: any;
  static version(): string;
  static warmUp(id?: string): void;
}

interface Window {
  DocsAPI?: {
    DocEditor: DocEditorConstructor;
  };
}
