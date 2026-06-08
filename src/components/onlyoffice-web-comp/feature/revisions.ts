export type RevisionItem = {
  Id: string;
  Index: number;
  Data?: unknown;
  Raw: unknown;
};

export type RevisionChangeHandlers = {
  onShowChanges?: (items: RevisionItem[]) => void;
  onTrackRevisionsChange?: (enabled: boolean) => void;
};

export function normalizeRevisionItem(raw: unknown, index: number): RevisionItem {
  const source = raw && typeof raw === "object" ? (raw as Record<string, unknown>) : {};
  const id = source.Id || source.id || source.Guid || source.guid || `rev-${index}`;

  return {
    Id: String(id),
    Index: index,
    Data: source.Data || source.data || source,
    Raw: raw,
  };
}

export function normalizeRevisionItems(raw: unknown): RevisionItem[] {
  if (!Array.isArray(raw)) return [];

  return raw.map((item, index) => normalizeRevisionItem(item, index));
}
