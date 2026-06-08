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

/** SDK 修订栈为空时，从 asc_GetTrackRevisionsReportByAuthors 拉平为 RevisionItem 列表。 */
export function flattenRevisionsReportByAuthors(
  report: unknown,
): RevisionItem[] {
  if (!report || typeof report !== "object") {
    return [];
  }

  const items: RevisionItem[] = [];
  let index = 0;

  for (const [author, changes] of Object.entries(
    report as Record<string, unknown>,
  )) {
    if (!Array.isArray(changes)) {
      continue;
    }

    for (let changeIndex = 0; changeIndex < changes.length; changeIndex++) {
      items.push({
        Id: `rev-${index}`,
        Index: index,
        Data: { author, changeIndex },
        Raw: changes[changeIndex],
      });
      index += 1;
    }
  }

  return items;
}
