export type CommentData = {
  Text?: string;
  UserName?: string;
  Time?: string | number | Date;
  Replies?: CommentData[];
  [key: string]: unknown;
};

export type CommentInput = CommentData | string;

export type CommentItem = {
  Id: string;
  Data: CommentData;
};

export type CommentChangeHandlers = {
  onAdd?: (id: string, data: CommentData) => void;
  onChange?: (id: string, data: CommentData) => void;
  onRemove?: (id: string) => void;
};

export function normalizeCommentInput(input: CommentInput): CommentData {
  if (typeof input === "string") {
    return { Text: input };
  }

  return input;
}

/** OnlyOffice pluginMethod_* / CZ.$be 识别的批注字段。 */
export function toPluginCommentPayload(data: CommentData): CommentData {
  const now = String(Date.now());

  return {
    Text: data.Text ?? "",
    UserName: data.UserName ?? "Guest",
    Time: data.Time != null ? String(data.Time) : now,
    QuoteText: data.QuoteText ?? "",
    Solved: data.Solved ?? false,
    UserData: data.UserData ?? "",
    Replies: Array.isArray(data.Replies) ? data.Replies : [],
    ...data,
  };
}

export function isResolvedComment(data: unknown) {
  if (!data || typeof data !== "object") {
    return false;
  }

  const comment = data as Record<string, unknown>;
  const solved =
    comment.Solved ??
    comment.solved ??
    comment.Resolved ??
    comment.resolved ??
    comment.Done ??
    comment.done;

  if (typeof solved === "boolean") {
    return solved;
  }

  const ascGetSolved = comment.asc_getSolved;
  if (typeof ascGetSolved === "function") {
    return !!ascGetSolved.call(data);
  }

  const getSolved = comment.get_Solved;
  if (typeof getSolved === "function") {
    return !!getSolved.call(data);
  }

  return false;
}
