export interface CommandEventLike {
  completed(): void;
}

export interface AsyncErrorLike {
  code?: string | number;
  name?: string;
  message?: string;
}

export interface DisplayedBodyLike {
  setAsync: (...args: any[]) => void;
}

export function getDisplayedBody(item: unknown): DisplayedBodyLike | null {
  const candidate = item as
    | { display?: { body?: { setAsync?: unknown } } }
    | null
    | undefined;
  const body = candidate?.display?.body;

  return body && typeof body.setAsync === "function"
    ? (body as DisplayedBodyLike)
    : null;
}

export function formatAsyncError(
  error: AsyncErrorLike | null | undefined,
): string {
  return `code=${String(error?.code ?? "unknown")}, name=${error?.name ?? "unknown"}, message=${error?.message ?? "unknown"}`;
}

export async function completeCommand(
  event: CommandEventLike,
  work: () => Promise<void>,
): Promise<void> {
  try {
    await work();
  } finally {
    event.completed();
  }
}
