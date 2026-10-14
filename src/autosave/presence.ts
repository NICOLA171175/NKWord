/**
 * Quali finestre di NKWord sono aperte adesso. Serve a non proporre il
 * recupero del documento di un'altra scheda ancora aperta: solo le copie
 * delle finestre chiuse (o andate in crash) sono da recuperare.
 */

export interface Presence {
  /** Id delle altre finestre che rispondono entro poco tempo. */
  others(): Promise<Set<string>>;
  dispose(): void;
}

const CHANNEL = 'nkoffice-finestre';

/** Con BroadcastChannel le finestre aperte rispondono a un "chi c'è?". */
export function createPresence(selfId: string, waitMs = 300): Presence {
  if (typeof BroadcastChannel === 'undefined') return { others: async () => new Set(), dispose() {} };
  const channel = new BroadcastChannel(CHANNEL);
  const seen = new Set<string>();
  channel.onmessage = (e: MessageEvent<{ type: string; id: string }>) => {
    if (e.data?.type === 'chi-c-e') channel.postMessage({ type: 'ci-sono', id: selfId });
    else if (e.data?.type === 'ci-sono' && e.data.id !== selfId) seen.add(e.data.id);
  };
  return {
    async others() {
      channel.postMessage({ type: 'chi-c-e', id: selfId });
      await new Promise((r) => setTimeout(r, waitMs));
      return new Set(seen);
    },
    dispose() {
      channel.close();
    },
  };
}
