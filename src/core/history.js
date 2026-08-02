export function createHistoryState() {
  return {
    past: [],
    future: [],
  };
}

export function canUndo(history) {
  return (history?.past?.length ?? 0) > 0;
}

export function canRedo(history) {
  return (history?.future?.length ?? 0) > 0;
}

export function pushHistorySnapshot(history, snapshot) {
  return {
    past: [...(history?.past ?? []), cloneSnapshot(snapshot)],
    future: [],
  };
}

export function undoHistory(history, currentSnapshot) {
  const past = history?.past ?? [];
  if (!past.length) {
    return {
      history: history ?? createHistoryState(),
      snapshot: null,
    };
  }
  return {
    history: {
      past: past.slice(0, -1),
      future: [cloneSnapshot(currentSnapshot), ...(history?.future ?? [])],
    },
    snapshot: cloneSnapshot(past.at(-1)),
  };
}

export function redoHistory(history, currentSnapshot) {
  const future = history?.future ?? [];
  if (!future.length) {
    return {
      history: history ?? createHistoryState(),
      snapshot: null,
    };
  }
  return {
    history: {
      past: [...(history?.past ?? []), cloneSnapshot(currentSnapshot)],
      future: future.slice(1),
    },
    snapshot: cloneSnapshot(future[0]),
  };
}

function cloneSnapshot(snapshot) {
  return JSON.parse(JSON.stringify(snapshot ?? null));
}
