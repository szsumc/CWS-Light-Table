export function createViewportState(overrides = {}) {
  return {
    scrollTop: normalizeNonNegativeNumber(overrides.scrollTop),
    scrollLeft: normalizeNonNegativeNumber(overrides.scrollLeft),
    width: normalizeNonNegativeNumber(overrides.width),
    height: normalizeNonNegativeNumber(overrides.height),
    contentWidth: normalizeNonNegativeNumber(overrides.contentWidth),
    contentHeight: normalizeNonNegativeNumber(overrides.contentHeight),
  };
}

export function readViewportState(host, overrides = {}) {
  const element = isElement(host) ? host : null;
  return createViewportState({
    scrollTop: element?.scrollTop ?? 0,
    scrollLeft: element?.scrollLeft ?? 0,
    width: element?.clientWidth ?? 0,
    height: element?.clientHeight ?? 0,
    contentWidth: overrides.contentWidth ?? element?.scrollWidth ?? 0,
    contentHeight: overrides.contentHeight ?? element?.scrollHeight ?? 0,
  });
}

export function updateViewportState(previousState, host, overrides = {}) {
  const current = readViewportState(host, overrides);
  if (!viewportStateChanged(previousState, current)) {
    return previousState;
  }
  return current;
}

export function viewportStateChanged(previousState, nextState) {
  const previous = createViewportState(previousState);
  const next = createViewportState(nextState);
  return (
    previous.scrollTop !== next.scrollTop ||
    previous.scrollLeft !== next.scrollLeft ||
    previous.width !== next.width ||
    previous.height !== next.height ||
    previous.contentWidth !== next.contentWidth ||
    previous.contentHeight !== next.contentHeight
  );
}

function normalizeNonNegativeNumber(value) {
  return Math.max(0, Math.round(Number(value) || 0));
}

function isElement(value) {
  return Boolean(value && typeof value === "object");
}
