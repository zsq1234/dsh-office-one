const OOXML_MAX_EXTENT_EMU = 2_147_483_647;
const EMU_PER_POINT = 12_700;
const OOXML_MAX_EXTENT_PT = OOXML_MAX_EXTENT_EMU / EMU_PER_POINT;

function normalizeExtent(transform, sizeKey, positionKey, flipKey) {
  const size = transform[sizeKey];
  if (typeof size !== 'number' || !Number.isFinite(size)) return;

  if (size < 0) {
    const position = transform[positionKey];
    if (typeof position === 'number' && Number.isFinite(position)) transform[positionKey] = position + size;
    transform[sizeKey] = Math.min(-size, OOXML_MAX_EXTENT_PT);
    transform[flipKey] = !Boolean(transform[flipKey]);
    return;
  }

  if (size > OOXML_MAX_EXTENT_PT) transform[sizeKey] = OOXML_MAX_EXTENT_PT;
}

function normalizePageElements(pages) {
  if (!pages || typeof pages !== 'object') return;
  for (const page of Object.values(pages)) {
    if (!page || typeof page !== 'object' || !page.elements || typeof page.elements !== 'object') continue;
    for (const element of Object.values(page.elements)) {
      if (!element || typeof element !== 'object' || !element.transform || typeof element.transform !== 'object') continue;
      normalizeExtent(element.transform, 'width', 'left', 'flipX');
      normalizeExtent(element.transform, 'height', 'top', 'flipY');
    }
  }
}

// Repairs presentation snapshots before OOXML export. Univer can preserve a
// negative width/height while an element is flipped or resized across its
// opposite edge, but OOXML extents must be unsigned 32-bit EMU values. Move the
// origin to the opposite edge, make the extent positive, and toggle the matching
// flip flag so the occupied bounds and visual orientation remain equivalent.
// Also repairs imported pages whose layout reference is missing or stale.
export function normalizeSlides(data) {
  const unit = data?.slide ?? data;
  if (!unit || typeof unit !== 'object') return data;

  for (const collection of [
    unit.slides,
    unit.layoutPages,
    unit.masterPages,
    unit.handoutMasterPages,
    unit.notesMasterPages,
  ]) normalizePageElements(collection);

  const { slides, layoutPageOrder, layoutPages, slideOrder } = unit;
  if (!slides || typeof slides !== 'object') return data;

  const layoutIds = new Set();
  if (Array.isArray(layoutPageOrder)) for (const id of layoutPageOrder) layoutIds.add(id);
  if (layoutPages && typeof layoutPages === 'object') for (const id of Object.keys(layoutPages)) layoutIds.add(id);
  if (layoutIds.size === 0) return data;

  const ordered = Array.isArray(slideOrder) && slideOrder.length > 0 ? slideOrder : Object.keys(slides);
  let fallback = null;
  for (const slideId of ordered) {
    const page = slides[slideId];
    if (page && typeof page === 'object' && typeof page.layoutPageId === 'string' && layoutIds.has(page.layoutPageId)) {
      fallback = page.layoutPageId;
      break;
    }
  }
  if (fallback === null) fallback = layoutIds.values().next().value;

  for (const page of Object.values(slides)) {
    if (!page || typeof page !== 'object') continue;
    if (typeof page.layoutPageId !== 'string' || !layoutIds.has(page.layoutPageId)) page.layoutPageId = fallback;
  }
  return data;
}
