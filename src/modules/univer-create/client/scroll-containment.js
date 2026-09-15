/**
 * Univer's Slides presentation dock keeps the active thumbnail visible with
 *
 *   element.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' })
 *
 * on every page turn. `scrollIntoView` scrolls *every* scrollable ancestor, and
 * `inline: 'center'` adds a horizontal component, so that call also drags the
 * surrounding DSH layout (the conversation panes, and eventually the document)
 * to the left — and leaves it there, because `overflow: hidden` ancestors are
 * still programmatically scrollable and expose no scrollbar to recover with.
 *
 * These helpers make the call harmless without losing the dock's own scrolling:
 * for any element inside the Univer mount point the scroll is applied only to
 * scroll containers that sit strictly between the element and the presentation
 * overlay (presentation mode) or the mount point (editor mode). Nothing above
 * the mount point can move.
 */

const PRESENTATION_OVERLAY_SELECTOR = '.univer-fixed.univer-inset-0'
const SCROLLABLE_OVERFLOW = /^(?:auto|scroll|hidden|clip)$/

/**
 * @param {CSSStyleDeclaration} style
 * @param {'x' | 'y'} axis
 * @returns {boolean}
 */
function axisAllowsScroll(style, axis) {
  return SCROLLABLE_OVERFLOW.test(axis === 'x' ? style.overflowX : style.overflowY)
}

/**
 * Distance the scroll container has to move so that the element satisfies the
 * requested alignment. Mirrors the `scrollIntoView` alignment keywords.
 *
 * @param {ScrollLogicalPosition} position
 * @param {number} elementStart
 * @param {number} elementEnd
 * @param {number} boxStart
 * @param {number} boxEnd
 * @returns {number}
 */
function resolveDelta(position, elementStart, elementEnd, boxStart, boxEnd) {
  switch (position) {
    case 'start':
      return elementStart - boxStart
    case 'end':
      return elementEnd - boxEnd
    case 'center':
      return (elementStart + elementEnd) / 2 - (boxStart + boxEnd) / 2
    default:
      if (elementStart < boxStart) return elementStart - boxStart
      if (elementEnd > boxEnd) return elementEnd - boxEnd
      return 0
  }
}

/**
 * @param {unknown} options
 * @returns {{ behavior: ScrollBehavior, block: ScrollLogicalPosition, inline: ScrollLogicalPosition }}
 */
function readOptions(options) {
  if (typeof options === 'object' && options !== null) {
    const value = /** @type {ScrollIntoViewOptions} */ (options)
    return {
      behavior: value.behavior === 'smooth' || value.behavior === 'instant' ? value.behavior : 'auto',
      block: value.block ?? 'start',
      inline: value.inline ?? 'nearest',
    }
  }
  // Legacy boolean form: `true` aligns to the top, `false` to the bottom.
  return { behavior: 'auto', block: options === false ? 'end' : 'start', inline: 'nearest' }
}

/**
 * Scroll only the containers inside `boundary`, never `boundary` itself and
 * never anything above it.
 *
 * @param {HTMLElement} boundary
 * @param {Element} element
 * @param {unknown} options
 */
function scrollWithin(boundary, element, options) {
  const { behavior, block, inline } = readOptions(options)
  for (let node = element.parentElement; node !== null && node !== boundary; node = node.parentElement) {
    const style = getComputedStyle(node)
    const canScrollX = axisAllowsScroll(style, 'x') && node.scrollWidth > node.clientWidth
    const canScrollY = axisAllowsScroll(style, 'y') && node.scrollHeight > node.clientHeight
    if (!canScrollX && !canScrollY) continue

    const rect = element.getBoundingClientRect()
    const box = node.getBoundingClientRect()
    const deltaX = canScrollX ? resolveDelta(inline, rect.left, rect.right, box.left, box.right) : 0
    const deltaY = canScrollY ? resolveDelta(block, rect.top, rect.bottom, box.top, box.bottom) : 0
    if (deltaX === 0 && deltaY === 0) continue

    const left = node.scrollLeft + deltaX
    const top = node.scrollTop + deltaY
    if (behavior === 'smooth') node.scrollTo({ left, top, behavior: 'smooth' })
    else if (deltaX !== 0 || deltaY !== 0) node.scrollTo({ left, top, behavior: 'auto' })
  }
}

/**
 * Contain `scrollIntoView` and `focus` for everything mounted under `container`.
 *
 * @param {HTMLElement} container
 * @returns {() => void} restore function
 */
export function installScrollContainment(container) {
  const originalScrollIntoView = Element.prototype.scrollIntoView
  const originalFocus = HTMLElement.prototype.focus
  const originalScrollIntoViewIfNeeded = /** @type {any} */ (Element.prototype).scrollIntoViewIfNeeded

  /**
   * @param {Element} element
   * @returns {HTMLElement | null}
   */
  const presentationOverlayFor = (element) => {
    const overlay = element.closest(PRESENTATION_OVERLAY_SELECTOR)
    return overlay instanceof HTMLElement && container.contains(overlay) ? overlay : null
  }

  Element.prototype.scrollIntoView = function (options) {
    if (!container.contains(this)) return originalScrollIntoView.call(this, options)
    scrollWithin(presentationOverlayFor(this) ?? container, this, options)
  }

  if (typeof originalScrollIntoViewIfNeeded === 'function') {
    /** @type {any} */ (Element.prototype).scrollIntoViewIfNeeded = function (centerIfNeeded) {
      if (!container.contains(this)) return originalScrollIntoViewIfNeeded.call(this, centerIfNeeded)
      const boundary = presentationOverlayFor(this) ?? container
      scrollWithin(boundary, this, { block: centerIfNeeded === true ? 'center' : 'nearest', inline: centerIfNeeded === true ? 'center' : 'nearest' })
    }
  }

  HTMLElement.prototype.focus = function (options) {
    if (!container.contains(this)) return originalFocus.call(this, options)
    // Univer focuses the full-screen presentation overlay after every page turn.
    // Without `preventScroll` the browser scrolls the host to reveal it.
    if (presentationOverlayFor(this) === null) return originalFocus.call(this, options)
    return originalFocus.call(this, { ...(options ?? {}), preventScroll: true })
  }

  return () => {
    Element.prototype.scrollIntoView = originalScrollIntoView
    HTMLElement.prototype.focus = originalFocus
    if (typeof originalScrollIntoViewIfNeeded === 'function') {
      /** @type {any} */ (Element.prototype).scrollIntoViewIfNeeded = originalScrollIntoViewIfNeeded
    }
  }
}
