import assert from 'node:assert/strict'
import test from 'node:test'

import { normalizeSlides } from '../server/normalize-slides.mjs'

function page(id, elements, layoutPageId = 'layout-blank') {
  return { id, layoutPageId, elements, elementOrder: Object.keys(elements) }
}

test('normalizes negative slide extents while preserving occupied bounds and orientation', () => {
  const snapshot = {
    slideOrder: ['GX5wTC'],
    slides: {
      GX5wTC: page('GX5wTC', {
        VpGNav: {
          id: 'VpGNav',
          transform: {
            left: 120,
            top: 200,
            width: -25,
            height: -43.125713333333294,
            flipX: true,
            flipY: false,
          },
        },
      }),
    },
  }

  assert.equal(normalizeSlides(snapshot), snapshot)
  const transform = snapshot.slides.GX5wTC.elements.VpGNav.transform
  assert.equal(transform.left, 95)
  assert.ok(Math.abs(transform.top - 156.87428666666672) < 1e-9)
  assert.equal(transform.width, 25)
  assert.equal(transform.height, 43.125713333333294)
  assert.equal(transform.flipX, false)
  assert.equal(transform.flipY, true)
})

test('normalizes all slide page collections and repairs stale layout references', () => {
  const snapshot = {
    slideOrder: ['slide-1'],
    slides: {
      'slide-1': page('slide-1', { shape: { transform: { left: 0, top: 10, width: 20, height: -5 } } }, 'missing'),
    },
    layoutPageOrder: ['layout-blank'],
    layoutPages: {
      'layout-blank': page('layout-blank', { placeholder: { transform: { left: 4, top: 0, width: -4, height: 10 } } }),
    },
    masterPages: {
      master: page('master', { logo: { transform: { left: 0, top: 0, width: 8, height: -2, flipY: true } } }),
    },
  }

  normalizeSlides({ slide: snapshot })

  assert.equal(snapshot.slides['slide-1'].layoutPageId, 'layout-blank')
  assert.deepEqual(snapshot.slides['slide-1'].elements.shape.transform, {
    left: 0,
    top: 5,
    width: 20,
    height: 5,
    flipY: true,
  })
  assert.deepEqual(snapshot.layoutPages['layout-blank'].elements.placeholder.transform, {
    left: 0,
    top: 0,
    width: 4,
    height: 10,
    flipX: true,
  })
  assert.deepEqual(snapshot.masterPages.master.elements.logo.transform, {
    left: 0,
    top: -2,
    width: 8,
    height: 2,
    flipY: false,
  })
})
