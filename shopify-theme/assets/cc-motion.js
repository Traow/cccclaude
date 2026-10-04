/* cc-motion — scroll reveals, split text, parallax, tilt, magnetic buttons,
   marquee velocity and the featured product picker. No dependencies. */
(() => {
  const reduce = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const finePointer = window.matchMedia('(any-pointer: fine)').matches;
  const designMode = document.documentElement.classList.contains('shopify-design-mode');

  /* ---------- Intro curtain: remember it was shown ---------- */
  try {
    if (document.documentElement.classList.contains('cc-intro')) {
      sessionStorage.setItem('cc-intro-seen', '1');
      setTimeout(() => document.documentElement.classList.remove('cc-intro'), 2400);
    }
  } catch (e) {}

  /* ---------- Reveal observer ---------- */
  const io =
    'IntersectionObserver' in window
      ? new IntersectionObserver(
          (entries) => {
            for (const entry of entries) {
              if (!entry.isIntersecting) continue;
              entry.target.classList.add('is-in');
              io.unobserve(entry.target);
            }
          },
          { rootMargin: '0px 0px -8% 0px', threshold: 0.12 }
        )
      : null;

  const observe = (el) => {
    if (!io || reduce || designMode) {
      el.classList.add('is-in');
      return;
    }
    io.observe(el);
  };

  /* ---------- Split text into lines/words/chars ---------- */
  function split(el) {
    if (el.dataset.ccSplitDone) return;
    el.dataset.ccSplitDone = '1';
    const lines = el.innerHTML.split(/<br\s*\/?>/i);
    let c = 0;
    el.innerHTML = lines
      .map((line) => {
        const tmp = document.createElement('div');
        tmp.innerHTML = line;
        const words = tmp.textContent.trim().split(/\s+/).filter(Boolean);
        const html = words
          .map(
            (w) =>
              `<span class="cc-split-word">${[...w]
                .map((ch) => `<span class="cc-split-char" style="--cc-c:${c++}">${ch.replace(/[&<>]/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[m])}</span>`)
                .join('')}</span>`
          )
          .join(' ');
        return `<span class="cc-split-line">${html}</span>`;
      })
      .join('');
  }

  /* ---------- Parallax on scroll ---------- */
  const parallaxEls = new Set();
  let lastY = window.scrollY;
  let velocity = 0;
  let ticking = false;

  function onScroll() {
    if (ticking) return;
    ticking = true;
    requestAnimationFrame(frame);
  }

  function frame() {
    ticking = false;
    const y = window.scrollY;
    velocity = y - lastY;
    lastY = y;

    const max = document.documentElement.scrollHeight - window.innerHeight;
    document.documentElement.style.setProperty('--cc-progress', max > 0 ? (y / max).toFixed(4) : 0);

    if (reduce) return;
    const vh = window.innerHeight;
    for (const el of parallaxEls) {
      const rect = el.getBoundingClientRect();
      if (rect.bottom < -200 || rect.top > vh + 200) continue;
      const speed = parseFloat(el.dataset.ccParallax) || 0.15;
      const offset = (rect.top + rect.height / 2 - vh / 2) * -speed;
      el.style.setProperty('--cc-py', `${offset.toFixed(1)}px`);
    }

    const skew = Math.max(-8, Math.min(8, velocity * 0.25));
    document.documentElement.style.setProperty('--cc-skew', `${skew.toFixed(2)}deg`);
    document.documentElement.style.setProperty('--cc-boost', Math.min(4, 1 + Math.abs(velocity) / 18).toFixed(2));
    clearTimeout(frame.reset);
    frame.reset = setTimeout(() => {
      document.documentElement.style.setProperty('--cc-skew', '0deg');
      document.documentElement.style.setProperty('--cc-boost', '1');
    }, 120);
  }

  window.addEventListener('scroll', onScroll, { passive: true });

  /* ---------- Pointer depth (hero) ---------- */
  function initDepth(root) {
    if (reduce || !finePointer) return;
    root.querySelectorAll('[data-cc-depth-scene]').forEach((scene) => {
      if (scene.dataset.ccBound) return;
      scene.dataset.ccBound = '1';
      let raf;
      scene.addEventListener('pointermove', (e) => {
        const r = scene.getBoundingClientRect();
        const x = (e.clientX - r.left) / r.width - 0.5;
        const y = (e.clientY - r.top) / r.height - 0.5;
        cancelAnimationFrame(raf);
        raf = requestAnimationFrame(() => {
          scene.style.setProperty('--cc-mx', x.toFixed(3));
          scene.style.setProperty('--cc-my', y.toFixed(3));
          scene.style.setProperty('--cc-gx', `${((x + 0.5) * 100).toFixed(1)}%`);
          scene.style.setProperty('--cc-gy', `${((y + 0.5) * 100).toFixed(1)}%`);
        });
      });
      scene.addEventListener('pointerleave', () => {
        scene.style.setProperty('--cc-mx', 0);
        scene.style.setProperty('--cc-my', 0);
      });
    });
  }

  /* ---------- 3D tilt ---------- */
  function initTilt(root) {
    if (reduce || !finePointer) return;
    root.querySelectorAll('[data-cc-tilt]').forEach((el) => {
      if (el.dataset.ccBound) return;
      el.dataset.ccBound = '1';
      el.addEventListener('pointermove', (e) => {
        const r = el.getBoundingClientRect();
        const x = (e.clientX - r.left) / r.width - 0.5;
        const y = (e.clientY - r.top) / r.height - 0.5;
        el.style.setProperty('--cc-rx', `${(-y * 10).toFixed(2)}deg`);
        el.style.setProperty('--cc-ry', `${(x * 12).toFixed(2)}deg`);
        el.style.setProperty('--cc-gx', `${((x + 0.5) * 100).toFixed(1)}%`);
        el.style.setProperty('--cc-gy', `${((y + 0.5) * 100).toFixed(1)}%`);
      });
      el.addEventListener('pointerleave', () => {
        el.style.setProperty('--cc-rx', '0deg');
        el.style.setProperty('--cc-ry', '0deg');
      });
    });
  }

  /* ---------- Magnetic buttons ---------- */
  function initMagnetic(root) {
    if (reduce || !finePointer) return;
    root.querySelectorAll('[data-cc-magnetic]').forEach((el) => {
      if (el.dataset.ccBound) return;
      el.dataset.ccBound = '1';
      el.addEventListener('pointermove', (e) => {
        const r = el.getBoundingClientRect();
        const x = e.clientX - r.left - r.width / 2;
        const y = e.clientY - r.top - r.height / 2;
        el.style.transform = `translate(${x * 0.25}px, ${y * 0.35}px)`;
      });
      el.addEventListener('pointerleave', () => {
        el.style.transform = '';
      });
    });
  }

  /* ---------- Featured product picker ---------- */
  function initShowcase(root) {
    root.querySelectorAll('[data-cc-showcase]').forEach((section) => {
      if (section.dataset.ccBound) return;
      section.dataset.ccBound = '1';

      const dataEl = section.querySelector('[data-cc-variants]');
      if (!dataEl) return;
      const variants = JSON.parse(dataEl.textContent);
      const idInput = section.querySelector('input[name="id"]');
      const submit = section.querySelector('[data-cc-submit]');
      const priceEl = section.querySelector('[data-cc-price]');
      const colorLabel = section.querySelector('[data-cc-color-label]');
      const state = {
        color: section.querySelector('[data-cc-color].is-active')?.dataset.ccColor,
        size: section.querySelector('[data-cc-size].is-active')?.dataset.ccSize,
      };

      function update() {
        const v = variants.find(
          (v) => (!state.color || v.color === state.color) && (!state.size || v.size === state.size)
        );
        section.querySelectorAll('[data-cc-slide]').forEach((s) => {
          s.classList.toggle('is-active', s.dataset.ccSlide === state.color);
        });
        section.style.setProperty('--cc-tint', section.querySelector(`[data-cc-color="${CSS.escape(state.color || '')}"]`)?.dataset.ccTint || 'transparent');
        if (colorLabel && state.color) colorLabel.textContent = state.color;

        section.querySelectorAll('[data-cc-size]').forEach((b) => {
          const match = variants.find((v) => v.color === state.color && v.size === b.dataset.ccSize);
          b.classList.toggle('is-unavailable', !match || !match.available);
        });

        if (!v) return;
        idInput.value = v.id;
        if (priceEl) priceEl.textContent = v.price;
        submit.disabled = !v.available;
        submit.querySelector('[data-cc-submit-label]').textContent = v.available
          ? submit.dataset.labelAdd
          : submit.dataset.labelSoldOut;
      }

      section.addEventListener('click', (e) => {
        const color = e.target.closest('[data-cc-color]');
        const size = e.target.closest('[data-cc-size]');
        if (color) {
          state.color = color.dataset.ccColor;
          section.querySelectorAll('[data-cc-color]').forEach((b) => {
            b.classList.toggle('is-active', b === color);
            b.setAttribute('aria-pressed', b === color);
          });
          update();
        }
        if (size) {
          state.size = size.dataset.ccSize;
          section.querySelectorAll('[data-cc-size]').forEach((b) => {
            b.classList.toggle('is-active', b === size);
            b.setAttribute('aria-pressed', b === size);
          });
          update();
        }
      });

      update();
    });
  }

  /* ---------- Colorway panels: tap to expand on touch ---------- */
  function initColorways(root) {
    root.querySelectorAll('[data-cc-colorways]').forEach((wrap) => {
      if (wrap.dataset.ccBound) return;
      wrap.dataset.ccBound = '1';
      const panels = [...wrap.querySelectorAll('[data-cc-panel]')];
      const activate = (p) => panels.forEach((x) => x.classList.toggle('is-open', x === p));
      panels.forEach((p) => {
        p.addEventListener('pointerenter', () => activate(p));
        p.addEventListener('focusin', () => activate(p));
      });
    });
  }

  /* ---------- Boot ---------- */
  function init(root = document) {
    root.querySelectorAll('[data-cc-split]').forEach((el) => {
      split(el);
      observe(el);
    });
    root.querySelectorAll('[data-cc-reveal]').forEach((el, i) => {
      if (!el.style.getPropertyValue('--cc-i') && el.dataset.ccStagger !== undefined) {
        el.style.setProperty('--cc-i', i % 6);
      }
      observe(el);
    });
    root.querySelectorAll('[data-cc-parallax]').forEach((el) => parallaxEls.add(el));

    // Product cards rendered by the theme's own sections
    root.querySelectorAll('main product-card, main .product-card').forEach((card) => {
      const item = card.closest('li') || card;
      if (item.classList.contains('cc-card-reveal')) return;
      const siblings = item.parentElement ? [...item.parentElement.children] : [item];
      item.classList.add('cc-card-reveal');
      item.style.setProperty('--cc-i', siblings.indexOf(item) % 4);
      observe(item);
    });

    // Headings of theme sections that aren't ours
    root.querySelectorAll('main .shopify-section:not([id*="cc_"]) :is(h1, h2, h3)').forEach((h) => {
      if (h.closest('.cc-section, product-card, .product-card') || h.hasAttribute('data-cc-reveal')) return;
      h.setAttribute('data-cc-reveal', '');
      observe(h);
    });

    initDepth(root);
    initTilt(root);
    initMagnetic(root);
    initShowcase(root);
    initColorways(root);
    frame();
  }

  if (!document.querySelector('.cc-progress')) {
    const bar = document.createElement('div');
    bar.className = 'cc-progress';
    bar.setAttribute('aria-hidden', 'true');
    document.body.appendChild(bar);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => init());
  } else {
    init();
  }

  // Theme editor & dynamically rendered sections (filters, recommendations)
  document.addEventListener('shopify:section:load', (e) => init(e.target));
  new MutationObserver((muts) => {
    for (const m of muts) {
      for (const n of m.addedNodes) {
        if (n.nodeType === 1 && (n.matches?.('product-card, .product-card, li') || n.querySelector?.('product-card, .product-card'))) {
          init(n.parentElement || document);
          return;
        }
      }
    }
  }).observe(document.querySelector('main') || document.body, { childList: true, subtree: true });
})();
