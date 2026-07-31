// support.js
// Minimal runtime for <x-dc> / <x-import> decks.
// <x-dc> is just a semantic wrapper.
// <x-import component-from-global-scope="X" from="./file.js" width height hint-size>
//   dynamically loads "file.js", which is expected to register a global
//   `window[X]` object exposing a `mount(hostEl, options)` function, then
//   calls it, handing off the <section> children as slides.
(function () {
  'use strict';

  if (!customElements.get('x-dc')) {
    customElements.define('x-dc', class extends HTMLElement {});
  }

  if (!customElements.get('x-import')) {
    customElements.define(
      'x-import',
      class extends HTMLElement {
        connectedCallback() {
          if (this._dcMounted) return;
          this._dcMounted = true;

          const globalName = this.getAttribute('component-from-global-scope');
          const src = this.getAttribute('from');
          const width = parseInt(this.getAttribute('width'), 10) || 1920;
          const height = parseInt(this.getAttribute('height'), 10) || 1080;
          const hintSize = this.getAttribute('hint-size') || '100%,100%';

          const boot = () => {
            const impl = globalName ? window[globalName] : null;
            if (impl && typeof impl.mount === 'function') {
              impl.mount(this, { width: width, height: height, hintSize: hintSize });
            } else {
              console.error('[x-import] component "' + globalName + '" from "' + src + '" did not register a mount() function on window.');
            }
          };

          if (globalName && window[globalName] && typeof window[globalName].mount === 'function') {
            boot();
            return;
          }

          if (!src) {
            console.error('[x-import] missing "from" attribute.');
            return;
          }

          const script = document.createElement('script');
          script.src = src;
          script.async = true;
          script.onload = boot;
          script.onerror = () => console.error('[x-import] failed to load script: ' + src);
          document.head.appendChild(script);
        }
      }
    );
  }
})();
