/* Desktop (Tauri) integration. Loaded after app.js; a no-op in a plain
   browser. Gives real file semantics: Open/Save As via native dialogs, Save
   writes back to the current path, and the open file is watched so external
   edits (another machine via OneDrive, Excel, etc.) auto-reload. */
(function () {
  'use strict';
  if (!window.__TAURI__) return;

  var dialog = window.__TAURI__.dialog;
  var fs = window.__TAURI__.fs;

  var XLSX_FILTER = [{ name: 'Excel workbook', extensions: ['xlsx'] }];
  var currentPath = null;   // absolute path of the open document (null = unsaved)
  var unwatch = null;       // stops the active directory watcher
  var suppressUntil = 0;    // ignore watch events this soon after our own write
  var reloading = false;

  function app() { return window.HeadwayApp; }
  function basename(p) { return String(p).replace(/^.*[\\/]/, ''); }
  function dirname(p) {
    var m = String(p).match(/^(.*)[\\/][^\\/]+$/);
    return m ? m[1] : p;
  }
  function samePath(a, b) {
    // watcher events may use the other separator style on Windows
    return String(a).replace(/\\/g, '/') === String(b).replace(/\\/g, '/');
  }

  function markTitle() {
    document.title = currentPath
      ? basename(currentPath) + ' — Headway' : 'Headway — Roadmap Planner';
    var t = document.getElementById('docTitle');
    if (t) t.title = currentPath || '';
  }

  function setPath(p) {
    currentPath = p;
    markTitle();
    rewatch();
  }

  // Watch the parent directory, not the file: editors and sync clients
  // (OneDrive included) replace files by rename, which kills a file watch.
  function rewatch() {
    if (unwatch) { try { unwatch(); } catch (e) { /* already gone */ } unwatch = null; }
    if (!currentPath) return;
    fs.watch(dirname(currentPath), function (event) {
      var paths = (event && event.paths) || [];
      var hit = paths.some(function (p) { return samePath(p, currentPath); });
      if (!hit || reloading || Date.now() < suppressUntil) return;
      reloadFromDisk();
    }, { delayMs: 800 }).then(function (un) {
      unwatch = un;
    }).catch(function (err) {
      app().toast('Could not watch for external changes: ' + err.message, 'err');
    });
  }

  function reloadFromDisk() {
    var p = currentPath;
    reloading = true;
    fs.readFile(p).then(function (bytes) {
      if (!bytes.length) throw new Error('file is empty (still syncing?)');
      return app().loadBuffer(bytes.buffer, basename(p), true);
    }).then(function () {
      app().toast('Reloaded “' + basename(p) + '” — changed on disk');
    }).catch(function (err) {
      app().toast('Auto-reload failed: ' + err.message, 'err');
    }).finally(function () {
      reloading = false;
    });
  }

  window.HeadwayDesktop = {
    // native open dialog → load → remember + watch the path
    openDialog: function () {
      dialog.open({ multiple: false, filters: XLSX_FILTER }).then(function (p) {
        if (!p) return;
        fs.readFile(p).then(function (bytes) {
          return app().loadBuffer(bytes.buffer, basename(p));
        }).then(function () {
          setPath(p);
        }).catch(function (err) {
          app().toast('Could not open: ' + err.message, 'err');
        });
      });
    },

    // write the workbook; dialog only when there's no path yet (or Save As).
    // Resolves to the saved path, or null if the user canceled.
    saveBlob: function (blob, suggestedName, forceDialog) {
      var target = (currentPath && !forceDialog)
        ? Promise.resolve(currentPath)
        : dialog.save({ defaultPath: suggestedName, filters: XLSX_FILTER });
      return target.then(function (p) {
        if (!p) return null;
        if (!/\.xlsx$/i.test(p)) p += '.xlsx';
        return blob.arrayBuffer().then(function (buf) {
          suppressUntil = Date.now() + 3000;
          return fs.writeFile(p, new Uint8Array(buf));
        }).then(function () {
          suppressUntil = Date.now() + 3000;
          if (!samePath(p, currentPath || '')) setPath(p);
          return p;
        });
      });
    },

    currentPath: function () { return currentPath; },
    basename: basename
  };

  // ------------------------------------------------------- window chrome
  // The header doubles as the titlebar. macOS: native titlebar hidden with
  // overlay traffic lights (see tauri.macos.conf.json) — the header gets
  // left padding to clear them. Windows: frameless window with Windows 11
  // style caption buttons on the right. Empty header space drags the window
  // (and double-click zooms/maximizes, per each platform's convention).
  (function chrome() {
    var isMac = navigator.platform.indexOf('Mac') === 0;
    var topbar = document.getElementById('topbar');
    if (!topbar) return;

    // dragging works from header background and passive elements (brand
    // mark, gaps around the view tabs) — never from actual controls
    Array.prototype.forEach.call(
      document.querySelectorAll('#topbar, .tb-brand, .tb-mark, .tb-mark span'),
      function (el) { el.setAttribute('data-tauri-drag-region', ''); }
    );
    document.body.classList.add(isMac ? 'chrome-mac' : 'chrome-win');
    if (isMac) return;

    // Windows caption buttons
    var win = window.__TAURI__.window.getCurrentWindow();
    var GLYPH = {
      min: '<svg viewBox="0 0 10 10" width="10" height="10"><path d="M0 5h10" stroke="currentColor" fill="none"/></svg>',
      max: '<svg viewBox="0 0 10 10" width="10" height="10"><rect x=".5" y=".5" width="9" height="9" stroke="currentColor" fill="none"/></svg>',
      restore: '<svg viewBox="0 0 10 10" width="10" height="10"><rect x=".5" y="2.5" width="7" height="7" stroke="currentColor" fill="none"/><path d="M2.5 2.5v-2h7v7h-2" stroke="currentColor" fill="none"/></svg>',
      close: '<svg viewBox="0 0 10 10" width="10" height="10"><path d="M0 0l10 10M10 0L0 10" stroke="currentColor" fill="none"/></svg>'
    };
    var bar = document.createElement('div');
    bar.className = 'win-caption';
    bar.innerHTML =
      '<button class="cap-min" title="Minimize" aria-label="Minimize">' + GLYPH.min + '</button>' +
      '<button class="cap-max" title="Maximize" aria-label="Maximize">' + GLYPH.max + '</button>' +
      '<button class="cap-close" title="Close" aria-label="Close">' + GLYPH.close + '</button>';
    topbar.appendChild(bar);

    function syncMaxGlyph() {
      win.isMaximized().then(function (max) {
        var b = bar.querySelector('.cap-max');
        b.innerHTML = max ? GLYPH.restore : GLYPH.max;
        b.title = max ? 'Restore' : 'Maximize';
        b.setAttribute('aria-label', b.title);
      });
    }
    bar.querySelector('.cap-min').addEventListener('click', function () { win.minimize(); });
    bar.querySelector('.cap-max').addEventListener('click', function () {
      win.toggleMaximize().then(syncMaxGlyph);
    });
    bar.querySelector('.cap-close').addEventListener('click', function () { win.close(); });
    win.onResized(syncMaxGlyph);
    syncMaxGlyph();
  })();

  // ------------------------------------------------------- macOS menu bar
  // Mirror the in-app File / Edit / View menus into the system menu bar and
  // hide the in-window menu buttons. Rebuilt (debounced, only when the spec
  // actually changes) so checkmarks and enabled states follow the app state.
  (function nativeMenu() {
    var menu = window.__TAURI__.menu;
    var isMac = navigator.platform.indexOf('Mac') === 0;
    if (!isMac || !menu) return;

    var menusNav = document.getElementById('menus');
    if (menusNav) menusNav.style.display = 'none';

    function accel(kbd) {
      if (!kbd || /scroll/.test(kbd)) return undefined;
      var mods = [];
      if (kbd.indexOf('⇧') >= 0) mods.push('Shift');
      if (kbd.indexOf('⌥') >= 0) mods.push('Alt');
      if (kbd.indexOf('⌘') >= 0) mods.push('CmdOrCtrl');
      var key = kbd.replace(/[⇧⌥⌘]/g, '');
      return key ? mods.concat(key.toUpperCase()).join('+') : undefined;
    }

    function fireItem(m) {
      // ⌘Z in a text field should stay the field's own undo, not the app's
      if (m.label === 'Undo' || m.label === 'Redo') {
        var ae = document.activeElement;
        if (ae && (/INPUT|TEXTAREA|SELECT/.test(ae.tagName) || ae.isContentEditable)) {
          document.execCommand(m.label.toLowerCase());
          return;
        }
      }
      m.fn();
    }

    function toNative(items) {
      return Promise.all(items.map(function (m) {
        if (m.sep) return menu.PredefinedMenuItem.new({ item: 'Separator' });
        var opts = {
          text: m.label,
          enabled: !m.disabled,
          accelerator: accel(m.kbd),
          action: function () { fireItem(m); }
        };
        return ('checked' in m)
          ? menu.CheckMenuItem.new(Object.assign({ checked: !!m.checked }, opts))
          : menu.MenuItem.new(opts);
      }));
    }

    function predefined(names) {
      return Promise.all(names.map(function (n) {
        return menu.PredefinedMenuItem.new({ item: n });
      }));
    }

    // app Edit menu + the system clipboard block (without it, ⌘C/⌘V would
    // stop working in the webview once the default menu is replaced)
    function buildEdit(items) {
      var cut = items.findIndex(function (m) { return m.sep; });
      return Promise.all([
        toNative(items.slice(0, cut)),
        predefined(['Separator', 'Cut', 'Copy', 'Paste', 'SelectAll']),
        toNative(items.slice(cut))
      ]).then(function (parts) {
        return menu.Submenu.new({ text: 'Edit', items: parts[0].concat(parts[1], parts[2]) });
      });
    }

    function buildMenu() {
      var M = window.HeadwayApp.menuItems;
      return Promise.all([
        predefined(['Hide', 'HideOthers', 'ShowAll', 'Separator', 'Quit']).then(function (items) {
          return menu.Submenu.new({ text: 'Headway', items: items });
        }),
        toNative(M('file')).then(function (items) {
          return menu.Submenu.new({ text: 'File', items: items });
        }),
        buildEdit(M('edit')),
        toNative(M('view')).then(function (items) {
          return menu.Submenu.new({ text: 'View', items: items });
        }),
        predefined(['Minimize', 'Maximize', 'Fullscreen', 'Separator', 'CloseWindow']).then(function (items) {
          return menu.Submenu.new({ text: 'Window', items: items });
        })
      ]).then(function (subs) {
        return menu.Menu.new({ items: subs });
      }).then(function (m) {
        return m.setAsAppMenu();
      });
    }

    var sig = '', building = false, dirty = false, timer = null;
    function specSig() {
      var M = window.HeadwayApp.menuItems;
      return JSON.stringify(['file', 'edit', 'view'].map(function (n) {
        return M(n).map(function (m) {
          return m.sep ? '-' : [m.label, !!m.disabled, !!m.checked, m.kbd || ''];
        });
      }));
    }
    function syncNow() {
      var s;
      try { s = specSig(); } catch (e) { return; }
      if (s === sig) return;
      if (building) { dirty = true; return; }
      sig = s; building = true;
      buildMenu().catch(function (err) {
        console.warn('menu build failed', err);
        sig = ''; // retry on the next sync
      }).finally(function () {
        building = false;
        if (dirty) { dirty = false; syncNow(); }
      });
    }

    window.HeadwayDesktop.syncMenu = function () {
      clearTimeout(timer);
      timer = setTimeout(syncNow, 200);
    };
    window.HeadwayDesktop.syncMenu();
  })();
})();
