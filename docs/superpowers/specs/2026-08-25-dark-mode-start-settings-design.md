# Dark mode, start page, disk-backed projects, settings redesign

Date: 2026-08-25 · Status: approved for implementation (user directive, autonomous session)

## Goals

1. **Dark mode** — three theme choices: System (follow OS), Light, Dark. Switchable from
   the app's menus and from a right-click context menu on app chrome, and from Personal
   settings.
2. **Start page** — the app opens to a start experience: pick a recent file, create a new
   project, open a file, or adjust personal settings. The editor only appears once a
   project is open.
3. **Disk-backed projects** — every new project is written to disk at creation time.
   Desktop: native Save dialog before the project exists. Browser: an immediate .xlsx
   download (closest a web page can get to "on disk").
4. **Setup redesign** — the Setup view becomes a settings surface with a vertical tab
   rail, split into **Project** settings (stored in the document) and **Personal**
   settings (stored per machine).
5. **macOS fullscreen** — native fullscreen hides the traffic lights, so the header's
   86 px left inset must collapse and the title area shift left.

## Architecture

### Theme system

- Preference stored in its own localStorage key `headway-theme-v1` (`system` | `light` |
  `dark`). Deliberately **not** part of `uiSnapshot()` — theme is personal and must not
  travel inside a shared .xlsx.
- JS theme manager (in app.js, loaded before first render): resolves the effective theme
  (`system` → `matchMedia('(prefers-color-scheme: dark)')`), stamps
  `document.documentElement.dataset.theme = 'light' | 'dark'`, and listens for OS theme
  changes while in `system`.
- CSS: all remaining hardcoded surface/ink colors in `css/app.css` move into `:root`
  tokens; a single `html[data-theme="dark"]` block redefines the palette. Workstream/bar
  colors and the PNG exporter's palette stay as-is (exports are always light for
  shareability).
- Theme entries appear in: View menu (and its macOS native mirror), a new app-chrome
  right-click context menu (topbar / setup / start page background), and Personal
  settings → Appearance.

### Start page

- New `#startPage` section in index.html; `body.start` hides `#topbar`/`#main` and shows
  it. App boots into the start page; `enterEditor()` removes it.
- Contents: brand header, primary actions (**New project**, **Open .xlsx…**), a recent
  projects list, and a settings (gear) button opening Personal settings as a modal.
- Desktop recents: localStorage `headway-recents-v1`, `[{path, title, at}]`, capped at 12,
  upserted whenever a file is opened or saved (`setPath` in desktop.js). Rows show title,
  path, relative time; an ✕ removes an entry; opening a missing file removes it and
  toasts. Replaces the old auto-reopen-last-file boot (LAST_KEY logic retired).
- Browser: no paths, so recents shows a single **Continue last session** card when
  `headway-v1` state exists in localStorage.
- Reload ergonomics: `sessionStorage['headway-in-editor']` — a refresh of a tab already
  in the editor skips the start page and restores the local session.
- Window chrome: the start page top strip is a Tauri drag region; on Windows the caption
  buttons also mount on the start page (fixed, top-right); on macOS content clears the
  traffic lights.

### Disk-backed project creation

- **New project** modal asks for a name → desktop: Save dialog (default `<name>.xlsx`),
  the blank state is exported and written before the editor opens; cancel = no project.
  Browser: blank state is created and its .xlsx download fires immediately, then the
  editor opens.
- File → "New roadmap" routes through the same flow. Auto-save (already default-on) keeps
  the disk copy current from the first edit on desktop since a path always exists.

### Setup redesign

- `renderSetup()` renders a two-column layout: a vertical rail (sticky) and a content
  pane. Rail groups:
  - **PROJECT**: Timeline (start/end, sprint numbering), Phases, Workstreams,
    Team (team types, capacity toggle), Sizing, Holidays.
  - **PERSONAL**: Appearance (theme picker), Preferences (snap, dependency arrows,
    critical path, capacity row, grouping, auto-order, auto-save).
- Active tab kept in the UI snapshot (`setupTab`); Resources "manage" jumps to Team.
- Existing delegated change/click handlers keep working (same element ids/data attrs);
  Personal controls write the same variables the View menu writes and re-render.
- Personal settings are also openable as a modal from the start page (same field
  builders, no document required).

### macOS fullscreen

- desktop.js: on resize events, `getCurrentWindow().isFullscreen()` toggles
  `body.fullscreen`; CSS `body.chrome-mac.fullscreen #topbar { padding-left: 14px }`.

## Error handling

- Recents open failure (moved/deleted file): toast + drop the entry, stay on start page.
- Save-dialog cancel during New project: modal stays open, nothing created.
- localStorage unavailable: start page still works; recents/continue simply absent.

## Testing

- `tests/smoke.test.js` adapts to the boot flow: assert the start page shows, then enter
  via Continue-last-session and run the existing suite. New assertions: theme stamping,
  setup rail tabs render, personal tab toggles.
- `tests/core.test.js` unaffected (no core changes).
