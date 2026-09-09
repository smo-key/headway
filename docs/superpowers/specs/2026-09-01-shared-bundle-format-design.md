# Shared bundle format — multi-user editing over a synced folder

Date: 2026-09-01 · Status: proposed (implemented on the `shared-bundle` branch for review)

## Goals

1. **Concurrent editing over file sync** — several people edit one roadmap at the same time
   through an ordinary synced folder (OneDrive/SharePoint, Dropbox, iCloud). No server.
   Conflicts are confined to *the same field of the same item edited inside the same sync
   window*; everything else merges without loss.
2. **A readable document** — the roadmap becomes a folder of small JSON files, one per
   entity. A team can open the folder and see what is there when something looks wrong.
3. **Deterministic merge** — two machines applying the same two changes in either order
   arrive at the same document. No prompts, no "keep mine / keep theirs".
4. **Shared history** — who changed what, across every editor, with no shared write.
5. **Presence** — see who else is in the roadmap and which rows they are on. Advisory only:
   locks cannot be enforced over eventually-consistent sync, so the app warns rather than
   blocks.
6. **Alternative plans stay** — the existing Options feature (parked alternative documents)
   survives as sub-bundles; which plan a person is looking at is theirs, not shared.
7. **Excel becomes an export** — `.xlsx` is a report generated on demand and re-opens as a
   standalone document. Edits made in Excel do not flow back into a shared roadmap.

Not real-time: sync latency is seconds to minutes. Two people dragging the same bar inside
that window still race; the loser's value is replaced deterministically and both changes
are recorded in history.

## Architecture

### Bundle layout

A shared roadmap is a folder named `<Title>.headway`:

```
<Title>.headway/
  headway.json                 { format:'headway-bundle-v1', docId, title, createdAt,
                                 plans:[{id, name, createdAt, updatedAt, deleted?}] }
  plans/<planId>/              every plan is a sub-bundle — the "main" one included
    meta.json                  meta.* plus wsOrder, wsColors, epicIcons, teamTypes, scopeCols
    phases/<uid>.json
    items/<uid>.json           stories inline (owned by the item; merged by story id)
    team/<uid>.json
    costs/<uid>.json
  history/<userId>.jsonl       append-only, one writer per file
  presence/<userId>.json       heartbeat {name, planId, editing:[uids], ts}
```

`headway.json` and `meta.json` are low-churn shared config (whole-key last-write-wins).
Everything that changes often has its own file.

### Entity envelope and merge (`js/bundle.js`, namespace `RMBundle`)

Every shard is an envelope, not the bare entity:

```json
{ "id":"i…", "rev":7, "updatedAt":"2026-09-01T15:37:23.412Z", "updatedBy":"alex-k3f9a",
  "deleted":false,
  "fields":{ "…entity fields…" },
  "fieldsAt":{ "size":"…", "notes":"…", "stories.<sid>":"…", "deps+<id>":"…", "deps-<id>":"…" } }
```

- `fieldsAt[k]` is the timestamp of the last change to field `k`, carried forward when the
  field is unchanged (`RMBundle.wrap(entity, prevEnv, userId, nowIso)`).
- `RMBundle.mergeEntity(a, b)` — per plain field the newer `fieldsAt` wins; equal stamps fall
  to the greater `updatedBy` string so the result is deterministic and commutative.
  `stories` merge by story id, then per story field. `deps` is an OR-set: a dependency is
  present iff its newest `deps+` stamp beats its newest `deps-` stamp — a plain union would
  resurrect a dependency someone deleted. A tombstone (`deleted:true`, `deletedAt`) wins
  over any edit older than `deletedAt`. `rev` and `updatedAt` take the max.
- **Canonical string** — `RMBundle.canonicalize(obj)`: sorted keys, volatile in-memory keys
  stripped (`holdPos`, `_idx`, `leadDays`), ascii-escaped via `RM.asciiJson` (moved into
  core from excel.js). Canonical equality is the "is this my own write echoing back" test —
  the same content-compare that already suppresses OneDrive's xlsx container rewrites
  (commit 9e3f797), applied per shard.

### Identity

- **Dependencies reference ids, not numbers.** Today `deps` holds human `num`s and `num` is
  max+1, so two people creating items concurrently both mint the same number and the
  collision renumber on load silently repoints dependencies. `RM.migrateDepsToIds(state)`
  runs inside `normalizeState` (idempotent): resolvable number → id; unresolvable → moved to
  `depsText` as `#n`; self-deps dropped. Legacy blobs, fixtures and Excel imports migrate on
  load.
- `num` stays the human label and the Excel column. `RM.dedupeNums(state)` renumbers
  duplicates deterministically (`RM.uidTime(id)` — the time36 prefix of `RM.uid` — then id).
  Safe because nothing references `num` any more. The Roadmap sheet still shows `#num`;
  `excel.js depCellText` maps ids to numbers on write.

### Ordering

- Array position was the only order. Shards have no array, so every entity carries an
  `order` string: `RM.orderBetween(a, b)` (base-62 fractional index), `RM.orderAfterAll`,
  `RM.sortByOrder`. `normalizeState` assigns missing keys from index and sorts each array by
  `order`, so the in-memory arrays stay canonical and existing splice code keeps working.
- A manual reorder (`RM.placeItem`, `RM.movePhaseTo`, story/team drags) writes ONE key.
- **Auto-order is a view sort.** `RM.viewItems(state, {autoOrder})` returns a sorted copy for
  rendering and export; `state.items` is no longer rewritten after every drag. Persisted
  `order` only changes when a person reorders on purpose.

### History and user identity

- `localStorage['headway-user-v2'] = {name, id}`; `id = slug(name) + '-' + rand5`, minted once
  and never changed (renaming yourself only changes the display name). Migrates the
  `headway-user-v1` plain string. One id serves history and presence.
- `state.history` leaves the document. Each commit appends one line to
  `history/<userId>.jsonl` (`{t,u,userId,planId,label,n,d,x,tl}`); coalescing of same-label
  edits within 5 minutes rewrites the writer's own tail line. Single writer per file — never
  a conflict. `RMBundle.mergeHistory` concatenates every user's file, sorts by `t`, caps at
  `RM.HISTORY_MAX`. Version history shows the merged feed filtered to the active plan.
  "Who's editing?" still asks once; it then stamps the writer's own anonymous lines.

### Plans (formerly Options)

- `activePlanId` lives in the UI snapshot (`headway-ui-v1`) — per machine, never shared.
- Switch = load another sub-bundle locally (undo/redo cleared, as today). New = copy the
  current plan's shards into `plans/<newId>/` (new files only → no conflicts) and append to
  `headway.json.plans`; the 12-option cap goes. Rename = LWW on the plan entry. Delete =
  tombstone the entry (files are collected later); anyone viewing it is moved to another plan
  on their next load. Compare reads the other plan from disk and invalidates on its events.
- `RM.splitOptions(state)` turns a standalone document with N parked options into N+1
  plans; Convert writes them all.

### Write path

`afterChange()` → `scheduleBundleFlush()` (1.5 s debounce, same as xlsx auto-save) →
`flushBundle()`:

1. `RMBundle.diffEntities(lastCanon, state)` — only entities whose canonical string changed.
2. `HeadwayDesktop.flushShards(dir, planId, changes)` — per change, read-merge-write: read
   the shard; if the disk `rev` is newer than my base, `mergeEntity` first; stamp; record
   the canonical string as my echo baseline *before* writing; write `x.json.tmp` then
   `rename` to `x.json` (atomic replace; rename retried on sharing violations, then falls
   back to a direct write).
3. Append the history line. `docSaved = true`.
4. Deletes are tombstone envelopes. A live shard is never removed; tombstones older than
   30 days are collected on load.

### Read path and sync

- One recursive `fs.watch` on the bundle folder (800 ms coalescing). Paths are classified
  (`plans` / `meta` / `item` / `phase` / `team` / `cost` / `history` / `presence` / `tmp`);
  `.tmp` is ignored; only the active plan's entity events are applied live.
- Each changed shard is read with the existing retry ladder (`[1200, 3000, 8000]` ms — sync
  clients announce before the bytes land). Canonical equal to my baseline → my own echo,
  dropped. Otherwise `applyExternalEntity(kind, env)`: merge into the live entity and apply
  with the same mechanics `activateOption` uses (`state = next; validate; saveLocal();
  render()`) — **no undo push, no history entry, no name prompt**. Changes from one watch
  event are batched into one render. An entity being dragged is deferred to pointer-up.
- **Undo is rebased.** Undo/redo are whole-document snapshots; a peer change merged into
  live state would be reverted by the next undo and then written back, clobbering the peer.
  On every incoming entity the merged envelope is patched into each stored snapshot, so
  undo only ever reverts the local person's own edits.

### Conflict copies

When two machines write the same shard inside the sync window, OneDrive keeps both: the
server copy under the canonical name, the local one renamed `<name>-<COMPUTERNAME>.json`
(up to five). Ids contain hyphens, so siblings are detected by **envelope `id` ≠ file
stem**, never by parsing the filename. A sibling is merged into the canonical shard, the
result written atomically, and the sibling removed. The local machine's in-memory state
still holds its edit, so read-merge-write on the next flush covers the same case from the
other side.

### Presence

- Every 30 s, and on selection change (throttled to 2 s), the app writes
  `presence/<userId>.json` with `{name, planId, editing:[selectedId, dragging id], ts}`.
  Entries older than 90 s are stale. Own id and other plans are ignored.
- Rows carry small initials chips (`.avatar.sm`, the existing avatar style) for peers on
  that item, patched in place without a full render. A flush that touches an item a peer is
  on toasts "*Name* is also editing #*num*" (once per item per minute).
- On window close (Tauri `onCloseRequested`) and on closing the bundle the presence file is
  removed. Presence never goes through `commit` or history.

### Excel

- File → **Export .xlsx…** writes today's workbook from `RMBundle.exportableState(state)`
  (bundle markers stripped) and always asks for a location. Opening an `.xlsx` always yields
  a standalone document, so an export cannot be mistaken for the live roadmap.
- File → **Convert to shared folder…** on an open `.xlsx`: pick a parent folder,
  `RMBundle.migrateFromState` (envelopes, order keys, deps→id, N+1 plans, history split),
  `HeadwayDesktop.createBundle`, then open it. The `.xlsx` is left untouched.
- `reconcileVisibleEdits` remains for standalone workbooks only; the bundle path never
  reaches it.

### Import from Excel (add-only)

- File → **Import from Excel…** (bundle only, desktop only) merges a workbook *into* the open
  shared roadmap. The workbook is read once through `RMExcel.importWorkbook` and never adopted:
  `currentPath`, the watcher and the bundle session are untouched. A foreign (template-layout)
  workbook is accepted and flagged in the preview.
- **The shared roadmap wins; the import only adds.** `RM.planImport(state, incoming)` pairs
  rows and never guesses: items by `id`, then by `num` + normalized title, then by title alone
  when it is unique on both sides; stories inside a matched item by `id` then title; team and
  phases by `id` then name. Unpaired rows are adds — a colliding id is re-minted, the item takes
  the next `num`, its phase maps by name (else the first phase), its dependencies resolve to
  merged ids or fall to `depsText` as `#num`.
- A paired row gains a value only for a field that is **empty** in the roadmap and non-empty in
  the workbook (`RM.IMPORT_FILL_FIELDS`, story `description`/`ac`, an empty dependency list).
  Where both hold a value and differ the difference is counted and shown as "left alone";
  nothing is overwritten. A fill re-checks emptiness at apply time.
- Idempotent: importing the same workbook again plans zero adds and zero fills ("Nothing new
  to import").
- The preview modal (counts + the first new feature titles) applies through
  `commit('import from Excel', …)` → `RM.applyImport`, so the import is one undo step, one
  history line, and flushes only the shards it changed.

### Menus, recents, start page

- File (and its macOS native mirror): **New shared roadmap…**, **Open shared roadmap…**,
  **Convert to shared folder…** (xlsx only), Save becomes **Export .xlsx…** in a bundle,
  **Import from Excel…** appears (bundle only), Auto save is hidden. Desktop only — the browser build keeps its `.xlsx` flow.
- Recents entries gain `kind` (`'bundle' | 'xlsx'`; missing = xlsx); folders show a
  `folder-open` icon and open through `openBundle`. The Save button reads **Sync** /
  **Synced ✓** in a bundle. `{docKind, bundleDir, activePlanId}` ride in the UI snapshot so
  a reload resumes the same plan.

### Capabilities

`src-tauri/capabilities/default.json` gains `fs:allow-rename`, `fs:allow-remove`,
`fs:allow-stat`, `fs:allow-write-text-file` and `core:window:allow-destroy`. The existing
`fs:scope` `**` already covers reads, directory listing and `mkdir` on user paths. This also
fixes `renameTo`, which has called `fs.rename` without the permission.

## Error handling

- A shard that fails to parse (partial upload, placeholder not yet hydrated) is retried on
  the ladder, then skipped with a warning; the open never aborts on one bad file. Only a
  missing or invalid `headway.json` fails an open.
- `rename` sharing violation (sync client holding the file): three retries, then a direct
  write. A plugin permission error surfaces as "Headway is missing capability
  fs:allow-<command>" rather than a silent no-op.
- A failed flush keeps `docSaved = false` and toasts; the Sync button retries.
- Someone else deletes the plan you are viewing: you are moved to another plan on your next
  load with a toast.
- Multi-item operations (⌘-drag ripple, auto-schedule) are N independent writes, not a
  transaction; a peer may briefly see part of the set. Dependencies are by id so nothing
  dangles. Auto-schedule warns that it rewrites every item.
- localStorage unavailable: user id is regenerated per session (history lines still carry
  the display name).

## Testing

- `tests/core.test.js` — new sections: `bundle: canonicalize`, `wrap/fieldsAt`, `merge same
  field`, `merge different fields`, `merge stories by id`, `deps OR-set`, `tombstone vs
  edit`, `merge is commutative`, `diffEntities`, `fractional order`, `mergePlanList`,
  `deps by id`, `order keys`, `history lines`, `plans`. Existing deps-by-number assertions
  adapt to ids.
- `tests/desktop.test.js` (new) with `tests/fake-tauri.js` — an in-memory Tauri fs
  (`Map` of paths, `watch` the test can fire, `rename` as move, denied commands rejecting
  with the plugin's plain-string error). Cases: open → counts; corrupt shard skipped with a
  warning; flush writes only changed shards; tmp-then-rename observed; own echo ignored;
  peer change merged while my unsaved edit on another field survives; undo after a peer
  change does not revert the peer; conflict sibling absorbed and removed; missing capability
  names the permission.
- `tests/smoke.test.js` — browser mode shows none of the shared-roadmap menu items or the
  start-page button; a Tauri-stubbed window shows all three; a `kind:'bundle'` recent renders
  the folder icon.
- Migration round-trip: workbook → bundle → export → import → normalized deep-equal (minus
  history and options).
- Live, on a real synced folder with two machines: concurrent edits to different items both
  land; same item, different fields both land; same field resolves LWW with both lines in
  history; a forced `-COMPUTERNAME` sibling is absorbed; killing one app mid-flush leaves no
  torn shard on the other.

## Implementation notes (2026-09-08)

Where the build refined the text above:

- `meta.json` is an envelope too (id `meta`), so workstream order, colors and icons merge
  whole-key through the same `mergeEntity` instead of a second code path.
- `deps` is an OR-set with add/remove stamps (`deps+<id>` / `deps-<id>`) rather than a plain
  set union — a union would resurrect a dependency someone deleted.
- Local edits carry per-field stamps taken at **commit** time, not at flush or apply time, so
  an edit made offline and synced later loses to a newer peer edit of the same field.
- The undo/redo rebase patches only the fields the peer changed into each snapshot; earlier
  local edits to the same entity stay undoable.
- One flush chain: a flush requested while another is in flight runs after it; closing,
  switching plans and opening an `.xlsx` wait for it. A flush that completes after a newer
  peer envelope was applied keeps that baseline and re-sends the local field merged.
- Conflict siblings are detected by envelope `id` ≠ file stem, never by filename (ids contain
  hyphens). The rename→in-place fallback re-reads and re-merges before writing.
- Presence: heartbeat 30 s, selection throttle 2 s, stale after 90 s, nudge once per row per
  60 s. Names from other machines are escaped.
- History coalescing is per plan; Convert / New refuse a `<title>.headway` that already exists.
- Plan-folder garbage collection is not implemented (tombstoned entries stay, files stay).
- Tests: `tests/desktop.test.js` (fake plugin fs, 182) and `tests/wiring.test.js` (the real
  app in jsdom against that fs, 347) join core (563) and smoke (600) in `make test`.
