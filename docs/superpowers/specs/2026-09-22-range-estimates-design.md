# Range estimates — design

## Goals

Some teams estimate a low and a high per feature (and per story) rather than one number, and want the
plan to show that spread. Headway sizes a row with one planned duration. This adds a **project-level**
choice — Setup → Sizing → Estimates — between the existing single estimate and a range estimate, so a
document either looks exactly as today or carries two numbers per row and shows the min/max span.

## Architecture

- `meta.estimateMode` `'single' | 'range'` (default single), `meta.estimateUnit` `'days' | 'points' |
  'hours'` (default days), `meta.daysPerUnit` (default per unit: 1, 1, 1/8), `meta.estimateBasis`
  `'high' | 'low'` (default high). All normalized in `normalizeState`; whole-key LWW in the bundle's meta
  shard like every other project setting.
- Items and stories gain `estLow` / `estHigh` (numbers in the project unit, quarter steps, null = not
  estimated; a reversed pair is swapped). Plain fields, so Excel state, the AI apply path and the shared
  bundle carry them with no format change.
- `RM.estRange(state, row)` → working days `{low, high, planned}` (a missing side falls back to the
  planned duration); `RM.rangeSpans` → grid ends via `stretchSpan`; `RM.basisDays` → what the planned
  bar should carry. `RM.estToDays` applies the unit rate.
- Planning: a `.bar-range` band (hatched, dashed border, low marker) rides behind a scheduled bar when a
  range is set and differs from the planned span; only in range mode. Tooltip carries the range.
- Panel: Low / High inputs (`data-f` / `data-stf` `estLow` `estHigh`) under the schedule block, labelled
  with the project unit. Editing either re-derives the planned duration from the basis (scheduled rows
  through `stretchSpan`); milestones keep zero.
- Excel: Roadmap sheet appends `Est. low` / `Est. high` after Tags; import reads them back by header.
- Validation: `EST_RANGE` (info) when a scheduled row's planned working days fall outside its range, range
  mode only.

## Error handling

Bad input parses to null (field cleared); a range with one side falls back to the planned duration for
the other; unit changes never rewrite stored numbers (they are in the unit the user typed) — they change
the conversion only.

## Testing

`tests/core.test.js` (`estimate range`): defaults, swap, story parsing, spans, basis, hours conversion,
EST_RANGE on/off, normalize round trip. `tests/smoke.test.js` (`range estimates`): Setup card, unit and
basis pickers, panel fields, planned bar follows basis, `.bar-range` drawn and hidden with the mode.
