/*
 * Headway Jira CSV export — a file for Jira Cloud's user-level importer
 * (work navigator → "Import issues from CSV"), which any user with the
 * Create work items + Make bulk changes permissions can run. That importer
 * has no Issue Id / Parent Id columns, so a row can only parent to an
 * issue that already exists: Parent / Blocked By carry the Jira keys typed
 * into Headway (item.jiraKey, story.jiraKey, state.epicJira) and stay
 * blank otherwise.
 *  - rows(state, opts) -> [{column: value}] (node-testable)
 *  - csv(state, opts)  -> CSV string (BOM, CRLF, RFC-4180 quoting)
 *  - fileName(state)   -> "<title>-jira.csv"
 * opts: features (bool), stories (bool), featureType, storyType.
 */
(function (root) {
  'use strict';

  var RM = root.RM || (typeof require !== 'undefined' ? require('./core.js') : null);
  var JR = {};

  JR.COLUMNS = ['Summary', 'Issue Type', 'Description', 'Parent', 'Labels',
    'Priority', 'Due Date', 'Start Date', 'End Date', 'Blocked By', 'Jira Key'];
  JR.DEFAULTS = { features: true, stories: false, featureType: 'Story', storyType: 'Sub-task' };

  function slug(s) {
    return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  }
  function labelList(parts) {
    return parts.filter(Boolean).join(' ');
  }
  function iso(meta, day) {
    return RM.fmtISO(RM.dayToDate(meta, day));
  }
  function section(title, html) {
    var t = RM.htmlToText(html);
    return t ? title + ':\n' + t : '';
  }
  function joinSections(parts) {
    return parts.filter(Boolean).join('\n\n');
  }

  JR.rows = function (state, opts) {
    var o = {};
    Object.keys(JR.DEFAULTS).forEach(function (k) {
      o[k] = opts && opts[k] != null && opts[k] !== '' ? opts[k] : JR.DEFAULTS[k];
    });
    var meta = state.meta;
    var phaseName = {};
    (state.phases || []).forEach(function (p) { phaseName[p.id] = p.name; });
    var keyByNum = {};
    state.items.forEach(function (it) { if (it.jiraKey) keyByNum[it.num] = it.jiraKey; });
    var out = [];

    state.items.forEach(function (it) {
      var phase = phaseName[it.phaseId] ? 'phase-' + slug(phaseName[it.phaseId]) : '';
      var ws = it.workstream ? 'ws-' + slug(it.workstream) : '';
      var sched = it.startDay != null && it.durDays != null;
      if (o.features) {
        var checklist = o.stories ? '' : it.stories.map(function (s) {
          return (s.done ? '[x] ' : '[ ] ') + s.title;
        }).join('\n');
        out.push({
          'Summary': it.feature,
          'Issue Type': o.featureType,
          'Description': joinSections([
            RM.htmlToText(it.description),
            checklist ? 'Stories:\n' + checklist : '',
            section('Enables', it.enables),
            section('Out of scope', it.outOfScope),
            section('External dependencies', it.extDeps),
            section('Notes', it.notes)
          ]),
          'Parent': (it.epic && state.epicJira[it.epic]) || '',
          'Labels': labelList([ws, phase, it.size ? 'size-' + slug(it.size) : '']),
          'Priority': it.priority || '',
          'Due Date': it.deadline || '',
          'Start Date': sched ? iso(meta, it.startDay) : '',
          'End Date': sched ? RM.fmtISO(RM.spanEndDate(meta, it.startDay, it.durDays)) : '',
          'Blocked By': it.deps.map(function (n) { return keyByNum[n]; }).filter(Boolean).join(' '),
          'Jira Key': it.jiraKey || ''
        });
      }
      if (o.stories) {
        it.stories.forEach(function (s) {
          var ssched = s.startDay != null && s.durDays > 0;
          out.push({
            'Summary': s.title,
            'Issue Type': o.storyType,
            'Description': joinSections([
              RM.htmlToText(s.description),
              section('Acceptance criteria', s.ac)
            ]),
            'Parent': it.jiraKey || '',
            'Labels': labelList(['feature-' + slug(it.feature), ws, phase]),
            'Priority': s.priority || '',
            'Due Date': s.deadline || '',
            'Start Date': ssched ? iso(meta, s.startDay) : '',
            'End Date': ssched ? RM.fmtISO(RM.spanEndDate(meta, s.startDay, s.durDays)) : '',
            'Blocked By': '',
            'Jira Key': s.jiraKey || ''
          });
        });
      }
    });
    return out;
  };

  function cell(v) {
    var s = v == null ? '' : String(v);
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  JR.csv = function (state, opts) {
    var lines = [JR.COLUMNS.map(cell).join(',')];
    JR.rows(state, opts).forEach(function (r) {
      lines.push(JR.COLUMNS.map(function (c) { return cell(r[c]); }).join(','));
    });
    return '\uFEFF' + lines.join('\r\n') + '\r\n';
  };

  JR.fileName = function (state) {
    return ((state.meta && state.meta.title) || 'roadmap') + '-jira.csv';
  };

  root.RM_JIRA = JR;
  if (typeof module !== 'undefined' && module.exports) module.exports = JR;
})(typeof window !== 'undefined' ? window : globalThis);
