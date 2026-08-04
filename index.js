/**
 * SINGLE-FILE VERSION — ported from your desktop PU scraper (Selenium/Tkinter)
 * to run on-device via a WebView. Paste into Expo Snack's App.js.
 *
 * Dependencies (Snack should auto-detect these from the imports; add
 * manually via the dependencies panel if it doesn't):
 *   react-native-webview
 *   xlsx
 *   expo-file-system
 *   expo-sharing
 *
 * ─────────────────────────────────────────────────────────────────────────
 * WHAT WAS PORTED FROM YOUR DESKTOP SCRIPT, 1:1 in spirit:
 *   - extract_flat_tables_html  -> extractAllTablesJS (runs inside the page,
 *     same rowspan/colspan grid-building + header-depth detection)
 *   - filter_table               -> filterTable() (first row "1"/"1.0", else
 *     keep only if a TOTAL cell exists, else discard the table)
 *   - Satker column insertion, _fix_rr_headers, epur/terkontrak first-row
 *     numeric strip -> all ported as-is
 *   - Tab mapping (Excel "Tabs" sheet: Nama, t1..tN) -> TAB_MAPPING config
 *     object below; click-and-wait-for-new-table logic ported to a WebView
 *     async injected script (document.evaluate XPath, same match rule)
 *   - Table-index routing (Excel "TableMap" sheet + terkontrak/epur
 *     defaults) -> TABLE_INDEX_MAP config below
 *   - Rencana/Realisasi sublink-following -> ported, RN drives navigation
 *     since following a link means leaving the page (JS context resets)
 *   - Manual-export domains -> ported as a visible-WebView "I've exported —
 *     continue" handoff (see MANUAL_EXPORT_DOMAINS + note below)
 *   - Multi-sheet Excel output with sheet-name grouping rules -> ported,
 *     built with SheetJS instead of pandas/openpyxl
 *
 * WHAT WAS DROPPED / CHANGED, and why:
 *   - Fast Mode (CDP resource blocking) — no CDP access from a mobile
 *     WebView; there's no direct equivalent, so it's just not there. Doesn't
 *     affect correctness, only desktop-Chrome scrape speed.
 *   - Excel mapping files -> replaced with the TAB_MAPPING /
 *     TABLE_INDEX_MAP / MANUAL_EXPORT_DOMAINS JS objects below, since
 *     picking & parsing an .xlsx mapping file on-device is extra plumbing.
 *     Same shape as your Excel sheets, just edited as JS instead.
 *   - The "pages" input list -> a PAGES array below (title/url), same as
 *     your Excel/CSV columns, edited directly rather than picked from
 *     device storage.
 *   - Manual export "wait for .xlsx to land in Downloads and stabilize" ->
 *     react-native-webview's onFileDownload only fires on Android. On iOS
 *     there's no equivalent hook; if you need this on iOS, check whether
 *     the export button actually calls an API endpoint under the hood
 *     (Network tab) and fetch that directly instead.
 *   - Retry-if-file-open-in-Excel / archive-folder rotation / settings
 *     persistence -> dropped; not meaningful on mobile (no open-in-Excel
 *     lock, no persistent settings file in this quick version).
 * ─────────────────────────────────────────────────────────────────────────
 */

import React, { useRef, useState, useCallback } from 'react';
import {
  SafeAreaView, View, Button, Text, ActivityIndicator,
  ScrollView, StyleSheet,
} from 'react-native';
import { WebView } from 'react-native-webview';
import * as XLSX from 'xlsx';
// SDK 54 moved the classic FileSystem.* API (cacheDirectory, downloadAsync,
// readAsStringAsync, EncodingType, etc.) here — the default 'expo-file-system'
// export is now the new File/Directory object API instead.
import * as FileSystem from 'expo-file-system/legacy';
import * as Sharing from 'expo-sharing';

// ============================================================================
// CONFIG — edit these for your sites (equivalent to your Excel inputs)
// ============================================================================

// Equivalent of your pages Excel/CSV ("title", "url" columns).
// Loaded from your List.xlsx (62 pages, sipp.pu.go.id + iemon.pu.go.id).
const PAGES = [
  { title: 'SIPP', url: 'https://sipp.pu.go.id/monevpro/pages/table_monitoring_percepatan/halaman_table_paket_kontraktual_gabungan.php?thn=2026' },
  { title: 'Latest Progress', url: 'https://iemon.pu.go.id/progres_balai_bm_sat2.php?kd=balaibm08&kdtematik=' },
  { title: 'Nasional Prognosis Balai', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693824&thang=2026' },
  { title: 'Pakar Rutin Jalan', url: 'https://iemon.pu.go.id/padat_karya_sat_bm?kd=PK0402' },
  { title: 'Pakar Rutin Jembatan', url: 'https://iemon.pu.go.id/padat_karya_sat_bm?kd=PK0403' },
  { title: 'Prognosis P2JN Bali', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693847&thang=2026' },
  { title: 'Prognosis P2JN Jatim', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693831&thang=2026' },
  { title: 'Prognosis PJN I Bali', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693912&thang=2026' },
  { title: 'Prognosis PJN I Jatim', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693899&thang=2026' },
  { title: 'Prognosis PJN II Bali', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693913&thang=2026' },
  { title: 'Prognosis PJN II Jatim', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693900&thang=2026' },
  { title: 'Prognosis PJN III Bali', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693866&thang=2026' },
  { title: 'Prognosis PJN III Jatim', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693863&thang=2026' },
  { title: 'Prognosis PJN IV Jatim', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693901&thang=2026' },
  { title: 'Prognosis SKPD Bali', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693772&thang=2026' },
  { title: 'Prognosis SKPD Jatim', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693752&thang=2026' },
  { title: 'Prognosis Suramadu', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693840&thang=2026' },
  { title: 'Prognosis Balai', url: 'https://iemon.pu.go.id/pep_prognosis2020?sat=04693824&thang=2026' },
  { title: 'Progress Balai', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693824&thang=2026' },
  { title: 'Progress P2JN Bali', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693847&thang=2026' },
  { title: 'Progress P2JN Jatim', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693831&thang=2026' },
  { title: 'Progress PJN I Bali', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693912&thang=2026' },
  { title: 'Progress PJN I Jatim', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693899&thang=2026' },
  { title: 'Progress PJN II Bali', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693913&thang=2026' },
  { title: 'Progress PJN II Jatim', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693900&thang=2026' },
  { title: 'Progress PJN III Bali', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693866&thang=2026' },
  { title: 'Progress PJN III Jatim', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693863&thang=2026' },
  { title: 'Progress PJN IV Jatim', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693901&thang=2026' },
  { title: 'Progress SKPD Bali', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693772&thang=2026' },
  { title: 'Progress SKPD Jatim', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693752&thang=2026' },
  { title: 'Progress Suramadu', url: 'https://iemon.pu.go.id/progres_per_paket?sat=04693840&thang=2026' },
  { title: 'Rekap', url: 'https://iemon.pu.go.id/progres_balai_bm2' },
  { title: 'Sp Progress Efisiensi', url: 'https://iemon.pu.go.id/progres_per_satker_efektif?thang=2026&x1=04&select=' },
  { title: 'Tender', url: 'https://iemon.pu.go.id/pep_kontrak2020bm_sat.php?kd=balaibm08' },
  { title: 'Terkontrak Balai', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693824' },
  { title: 'Terkontrak P2JN Bali', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693847' },
  { title: 'Terkontrak P2JN Jatim', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693831' },
  { title: 'Terkontrak PJN I Bali', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693912' },
  { title: 'Terkontrak PJN I Jatim', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693899' },
  { title: 'Terkontrak PJN II Bali', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693913' },
  { title: 'Terkontrak PJN II Jatim', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693900' },
  { title: 'Terkontrak PJN III Bali', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693866' },
  { title: 'Terkontrak PJN III Jatim', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693863' },
  { title: 'Terkontrak PJN IV Jatim', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693901' },
  { title: 'Terkontrak SKPD Jatim', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693752' },
  { title: 'Terkontrak SKPD Bali', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693772' },
  { title: 'Terkontrak Suramadu', url: 'https://iemon.pu.go.id/kontrak_sat?sat=04693840' },
  { title: 'Epur Balai', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693824' },
  { title: 'Epur P2JN Bali', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693847' },
  { title: 'Epur P2JN Jatim', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693831' },
  { title: 'Epur PJN I Bali', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693912' },
  { title: 'Epur PJN I Jatim', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693899' },
  { title: 'Epur PJN II Bali', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693913' },
  { title: 'Epur PJN II Jatim', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693900' },
  { title: 'Epur PJN III Bali', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693866' },
  { title: 'Epur PJN III Jatim', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693863' },
  { title: 'Epur PJN IV Jatim', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693901' },
  { title: 'Epur SKPD Jatim', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693752' },
  { title: 'Epur SKPD Bali', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693772' },
  { title: 'Epur Suramadu', url: 'https://iemon.pu.go.id/nontender_sat?sat=04693840' },
  { title: 'PU_total', url: 'https://iemon.pu.go.id/progres_kemen?jn=' },
  { title: 'PAKET', url: 'https://iemon.pu.go.id/pep_paket_all_bm' },
];

// Equivalent of the Excel "Tabs" sheet: Nama (lowercase first word of
// title) -> ordered list of tab labels to click. Your tab_mapping.xlsx
// "Tabs" sheet is empty, so this stays empty too — terkontrak/epur route
// through TABLE_INDEX_MAP below instead, same as your desktop default.
const TAB_MAPPING = {
};

// Equivalent of the Excel "TableMap" sheet — loaded from your
// tab_mapping.xlsx exactly as given (your epur mapping has 4 entries,
// not the 2-entry default from the desktop script).
// key -> { tableIndex(1-based, in page order): sheetName }
const TABLE_INDEX_MAP = {
  terkontrak: { 3: 'Terkontrak', 4: 'Persiapan', 5: 'Belum', 6: 'Proses', 7: 'Belum' },
  epur: { 3: 'Terkontrak', 4: 'Proses', 5: 'Belum', 6: 'Belum' },
};

// Equivalent of your "manual export" domain checkboxes. Set to sipp.pu.go.id
// since your desktop script's own tooltip named SIPP as the standard
// manual-export example — remove this line if SIPP should scrape normally.
const MANUAL_EXPORT_DOMAINS = new Set([
  'sipp.pu.go.id',
]);

const EXPECTED_RR_HEADERS = [
  'Satker', 'NO', 'Kode', 'Satker/Output/ Komponen/ Sub Komponen',
  'Target Output Vol', 'Target Output Satuan',
  'Target Paket Vol', 'Target Paket Satuan',
  'Lokasi', 'Pagu (Rp Ribu)', 'Realisasi (Rp Ribu)',
  'Keuangan RN', 'Keuangan RL', 'Fisik RN', 'Fisik RL', 'Kinerja',
];

// ============================================================================
// Injected JS — runs inside the WebView (ported from BeautifulSoup logic)
// ============================================================================

// Shared helper functions injected into every script below.
const JS_HELPERS = `
  function __sleep(ms) { return new Promise(function(r){ setTimeout(r, ms); }); }
  function __countTables() { return document.querySelectorAll('table').length; }

  function __extractAllTables() {
    var tables = Array.prototype.slice.call(document.querySelectorAll('table'));
    var out = [];
    tables.forEach(function(table) {
      var rows = Array.prototype.slice.call(table.querySelectorAll('tr'));
      var grid = [];
      var rowSpans = {};
      rows.forEach(function(tr, rowIdx) {
        var row = [];
        var colIdx = 0;
        while (rowSpans[rowIdx + ',' + colIdx]) {
          var rs = rowSpans[rowIdx + ',' + colIdx];
          row.push(rs.val);
          delete rowSpans[rowIdx + ',' + colIdx];
          if (rs.remain > 1) rowSpans[(rowIdx + 1) + ',' + colIdx] = { val: rs.val, remain: rs.remain - 1 };
          colIdx++;
        }
        var cells = Array.prototype.slice.call(tr.querySelectorAll('th,td'));
        cells.forEach(function(cell) {
          var text = (cell.textContent || '').replace(/\\u00a0/g, ' ').trim();
          var colspan = parseInt(cell.getAttribute('colspan') || '1', 10) || 1;
          var rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10) || 1;
          row.push(text);
          for (var i = 1; i < colspan; i++) row.push(text);
          if (rowspan > 1) {
            for (var r = 1; r < rowspan; r++) {
              for (var c = 0; c < colspan; c++) {
                rowSpans[(rowIdx + r) + ',' + (colIdx + c)] = { val: text, remain: rowspan - r };
              }
            }
          }
          colIdx += colspan;
        });
        grid.push(row);
      });
      if (grid.length === 0) return;
      var maxCols = 0;
      grid.forEach(function(r) { if (r.length > maxCols) maxCols = r.length; });
      grid.forEach(function(r) { while (r.length < maxCols) r.push(''); });

      var headerDepth = 0;
      for (var i = 0; i < rows.length; i++) {
        var cells2 = Array.prototype.slice.call(rows[i].querySelectorAll('th,td'));
        if (cells2.length > 0 && cells2.every(function(c) { return c.tagName === 'TH'; })) headerDepth++;
        else break;
      }

      var headers = null, bodyRows;
      if (headerDepth > 0) {
        headers = [];
        for (var c2 = 0; c2 < maxCols; c2++) {
          var parts = [];
          for (var r2 = 0; r2 < headerDepth; r2++) if (grid[r2][c2]) parts.push(grid[r2][c2]);
          headers.push(parts.join(' ').trim() || ('col_' + (c2 + 1)));
        }
        bodyRows = grid.slice(headerDepth);
      } else {
        bodyRows = grid;
      }
      out.push({ headers: headers, rows: bodyRows });
    });
    return out;
  }

  function __findTabElements(labelText) {
    var xpath = "//a[normalize-space()='" + labelText + "'] | //button[normalize-space()='" + labelText + "'] | " +
                "//li[a[normalize-space()='" + labelText + "']]//a | //li[button[normalize-space()='" + labelText + "']]//button";
    var res = document.evaluate(xpath, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
    var els = [];
    for (var i = 0; i < res.snapshotLength; i++) els.push(res.snapshotItem(i));
    return els;
  }
`;

// Non-tabbed page: wait for tables to appear + stabilize, extract all,
// optionally collect sublinks found inside tables (Rencana/Realisasi).
function buildPageExtractScript(collectSubLinks) {
  return `
    ${JS_HELPERS}
    (async function() {
      try {
        var limit = Date.now() + 20000;
        while (__countTables() === 0 && Date.now() < limit) await __sleep(400);
        var prev = -1, stableUntil = Date.now() + 8000;
        while (Date.now() < stableUntil) {
          var c = __countTables();
          if (c === prev && c > 0) break;
          prev = c;
          await __sleep(400);
        }
        var tables = __extractAllTables();
        var subLinks = [];
        ${collectSubLinks ? `
        var seen = {};
        document.querySelectorAll('table a[href]').forEach(function(a) {
          try {
            var abs = new URL(a.getAttribute('href'), window.location.href).href;
            if (!seen[abs]) { seen[abs] = true; subLinks.push({ url: abs, text: (a.textContent || '').trim() }); }
          } catch (e) {}
        });
        ` : ''}
        window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'PAGE_RESULT', tables: tables, subLinks: subLinks }));
      } catch (e) {
        window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'ERROR', message: String(e) }));
      }
    })();
    true;
  `;
}

// Sub-page (a link found inside a Rencana/Realisasi table): extract only.
function buildSubPageExtractScript() {
  return `
    ${JS_HELPERS}
    (async function() {
      try {
        var limit = Date.now() + 20000;
        while (__countTables() === 0 && Date.now() < limit) await __sleep(400);
        var tables = __extractAllTables();
        window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'SUBPAGE_RESULT', tables: tables }));
      } catch (e) {
        window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'ERROR', message: String(e) }));
      }
    })();
    true;
  `;
}

// Tabbed page: click each tab in order, wait for new tables, extract after each.
function buildTabFlowScript(tabList) {
  return `
    ${JS_HELPERS}
    (async function() {
      try {
        var tabs = ${JSON.stringify(tabList)};
        var tabResults = [];
        for (var i = 0; i < tabs.length; i++) {
          var label = tabs[i].label;
          var before = __countTables();
          var els = __findTabElements(label);
          for (var j = 0; j < els.length; j++) {
            try {
              els[j].scrollIntoView({ block: 'center' });
              els[j].click();
              var limit = Date.now() + 12000;
              while (Date.now() < limit) {
                if (__countTables() > before) break;
                await __sleep(400);
              }
              break;
            } catch (e) {}
          }
          var tables = __extractAllTables();
          tabResults.push({ label: label, tables: tables });
        }
        window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'TAB_FLOW_RESULT', tabResults: tabResults }));
      } catch (e) {
        window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'ERROR', message: String(e) }));
      }
    })();
    true;
  `;
}

// ============================================================================
// RN-side processing (ported from the pandas/BeautifulSoup post-processing)
// ============================================================================

function dedupeColumns(cols) {
  var seen = {}, out = [];
  cols.forEach(function (c) {
    c = String(c);
    if (seen[c] !== undefined) { seen[c] += 1; out.push(c + '.' + seen[c]); }
    else { seen[c] = 0; out.push(c); }
  });
  return out;
}

// { headers: string[]|null, rows: string[][] } -> { headers: string[], rows: object[] }
function tableToRowObjects(table) {
  var maxCols = table.rows.length ? table.rows[0].length : (table.headers ? table.headers.length : 0);
  var headers = table.headers
    ? dedupeColumns(table.headers)
    : Array.from({ length: maxCols }, (_, i) => 'col_' + (i + 1));
  var rows = table.rows.map(function (r) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = (r[i] !== undefined ? r[i] : '').toString(); });
    return obj;
  });
  return { headers: headers, rows: rows };
}

// Ported filter_table: keep from first row where col0 is "1"/"1.0"; else
// keep the whole table only if a TOTAL cell exists anywhere; else discard.
function filterTable(headers, rows) {
  if (!headers.length || !rows.length) return null;
  var firstCol = headers[0];
  var startIdx = -1;
  for (var i = 0; i < rows.length; i++) {
    var v = String(rows[i][firstCol] || '').replace(/\s+/g, ' ').trim();
    if (/^1(\.0+)?$/.test(v)) { startIdx = i; break; }
  }
  if (startIdx >= 0) return rows.slice(startIdx);
  var hasTotal = rows.some(function (row) {
    return headers.some(function (h) { return /\bTOTAL\b/i.test(String(row[h] || '')); });
  });
  return hasTotal ? rows.slice() : null;
}

function insertSatker(headers, rows, satker) {
  var newHeaders = ['Satker'].concat(headers);
  var newRows = rows.map(function (r) {
    var nr = { Satker: satker };
    headers.forEach(function (h) { nr[h] = r[h]; });
    return nr;
  });
  return { headers: newHeaders, rows: newRows };
}

// Ported _fix_rr_headers: overwrite headers positionally for Rencana/Realisasi pages.
function fixRRHeaders(headers, rows) {
  if (headers.length < EXPECTED_RR_HEADERS.length) return { headers, rows };
  var newHeaders = EXPECTED_RR_HEADERS.concat(headers.slice(EXPECTED_RR_HEADERS.length));
  var renamed = rows.map(function (r) {
    var nr = {};
    newHeaders.forEach(function (h, i) { nr[h] = r[headers[i]]; });
    return nr;
  });
  return { headers: newHeaders, rows: renamed };
}

function isNumericLike(s) {
  s = String(s || '').trim();
  if (s === '') return false;
  if (/^\d+$/.test(s)) return true;
  if (/^\d+\.0$/.test(s)) return true;
  return false;
}

// Ported epur/terkontrak first-row strip: if row0 cols[1..9] are all
// numeric-looking, drop it; discard whole table if <=1 row remains.
function stripNumericFirstRow(headers, rows) {
  if (rows.length === 0) return rows;
  var checkCols = headers.slice(1, 10);
  var allNumeric = checkCols.length > 0 && checkCols.every(function (h) { return isNumericLike(rows[0][h]); });
  var out = allNumeric ? rows.slice(1) : rows;
  return out.length <= 1 ? null : out;
}

function domainOf(url) {
  try { return url.split('/')[2]; } catch (e) { return ''; }
}

// Sheet-name grouping, ported from the Python merge step.
function linksSheetFor(baseToken) {
  var tok = baseToken.toLowerCase();
  if (tok.startsWith('rencana')) return 'Rencana_Links';
  if (tok.startsWith('realisasi')) return 'Realisasi_Links';
  return tok + '_Sub';
}

// ============================================================================
// Component
// ============================================================================

export default function App() {
  const webviewRef = useRef(null);
  const [phase, setPhase] = useState('SETUP'); // SETUP -> LOGIN -> RUNNING -> DONE
  const [manualExportActive, setManualExportActive] = useState(false);
  const [downloadCaptured, setDownloadCaptured] = useState(false);
  const [log, setLog] = useState([]);
  const [progress, setProgress] = useState({ done: 0, total: PAGES.length });
  const [resultUri, setResultUri] = useState(null);

  // Pending-promise bridge for WebView round trips.
  const pendingResolveRef = useRef(null);
  const pendingDownloadUrlRef = useRef(null);
  const capturedBlobRef = useRef(null); // { filename, base64 } from the blob-capture hook

  const groupedRef = useRef({}); // sheetName -> array of {headers, rows}
  const skippedRef = useRef([]);
  // resolved by the "Continue" / "I've exported" buttons rendered below
  const loginResolveRef = useRef(null);
  const exportResolveRef = useRef(null);

  const appendLog = useCallback((msg) => {
    setLog((prev) => [...prev.slice(-200), msg]);
  }, []);

  const waitForMessage = useCallback(() => {
    return new Promise((resolve) => { pendingResolveRef.current = resolve; });
  }, []);

  const runScript = useCallback((script) => {
    const p = waitForMessage();
    webviewRef.current?.injectJavaScript(script);
    return p;
  }, [waitForMessage]);

  const navigateAndWaitLoad = useCallback((url) => {
    return new Promise((resolve) => {
      pendingResolveRef.current = () => resolve();
      webviewRef.current?.injectJavaScript(`window.location.href = ${JSON.stringify(url)}; true;`);
    });
  }, []);

  const handleMessage = useCallback((event) => {
    let msg;
    try { msg = JSON.parse(event.nativeEvent.data); } catch { return; }
    if (msg.type === 'BLOB_EXPORT') {
      capturedBlobRef.current = { filename: msg.filename, base64: msg.base64 };
      setDownloadCaptured(true);
      appendLog(`Download captured via blob hook: ${msg.filename} (${Math.round(msg.base64.length / 1370)}KB)`);
      return;
    }
    if (msg.type === 'HOOK_READY' || msg.type === 'HOOK_INJECTING') {
      appendLog(msg.type);
      return;
    }
    if (msg.type === 'DEBUG') {
      appendLog(`DEBUG: ${msg.message}`);
      return;
    }
    if (msg.type === 'ERROR' && !pendingResolveRef.current) {
      // No page-scrape RPC is waiting on this (e.g. it fired during the
      // manual-export wait) — log it directly so it's actually visible.
      appendLog(`ERROR: ${msg.message}`);
      return;
    }
    const resolve = pendingResolveRef.current;
    pendingResolveRef.current = null;
    if (resolve) resolve(msg);
  }, [appendLog]);

  // Injected into the export page: intercepts blob: URL downloads (the
  // common case when a site builds the export file client-side in JS,
  // e.g. the DataTables "Buttons" excelHtml5 export, rather than serving
  // it from the server) since Android's native onFileDownload never fires
  // for those. DataTables-style exports often build a throwaway <a> and
  // call .click() on it in JS WITHOUT appending it to the document, which
  // a document-level click listener can miss entirely — so this hooks the
  // actual APIs involved instead: URL.createObjectURL (to capture the real
  // Blob the moment it's created, keyed by its generated URL) and
  // HTMLAnchorElement.prototype.click (to catch programmatic clicks on
  // detached elements). A real click-event listener and window.open
  // override are kept too, as extra safety nets for other export styles.
  const BLOB_CAPTURE_HOOK = `
    (function() {
      window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'HOOK_INJECTING' }));
      try {
        function dbg(msg) {
          try { window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'DEBUG', message: msg })); } catch (e) {}
        }

        var blobMap = {};
        var origCreateObjectURL = URL.createObjectURL.bind(URL);
        URL.createObjectURL = function(blob) {
          var url = origCreateObjectURL(blob);
          try { blobMap[url] = blob; } catch (e) {}
          dbg('URL.createObjectURL called, type=' + (blob && blob.type) + ' size=' + (blob && blob.size));
          return url;
        };

        function toBase64(blob) {
          return new Promise(function(resolve, reject) {
            var reader = new FileReader();
            reader.onloadend = function() { resolve(reader.result.split(',')[1]); };
            reader.onerror = reject;
            reader.readAsDataURL(blob);
          });
        }
        function sendBlob(blob, filename) {
          toBase64(blob).then(function(b64) {
            window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'BLOB_EXPORT', filename: filename || 'export.xlsx', base64: b64 }));
          }).catch(function(e) {
            window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'ERROR', message: 'blob read failed: ' + e }));
          });
        }
        function sendDataUri(href, filename) {
          var commaIdx = href.indexOf(',');
          if (commaIdx === -1) { dbg('data URI had no comma: ' + href.slice(0, 60)); return; }
          var base64 = href.slice(commaIdx + 1);
          window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'BLOB_EXPORT', filename: filename || 'export.xlsx', base64: base64 }));
        }
        function resolveBlob(href) {
          if (blobMap[href]) return Promise.resolve(blobMap[href]);
          return fetch(href).then(function(r) { return r.blob(); });
        }
        // Handles blob:, data:, or anything else — logs what it saw either way.
        function handleHref(href, filename) {
          if (href.indexOf('blob:') === 0) {
            resolveBlob(href).then(function(blob) { sendBlob(blob, filename); })
              .catch(function(err) { window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'ERROR', message: 'blob resolve failed: ' + err })); });
            return true;
          }
          if (href.indexOf('data:') === 0) {
            sendDataUri(href, filename);
            return true;
          }
          return false;
        }

        // Catches DataTables-style: document.createElement('a'); a.click();
        // — works even if the anchor was never appended to the document.
        // Logs EVERY anchor .click() call, matched or not, for diagnosis.
        var origAnchorClick = HTMLAnchorElement.prototype.click;
        HTMLAnchorElement.prototype.click = function() {
          var href = this.getAttribute('href') || '';
          dbg('anchor.click() href=' + href.slice(0, 80) + ' download=' + this.getAttribute('download'));
          var filename = this.getAttribute('download') || 'export.xlsx';
          if (handleHref(href, filename)) return;
          return origAnchorClick.apply(this, arguments);
        };

        // Catches FileSaver.js-style: node.dispatchEvent(new MouseEvent('click'))
        // — this is what DataTables Buttons uses when FileSaver.js is loaded
        // (window.saveAs exists), and it neither calls .click() nor requires
        // the node to be attached to the document, so neither hook above
        // would see it without this.
        var origDispatchEvent = HTMLAnchorElement.prototype.dispatchEvent;
        HTMLAnchorElement.prototype.dispatchEvent = function(event) {
          if (event && event.type === 'click') {
            var href = this.getAttribute('href') || '';
            dbg('anchor.dispatchEvent(click) href=' + href.slice(0, 80) + ' download=' + this.getAttribute('download'));
            var filename = this.getAttribute('download') || 'export.xlsx';
            if (handleHref(href, filename)) return true;
          }
          return origDispatchEvent.call(this, event);
        };

        // Fallback: a real user click on a visible link/button.
        document.addEventListener('click', function(e) {
          var a = e.target && e.target.closest ? e.target.closest('a') : null;
          if (!a) return;
          var href = a.getAttribute('href') || '';
          dbg('document click on <a> href=' + href.slice(0, 80));
          if (href.indexOf('blob:') === 0 || href.indexOf('data:') === 0 || a.hasAttribute('download')) {
            e.preventDefault();
            e.stopPropagation();
            handleHref(href, a.getAttribute('download') || 'export.xlsx');
          }
        }, true);

        // Fallback: window.open(blobUrl / dataUrl) style export.
        var origOpen = window.open;
        window.open = function(url) {
          dbg('window.open(' + String(url).slice(0, 80) + ')');
          if (typeof url === 'string' && handleHref(url, 'export.xlsx')) return null;
          return origOpen.apply(window, arguments);
        };

        window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'HOOK_READY' }));
      } catch (e) {
        window.ReactNativeWebView.postMessage(JSON.stringify({ type: 'ERROR', message: 'hook install failed: ' + (e && e.message ? e.message : String(e)) }));
      }
    })();
    true;
  `;

  // Fires when navigateAndWaitLoad() is waiting on an actual page load
  // (as opposed to runScript()'s waitForMessage(), which resolves via
  // handleMessage instead). Since the two are always awaited sequentially
  // in the pipeline below, sharing one pending-resolver ref is safe.
  const handleLoadEnd = useCallback(() => {
    const resolve = pendingResolveRef.current;
    if (resolve) {
      pendingResolveRef.current = null;
      resolve();
    }
  }, []);

  // ---- process one raw extracted table into the grouping structure ----
  const processMainTable = useCallback((rawTable, baseToken, satker, key, sheetKeyOverride, tIdx, origin) => {
    let { headers, rows } = tableToRowObjects(rawTable);
    let kept = filterTable(headers, rows);
    if (kept === null) return { keptRowCount: 0, reason: `filter_table discarded (no row-1 start, no TOTAL cell; ${rows.length} raw rows, headers: ${headers.slice(0, 6).join(', ')})` };
    ({ headers, rows: kept } = insertSatker(headers, kept, satker));

    const firstLower = baseToken.split('_')[0].toLowerCase();
    if (firstLower.startsWith('rencana') || firstLower.startsWith('realisasi')) {
      ({ headers, rows: kept } = fixRRHeaders(headers, kept));
    }
    if (firstLower === 'epur' || firstLower === 'terkontrak') {
      kept = stripNumericFirstRow(headers, kept);
      if (kept === null) return { keptRowCount: 0, reason: 'stripNumericFirstRow left <=1 row' };
    }

    let sheetName = sheetKeyOverride;
    if (origin === 'main' && tIdx != null) {
      const mp = TABLE_INDEX_MAP[key];
      if (mp && mp[tIdx]) sheetName = mp[tIdx];
      else if (key === 'terkontrak' || key === 'epur') sheetName = key[0].toUpperCase() + key.slice(1);
    }
    if (origin === 'sub') sheetName = linksSheetFor(baseToken.split('_')[0]);

    if (!groupedRef.current[sheetName]) groupedRef.current[sheetName] = [];
    groupedRef.current[sheetName].push({ headers, rows: kept });
    return { keptRowCount: kept.length, sheetName };
  }, []);

  const startPipeline = useCallback(async () => {
    setPhase('LOGIN');
    groupedRef.current = {};
    skippedRef.current = [];

    // ---- Step 1: sequential per-domain manual login ----
    const domains = [...new Set(PAGES.map((p) => domainOf(p.url)))];
    for (const dom of domains) {
      const pageUrl = PAGES.find((p) => domainOf(p.url) === dom).url;
      appendLog(`Login: ${dom}`);
      await new Promise((resolve) => {
        webviewRef.current?.stopLoading?.();
        webviewRef.current?.injectJavaScript(`window.location.href = ${JSON.stringify(pageUrl)}; true;`);
        // handoff button (rendered below) resolves this via loginResolveRef
        loginResolveRef.current = resolve;
      });
    }

    // ---- Step 2: scrape each page ----
    setPhase('RUNNING');
    for (let idx = 0; idx < PAGES.length; idx++) {
      const page = PAGES[idx];
      const title = page.title || `Page_${idx + 1}`;
      const base = title.replace(/ /g, '_');
      const key = (title.split(' ')[0] || base.split('_')[0]).toLowerCase();
      const parts = base.split(/_(.+)/); // split on first underscore
      const satker = parts.length > 1 ? parts[1] : '';
      const dom = domainOf(page.url);

      setProgress({ done: idx, total: PAGES.length });
      appendLog(`[${idx + 1}/${PAGES.length}] ${title}`);

      // ---- manual export domain ----
      if (MANUAL_EXPORT_DOMAINS.has(dom)) {
        pendingDownloadUrlRef.current = null;
        capturedBlobRef.current = null;
        setDownloadCaptured(false);
        await navigateAndWaitLoad(page.url);
        webviewRef.current?.injectJavaScript(BLOB_CAPTURE_HOOK);
        appendLog(`Waiting for you to tap the export button, then "Continue"…`);
        setManualExportActive(true);
        try {
          await new Promise((resolve) => { exportResolveRef.current = resolve; });
        } finally {
          setManualExportActive(false);
        }

        if (capturedBlobRef.current) {
          try {
            const { base64 } = capturedBlobRef.current;
            const wb = XLSX.read(base64, { type: 'base64' });
            const firstSheet = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
            if (!groupedRef.current[key]) groupedRef.current[key] = [];
            groupedRef.current[key].push({ headers: Object.keys(rows[0] || {}), rows });
          } catch (e) {
            skippedRef.current.push({ 'Sheet Name': base.slice(0, 31), 'Page Title': title, 'Page URL': page.url, Reason: `Manual export (blob) failed: ${e}` });
          }
          continue;
        }

        const downloadUrl = pendingDownloadUrlRef.current;
        if (!downloadUrl) {
          skippedRef.current.push({ 'Sheet Name': base.slice(0, 31), 'Page Title': title, 'Page URL': page.url, Reason: 'No download detected — neither the blob hook nor onFileDownload caught anything. Check whether the export actually navigates the page instead of downloading.' });
          continue;
        }
        try {
          const dest = FileSystem.cacheDirectory + `export_${idx}.xlsx`;
          await FileSystem.downloadAsync(downloadUrl, dest);
          const b64 = await FileSystem.readAsStringAsync(dest, { encoding: FileSystem.EncodingType.Base64 });
          const wb = XLSX.read(b64, { type: 'base64' });
          const firstSheet = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
          if (!groupedRef.current[key]) groupedRef.current[key] = [];
          groupedRef.current[key].push({ headers: Object.keys(rows[0] || {}), rows });
        } catch (e) {
          skippedRef.current.push({ 'Sheet Name': base.slice(0, 31), 'Page Title': title, 'Page URL': page.url, Reason: `Manual export failed: ${e}` });
        }
        continue;
      }

      // ---- tabbed page ----
      if (TAB_MAPPING[key]) {
        await navigateAndWaitLoad(page.url);
        const msg = await runScript(buildTabFlowScript(TAB_MAPPING[key]));
        if (msg.type === 'ERROR') {
          skippedRef.current.push({ 'Sheet Name': base.slice(0, 31), 'Page Title': title, 'Page URL': page.url, Reason: msg.message });
          continue;
        }
        for (const tabResult of msg.tabResults) {
          let matched = false;
          for (const rawTable of tabResult.tables) {
            const { headers, rows } = tableToRowObjects(rawTable);
            const kept = filterTable(headers, rows);
            if (kept === null) continue;
            const outcome = processMainTable(rawTable, base, satker, key, tabResult.label, null, 'main');
            appendLog(`  [${tabResult.label}] kept ${outcome.keptRowCount} rows -> ${outcome.sheetName}`);
            matched = true;
            break; // first valid table per tab, same as desktop version
          }
          if (!matched) appendLog(`  [${tabResult.label}] ${tabResult.tables.length} raw tables, none passed filter`);
        }
        continue;
      }

      // ---- normal page ----
      const followSublinks = key.startsWith('rencana') || key.startsWith('realisasi');
      await navigateAndWaitLoad(page.url);
      const msg = await runScript(buildPageExtractScript(followSublinks));
      if (msg.type === 'ERROR') {
        skippedRef.current.push({ 'Sheet Name': base.slice(0, 31), 'Page Title': title, 'Page URL': page.url, Reason: msg.message });
        continue;
      }
      appendLog(`  raw tables found: ${msg.tables.length}`);
      msg.tables.forEach((rawTable, tIdx) => {
        const outcome = processMainTable(rawTable, base, satker, key, key, tIdx + 1, 'main');
        if (outcome.keptRowCount > 0) appendLog(`  table ${tIdx + 1}: kept ${outcome.keptRowCount} rows -> ${outcome.sheetName}`);
        else appendLog(`  table ${tIdx + 1}: discarded (${outcome.reason})`);
      });

      // ---- follow sublinks (Rencana/Realisasi) ----
      if (followSublinks && msg.subLinks && msg.subLinks.length) {
        for (const sub of msg.subLinks) {
          await navigateAndWaitLoad(sub.url);
          const subMsg = await runScript(buildSubPageExtractScript());
          if (subMsg.type === 'ERROR') {
            skippedRef.current.push({ 'Sheet Name': base.slice(0, 31), 'Page Title': `${title} - ${sub.text}`, 'Page URL': sub.url, Reason: subMsg.message });
            continue;
          }
          subMsg.tables.forEach((rawTable) => {
            processMainTable(rawTable, base, satker, key, sub.text || 'Sub', null, 'sub');
          });
        }
      }
    }

    setProgress({ done: PAGES.length, total: PAGES.length });
    setPhase('DONE');
  }, [appendLog, navigateAndWaitLoad, runScript, processMainTable]);

  const [exportError, setExportError] = useState(null);

  const exportXlsx = useCallback(async () => {
    setExportError(null);
    try {
      const wb = XLSX.utils.book_new();
      Object.entries(groupedRef.current).forEach(([sheetName, tables]) => {
        const allRows = tables.flatMap((t) => t.rows);
        if (!allRows.length) return;
        const ws = XLSX.utils.json_to_sheet(allRows, { dense: true });
        XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
      });
      if (skippedRef.current.length) {
        const ws = XLSX.utils.json_to_sheet(skippedRef.current, { dense: true });
        XLSX.utils.book_append_sheet(wb, ws, 'Info');
      }
      if (wb.SheetNames.length === 0) {
        // SheetJS throws on write if there are zero sheets — guard with a
        // placeholder so a "nothing scraped" run still produces a file
        // instead of silently failing.
        const ws = XLSX.utils.json_to_sheet([{ Note: 'No tables were scraped this run.' }], { dense: true });
        XLSX.utils.book_append_sheet(wb, ws, 'Info');
      }
      const b64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
      const uri = FileSystem.cacheDirectory + 'scrape_result.xlsx';
      await FileSystem.writeAsStringAsync(uri, b64, { encoding: FileSystem.EncodingType.Base64 });
      setResultUri(uri);
      appendLog(`Exported ${wb.SheetNames.length} sheet(s) to ${uri}`);
      if (await Sharing.isAvailableAsync()) {
        await Sharing.shareAsync(uri);
      } else {
        appendLog('Sharing not available on this device — file is saved at the path above.');
      }
    } catch (e) {
      setExportError(String(e));
      appendLog(`Export failed: ${e}`);
    }
  }, [appendLog]);

  return (
    <SafeAreaView style={styles.container}>
      <Text style={styles.title}>PU Table Scraper</Text>

      {phase === 'SETUP' && (
        <>
          <Text style={styles.hint}>{PAGES.length} pages configured.</Text>
          <Button title="Start" onPress={startPipeline} disabled={PAGES.length === 0} />
        </>
      )}

      {phase === 'LOGIN' && (
        <>
          <Text style={styles.hint}>Log in on the site below, then tap Continue.</Text>
          <Button title="Continue" onPress={() => loginResolveRef.current?.()} />
        </>
      )}

      {phase === 'RUNNING' && !manualExportActive && (
        <>
          <ActivityIndicator />
          <Text style={styles.hint}>{progress.done}/{progress.total} pages processed</Text>
        </>
      )}

      {manualExportActive && (
        <>
          <Text style={styles.hint}>
            {downloadCaptured
              ? '✓ Download captured — tap Continue.'
              : 'Tap the export button on the page below. This will confirm here once a file download is detected.'}
          </Text>
          <Button title="I've exported — continue" onPress={() => exportResolveRef.current?.()} />
        </>
      )}

      {phase === 'DONE' && (
        <>
          <Text style={styles.hint}>Done. {skippedRef.current.length} skipped.</Text>
          <Button title="Export & share XLSX" onPress={exportXlsx} />
          {resultUri && <Text style={styles.hint}>Saved.</Text>}
          {exportError && <Text style={styles.error}>{exportError}</Text>}
        </>
      )}

      <ScrollView style={styles.logBox}>
        {log.map((l, i) => <Text key={i} style={styles.logLine}>{l}</Text>)}
      </ScrollView>

      <View style={(phase === 'LOGIN' || manualExportActive) ? styles.visibleWebview : styles.hiddenWebview}>
        <WebView
          ref={webviewRef}
          source={{ uri: PAGES[0]?.url || 'about:blank' }}
          onMessage={handleMessage}
          onLoadEnd={handleLoadEnd}
          onFileDownload={({ nativeEvent }) => {
            pendingDownloadUrlRef.current = nativeEvent.downloadUrl;
            setDownloadCaptured(true);
            appendLog(`Download captured: ${nativeEvent.downloadUrl}`);
          }}
          sharedCookiesEnabled
          thirdPartyCookiesEnabled
          javaScriptEnabled
          domStorageEnabled
        />
      </View>
    </SafeAreaView>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, padding: 16, paddingTop: 40 },
  title: { fontSize: 20, fontWeight: '600', marginBottom: 8 },
  hint: { marginVertical: 8, color: '#555' },
  error: { color: 'red', marginVertical: 8 },
  logBox: { maxHeight: 160, marginTop: 8 },
  logLine: { fontSize: 11, color: '#333' },
  visibleWebview: { flex: 1, marginTop: 8, borderWidth: 1, borderColor: '#ddd' },
  hiddenWebview: { position: 'absolute', width: 1, height: 1, opacity: 0, top: -9999 },
});