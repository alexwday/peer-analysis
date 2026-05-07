#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import json
import mimetypes
from pathlib import Path
from typing import Any

from xlsxwriter.utility import xl_col_to_name

from create_attributed_growth_by_segment_template import (
    DEFAULT_INPUT_PATH,
    INPUT_LAYOUT,
    REPO_ROOT,
    load_config,
    pct_label,
    png_dimensions,
    resolve_logo_path,
    section_header,
    segment_value,
)

HTML_OUTPUT_PATH = (
    REPO_ROOT
    / "outputs"
    / "attributed-growth-by-segment-template"
    / "attributed_growth_by_segment_template.html"
)


def image_data_uri(path: Path | None) -> str | None:
    if path is None or not path.is_file():
        return None
    mime_type = mimetypes.guess_type(path.name)[0] or "image/png"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def source_refs(config: dict[str, Any]) -> dict[str, dict[str, str]]:
    refs: dict[str, dict[str, str]] = {}
    sheet = INPUT_LAYOUT["sheet_name"]
    for bank_index, bank in enumerate(config["banks"]):
        row = INPUT_LAYOUT["data_start_row"] + bank_index
        bank_refs = {
            "bank": f"{sheet}!A{row}",
            "absoluteGrowthBn": f"{sheet}!B{row}",
            "totalGrowthPct": f"{sheet}!C{row}",
        }
        for segment_index, segment in enumerate(config["segments"]):
            col = xl_col_to_name(INPUT_LAYOUT["data_segment_start_col"] + segment_index)
            bank_refs[segment["key"]] = f"{sheet}!{col}{row}"
        refs[bank["bank"]] = bank_refs
    return refs


def app_data(config: dict[str, Any], input_path: Path) -> dict[str, Any]:
    refs = source_refs(config)
    banks = []
    source_rows = []

    for bank in config["banks"]:
        logo_path = resolve_logo_path(bank, config)
        logo_dimensions = png_dimensions(logo_path) if logo_path else None
        values = {
            segment["key"]: segment_value(bank, segment["key"])
            for segment in config["segments"]
        }
        bank_record = {
            "bank": bank["bank"],
            "shortName": bank.get("shortName") or bank["bank"],
            "absoluteGrowthBn": bank["absoluteGrowthBn"],
            "totalGrowthPct": bank["totalGrowthPct"],
            "attributed": bool(bank.get("attributed")),
            "brandColor": bank.get("brandColor", "#111827"),
            "segments": values,
            "logo": image_data_uri(logo_path),
            "logoWidth": logo_dimensions[0] if logo_dimensions else 120,
            "logoHeight": logo_dimensions[1] if logo_dimensions else 40,
            "refs": refs[bank["bank"]],
        }
        banks.append(bank_record)

        source_rows.extend(
            [
                {
                    "bank": bank["bank"],
                    "metric": "Absolute Growth ($BN)",
                    "value": bank["absoluteGrowthBn"],
                    "display": f'{bank["absoluteGrowthBn"]:+.1f}BN',
                    "sourceCell": refs[bank["bank"]]["absoluteGrowthBn"],
                    "role": "Top label",
                },
                {
                    "bank": bank["bank"],
                    "metric": "Reported YoY Growth",
                    "value": bank["totalGrowthPct"],
                    "display": pct_label(bank["totalGrowthPct"]),
                    "sourceCell": refs[bank["bank"]]["totalGrowthPct"],
                    "role": "Top label",
                },
            ]
        )
        for segment in config["segments"]:
            value = values[segment["key"]]
            source_rows.append(
                {
                    "bank": bank["bank"],
                    "metric": segment["label"],
                    "value": value,
                    "display": pct_label(value),
                    "sourceCell": refs[bank["bank"]][segment["key"]],
                    "role": "Chart segment",
                }
            )

    return {
        "title": section_header(config),
        "sectionTitle": config["sectionTitle"],
        "attributedLabel": config.get("attributedLabel", "Attributed Growth/Contribution"),
        "inputPath": str(input_path.resolve()),
        "inputSheet": INPUT_LAYOUT["sheet_name"],
        "segments": config["segments"],
        "banks": banks,
        "sourceRows": source_rows,
    }


HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Peer Analysis</title>
<style>
:root {
  --ink: #101820;
  --muted: #64707d;
  --line: #d7dde3;
  --panel: #ffffff;
  --panel-soft: #f5f7f9;
  --navy: #071734;
  --blue: #0d6786;
  --focus: #2b6cb0;
  --danger: #9b1c1c;
  --shadow: 0 10px 28px rgba(16, 24, 32, 0.09);
}
* { box-sizing: border-box; }
body {
  margin: 0;
  background: #eef2f5;
  color: var(--ink);
  font-family: Arial, Helvetica, sans-serif;
  font-size: 14px;
}
button, input, select { font: inherit; }
button {
  border: 1px solid #b9c3cc;
  background: #fff;
  color: #111827;
  border-radius: 6px;
  padding: 7px 11px;
  font-weight: 700;
  cursor: pointer;
}
button:hover { border-color: #7b8794; background: #f8fafc; }
button.primary { background: var(--navy); border-color: var(--navy); color: #fff; }
button.ghost { background: transparent; }
.icon-btn {
  width: 34px;
  height: 34px;
  padding: 0;
  display: inline-grid;
  place-items: center;
}
.icon-btn svg {
  width: 18px;
  height: 18px;
  stroke: currentColor;
  stroke-width: 2.2;
  fill: none;
  stroke-linecap: round;
  stroke-linejoin: round;
}
.app-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  padding: 14px 22px;
  background: #ffffff;
  border-bottom: 1px solid var(--line);
  position: sticky;
  top: 0;
  z-index: 20;
}
.title-block { min-width: 280px; }
.eyebrow {
  color: var(--muted);
  font-size: 11px;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: 0;
}
h1 {
  margin: 2px 0 0;
  font-size: 20px;
  line-height: 1.15;
}
.actions {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
  justify-content: flex-end;
}
.workspace {
  display: grid;
  gap: 16px;
  padding: 16px;
  max-width: 1500px;
  margin: 0 auto;
}
.panel {
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 8px;
  box-shadow: var(--shadow);
  min-width: 0;
}
.main-stack {
  display: grid;
  gap: 12px;
  min-width: 0;
}
.chart-panel { overflow: hidden; }
.chart-wrap {
  padding: 12px 14px 8px;
  overflow-x: auto;
}
#chartSvg {
  display: block;
  width: 100%;
  min-width: 980px;
  height: auto;
  background: #fff;
}
.table-panel {
  margin: 0 16px 16px;
  overflow: hidden;
}
.drawer {
  overflow: hidden;
}
.drawer > summary {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 12px 14px;
  cursor: pointer;
  list-style: none;
  font-weight: 900;
  background: #fbfcfd;
  border-bottom: 1px solid transparent;
}
.drawer[open] > summary {
  border-bottom-color: var(--line);
}
.drawer > summary::-webkit-details-marker { display: none; }
.drawer-title {
  display: flex;
  align-items: center;
  gap: 9px;
}
.drawer-title::before {
  content: "▸";
  color: var(--muted);
  font-size: 12px;
}
.drawer[open] .drawer-title::before {
  content: "▾";
}
.drawer-meta {
  color: var(--muted);
  font-size: 12px;
  font-weight: 700;
}
.drawer-body {
  padding: 12px;
}
.drawer-controls {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
  margin-bottom: 10px;
}
.source-controls { justify-content: flex-end; }
.calc-controls { justify-content: space-between; }
.calc-controls select { margin-right: auto; }
.drawer-controls select {
  height: 34px;
  min-width: 210px;
  border: 1px solid #cbd5df;
  border-radius: 6px;
  background: #fff;
  padding: 4px 8px;
  font-weight: 700;
}
.adjusted-cell {
  background: #fff3c4;
}
.comment-mark {
  color: #8a5a00;
  font-size: 11px;
  font-weight: 900;
}
.cell-display {
  width: 122px;
  height: 30px;
  display: inline-grid;
  grid-template-columns: minmax(54px, 1fr) 48px;
  align-items: center;
  gap: 4px;
  border: 1px solid #d6dee7;
  background: #ffffff;
  border-radius: 5px;
  padding: 3px 5px;
  cursor: pointer;
}
.cell-display:hover {
  border-color: var(--focus);
}
.cell-value {
  text-align: right;
  font-weight: 800;
}
.delta-slot {
  min-width: 0;
  text-align: left;
  white-space: nowrap;
}
.delta {
  display: inline-flex;
  align-items: center;
  gap: 1px;
  font-size: 10px;
  font-weight: 900;
  line-height: 1;
}
.delta.up { color: #157347; }
.delta.down { color: #b42318; }
.adjustment-log {
  margin-top: 10px;
  padding: 9px 10px;
  border: 1px solid #ead99b;
  border-radius: 7px;
  background: #fff9df;
  color: #5f4700;
  font-size: 12px;
  line-height: 1.35;
}
.supporting-table-wrap {
  max-height: 420px;
  overflow: auto;
  border: 1px solid #c9d3dc;
  background: #fff;
}
.table-header, .side-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 11px 12px;
  border-bottom: 1px solid var(--line);
  background: #fbfcfd;
}
.table-header h2, .side-header h2 {
  margin: 0;
  font-size: 14px;
}
.table-shell { overflow: auto; max-height: 325px; }
table {
  width: 100%;
  border-collapse: collapse;
  font-size: 13px;
}
table.excel-grid {
  border-collapse: collapse;
  font-size: 12.5px;
}
table.excel-grid th {
  background: #eef3f7;
  color: #111827;
  border: 1px solid #c9d3dc;
  border-top: 0;
  padding: 7px 8px;
}
table.excel-grid td {
  border: 1px solid #d7dde3;
  padding: 6px 8px;
  background: #fff;
}
table.excel-grid tbody tr:hover td {
  background: #f7fafc;
}
th {
  position: sticky;
  top: 0;
  background: #0d6786;
  color: #fff;
  z-index: 1;
  text-align: left;
  font-weight: 800;
  padding: 7px 8px;
  white-space: nowrap;
}
td {
  border-bottom: 1px solid #e3e8ee;
  padding: 5px 8px;
  vertical-align: middle;
  white-space: nowrap;
}
tbody tr:hover { background: #f6f9fb; }
td.numeric { text-align: right; }
td.ref { color: var(--muted); font-size: 12px; }
.metric-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 8px;
  padding: 12px;
}
.metric {
  border: 1px solid var(--line);
  border-radius: 7px;
  padding: 9px;
  background: #fff;
}
.metric .label {
  color: var(--muted);
  font-size: 11px;
  font-weight: 800;
  text-transform: uppercase;
}
.metric .value {
  margin-top: 3px;
  font-size: 18px;
  font-weight: 900;
}
.audit-body { padding: 12px; }
.audit-title {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 16px;
  font-weight: 900;
}
.audit-swatch {
  width: 13px;
  height: 13px;
  border: 1px solid rgba(0,0,0,0.12);
}
.audit-kv {
  display: grid;
  grid-template-columns: 120px 1fr;
  gap: 8px 10px;
  margin-top: 12px;
  font-size: 13px;
}
.audit-kv div:nth-child(odd) {
  color: var(--muted);
  font-weight: 800;
}
.tabs {
  display: flex;
  gap: 4px;
  padding: 8px;
  border-bottom: 1px solid var(--line);
  background: #fbfcfd;
}
.tab {
  border: 1px solid transparent;
  background: transparent;
  padding: 6px 8px;
  font-size: 12px;
}
.tab.active {
  border-color: var(--line);
  background: #fff;
}
.tab-pane { display: none; padding: 10px 12px; }
.tab-pane.active { display: block; }
.source-table { max-height: 330px; overflow: auto; }
.small { color: var(--muted); font-size: 12px; line-height: 1.35; }
.modal-backdrop[hidden] { display: none; }
.modal-backdrop {
  position: fixed;
  inset: 0;
  z-index: 60;
  display: grid;
  place-items: center;
  background: rgba(15, 23, 42, 0.42);
  padding: 18px;
}
.modal {
  width: min(480px, 100%);
  background: #fff;
  border: 1px solid var(--line);
  border-radius: 8px;
  box-shadow: 0 22px 60px rgba(15, 23, 42, 0.28);
  padding: 16px;
}
.modal-header {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  gap: 14px;
  margin-bottom: 14px;
}
.modal-title {
  margin: 0;
  font-size: 17px;
  line-height: 1.25;
}
.modal-grid {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  gap: 10px;
  margin-bottom: 12px;
}
.modal-field label,
.comment-field label {
  display: block;
  color: var(--muted);
  font-size: 11px;
  font-weight: 900;
  text-transform: uppercase;
  margin-bottom: 5px;
}
.modal-value,
.modal-field input,
.comment-field textarea {
  width: 100%;
  border: 1px solid #cbd5df;
  border-radius: 6px;
  background: #fff;
  padding: 8px;
  font-weight: 800;
}
.modal-value {
  min-height: 36px;
  background: #f8fafc;
}
.comment-field textarea {
  min-height: 86px;
  resize: vertical;
  font-weight: 500;
}
.modal-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  margin-top: 12px;
}
@media (max-width: 760px) {
  .app-header { align-items: flex-start; flex-direction: column; }
  .actions { justify-content: flex-start; }
  .workspace { padding: 10px; }
  .table-panel { margin: 0 10px 10px; }
  .modal-grid { grid-template-columns: 1fr; }
}
</style>
</head>
<body>
<header class="app-header">
  <div class="title-block">
    <div class="eyebrow">Peer Analysis</div>
    <h1 id="pageTitle"></h1>
  </div>
  <div class="actions">
    <button class="icon-btn primary" id="copyPngBtn" title="Copy graph as PNG" aria-label="Copy graph as PNG">
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <rect x="9" y="9" width="13" height="13" rx="2"></rect>
        <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
      </svg>
    </button>
    <button class="icon-btn" id="downloadJpgBtn" title="Download graph as JPG" aria-label="Download graph as JPG">
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <path d="M7 10l5 5 5-5"></path>
        <path d="M12 15V3"></path>
      </svg>
    </button>
    <button class="icon-btn" id="saveHtmlBtn" title="Save current HTML" aria-label="Save current HTML">
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2Z"></path>
        <path d="M17 21v-8H7v8"></path>
        <path d="M7 3v5h8"></path>
      </svg>
    </button>
  </div>
</header>

<main class="workspace">
  <section class="main-stack">
    <div class="panel chart-panel">
      <div class="chart-wrap">
        <svg id="chartSvg" viewBox="0 0 1180 650" role="img" aria-label="Attributed growth by segment chart"></svg>
      </div>
    </div>

    <details class="panel drawer" id="chartSourceDrawer">
      <summary>
        <span class="drawer-title">Chart Source</span>
      </summary>
      <div class="drawer-body">
        <div class="drawer-controls source-controls">
          <button class="icon-btn ghost" id="resetSourceBtn" title="Reset chart source" aria-label="Reset chart source">
            <svg viewBox="0 0 24 24" aria-hidden="true">
              <path d="M3 12a9 9 0 1 0 3-6.7"></path>
              <path d="M3 3v6h6"></path>
            </svg>
          </button>
          <button class="icon-btn" id="copySourceBtn" title="Copy chart source table" aria-label="Copy chart source table">
            <svg viewBox="0 0 24 24" aria-hidden="true">
              <rect x="9" y="9" width="13" height="13" rx="2"></rect>
              <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
            </svg>
          </button>
        </div>
        <div class="table-shell">
          <table id="editTable"></table>
        </div>
        <div class="adjustment-log" id="adjustmentLog"></div>
      </div>
    </details>

    <details class="panel drawer" id="supportingDataDrawer">
      <summary>
        <span class="drawer-title">Work & Calculations</span>
      </summary>
      <div class="drawer-body">
        <div class="drawer-controls calc-controls">
          <select id="supportingDataSelect" aria-label="Supporting data view">
            <option value="raw">1. Raw data</option>
            <option value="segment">2. Segment calculations</option>
            <option value="bank">3. Bank totals</option>
          </select>
          <button class="icon-btn" id="copySupportingBtn" title="Copy selected calculations table" aria-label="Copy selected calculations table">
            <svg viewBox="0 0 24 24" aria-hidden="true">
              <rect x="9" y="9" width="13" height="13" rx="2"></rect>
              <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
            </svg>
          </button>
        </div>
        <div id="supportingDataView"></div>
      </div>
    </details>
  </section>
</main>

<div class="modal-backdrop" id="adjustModal" hidden>
  <div class="modal" role="dialog" aria-modal="true" aria-labelledby="adjustTitle">
    <div class="modal-header">
      <h2 class="modal-title" id="adjustTitle">Adjust value</h2>
      <button class="icon-btn ghost" id="closeAdjustBtn" title="Close" aria-label="Close adjustment dialog">
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path d="M18 6 6 18"></path>
          <path d="m6 6 12 12"></path>
        </svg>
      </button>
    </div>
    <div class="modal-grid">
      <div class="modal-field">
        <label for="adjustOriginal">Original</label>
        <div class="modal-value" id="adjustOriginal"></div>
      </div>
      <div class="modal-field">
        <label for="adjustValueInput">Adjustment value</label>
        <input id="adjustValueInput" type="text" inputmode="decimal" autocomplete="off">
      </div>
      <div class="modal-field">
        <label for="adjustDelta">Change</label>
        <div class="modal-value" id="adjustDelta"></div>
      </div>
    </div>
    <div class="comment-field">
      <label for="adjustComment">Comment</label>
      <textarea id="adjustComment" placeholder="Reason for adjustment"></textarea>
    </div>
    <div class="modal-actions">
      <button id="cancelAdjustBtn">Cancel</button>
      <button class="primary" id="doneAdjustBtn">Done</button>
    </div>
  </div>
</div>

<script id="savedState" type="application/json">null</script>
<script>
const initialData = __APP_DATA__;
const savedSnapshot = readSavedSnapshot();
let state = savedSnapshot?.state || JSON.parse(JSON.stringify(initialData));
let comments = savedSnapshot?.comments || {};
let activeAdjustment = null;

const svgNS = 'http://www.w3.org/2000/svg';
const el = id => document.getElementById(id);
const fmtPct = value => `${Math.round(Number(value) || 0)}%`;
const fmtAbs = value => `${Number(value) >= 0 ? '+' : '-'}${Math.abs(Number(value) || 0).toFixed(1)}BN`;
const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
const sum = values => values.reduce((acc, value) => acc + (Number(value) || 0), 0);

function readSavedSnapshot() {
  const node = document.getElementById('savedState');
  if (!node) return null;
  const raw = node.textContent.trim();
  if (!raw || raw === 'null') return null;
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

function setStatus(text) {
  if (text) console.info(text);
}

function svgEl(name, attrs = {}, parent) {
  const node = document.createElementNS(svgNS, name);
  for (const [key, value] of Object.entries(attrs)) {
    if (value !== undefined && value !== null) node.setAttribute(key, String(value));
  }
  if (parent) parent.appendChild(node);
  return node;
}

function segmentColor(key) {
  return state.segments.find(segment => segment.key === key)?.color || '#333333';
}

function segmentTextColor(key) {
  return state.segments.find(segment => segment.key === key)?.textColor || '#ffffff';
}

function bankPositiveTotal(bank) {
  return sum(state.segments.map(segment => Math.max(0, Number(bank.segments[segment.key]) || 0)));
}

function bankNegativeTotal(bank) {
  return sum(state.segments.map(segment => Math.max(0, -(Number(bank.segments[segment.key]) || 0))));
}

function renderChart() {
  const svg = el('chartSvg');
  svg.innerHTML = '';

  const width = 1180;
  const height = 650;
  const plot = { left: 55, right: 1125, top: 150, bottom: 462 };
  const logoY = 516;
  const barWidth = 104;
  const labelW = 46;
  const labelH = 28;
  const maxPos = Math.max(1, ...state.banks.map(bankPositiveTotal));
  const maxNeg = Math.max(0, ...state.banks.map(bankNegativeTotal));
  const axisMax = maxPos * 1.08;
  const axisMin = maxNeg > 0 ? -Math.max(1.4, maxNeg * 1.4) : 0;
  const yFor = value => plot.top + ((axisMax - value) / (axisMax - axisMin)) * (plot.bottom - plot.top);
  const zeroY = yFor(0);
  const bankStep = (plot.right - plot.left) / state.banks.length;

  svgEl('rect', { x: 0, y: 0, width, height, fill: '#ffffff' }, svg);
  svgEl('text', { x: 18, y: 30, fill: '#ffffff', 'font-size': 1 }, svg).textContent = state.title;
  drawSvgLegend(svg, 16, 29);

  const attributedBanks = state.banks.map((bank, index) => bank.attributed ? index : -1).filter(index => index >= 0);
  if (attributedBanks.length) {
    const bannerW = 318;
    const bannerH = 26;
    const bannerX = width - bannerW - 22;
    const bannerY = 14;
    svgEl('rect', { x: bannerX, y: bannerY, width: bannerW, height: bannerH, fill: '#bdbdbd' }, svg);
    const text = svgEl('text', { x: bannerX + bannerW / 2, y: bannerY + 18, 'text-anchor': 'middle', fill: '#111827', 'font-size': 17, 'font-weight': 800 }, svg);
    text.textContent = state.attributedLabel;
  }

  svgEl('line', { x1: plot.left - 25, x2: plot.right + 25, y1: zeroY, y2: zeroY, stroke: '#8c8c8c', 'stroke-width': 2 }, svg);

  state.banks.forEach((bank, bankIndex) => {
    const center = plot.left + bankIndex * bankStep + bankStep / 2;
    const x = center - barWidth / 2;
    let posStack = 0;
    let negStack = 0;
    const labels = [];

    const absText = svgEl('text', { x: center, y: 72, 'text-anchor': 'middle', fill: '#000', 'font-size': 27, 'font-weight': 900 }, svg);
    absText.textContent = fmtAbs(bank.absoluteGrowthBn);
    const pctText = svgEl('text', { x: center, y: 110, 'text-anchor': 'middle', fill: '#000', 'font-size': 24, 'font-weight': 900 }, svg);
    pctText.textContent = fmtPct(bank.totalGrowthPct);

    state.segments.forEach((segment, segmentIndex) => {
      const rawValue = Number(bank.segments[segment.key]) || 0;
      let rectY;
      let rectH;
      let labelY;
      let connectorY;

      if (rawValue > 0) {
        const yTop = yFor(posStack + rawValue);
        const yBase = yFor(posStack);
        rectY = yTop;
        rectH = Math.max(1, yBase - yTop);
        labelY = yTop + rectH / 2;
        connectorY = labelY;
        posStack += rawValue;
      } else if (rawValue < 0) {
        const yTop = yFor(-negStack);
        const yBase = yFor(-negStack + rawValue);
        rectY = yTop;
        rectH = Math.max(1, yBase - yTop);
        labelY = yTop + rectH / 2;
        connectorY = labelY;
        negStack += Math.abs(rawValue);
      } else {
        const yLine = yFor(posStack > 0 ? posStack : -negStack);
        rectY = yLine - 2;
        rectH = 4;
        labelY = yLine;
        connectorY = yLine;
      }

      svgEl('rect', {
        x,
        y: rectY,
        width: barWidth,
        height: rectH,
        fill: segment.color,
        rx: 0,
        'data-bank-index': bankIndex,
        'data-segment-key': segment.key
      }, svg);

      labels.push({
        bankIndex,
        segment,
        segmentIndex,
        value: rawValue,
        x: center,
        y: labelY,
        connectorY,
        segmentHeight: rectH,
        text: fmtPct(rawValue)
      });
    });

    placeAndDrawLabels(svg, labels, x, barWidth, labelW, labelH);

    if (bank.logo) {
      const maxLogoW = bank.bank === 'CIBC' ? 150 : 124;
      const maxLogoH = 36;
      const scale = Math.min(maxLogoW / bank.logoWidth, maxLogoH / bank.logoHeight);
      const logoW = bank.logoWidth * scale;
      const logoH = bank.logoHeight * scale;
      svgEl('image', {
        href: bank.logo,
        x: center - logoW / 2,
        y: logoY - logoH / 2,
        width: logoW,
        height: logoH,
        preserveAspectRatio: 'xMidYMid meet'
      }, svg);
    } else {
      const logoText = svgEl('text', { x: center, y: logoY + 6, 'text-anchor': 'middle', fill: bank.brandColor, 'font-size': 24, 'font-weight': 900 }, svg);
      logoText.textContent = bank.shortName;
    }
  });
}

function drawSvgLegend(svg, x, y) {
  let cursorX = x;
  state.segments.forEach(segment => {
    svgEl('rect', { x: cursorX, y: y - 12, width: 10, height: 10, fill: segment.color }, svg);
    const text = svgEl('text', { x: cursorX + 16, y: y - 3, fill: '#111827', 'font-size': 16, 'font-weight': 900 }, svg);
    text.textContent = segment.label;
    cursorX += Math.max(58, segment.label.length * 9 + 30);
  });
}

function placeAndDrawLabels(svg, labels, barX, barWidth, labelW, labelH) {
  const sorted = [...labels].sort((a, b) => a.y - b.y);
  const clusters = [];
  let cluster = [];
  const collisionGap = labelH + 3;

  sorted.forEach(label => {
    if (!cluster.length || Math.abs(label.y - cluster[cluster.length - 1].y) < collisionGap) {
      cluster.push(label);
    } else {
      clusters.push(cluster);
      cluster = [label];
    }
  });
  if (cluster.length) clusters.push(cluster);

  clusters.forEach(group => {
    const placements = inBarPlacements(group, barX, barWidth, labelW);
    group.forEach(label => {
      const boxX = placements.get(label) ?? barX + (barWidth - labelW) / 2;
      drawLabel(svg, label, boxX, label.y - labelH / 2, labelW, labelH);
    });
  });
}

function inBarPlacements(group, barX, barWidth, labelW) {
  const pad = 3;
  const left = barX + pad;
  const right = barX + barWidth - labelW - pad;
  const center = barX + (barWidth - labelW) / 2;
  const placements = new Map();

  if (group.length <= 1) {
    placements.set(group[0], center);
    return placements;
  }

  if (group.length === 2) {
    placements.set(group[0], left);
    placements.set(group[1], right);
    return placements;
  }

  const centerLabel = [...group].sort(
    (a, b) => Math.abs(b.value) - Math.abs(a.value) || b.segmentHeight - a.segmentHeight
  )[0];
  placements.set(centerLabel, center);

  const sideLabels = group.filter(label => label !== centerLabel);
  sideLabels.forEach((label, index) => {
    placements.set(label, index % 2 === 0 ? left : right);
  });
  return placements;
}

function drawLabel(svg, label, boxX, boxY, labelW, labelH) {
    svgEl('rect', {
      x: boxX,
      y: boxY,
      width: labelW,
      height: labelH,
      fill: label.segment.color,
      stroke: label.segment.color,
      rx: 3,
      'data-bank-index': label.bankIndex,
      'data-segment-key': label.segment.key
    }, svg);
    const text = svgEl('text', { x: boxX + labelW / 2, y: boxY + 21, 'text-anchor': 'middle', fill: label.segment.textColor || '#fff', 'font-size': 20, 'font-weight': 900, 'pointer-events': 'none' }, svg);
    text.textContent = label.text;
}

function renderEditTable() {
  const headers = ['Bank', 'Abs Growth', 'YoY Growth', ...state.segments.map(segment => segment.label), 'Source'];
  const rows = state.banks.map((bank, bankIndex) => `
    <tr>
      <td><strong>${esc(bank.bank)}</strong></td>
      ${sourceCellHtml(bankIndex, 'absoluteGrowthBn', bank.absoluteGrowthBn)}
      ${sourceCellHtml(bankIndex, 'totalGrowthPct', bank.totalGrowthPct)}
      ${state.segments.map(segment => sourceCellHtml(bankIndex, segment.key, bank.segments[segment.key], true)).join('')}
      <td class="ref">${esc(bank.refs.bank)}</td>
    </tr>
  `).join('');
  el('editTable').innerHTML = `<thead><tr>${headers.map(header => `<th>${esc(header)}</th>`).join('')}</tr></thead><tbody>${rows}</tbody>`;

  el('editTable').querySelectorAll('.cell-display').forEach(button => {
    button.addEventListener('click', event => {
      const target = event.currentTarget;
      const bankIndex = Number(target.dataset.bank);
      const metric = target.dataset.metric;
      openAdjustmentModal(bankIndex, metric, target.dataset.segment === 'true');
    });
  });
  renderAdjustmentLog();
}

function sourceCellHtml(bankIndex, metric, value, isSegment = false) {
  const key = cellKey(bankIndex, metric);
  const adjusted = isAdjusted(bankIndex, metric);
  const delta = currentMetricValue(bankIndex, metric) - originalMetricValue(bankIndex, metric);
  const note = comments[key] ? '<span class="comment-mark" title="Adjustment comment">●</span>' : '';
  return `
    <td class="numeric ${adjusted ? 'adjusted-cell' : ''}">
      <button class="cell-display" type="button" data-bank="${bankIndex}" data-metric="${esc(metric)}" data-segment="${isSegment ? 'true' : 'false'}" title="Adjust value">
        <span class="cell-value">${esc(formatMetricValue(metric, value))}</span>
        <span class="delta-slot">${adjusted ? deltaHtml(metric, delta) : ''}${note}</span>
      </button>
    </td>
  `;
}

function cellKey(bankIndex, metric) {
  return `${bankIndex}:${metric}`;
}

function originalMetricValue(bankIndex, metric) {
  const bank = initialData.banks[bankIndex];
  if (!bank) return 0;
  return metric in bank ? Number(bank[metric]) || 0 : Number(bank.segments[metric]) || 0;
}

function currentMetricValue(bankIndex, metric) {
  const bank = state.banks[bankIndex];
  if (!bank) return 0;
  return metric in bank ? Number(bank[metric]) || 0 : Number(bank.segments[metric]) || 0;
}

function isAdjusted(bankIndex, metric) {
  return currentMetricValue(bankIndex, metric) !== originalMetricValue(bankIndex, metric);
}

function metricLabel(metric) {
  if (metric === 'absoluteGrowthBn') return 'Abs Growth';
  if (metric === 'totalGrowthPct') return 'YoY Growth';
  return state.segments.find(segment => segment.key === metric)?.label || metric;
}

function renderAdjustmentLog() {
  const adjusted = adjustedEntries();
  const node = el('adjustmentLog');
  if (!adjusted.length) {
    node.innerHTML = 'No adjustments. Values are currently aligned to the generated source.';
    return;
  }
  node.innerHTML = `
    <strong>${adjusted.length} adjusted value${adjusted.length === 1 ? '' : 's'}</strong>
    <table style="margin-top:6px">
      <thead><tr><th>Cell</th><th>Original</th><th>Current</th><th>Comment</th></tr></thead>
      <tbody>${adjusted.map(item => `
        <tr>
          <td>${esc(item.label)}</td>
          <td class="numeric">${esc(item.original)}</td>
          <td class="numeric">${esc(item.current)}</td>
          <td>${esc(item.comment || '')}</td>
        </tr>
      `).join('')}</tbody>
    </table>
  `;
}

function adjustedEntries() {
  const entries = [];
  state.banks.forEach((bank, bankIndex) => {
    ['absoluteGrowthBn', 'totalGrowthPct', ...state.segments.map(segment => segment.key)].forEach(metric => {
      if (!isAdjusted(bankIndex, metric)) return;
      entries.push({
        label: `${bank.bank} ${metricLabel(metric)}`,
        original: formatMetricValue(metric, originalMetricValue(bankIndex, metric)),
        current: formatMetricValue(metric, currentMetricValue(bankIndex, metric)),
        comment: comments[cellKey(bankIndex, metric)] || '',
      });
    });
  });
  return entries;
}

function formatMetricValue(metric, value) {
  return metric === 'absoluteGrowthBn' ? fmtAbs(value) : fmtPct(value);
}

function formatRawValue(value) {
  const number = Number(value) || 0;
  return Number.isInteger(number) ? String(number) : String(Number(number.toFixed(1)));
}

function formatDeltaValue(metric, delta) {
  const suffix = metric === 'absoluteGrowthBn' ? 'BN' : '%';
  const rounded = Math.abs(delta) < 0.05 ? 0 : Number(delta.toFixed(1));
  const absValue = Math.abs(rounded);
  const display = Number.isInteger(absValue) ? String(absValue) : String(absValue.toFixed(1));
  return `${rounded >= 0 ? '+' : '-'}${display}${suffix}`;
}

function formatCompactDeltaValue(metric, delta) {
  const suffix = metric === 'absoluteGrowthBn' ? 'BN' : '%';
  const rounded = Math.abs(delta) < 0.05 ? 0 : Number(delta.toFixed(1));
  const absValue = Math.abs(rounded);
  const display = Number.isInteger(absValue) ? String(absValue) : String(absValue.toFixed(1));
  return `${display}${suffix}`;
}

function deltaHtml(metric, delta) {
  const rounded = Number(delta.toFixed(1));
  if (Math.abs(rounded) < 0.0001) return '';
  const direction = rounded > 0 ? 'up' : 'down';
  const arrow = rounded > 0 ? '&#9650;' : '&#9660;';
  return `<span class="delta ${direction}" title="${esc(formatDeltaValue(metric, rounded))}"><span>${arrow}</span><span>${esc(formatCompactDeltaValue(metric, rounded))}</span></span>`;
}

function openAdjustmentModal(bankIndex, metric, isSegment) {
  const bank = state.banks[bankIndex];
  if (!bank) return;
  activeAdjustment = { bankIndex, metric, isSegment };
  const key = cellKey(bankIndex, metric);
  const original = originalMetricValue(bankIndex, metric);
  const current = currentMetricValue(bankIndex, metric);
  el('adjustTitle').textContent = `${bank.bank} ${metricLabel(metric)}`;
  el('adjustOriginal').textContent = formatMetricValue(metric, original);
  el('adjustValueInput').value = formatRawValue(current);
  el('adjustComment').value = comments[key] || '';
  el('adjustModal').hidden = false;
  updateAdjustmentDelta();
  window.setTimeout(() => {
    el('adjustValueInput').focus();
    el('adjustValueInput').select();
  }, 0);
}

function updateAdjustmentDelta() {
  if (!activeAdjustment) return;
  const value = Number(el('adjustValueInput').value);
  const node = el('adjustDelta');
  if (!Number.isFinite(value)) {
    node.innerHTML = '<span style="color:#b42318">Invalid number</span>';
    return;
  }
  const original = originalMetricValue(activeAdjustment.bankIndex, activeAdjustment.metric);
  const delta = value - original;
  node.innerHTML = Math.abs(Number(delta.toFixed(1))) < 0.0001 ? 'No change' : deltaHtml(activeAdjustment.metric, delta);
}

function closeAdjustmentModal() {
  activeAdjustment = null;
  el('adjustModal').hidden = true;
}

function commitAdjustment() {
  if (!activeAdjustment) return;
  const value = Number(el('adjustValueInput').value);
  if (!Number.isFinite(value)) {
    setStatus('Enter a valid number');
    return;
  }

  const { bankIndex, metric, isSegment } = activeAdjustment;
  const bank = state.banks[bankIndex];
  if (isSegment) {
    bank.segments[metric] = value;
  } else {
    bank[metric] = value;
  }

  const key = cellKey(bankIndex, metric);
  if (isAdjusted(bankIndex, metric)) {
    const note = el('adjustComment').value.trim();
    if (note) comments[key] = note;
    else delete comments[key];
  } else {
    delete comments[key];
  }

  closeAdjustmentModal();
  renderAll();
  setStatus('Adjustment applied');
}

function renderSupportingData() {
  const view = el('supportingDataSelect').value;
  const rows = supportingRows(view);
  const headers = Object.keys(rows[0] || {});
  el('supportingDataView').innerHTML = `
    <div class="small" style="margin-bottom:8px">${supportingDescription(view)}</div>
    <div class="supporting-table-wrap">
      <table class="excel-grid">
        <thead><tr>${headers.map(header => `<th>${esc(header)}</th>`).join('')}</tr></thead>
        <tbody>${rows.map(row => `<tr>${headers.map(header => `<td>${esc(row[header])}</td>`).join('')}</tr>`).join('')}</tbody>
      </table>
    </div>
  `;
}

function supportingDescription(view) {
  if (view === 'raw') return 'Example raw pull rows. The real pipeline can inject transaction/product/geography source records here.';
  if (view === 'segment') return 'Example segment calculations that combine raw metrics into segment contributions.';
  return 'Example bank-wide totals combining segment metrics into top-line growth figures.';
}

function supportingRows(view) {
  if (view === 'raw') {
    return state.banks.flatMap(bank => state.segments.slice(0, 3).map(segment => ({
      Bank: bank.bank,
      Source: 'Example source extract',
      Metric: `${segment.label} revenue growth proxy`,
      Value: fmtPct((Number(bank.segments[segment.key]) || 0) + 1),
      Reference: bank.refs[segment.key],
    })));
  }
  if (view === 'segment') {
    return state.banks.flatMap(bank => state.segments.map(segment => ({
      Bank: bank.bank,
      Segment: segment.label,
      Input: fmtPct(Number(bank.segments[segment.key]) || 0),
      Calculation: 'Revenue contribution + mix impact + normalization',
      Output: fmtPct(Number(bank.segments[segment.key]) || 0),
    })));
  }
  return state.banks.map(bank => {
    const segmentTotal = sum(state.segments.map(segment => Number(bank.segments[segment.key]) || 0));
    return {
      Bank: bank.bank,
      'Segment total': fmtPct(segmentTotal),
      'Reported YoY': fmtPct(bank.totalGrowthPct),
      'Abs growth': fmtAbs(bank.absoluteGrowthBn),
      Calculation: 'Segment totals reconciled to reported bank total',
    };
  });
}

function supportingDataTable() {
  const view = el('supportingDataSelect').value;
  const rows = supportingRows(view);
  const headers = Object.keys(rows[0] || {});
  return {
    name: supportingTableName(view),
    headers,
    rows: rows.map(row => headers.map(header => row[header])),
  };
}

function supportingTableName(view) {
  if (view === 'raw') return 'Raw data';
  if (view === 'segment') return 'Segment calculations';
  return 'Bank totals';
}

function renderAll() {
  renderChart();
  renderEditTable();
  renderSupportingData();
}

function chartSvgBlob() {
  const svg = el('chartSvg').cloneNode(true);
  svg.setAttribute('xmlns', svgNS);
  const text = new XMLSerializer().serializeToString(svg);
  return new Blob([text], { type: 'image/svg+xml;charset=utf-8' });
}

async function svgToImageBlob(type = 'image/png', quality = 1) {
  const blob = chartSvgBlob();
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.decoding = 'async';
    const loaded = new Promise((resolve, reject) => {
      img.onload = resolve;
      img.onerror = reject;
    });
    img.src = url;
    await loaded;
    const canvas = document.createElement('canvas');
    canvas.width = 2360;
    canvas.height = 1300;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    return await new Promise(resolve => canvas.toBlob(resolve, type, quality));
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function svgToPngBlob() {
  return svgToImageBlob('image/png', 1);
}

async function svgToJpgBlob() {
  return svgToImageBlob('image/jpeg', 0.95);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function copyPng() {
  const blob = await svgToPngBlob();
  if (navigator.clipboard && window.ClipboardItem) {
    try {
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
      setStatus('PNG copied');
      return;
    } catch (error) {
      downloadBlob(blob, 'attributed-growth-by-segment-chart.png');
      setStatus('Clipboard blocked; PNG downloaded');
      return;
    }
  }
  downloadBlob(blob, 'attributed-growth-by-segment-chart.png');
  setStatus('PNG downloaded');
}

async function downloadJpg() {
  const blob = await svgToJpgBlob();
  downloadBlob(blob, 'attributed-growth-by-segment-chart.jpg');
  setStatus('JPG downloaded');
}

function chartSourceTable() {
  const headers = ['Bank', 'Abs Growth', 'YoY Growth', ...state.segments.map(segment => segment.label), 'Source'];
  const rows = state.banks.map(bank => [
    bank.bank,
    formatMetricValue('absoluteGrowthBn', bank.absoluteGrowthBn),
    formatMetricValue('totalGrowthPct', bank.totalGrowthPct),
    ...state.segments.map(segment => formatMetricValue(segment.key, bank.segments[segment.key])),
    bank.refs.bank,
  ]);
  return { headers, rows };
}

function clipboardCell(value) {
  return String(value ?? '').replace(/\t/g, ' ').replace(/\r?\n/g, ' ');
}

async function copySourceTable() {
  const table = chartSourceTable();
  const text = [table.headers, ...table.rows]
    .map(row => row.map(clipboardCell).join('\t'))
    .join('\n');

  if (navigator.clipboard) {
    try {
      await navigator.clipboard.writeText(text);
      setStatus('Chart source table copied');
      return;
    } catch (error) {
      // Fall through to the textarea fallback below.
    }
  }

  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.setAttribute('readonly', '');
  textarea.style.position = 'fixed';
  textarea.style.left = '-9999px';
  document.body.appendChild(textarea);
  textarea.select();
  const copied = document.execCommand('copy');
  textarea.remove();
  setStatus(copied ? 'Chart source table copied' : 'Copy blocked');
}

async function copySupportingTable() {
  const table = supportingDataTable();
  const text = [table.headers, ...table.rows]
    .map(row => row.map(clipboardCell).join('\t'))
    .join('\n');

  if (navigator.clipboard) {
    try {
      await navigator.clipboard.writeText(text);
      setStatus(`${table.name} copied`);
      return;
    } catch (error) {
      // Fall through to the textarea fallback below.
    }
  }

  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.setAttribute('readonly', '');
  textarea.style.position = 'fixed';
  textarea.style.left = '-9999px';
  document.body.appendChild(textarea);
  textarea.select();
  const copied = document.execCommand('copy');
  textarea.remove();
  setStatus(copied ? `${table.name} copied` : 'Copy blocked');
}

function resetData() {
  closeAdjustmentModal();
  state = JSON.parse(JSON.stringify(initialData));
  comments = {};
  renderAll();
  setStatus('Reset to input');
}

function saveCurrentHtml() {
  const savedState = {
    savedAt: new Date().toISOString(),
    state,
    comments,
  };
  const savedNode = document.getElementById('savedState');
  savedNode.textContent = JSON.stringify(savedState).replace(/</g, '\\u003c');
  const html = `<!doctype html>\n${document.documentElement.outerHTML}`;
  downloadBlob(new Blob([html], { type: 'text/html;charset=utf-8' }), 'attributed_growth_by_segment_adjusted.html');
  setStatus('Adjusted HTML saved');
}

function init() {
  el('pageTitle').textContent = state.title;
  renderAll();
  el('copyPngBtn').addEventListener('click', copyPng);
  el('downloadJpgBtn').addEventListener('click', downloadJpg);
  el('saveHtmlBtn').addEventListener('click', saveCurrentHtml);
  el('copySourceBtn').addEventListener('click', copySourceTable);
  el('resetSourceBtn').addEventListener('click', resetData);
  el('supportingDataSelect').addEventListener('change', renderSupportingData);
  el('copySupportingBtn').addEventListener('click', copySupportingTable);
  el('adjustValueInput').addEventListener('input', updateAdjustmentDelta);
  el('closeAdjustBtn').addEventListener('click', closeAdjustmentModal);
  el('cancelAdjustBtn').addEventListener('click', closeAdjustmentModal);
  el('doneAdjustBtn').addEventListener('click', commitAdjustment);
  el('adjustModal').addEventListener('click', event => {
    if (event.target === el('adjustModal')) closeAdjustmentModal();
  });
  document.addEventListener('keydown', event => {
    if (event.key === 'Escape' && !el('adjustModal').hidden) closeAdjustmentModal();
  });
}

init();
</script>
</body>
</html>
"""


def write_html_output(config: dict[str, Any], input_path: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    data = json.dumps(app_data(config, input_path), ensure_ascii=True, separators=(",", ":"))
    output_path.write_text(HTML_TEMPLATE.replace("__APP_DATA__", data), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build the interactive HTML output.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT_PATH)
    parser.add_argument("--output", type=Path, default=HTML_OUTPUT_PATH)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config = load_config(args.input)
    output_path = args.output.resolve()
    write_html_output(config, args.input, output_path)
    print(output_path)


if __name__ == "__main__":
    main()
