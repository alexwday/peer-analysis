import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const SCRIPT_DIR = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_BANK_ORDER = ["RBC", "BMO", "Scotia", "CIBC", "National Bank", "TD"];
const DEFAULT_LOGO_DIR = path.resolve(SCRIPT_DIR, "../assets/bank-logos");
const DEFAULT_LOGO_FILES = {
  RBC: "ry.png",
  BMO: "bmo.png",
  Scotia: "bns.png",
  CIBC: "cm.png",
  "National Bank": "na.png",
  TD: "td.png",
};
const CENTER_COLS = ["C", "E", "G", "I", "K", "M"];
const CENTER_COL_INDEXES = [2, 4, 6, 8, 10, 12];
const LAYOUT = {
  spacerRow: 3,
  absoluteLabelRow: 4,
  totalLabelRow: 5,
  chartTopCell: "A6",
  chartBottomCell: "N29",
  chartTopRowIndex: 5,
  chartBottomRowIndex: 29,
  logoTopRow: 30,
  logoBottomRow: 31,
  logoAnchorRowIndex: 29,
  sourceTableStartRow: 4,
  previewRange: "A1:Y32",
};
const INPUT_LAYOUT = {
  sheetName: "Data Input",
  segmentStartRow: 5,
  segmentMaxRows: 5,
  dataStartRow: 12,
  dataMaxRows: 6,
  dataSegmentStartColIndex: 3,
};
const BANK_ALIASES = new Map(
  [
    ["rbc", "RBC"],
    ["ry", "RBC"],
    ["royal bank of canada", "RBC"],
    ["bmo", "BMO"],
    ["bank of montreal", "BMO"],
    ["scotia", "Scotia"],
    ["scotiabank", "Scotia"],
    ["bns", "Scotia"],
    ["bank of nova scotia", "Scotia"],
    ["cibc", "CIBC"],
    ["cm", "CIBC"],
    ["canadian imperial bank of commerce", "CIBC"],
    ["national", "National Bank"],
    ["national bank", "National Bank"],
    ["national bank of canada", "National Bank"],
    ["na", "National Bank"],
    ["nbc", "National Bank"],
    ["td", "TD"],
    ["td bank", "TD"],
    ["toronto dominion", "TD"],
    ["toronto dominion bank", "TD"],
  ].map(([alias, bank]) => [normalizeText(alias), bank]),
);

const DEFAULT_CONFIG = {
  sectionNumber: "1",
  sectionTitle: "Revenue Growth by segment",
  sheetName: "Revenue Growth",
  attributedLabel: "Attributed Growth/Contribution",
  zeroLineColor: "#8C8C8C",
  zeroLineWidth: 12700,
  bankOrder: DEFAULT_BANK_ORDER,
  requireCompleteBankOrder: true,
  useLogos: true,
  logoDir: DEFAULT_LOGO_DIR,
  logoFiles: DEFAULT_LOGO_FILES,
  segments: [
    { key: "CB", label: "CB", color: "#40484A", textColor: "#FFFFFF" },
    { key: "IB", label: "IB", color: "#001A52", textColor: "#FFFFFF" },
    { key: "CM", label: "CM", color: "#2D5967", textColor: "#FFFFFF" },
    { key: "WM", label: "WM", color: "#F28C00", textColor: "#000000" },
    { key: "CS_INS", label: "CS + Ins.", color: "#000000", textColor: "#FFFFFF" },
  ],
  // Values are estimates from the screenshot. Replace `banks` with your
  // upstream metric output when exact source values are available.
  banks: [
    {
      bank: "RBC",
      absoluteGrowthBn: 7.4,
      totalGrowthPct: 12,
      highlight: true,
      brandColor: "#0B3A78",
      segments: { CB: 5, IB: 1, CM: 3, WM: 3, CS_INS: 0 },
    },
    {
      bank: "BMO",
      absoluteGrowthBn: 3.0,
      totalGrowthPct: 9,
      brandColor: "#0072A8",
      segments: { CB: 2, IB: 1, CM: 2, WM: 2, CS_INS: 2 },
    },
    {
      bank: "Scotia",
      absoluteGrowthBn: 3.7,
      totalGrowthPct: 11,
      brandColor: "#D71920",
      segments: { CB: 1, IB: 0, CM: 3, WM: 2, CS_INS: 4 },
    },
    {
      bank: "CIBC",
      absoluteGrowthBn: 3.6,
      totalGrowthPct: 13,
      brandColor: "#8B1020",
      segments: { CB: 5, IB: 1, CM: 6, WM: 2, CS_INS: -1 },
    },
    {
      bank: "National Bank",
      shortName: "National",
      absoluteGrowthBn: 1.7,
      totalGrowthPct: 15,
      attributed: true,
      brandColor: "#C6002B",
      segments: { CB: 2, IB: 1, CM: 7, WM: 2, CS_INS: 3 },
    },
    {
      bank: "TD",
      absoluteGrowthBn: 5.9,
      totalGrowthPct: 12,
      attributed: true,
      brandColor: "#2E7D32",
      segments: { CB: 2, IB: 2, CM: 3, WM: 3, CS_INS: 2 },
    },
  ],
};

function normalizeText(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function parseArgs(argv) {
  const args = {
    input: null,
    output: null,
    preview: null,
    createInputTemplate: null,
    createBlankTemplate: null,
  };

  for (let i = 2; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--input") {
      args.input = argv[++i];
    } else if (arg === "--output") {
      args.output = argv[++i];
    } else if (arg === "--preview") {
      args.preview = argv[++i];
    } else if (arg === "--create-input-template") {
      const next = argv[i + 1];
      args.createInputTemplate =
        next && !next.startsWith("--")
          ? argv[++i]
          : path.resolve(
              SCRIPT_DIR,
              "../inputs/revenue-growth-by-segment/revenue_growth_by_segment_input.xlsx",
            );
    } else if (arg === "--create-blank-template") {
      const next = argv[i + 1];
      args.createBlankTemplate =
        next && !next.startsWith("--")
          ? argv[++i]
          : path.resolve(
              SCRIPT_DIR,
              "../templates/revenue-growth-by-segment/revenue_growth_by_segment_blank_template.xlsx",
            );
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }

  return args;
}

async function loadConfig(inputPath) {
  if (!inputPath) {
    return prepareConfig({});
  }

  if (path.extname(inputPath).toLowerCase() === ".xlsx") {
    return prepareConfig(await loadConfigFromInputWorkbook(inputPath));
  }

  const raw = await fs.readFile(inputPath, "utf8");
  return prepareConfig(JSON.parse(raw));
}

function prepareConfig(parsed) {
  const segments = (parsed.segments ?? DEFAULT_CONFIG.segments).map((segment) => ({
    ...(DEFAULT_CONFIG.segments.find(
      (defaultSegment) =>
        defaultSegment.key === segment.key || defaultSegment.label === segment.label,
    ) ?? {}),
    ...segment,
  }));
  const rawBanks = parsed.data
    ? normalizeDataRows(parsed.data, segments)
    : (parsed.banks ?? DEFAULT_CONFIG.banks);
  const bankOrder = parsed.bankOrder ?? DEFAULT_CONFIG.bankOrder;
  const config = {
    ...DEFAULT_CONFIG,
    ...parsed,
    segments,
    bankOrder,
    banks: orderBanks(rawBanks, bankOrder, parsed.requireCompleteBankOrder),
  };

  return {
    ...config,
    logoFiles: { ...DEFAULT_LOGO_FILES, ...normalizeLogoFiles(parsed.logoFiles ?? {}) },
  };
}

function canonicalBankName(name) {
  const normalized = normalizeText(name);
  return BANK_ALIASES.get(normalized) ?? String(name ?? "").trim();
}

function normalizeLogoFiles(logoFiles) {
  return Object.fromEntries(
    Object.entries(logoFiles).map(([bankName, logoFile]) => [
      canonicalBankName(bankName),
      logoFile,
    ]),
  );
}

function normalizeBankRecord(bank) {
  const canonicalName = canonicalBankName(bank.bank);
  const defaults = DEFAULT_CONFIG.banks.find((candidate) => candidate.bank === canonicalName) ?? {};

  return {
    ...defaults,
    ...bank,
    bank: canonicalName,
    shortName: bank.shortName ?? defaults.shortName,
    brandColor: bank.brandColor ?? defaults.brandColor ?? "#111827",
    segments: {
      ...(defaults.segments ?? {}),
      ...(bank.segments ?? {}),
    },
  };
}

function orderBanks(banks, bankOrder, requireCompleteBankOrder = DEFAULT_CONFIG.requireCompleteBankOrder) {
  const normalizedBanks = banks.map(normalizeBankRecord);
  const order = bankOrder.map(canonicalBankName);
  const orderSet = new Set(order);
  const banksByName = new Map(normalizedBanks.map((bank) => [bank.bank, bank]));
  const missing = order.filter((bankName) => !banksByName.has(bankName));
  const unknown = normalizedBanks.filter((bank) => !orderSet.has(bank.bank));

  if (missing.length > 0 && requireCompleteBankOrder !== false) {
    throw new Error(`Missing required bank rows: ${missing.join(", ")}`);
  }
  if (unknown.length > 0) {
    throw new Error(
      `Unknown bank rows: ${unknown.map((bank) => bank.bank).join(", ")}. Add them to bankOrder before rendering.`,
    );
  }
  if (order.length > CENTER_COLS.length) {
    throw new Error(`The fixed chart layout supports up to ${CENTER_COLS.length} banks.`);
  }

  return order.map((bankName) => banksByName.get(bankName)).filter(Boolean);
}

function toNumber(value, fallback = 0) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "string" && value.trim() !== "") {
    const parsed = Number(value.replace(/[%$,BNbn]/g, ""));
    return Number.isFinite(parsed) ? parsed : fallback;
  }
  return fallback;
}

function toBoolean(value) {
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "string") {
    return ["true", "yes", "y", "1"].includes(value.trim().toLowerCase());
  }
  return Boolean(value);
}

function rowValue(row, keys) {
  for (const key of keys) {
    if (Object.hasOwn(row, key)) {
      return row[key];
    }
  }
  return undefined;
}

function normalizeDataRows(rows, segments) {
  return rows.filter((row) => rowValue(row, ["bank", "Bank"])).map((row) => {
    const segmentValues = {};
    for (const segment of segments) {
      segmentValues[segment.key] = toNumber(
        row.segments?.[segment.key] ??
          row.segments?.[segment.label] ??
          rowValue(row, [segment.key, segment.label]),
      );
    }

    const normalized = {
      bank: rowValue(row, ["bank", "Bank"]),
      absoluteGrowthBn: toNumber(
        rowValue(row, ["absoluteGrowthBn", "Absolute Growth ($BN)", "Absolute Growth"]),
      ),
      totalGrowthPct: toNumber(
        rowValue(row, ["totalGrowthPct", "Reported YoY Growth", "YoY Growth"]),
      ),
      segments: segmentValues,
    };

    const optionalFields = [
      ["shortName", ["shortName", "Short Name", "ShortName"]],
      ["brandColor", ["brandColor", "Brand Color"]],
      ["logoPath", ["logoPath", "Logo Path"]],
      ["logoFile", ["logoFile", "Logo File"]],
    ];
    for (const [targetKey, sourceKeys] of optionalFields) {
      const value = rowValue(row, sourceKeys);
      if (value !== undefined && value !== "") {
        normalized[targetKey] = value;
      }
    }

    const attributed = rowValue(row, ["attributed", "Attributed"]);
    if (attributed !== undefined && attributed !== "") {
      normalized.attributed = toBoolean(attributed);
    }

    return normalized;
  });
}

function columnName(indexZeroBased) {
  let index = indexZeroBased + 1;
  let name = "";
  while (index > 0) {
    const remainder = (index - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    index = Math.floor((index - 1) / 26);
  }
  return name;
}

function xmlUnescape(value) {
  return String(value ?? "")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function xmlAttribute(xml, name) {
  const match = xml.match(new RegExp(`${name}="([^"]*)"`));
  return match ? xmlUnescape(match[1]) : undefined;
}

function cellValue(cells, address) {
  const value = cells.get(address);
  if (typeof value === "string") {
    return value.trim();
  }
  return value;
}

async function readSharedStrings(zip) {
  const file = zip.file("xl/sharedStrings.xml");
  if (!file) {
    return [];
  }

  const xml = await file.async("string");
  return [...xml.matchAll(/<(?:\w+:)?si>([\s\S]*?)<\/(?:\w+:)?si>/g)].map((match) => {
    return [...match[1].matchAll(/<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>/g)]
      .map((textMatch) => xmlUnescape(textMatch[1]))
      .join("");
  });
}

async function readXlsxSheetCells(inputPath, sheetName) {
  const zip = await JSZip.loadAsync(await fs.readFile(inputPath));
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!workbookXml || !relsXml) {
    throw new Error(`Could not read workbook metadata from ${inputPath}`);
  }

  const relTargets = new Map(
    [...relsXml.matchAll(/<(?:\w+:)?Relationship\b([^>]*)\/>/g)].map((match) => [
      xmlAttribute(match[1], "Id"),
      xmlAttribute(match[1], "Target"),
    ]),
  );
  const sheetMatch = [...workbookXml.matchAll(/<(?:\w+:)?sheet\b([^>]*)\/>/g)].find(
    (match) => xmlAttribute(match[1], "name") === sheetName,
  );
  if (!sheetMatch) {
    throw new Error(`Could not find sheet "${sheetName}" in ${inputPath}`);
  }

  const relId = xmlAttribute(sheetMatch[1], "r:id");
  const target = relTargets.get(relId);
  if (!target) {
    throw new Error(`Could not resolve sheet "${sheetName}" in ${inputPath}`);
  }

  const sheetPath = target.startsWith("/")
    ? target.replace(/^\//, "")
    : `xl/${target.replace(/^\.\.\//, "")}`;
  const sheetXml = await zip.file(sheetPath)?.async("string");
  if (!sheetXml) {
    throw new Error(`Could not read ${sheetPath} from ${inputPath}`);
  }

  const sharedStrings = await readSharedStrings(zip);
  const cells = new Map();
  const populatedCellXml = sheetXml.replace(/<(?:\w+:)?c\b[^>]*\/>/g, "");
  for (const match of populatedCellXml.matchAll(/<(?:\w+:)?c\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?c>/g)) {
    const attrs = match[1];
    const body = match[2];
    const ref = xmlAttribute(attrs, "r");
    if (!ref) {
      continue;
    }

    const type = xmlAttribute(attrs, "t");
    const valueMatch = body.match(/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/);
    const inlineMatch = body.match(
      /<(?:\w+:)?is>[\s\S]*?<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>[\s\S]*?<\/(?:\w+:)?is>/,
    );
    let value;
    if (type === "s") {
      value = sharedStrings[Number(valueMatch?.[1] ?? 0)] ?? "";
    } else if (type === "str" || type === "inlineStr") {
      value = xmlUnescape(inlineMatch?.[1] ?? valueMatch?.[1] ?? "");
    } else if (type === "b") {
      value = valueMatch?.[1] === "1";
    } else if (valueMatch) {
      const raw = xmlUnescape(valueMatch[1]);
      const numeric = Number(raw);
      value = Number.isFinite(numeric) ? numeric : raw;
    } else if (inlineMatch) {
      value = xmlUnescape(inlineMatch[1]);
    }

    if (value !== undefined) {
      cells.set(ref, value);
    }
  }

  return cells;
}

async function loadConfigFromInputWorkbook(inputPath) {
  const cells = await readXlsxSheetCells(inputPath, INPUT_LAYOUT.sheetName);
  const config = {
    sectionNumber: cellValue(cells, "B1") ?? DEFAULT_CONFIG.sectionNumber,
    sectionTitle: cellValue(cells, "B2") ?? DEFAULT_CONFIG.sectionTitle,
  };

  const segments = [];
  for (let offset = 0; offset < INPUT_LAYOUT.segmentMaxRows; offset += 1) {
    const row = INPUT_LAYOUT.segmentStartRow + offset;
    const key = cellValue(cells, `A${row}`);
    if (!key) {
      continue;
    }
    segments.push({
      key: String(key).trim(),
      label: cellValue(cells, `B${row}`) || String(key).trim(),
    });
  }
  if (segments.length > 0) {
    config.segments = segments;
  }

  const activeSegments = config.segments ?? DEFAULT_CONFIG.segments;
  const rows = [];
  for (let offset = 0; offset < INPUT_LAYOUT.dataMaxRows; offset += 1) {
    const rowNumber = INPUT_LAYOUT.dataStartRow + offset;
    const bank = cellValue(cells, `A${rowNumber}`);
    if (!bank) {
      continue;
    }

    const row = {
      bank,
      absoluteGrowthBn: cellValue(cells, `B${rowNumber}`),
      totalGrowthPct: cellValue(cells, `C${rowNumber}`),
    };

    activeSegments.forEach((segment, index) => {
      const col = columnName(INPUT_LAYOUT.dataSegmentStartColIndex + index);
      row[segment.key] = cellValue(cells, `${col}${rowNumber}`);
    });
    rows.push(row);
  }

  return {
    ...config,
    bankOrder: rows.map((row) => row.bank),
    data: rows,
  };
}

function segmentValue(bank, segmentKey) {
  const value = bank.segments?.[segmentKey];
  return Number.isFinite(value) ? value : 0;
}

function absGrowthLabel(value) {
  return `${value >= 0 ? "+" : "-"}${Math.abs(value).toFixed(1)}BN`;
}

function pctLabel(value) {
  return `${Math.round(value)}%`;
}

async function fileExists(filePath) {
  try {
    const stat = await fs.stat(filePath);
    return stat.isFile();
  } catch {
    return false;
  }
}

async function resolveLogoPath(bank, config) {
  if (config.useLogos === false) {
    return null;
  }

  if (bank.logoPath) {
    const explicitPath = path.resolve(bank.logoPath);
    return (await fileExists(explicitPath)) ? explicitPath : null;
  }

  const logoFile = bank.logoFile ?? config.logoFiles?.[bank.bank];
  if (!logoFile || !config.logoDir) {
    return null;
  }

  const candidate = path.resolve(config.logoDir, logoFile);
  return (await fileExists(candidate)) ? candidate : null;
}

function setColumnWidths(sheet) {
  const widths = {
    A: 4,
    B: 12,
    C: 12,
    D: 12,
    E: 12,
    F: 12,
    G: 12,
    H: 12,
    I: 12,
    J: 12,
    K: 12,
    L: 12,
    M: 12,
    N: 12,
    O: 3,
    P: 15,
    Q: 13,
    R: 11,
    S: 10,
    T: 10,
    U: 10,
    V: 10,
    W: 10,
    X: 11,
    Y: 11,
  };

  for (const [column, width] of Object.entries(widths)) {
    sheet.getRange(`${column}:${column}`).format.columnWidth = width;
  }
}

function sectionHeader(config) {
  return config.sectionNumber
    ? `${config.sectionNumber}. ${config.sectionTitle}`
    : config.sectionTitle;
}

function safeSheetName(config) {
  const rawName = config.sheetName ?? config.sectionTitle ?? "Segment Growth";
  return String(rawName).replace(/[\[\]:*?/\\]/g, " ").slice(0, 31).trim() || "Segment Growth";
}

function styleSectionHeader(sheet, config) {
  sheet.showGridLines = false;

  sheet.mergeCells("A1:Y1");
  sheet.getRange("A1").values = [[sectionHeader(config)]];
  sheet.getRange("A1:Y1").format.fill = { type: "solid", color: "#071734" };
  sheet.getRange("A1:Y1").format.font = { color: "#FFFFFF", bold: true, size: 18 };
  sheet.getRange("A1:Y1").format.rowHeightPx = 36;
  sheet.getRange("A1:Y1").format.verticalAlignment = "middle";
}

function addLegend(sheet, config) {
  const legendCells = ["B2", "C2", "D2", "E2", "F2"];
  config.segments.forEach((segment, index) => {
    const cell = sheet.getRange(legendCells[index]);
    cell.values = [[`\u25A0 ${segment.label}`]];
    cell.format.font = { color: segment.color, size: 11 };
    cell.format.horizontalAlignment = "left";
    cell.format.verticalAlignment = "middle";
  });
  sheet.getRange("A2:Y2").format.rowHeightPx = 24;
  sheet.getRange(`A${LAYOUT.spacerRow}:Y${LAYOUT.spacerRow}`).format.rowHeightPx = 10;
}

function addStackedColumnChart(sheet, config) {
  const chart = sheet.charts.add("bar", {
    title: "",
    titlePlacement: "none",
    hasLegend: false,
    categories: config.banks.map((_, index) => " ".repeat(index + 1)),
    series: config.segments.map((segment) => ({
      name: segment.label,
      values: config.banks.map((bank) => segmentValue(bank, segment.key)),
      fill: { type: "solid", color: segment.color },
      border: { color: "#FFFFFF", width: 0.5 },
    })),
    from: { row: 5, col: 0 },
    extent: { widthPx: 980, heightPx: 470 },
    barOptions: { direction: "column", grouping: "stacked", gapWidth: 70 },
    chartFill: { type: "solid", color: "#FFFFFF" },
    plotAreaFill: { type: "solid", color: "#FFFFFF" },
    dataLabels: {
      showValue: true,
      position: "center",
      numberFormatCode: '0"%"',
      textStyle: { fontSize: 12, bold: false, fill: "#FFFFFF" },
    },
    xAxis: {
      majorGridlines: null,
      line: { fill: "#BDBDBD", style: "solid", width: 1 },
      textStyle: { fontSize: 1, fill: "#FFFFFF" },
    },
    yAxis: {
      minimumScale: -1.5,
      maximumScale: 17,
      majorUnit: 5,
      majorGridlines: null,
      textStyle: { fill: "#FFFFFF" },
      line: { fill: "#FFFFFF", style: "solid", width: 0 },
      numberFormatCode: '0"%"',
    },
  });

  chart.setPosition(LAYOUT.chartTopCell, LAYOUT.chartBottomCell);

  for (let i = 0; i < config.segments.length; i += 1) {
    const series = chart.series.items[i];
    if (!series) {
      continue;
    }
    series.fill = { type: "solid", color: config.segments[i].color };
    series.dataLabels = {
      showValue: true,
      position: "center",
      numberFormatCode: '0"%"',
      textStyle: {
        fontSize: 12,
        bold: false,
        fill: config.segments[i].textColor,
      },
    };
  }
}

function addTopLabels(sheet, config) {
  config.banks.forEach((bank, index) => {
    const col = CENTER_COLS[index];
    if (!col) {
      return;
    }
    const range = `${col}${LAYOUT.absoluteLabelRow}`;
    sheet.getRange(range).values = [[config.templateMode ? "" : absGrowthLabel(bank.absoluteGrowthBn)]];
    sheet.getRange(range).format.font = { bold: true, color: "#000000", size: 14 };
    sheet.getRange(range).format.horizontalAlignment = "center";

    const totalRange = `${col}${LAYOUT.totalLabelRow}`;
    sheet.getRange(totalRange).values = [[config.templateMode ? "" : pctLabel(bank.totalGrowthPct)]];
    sheet.getRange(totalRange).format.font = { bold: true, color: "#000000", size: 12 };
    sheet.getRange(totalRange).format.horizontalAlignment = "center";
  });

  const attributedBanks = config.banks
    .map((bank, index) => ({ bank, index }))
    .filter(({ bank }) => bank.attributed);
  if (attributedBanks.length > 0) {
    const first = CENTER_COLS[attributedBanks[0].index];
    const last = CENTER_COLS[attributedBanks.at(-1).index];
    if (!first || !last) {
      return;
    }
    const labelFirst = first;
    const labelLast = String.fromCharCode(last.charCodeAt(0) + 1);
    sheet.mergeCells(`${labelFirst}2:${labelLast}2`);
    sheet.getRange(`${labelFirst}2:${labelLast}2`).values = [[config.attributedLabel]];
    sheet.getRange(`${labelFirst}2:${labelLast}2`).format.fill = {
      type: "solid",
      color: "#BDBDBD",
    };
    sheet.getRange(`${labelFirst}2:${labelLast}2`).format.font = {
      bold: true,
      color: "#111827",
      size: 13,
    };
    sheet.getRange(`${labelFirst}2:${labelLast}2`).format.horizontalAlignment = "center";
  }
}

async function addBankBranding(sheet, config) {
  for (const [index, bank] of config.banks.entries()) {
    const col = CENTER_COLS[index];
    if (!col) {
      continue;
    }
    const range = `${col}${LAYOUT.logoTopRow}:${col}${LAYOUT.logoBottomRow}`;
    const logoPath = await resolveLogoPath(bank, config);
    bank.logoPath = logoPath ?? undefined;

    sheet.mergeCells(range);
    sheet.getRange(range).values = [[logoPath ? "" : (bank.shortName ?? bank.bank)]];
    sheet.getRange(range).format.font = {
      bold: true,
      color: bank.brandColor,
      size: bank.bank === "National Bank" ? 14 : 18,
    };
    sheet.getRange(range).format.horizontalAlignment = "center";
    sheet.getRange(range).format.verticalAlignment = "middle";
  }
}

function writeSourceTable(sheet, config) {
  const header = [
    "Bank",
    "Absolute Growth ($BN)",
    "Reported YoY Growth",
    ...config.segments.map((segment) => segment.label),
    "Attributed",
  ];
  const rows = config.banks.map((bank) => {
    return [
      bank.bank,
      config.templateMode ? "" : bank.absoluteGrowthBn,
      config.templateMode ? "" : bank.totalGrowthPct,
      ...config.segments.map((segment) =>
        config.templateMode ? "" : segmentValue(bank, segment.key),
      ),
      bank.attributed ? "Yes" : "",
    ];
  });

  const startRow = LAYOUT.sourceTableStartRow;
  const startCol = 15;
  sheet.getRangeByIndexes(startRow - 1, startCol, rows.length + 1, header.length).values = [
    header,
    ...rows,
  ];

  const endCol = String.fromCharCode(65 + startCol + header.length - 1);
  const endRow = startRow + rows.length;
  const tableRange = `P${startRow}:${endCol}${endRow}`;
  const table = sheet.tables.add(tableRange, true, "RevenueGrowthBySegmentSource");
  table.style = "TableStyleMedium2";

  sheet.getRange(`Q${startRow + 1}:Q${endRow}`).setNumberFormat('+0.0"BN";-0.0"BN";0.0"BN"');
  sheet.getRange(`R${startRow + 1}:W${endRow}`).setNumberFormat('0"%"');
  sheet.getRange(`P${startRow}:${endCol}${endRow}`).format.font = { size: 9 };
  sheet.getRange(`P${startRow}:${endCol}${startRow}`).format.font = {
    bold: true,
    color: "#FFFFFF",
    size: 9,
  };
  sheet.getRange(`P:${endCol}`).format.autofitColumns();
}

function addInputTemplateSheet(workbook, config) {
  const sheet = workbook.worksheets.add(INPUT_LAYOUT.sheetName);
  const inputWidthByColumn = {
    A: 20,
    B: 18,
    C: 18,
    D: 10,
    E: 10,
    F: 10,
    G: 10,
    H: 10,
  };
  for (const [column, width] of Object.entries(inputWidthByColumn)) {
    sheet.getRange(`${column}:${column}`).format.columnWidth = width;
  }

  sheet.getRange("A1:B2").values = [
    ["Section Number", config.sectionNumber],
    ["Section Title", config.sectionTitle],
  ];
  sheet.getRange("A1:A2").format.fill = { type: "solid", color: "#D9EAF7" };
  sheet.getRange("A1:A2").format.font = { bold: true, color: "#111827" };

  const segmentHeaderRow = INPUT_LAYOUT.segmentStartRow - 1;
  sheet.getRange(`A${segmentHeaderRow}:B${segmentHeaderRow}`).values = [
    ["Segment Key", "Segment Label"],
  ];
  sheet.getRange(`A${segmentHeaderRow}:B${segmentHeaderRow}`).format.fill = {
    type: "solid",
    color: "#0D6786",
  };
  sheet.getRange(`A${segmentHeaderRow}:B${segmentHeaderRow}`).format.font = {
    bold: true,
    color: "#FFFFFF",
  };
  const segmentRows = config.segments.map((segment) => [
    segment.key,
    segment.label,
  ]);
  sheet.getRangeByIndexes(INPUT_LAYOUT.segmentStartRow - 1, 0, segmentRows.length, 2).values =
    segmentRows;
  sheet.tables.add(
    `A${segmentHeaderRow}:B${segmentHeaderRow + segmentRows.length}`,
    true,
    "RevenueGrowthSegmentsInput",
  );

  const dataHeaders = [
    "Bank",
    "Absolute Growth ($BN)",
    "Reported YoY Growth",
    ...config.segments.map((segment) => segment.key),
  ];
  const dataRows = config.banks.map((bank) => [
    bank.bank,
    bank.absoluteGrowthBn,
    bank.totalGrowthPct,
    ...config.segments.map((segment) => segmentValue(bank, segment.key)),
  ]);
  sheet.getRangeByIndexes(INPUT_LAYOUT.dataStartRow - 2, 0, 1, dataHeaders.length).values = [
    dataHeaders,
  ];
  sheet.getRangeByIndexes(INPUT_LAYOUT.dataStartRow - 1, 0, dataRows.length, dataHeaders.length).values =
    dataRows;
  const dataEndCol = columnName(dataHeaders.length - 1);
  const dataEndRow = INPUT_LAYOUT.dataStartRow + dataRows.length - 1;
  sheet.getRange(`A${INPUT_LAYOUT.dataStartRow - 1}:${dataEndCol}${INPUT_LAYOUT.dataStartRow - 1}`).format.fill = {
    type: "solid",
    color: "#0D6786",
  };
  sheet.getRange(`A${INPUT_LAYOUT.dataStartRow - 1}:${dataEndCol}${INPUT_LAYOUT.dataStartRow - 1}`).format.font = {
    bold: true,
    color: "#FFFFFF",
  };
  sheet.tables.add(
    `A${INPUT_LAYOUT.dataStartRow - 1}:${dataEndCol}${dataEndRow}`,
    true,
    "RevenueGrowthBankDataInput",
  );
  sheet.getRange(`B${INPUT_LAYOUT.dataStartRow}:B${dataEndRow}`).setNumberFormat('+0.0"BN";-0.0"BN";0.0"BN"');
  sheet.getRange(`C${INPUT_LAYOUT.dataStartRow}:C${dataEndRow}`).setNumberFormat('0"%"');
  sheet
    .getRange(`D${INPUT_LAYOUT.dataStartRow}:${dataEndCol}${dataEndRow}`)
    .setNumberFormat('0"%"');
  sheet.getRange(`A1:${dataEndCol}${dataEndRow}`).format.font = { size: 10 };

  return sheet;
}

async function createInputTemplateWorkbook(outputPath) {
  const config = prepareConfig({});
  const workbook = Workbook.create();
  addInputTemplateSheet(workbook, config);
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);
  console.log(outputPath);
}

function blankTemplateConfig() {
  const segments = DEFAULT_CONFIG.segments.map((segment) => ({ ...segment }));
  const emptySegments = Object.fromEntries(segments.map((segment) => [segment.key, 0]));
  return prepareConfig({
    banks: DEFAULT_CONFIG.banks.map((bank) => ({
      ...bank,
      absoluteGrowthBn: 0,
      totalGrowthPct: 0,
      segments: { ...emptySegments },
    })),
    segments,
    templateMode: true,
  });
}

async function renderPreview(workbook, previewPath, sheetName) {
  const image = await workbook.render({
    sheetName,
    range: LAYOUT.previewRange,
    scale: 1.2,
  });
  await fs.writeFile(previewPath, Buffer.from(await image.arrayBuffer()));
}

function removeAxisGridlines(axisXml) {
  return axisXml.replace(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/g, "");
}

function setOrInsertTickLabelPosition(axisXml, value) {
  if (axisXml.includes("<c:tickLblPos")) {
    return axisXml.replace(/<c:tickLblPos val="[^"]*" \/>/, `<c:tickLblPos val="${value}" />`);
  }

  return axisXml.replace(
    /(<c:minorTickMark val="none" \/>)/,
    `$1<c:tickLblPos val="${value}" />`,
  );
}

function setOrInsertAxisLine(axisXml, color, width = 9525) {
  const cleanColor = String(color).replace("#", "");
  const spPr = `<c:spPr><a:ln w="${width}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:srgbClr val="${cleanColor}" /></a:solidFill><a:prstDash val="solid" /></a:ln></c:spPr>`;
  if (axisXml.includes("<c:spPr>")) {
    return axisXml.replace(/<c:spPr>[\s\S]*?<\/c:spPr>/, spPr);
  }

  return axisXml.replace(/(<c:axPos val="[^"]+" \/>)/, `$1${spPr}`);
}

function updateAxisXml(xml, axisName, updater) {
  const pattern = new RegExp(`(<c:${axisName}>[\\s\\S]*?<\\/c:${axisName}>)`);
  return xml.replace(pattern, (axisXml) => updater(axisXml));
}

function labelTextXml(text, textColor) {
  const color = textColor.replace("#", "");
  return `<c:tx><c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" /><a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" /><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:rPr lang="en-US" sz="1200"><a:solidFill><a:srgbClr val="${color}" /></a:solidFill></a:rPr><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx><c:txPr><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" /><a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" /><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:pPr><a:defRPr lang="en-US" sz="1200"><a:solidFill><a:srgbClr val="${color}" /></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>`;
}

function pointLabelXml(index, text, textColor) {
  return `<c:dLbl><c:idx val="${index}" />${labelTextXml(text, textColor)}<c:dLblPos val="ctr" /><c:showLegendKey val="0" /><c:showVal val="0" /><c:showCatName val="0" /><c:showSerName val="0" /><c:showPercent val="0" /><c:showBubbleSize val="0" /><c:showLeaderLines val="0" /></c:dLbl>`;
}

function seriesLabelsXml(segment, config) {
  const labels = config.banks
    .map((bank, index) => {
      const value = segmentValue(bank, segment.key);
      if (Math.round(value) === 0 || value < 0) {
        return "";
      }
      return pointLabelXml(index, pctLabel(value), segment.textColor);
    })
    .join("");
  return `<c:dLbls>${labels}<c:showLegendKey val="0" /><c:showVal val="0" /><c:showCatName val="0" /><c:showSerName val="0" /><c:showPercent val="0" /><c:showBubbleSize val="0" /><c:showLeaderLines val="0" /></c:dLbls>`;
}

function negativePointFormattingXml(segment, config) {
  const fillColor = String(segment.color).replace("#", "");
  return config.banks
    .map((bank, index) => {
      const value = segmentValue(bank, segment.key);
      if (value >= 0) {
        return "";
      }

      return `<c:dPt><c:idx val="${index}" /><c:spPr><a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="${fillColor}" /></a:solidFill></c:spPr></c:dPt>`;
    })
    .join("");
}

function addExplicitPointLabels(xml, config) {
  let seriesIndex = 0;
  return xml.replace(/<c:ser>([\s\S]*?)<\/c:ser>/g, (seriesXml) => {
    const segment = config.segments[seriesIndex];
    seriesIndex += 1;

    if (!segment) {
      return seriesXml;
    }

    const withoutLabels = seriesXml.replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/g, "");
    const labels = seriesLabelsXml(segment, config);
    if (withoutLabels.includes("<c:cat>")) {
      return withoutLabels.replace("<c:cat>", `${labels}<c:cat>`);
    }

    return withoutLabels.replace("</c:ser>", `${labels}</c:ser>`);
  });
}

function removeChartBorder(xml) {
  return xml.replace(
    /<c:spPr><a:ln w="9525" xmlns:a="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/main"><a:solidFill><a:srgbClr val="D9D9D9" \/><\/a:solidFill><a:prstDash val="solid" \/><\/a:ln><\/c:spPr><\/c:chartSpace>/,
    "</c:chartSpace>",
  );
}

function blankCategoryLiteralXml(config) {
  const points = config.banks
    .map((_, index) => `<c:pt idx="${index}"><c:v>${" ".repeat(index + 1)}</c:v></c:pt>`)
    .join("");
  return `<c:cat><c:strLit><c:ptCount val="${config.banks.length}" />${points}</c:strLit></c:cat>`;
}

function zeroValueLiteralXml(config) {
  const points = config.banks
    .map((_, index) => `<c:pt idx="${index}"><c:v>0</c:v></c:pt>`)
    .join("");
  return `<c:val><c:numLit><c:formatCode></c:formatCode><c:ptCount val="${config.banks.length}" />${points}</c:numLit></c:val>`;
}

function zeroLineChartXml(config, catAxisId, valAxisId) {
  const cleanColor = String(config.zeroLineColor).replace("#", "");
  const seriesIndex = config.segments.length;
  return [
    "<c:lineChart>",
    '<c:grouping val="standard" />',
    '<c:varyColors val="0" />',
    "<c:ser>",
    `<c:idx val="${seriesIndex}" />`,
    `<c:order val="${seriesIndex}" />`,
    "<c:tx><c:v>Zero line</c:v></c:tx>",
    `<c:spPr><a:ln w="${config.zeroLineWidth}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:srgbClr val="${cleanColor}" /></a:solidFill><a:prstDash val="solid" /></a:ln></c:spPr>`,
    '<c:marker><c:symbol val="none" /></c:marker>',
    blankCategoryLiteralXml(config),
    zeroValueLiteralXml(config),
    '<c:smooth val="0" />',
    "</c:ser>",
    `<c:axId val="${catAxisId}" />`,
    `<c:axId val="${valAxisId}" />`,
    "</c:lineChart>",
  ].join("");
}

function addZeroLineSeries(xml, config) {
  const barChartMatch = xml.match(/<c:barChart>[\s\S]*?<\/c:barChart>/);
  if (!barChartMatch) {
    return xml;
  }

  const axisIds = [...barChartMatch[0].matchAll(/<c:axId val="([^"]+)" \/>/g)].map(
    (match) => match[1],
  );
  if (axisIds.length < 2) {
    return xml;
  }

  return xml.replace(
    "</c:barChart>",
    `</c:barChart>${zeroLineChartXml(config, axisIds[0], axisIds[1])}`,
  );
}

function setBarSeriesOverlap(xml, overlap = 100) {
  return xml.replace(/<c:barChart>[\s\S]*?<\/c:barChart>/, (barChartXml) => {
    if (barChartXml.includes("<c:overlap")) {
      return barChartXml.replace(/<c:overlap val="[^"]*" \/>/, `<c:overlap val="${overlap}" />`);
    }

    return barChartXml.replace(
      /(<c:gapWidth val="[^"]*" \/>)/,
      `$1<c:overlap val="${overlap}" />`,
    );
  });
}

function disableBarNegativeInversion(xml) {
  return xml.replace(/<c:barChart>[\s\S]*?<\/c:barChart>/, (barChartXml) => {
    return barChartXml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, (seriesXml) => {
      if (seriesXml.includes("<c:invertIfNegative")) {
        return seriesXml.replace(
          /<c:invertIfNegative val="[^"]*" \/>/,
          '<c:invertIfNegative val="0" />',
        );
      }

      return seriesXml.replace(/(<c:spPr>[\s\S]*?<\/c:spPr>)/, '$1<c:invertIfNegative val="0" />');
    });
  });
}

function addNegativePointFormatting(xml, config) {
  let seriesIndex = 0;
  return xml.replace(/<c:barChart>[\s\S]*?<\/c:barChart>/, (barChartXml) => {
    return barChartXml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, (seriesXml) => {
      const segment = config.segments[seriesIndex];
      seriesIndex += 1;

      if (!segment) {
        return seriesXml;
      }

      const pointFormatting = negativePointFormattingXml(segment, config);
      if (!pointFormatting) {
        return seriesXml;
      }

      const withoutExistingPointFormatting = seriesXml.replace(/<c:dPt>[\s\S]*?<\/c:dPt>/g, "");
      if (withoutExistingPointFormatting.includes("<c:invertIfNegative")) {
        return withoutExistingPointFormatting.replace(
          /(<c:invertIfNegative val="[^"]*" \/>)/,
          `$1${pointFormatting}`,
        );
      }

      if (withoutExistingPointFormatting.includes("<c:dLbls>")) {
        return withoutExistingPointFormatting.replace("<c:dLbls>", `${pointFormatting}<c:dLbls>`);
      }

      return withoutExistingPointFormatting.replace(
        /(<c:spPr>[\s\S]*?<\/c:spPr>)/,
        `$1${pointFormatting}`,
      );
    });
  });
}

function zeroNegativeBarValues(xml, config) {
  let seriesIndex = 0;
  return xml.replace(/<c:barChart>[\s\S]*?<\/c:barChart>/, (barChartXml) => {
    return barChartXml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, (seriesXml) => {
      const segment = config.segments[seriesIndex];
      seriesIndex += 1;

      if (!segment) {
        return seriesXml;
      }

      return seriesXml.replace(/<c:val>[\s\S]*?<\/c:val>/, (valXml) => {
        let updated = valXml;
        config.banks.forEach((bank, bankIndex) => {
          if (segmentValue(bank, segment.key) >= 0) {
            return;
          }

          const pointPattern = new RegExp(
            `(<c:pt idx="${bankIndex}"><c:v>)([^<]*)(<\\/c:v><\\/c:pt>)`,
          );
          updated = updated.replace(pointPattern, (_match, prefix, _value, suffix) => {
            return `${prefix}0${suffix}`;
          });
        });
        return updated;
      });
    });
  });
}

function xmlEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function pxToEmu(px) {
  return Math.round(px * 9525);
}

async function pngDimensions(filePath) {
  const bytes = await fs.readFile(filePath);
  if (
    bytes.length >= 24 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47
  ) {
    return {
      width: bytes.readUInt32BE(16),
      height: bytes.readUInt32BE(20),
    };
  }
  return null;
}

async function logoLayout(filePath) {
  const maxWidthPx = 130;
  const maxHeightPx = 40;
  const dimensions = await pngDimensions(filePath);
  if (!dimensions || dimensions.width <= 0 || dimensions.height <= 0) {
    return { widthPx: 96, heightPx: 34 };
  }

  const scale = Math.min(maxWidthPx / dimensions.width, maxHeightPx / dimensions.height);
  return {
    widthPx: Math.round(dimensions.width * scale),
    heightPx: Math.round(dimensions.height * scale),
  };
}

async function logoAnchor(logoPath, bankIndex) {
  const centerCol = CENTER_COL_INDEXES[bankIndex];
  const { widthPx, heightPx } = await logoLayout(logoPath);
  const defaultColumnWidthPx = 84;
  const totalLabelHeightPx = 44;
  let col = centerCol;
  let colOffsetPx = Math.round(defaultColumnWidthPx / 2 - widthPx / 2);

  if (colOffsetPx < 0) {
    col -= 1;
    colOffsetPx += defaultColumnWidthPx;
  }

  return {
    col,
    colOffsetPx,
    row: LAYOUT.logoAnchorRowIndex,
    rowOffsetPx: Math.max(0, Math.round((totalLabelHeightPx - heightPx) / 2)),
    widthPx,
    heightPx,
  };
}

function logoAnchorXml({ relId, picId, name, anchor }) {
  return [
    '<xdr:oneCellAnchor>',
    '<xdr:from>',
    `<xdr:col>${anchor.col}</xdr:col>`,
    `<xdr:colOff>${pxToEmu(anchor.colOffsetPx)}</xdr:colOff>`,
    `<xdr:row>${anchor.row}</xdr:row>`,
    `<xdr:rowOff>${pxToEmu(anchor.rowOffsetPx)}</xdr:rowOff>`,
    '</xdr:from>',
    `<xdr:ext cx="${pxToEmu(anchor.widthPx)}" cy="${pxToEmu(anchor.heightPx)}" />`,
    '<xdr:pic>',
    '<xdr:nvPicPr>',
    `<xdr:cNvPr id="${picId}" name="${xmlEscape(name)}" />`,
    '<xdr:cNvPicPr><a:picLocks noChangeAspect="1" /></xdr:cNvPicPr>',
    '</xdr:nvPicPr>',
    '<xdr:blipFill>',
    `<a:blip r:embed="${relId}" />`,
    '<a:stretch><a:fillRect /></a:stretch>',
    '</xdr:blipFill>',
    '<xdr:spPr>',
    '<a:xfrm><a:off x="0" y="0" /><a:ext cx="0" cy="0" /></a:xfrm>',
    '<a:prstGeom prst="rect"><a:avLst /></a:prstGeom>',
    '</xdr:spPr>',
    '</xdr:pic>',
    '<xdr:clientData />',
    '</xdr:oneCellAnchor>',
  ].join("");
}

function rowHeightPx(rowIndex) {
  if (rowIndex === 0) {
    return 36;
  }
  if (rowIndex === 1) {
    return 24;
  }
  if (rowIndex === 2) {
    return 10;
  }
  return 20;
}

function columnWidthPx(columnIndex) {
  if (columnIndex === 0) {
    return 28;
  }
  if (columnIndex >= 1 && columnIndex <= 13) {
    return 84;
  }
  if (columnIndex === 14) {
    return 21;
  }
  return 64;
}

function rowTopPx(rowIndex) {
  let top = 0;
  for (let row = 0; row < rowIndex; row += 1) {
    top += rowHeightPx(row);
  }
  return top;
}

function columnLeftPx(columnIndex) {
  let left = 0;
  for (let column = 0; column < columnIndex; column += 1) {
    left += columnWidthPx(column);
  }
  return left;
}

function pixelToRowAnchor(yPx) {
  let row = 0;
  let remaining = Math.max(0, yPx);
  while (remaining >= rowHeightPx(row)) {
    remaining -= rowHeightPx(row);
    row += 1;
  }
  return { row, rowOffsetPx: Math.round(remaining) };
}

function pixelToColumnAnchor(xPx) {
  let col = 0;
  let remaining = Math.max(0, xPx);
  while (remaining >= columnWidthPx(col)) {
    remaining -= columnWidthPx(col);
    col += 1;
  }
  return { col, colOffsetPx: Math.round(remaining) };
}

function anchorFromPixels(xPx, yPx, widthPx, heightPx) {
  return {
    ...pixelToColumnAnchor(xPx),
    ...pixelToRowAnchor(yPx),
    widthPx: Math.max(1, Math.round(widthPx)),
    heightPx: Math.max(1, Math.round(heightPx)),
  };
}

function rectangleShapeXml({
  shapeId,
  name,
  anchor,
  fillColor = "#FFFFFF",
  lineColor = "#FFFFFF",
  lineWidthPx = 0,
  textColor,
  label,
  fontSize = 1200,
  bold = true,
}) {
  const fill = fillColor === null ? null : String(fillColor).replace("#", "");
  const line = lineColor === null ? null : String(lineColor).replace("#", "");
  const lineWidth = Math.round(lineWidthPx * 9525);
  const textFill = textColor ? String(textColor).replace("#", "") : null;
  const fillXml = fill ? `<a:solidFill><a:srgbClr val="${fill}" /></a:solidFill>` : "<a:noFill />";
  const lineXml = line
    ? `<a:ln w="${lineWidth}"><a:solidFill><a:srgbClr val="${line}" /></a:solidFill></a:ln>`
    : '<a:ln w="0"><a:noFill /></a:ln>';
  const textBody =
    label && textFill
      ? [
          '<xdr:txBody>',
          '<a:bodyPr anchor="ctr" lIns="0" tIns="0" rIns="0" bIns="0" />',
          '<a:lstStyle />',
          '<a:p>',
          `<a:pPr algn="ctr"><a:defRPr sz="${fontSize}"${bold ? ' b="1"' : ""}><a:solidFill><a:srgbClr val="${textFill}" /></a:solidFill></a:defRPr></a:pPr>`,
          `<a:r><a:rPr lang="en-US" sz="${fontSize}"${bold ? ' b="1"' : ""}><a:solidFill><a:srgbClr val="${textFill}" /></a:solidFill></a:rPr><a:t>${xmlEscape(label)}</a:t></a:r>`,
          '</a:p>',
          '</xdr:txBody>',
        ].join("")
      : "";

  return [
    '<xdr:oneCellAnchor>',
    '<xdr:from>',
    `<xdr:col>${anchor.col}</xdr:col>`,
    `<xdr:colOff>${pxToEmu(anchor.colOffsetPx)}</xdr:colOff>`,
    `<xdr:row>${anchor.row}</xdr:row>`,
    `<xdr:rowOff>${pxToEmu(anchor.rowOffsetPx)}</xdr:rowOff>`,
    '</xdr:from>',
    `<xdr:ext cx="${pxToEmu(anchor.widthPx)}" cy="${pxToEmu(anchor.heightPx)}" />`,
    '<xdr:sp>',
    '<xdr:nvSpPr>',
    `<xdr:cNvPr id="${shapeId}" name="${xmlEscape(name)}" />`,
    '<xdr:cNvSpPr />',
    '</xdr:nvSpPr>',
    '<xdr:spPr>',
    '<a:prstGeom prst="rect"><a:avLst /></a:prstGeom>',
    fillXml,
    lineXml,
    '</xdr:spPr>',
    textBody,
    '</xdr:sp>',
    '<xdr:clientData />',
    '</xdr:oneCellAnchor>',
  ].join("");
}

function segmentStackExtents(config) {
  return config.banks.reduce(
    (extents, bank) => {
      let positive = 0;
      let negative = 0;
      for (const segment of config.segments) {
        const value = segmentValue(bank, segment.key);
        if (value > 0) {
          positive += value;
        } else if (value < 0) {
          negative += Math.abs(value);
        }
      }
      return {
        positive: Math.max(extents.positive, positive),
        negative: Math.max(extents.negative, negative),
      };
    },
    { positive: 0, negative: 0 },
  );
}

function generatedChartGeometry(config) {
  const plotLeftPx = columnLeftPx(0);
  const plotRightPx = columnLeftPx(14);
  const plotTopPx = rowTopPx(LAYOUT.chartTopRowIndex);
  const plotBottomPx = rowTopPx(LAYOUT.chartBottomRowIndex);
  const extents = segmentStackExtents(config);
  const axisMax = Math.max(1, extents.positive * 1.2);
  const axisMin = extents.negative > 0 ? -Math.max(1.5, extents.negative * 1.5) : 0;
  const unitPx = (plotBottomPx - plotTopPx) / (axisMax - axisMin);

  return {
    plotLeftPx,
    plotTopPx,
    plotWidthPx: plotRightPx - plotLeftPx,
    plotHeightPx: plotBottomPx - plotTopPx,
    baselinePx: plotTopPx + axisMax * unitPx,
    unitPx,
  };
}

function generatedStackedBarChartShapesXml(config) {
  const geometry = generatedChartGeometry(config);
  const barWidthPx = 130;
  const zeroMarkerHeightPx = 3;
  const zeroLabelWidthPx = 34;
  const zeroLabelHeightPx = 15;
  let shapeId = 200;
  const shapes = [];
  const zeroMarkers = [];

  if (config.templateMode) {
    return rectangleShapeXml({
      shapeId,
      name: "Zero percent axis",
      anchor: anchorFromPixels(
        geometry.plotLeftPx,
        geometry.baselinePx - 0.5,
        geometry.plotWidthPx,
        1,
      ),
      fillColor: config.zeroLineColor,
      lineColor: config.zeroLineColor,
      lineWidthPx: 0,
    });
  }

  config.banks.forEach((bank, bankIndex) => {
    const centerCol = CENTER_COL_INDEXES[bankIndex];
    if (centerCol === undefined) {
      return;
    }

    const centerPx = columnLeftPx(centerCol) + columnWidthPx(centerCol) / 2;
    const leftPx = centerPx - barWidthPx / 2;
    let positiveStack = 0;
    let negativeStack = 0;
    const zeroMarkerCounts = new Map();

    for (const segment of config.segments) {
      const value = segmentValue(bank, segment.key);
      if (value === 0) {
        const markerBaseY =
          positiveStack > 0
            ? geometry.baselinePx - positiveStack * geometry.unitPx
            : geometry.baselinePx + negativeStack * geometry.unitPx;
        const markerKey = Math.round(markerBaseY);
        const markerCount = zeroMarkerCounts.get(markerKey) ?? 0;
        zeroMarkerCounts.set(markerKey, markerCount + 1);
        const markerY = markerBaseY + markerCount * 5;
        zeroMarkers.push(
          rectangleShapeXml({
            shapeId,
            name: `${bank.bank} ${segment.label} zero marker`,
            anchor: anchorFromPixels(
              leftPx,
              markerY - zeroMarkerHeightPx / 2,
              barWidthPx,
              zeroMarkerHeightPx,
            ),
            fillColor: segment.color,
            lineColor: segment.color,
            lineWidthPx: 0,
          }),
        );
        shapeId += 1;
        zeroMarkers.push(
          rectangleShapeXml({
            shapeId,
            name: `${bank.bank} ${segment.label} zero label`,
            anchor: anchorFromPixels(
              centerPx - zeroLabelWidthPx / 2,
              markerY - zeroLabelHeightPx / 2,
              zeroLabelWidthPx,
              zeroLabelHeightPx,
            ),
            fillColor: segment.color,
            lineColor: segment.color,
            textColor: segment.textColor,
            label: "0%",
            fontSize: 900,
          }),
        );
        shapeId += 1;
        continue;
      }

      const heightPx = Math.max(1, Math.abs(value) * geometry.unitPx);
      const topPx =
        value > 0
          ? geometry.baselinePx - (positiveStack + value) * geometry.unitPx
          : geometry.baselinePx + negativeStack * geometry.unitPx;
      const roundedValue = Math.round(value);
      const label = roundedValue === 0 ? "" : pctLabel(value);

      shapes.push(
        rectangleShapeXml({
          shapeId,
          name: `${bank.bank} ${segment.label} segment`,
          anchor: anchorFromPixels(leftPx, topPx, barWidthPx, heightPx),
          fillColor: segment.color,
          lineColor: segment.color,
          lineWidthPx: 0,
          textColor: segment.textColor,
          label: heightPx >= 16 ? label : "",
        }),
      );
      shapeId += 1;

      if (value > 0) {
        positiveStack += value;
      } else {
        negativeStack += Math.abs(value);
      }
    }
  });

  shapes.push(
    rectangleShapeXml({
      shapeId,
      name: "Zero percent axis",
      anchor: anchorFromPixels(
        geometry.plotLeftPx,
        geometry.baselinePx - 0.5,
        geometry.plotWidthPx,
        1,
      ),
      fillColor: config.zeroLineColor,
      lineColor: config.zeroLineColor,
      lineWidthPx: 0,
    }),
  );
  shapeId += 1;
  shapes.push(...zeroMarkers);

  return shapes.join("");
}

function removeNativeChartAnchors(xml) {
  return xml.replace(
    /<xdr:twoCellAnchor>[\s\S]*?<c:chart[\s\S]*?<\/xdr:twoCellAnchor>/g,
    "",
  );
}

function addDrawingNamespaces(xml) {
  return xml.replace(/<xdr:wsDr([^>]*)>/, (match, attrs) => {
    let nextAttrs = attrs;
    if (!nextAttrs.includes("xmlns:a=")) {
      nextAttrs += ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    }
    if (!nextAttrs.includes("xmlns:r=")) {
      nextAttrs += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    }
    return `<xdr:wsDr${nextAttrs}>`;
  });
}

function ensureContentType(contentTypesXml, extension, contentType) {
  const defaultPattern = new RegExp(`<Default Extension="${extension}"\\s+ContentType="[^"]+"\\s*/>`);
  if (defaultPattern.test(contentTypesXml)) {
    return contentTypesXml;
  }

  return contentTypesXml.replace(
    /<Types([^>]*)>/,
    `<Types$1><Default Extension="${extension}" ContentType="${contentType}" />`,
  );
}

function logoContentType(filePath) {
  const extension = path.extname(filePath).replace(".", "").toLowerCase();
  if (extension === "png") {
    return { extension, contentType: "image/png" };
  }
  if (extension === "jpg" || extension === "jpeg") {
    return { extension, contentType: "image/jpeg" };
  }
  return null;
}

async function postProcessChart(outputPath, config) {
  const zip = await JSZip.loadAsync(await fs.readFile(outputPath));
  const chartPath = "xl/drawings/charts/chart1.xml";
  const chartFile = zip.file(chartPath);

  if (!chartFile) {
    throw new Error(`Could not find ${chartPath} in exported workbook.`);
  }

  let xml = await chartFile.async("string");
  xml = updateAxisXml(xml, "catAx", (axisXml) => {
    let updated = removeAxisGridlines(axisXml);
    updated = setOrInsertTickLabelPosition(updated, "none");
    updated = setOrInsertAxisLine(updated, "FFFFFF", 0);
    updated = updated.replace(/<c:crosses val="[^"]+" \/>/, '<c:crossesAt val="0" />');
    return updated;
  });
  xml = updateAxisXml(xml, "valAx", (axisXml) => {
    let updated = removeAxisGridlines(axisXml);
    updated = setOrInsertTickLabelPosition(updated, "none");
    updated = setOrInsertAxisLine(updated, "FFFFFF", 0);
    return updated;
  });
  xml = addExplicitPointLabels(xml, config);
  xml = setBarSeriesOverlap(xml);
  xml = disableBarNegativeInversion(xml);
  xml = addNegativePointFormatting(xml, config);
  xml = zeroNegativeBarValues(xml, config);
  xml = addZeroLineSeries(xml, config);
  xml = removeChartBorder(xml);

  zip.file(chartPath, xml);
  await fs.writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
}

async function insertBankLogos(outputPath, config) {
  const logoBanks = config.banks
    .map((bank, index) => ({ bank, index }))
    .filter(
      ({ bank, index }) =>
        config.useLogos !== false && bank.logoPath && CENTER_COL_INDEXES[index] !== undefined,
    );

  const zip = await JSZip.loadAsync(await fs.readFile(outputPath));
  const drawingPath = "xl/drawings/drawing1.xml";
  const drawingRelsPath = "xl/drawings/_rels/drawing1.xml.rels";
  const contentTypesPath = "[Content_Types].xml";
  const drawingFile = zip.file(drawingPath);
  const drawingRelsFile = zip.file(drawingRelsPath);
  const contentTypesFile = zip.file(contentTypesPath);

  if (!drawingFile || !drawingRelsFile || !contentTypesFile) {
    throw new Error("Could not find workbook drawing parts needed to insert bank logos.");
  }

  let drawingXml = addDrawingNamespaces(await drawingFile.async("string"));
  let drawingRelsXml = await drawingRelsFile.async("string");
  let contentTypesXml = await contentTypesFile.async("string");
  const anchors = [];

  for (const { bank, index } of logoBanks) {
    const typeInfo = logoContentType(bank.logoPath);
    if (!typeInfo) {
      continue;
    }

    const mediaName = `bank_logo_${index + 1}.${typeInfo.extension}`;
    const relId = `rIdBankLogo${index + 1}`;
    const picId = 100 + index + 1;
    const anchor = await logoAnchor(bank.logoPath, index);
    anchors.push(
      logoAnchorXml({
        relId,
        picId,
        name: `${bank.bank} Logo`,
        anchor,
      }),
    );
    zip.file(`xl/media/${mediaName}`, await fs.readFile(bank.logoPath));
    drawingRelsXml = drawingRelsXml.replace(
      "</Relationships>",
      `<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/xl/media/${mediaName}" Id="${relId}" /></Relationships>`,
    );
    contentTypesXml = ensureContentType(
      contentTypesXml,
      typeInfo.extension,
      typeInfo.contentType,
    );
  }

  drawingXml = removeNativeChartAnchors(drawingXml);
  drawingXml = drawingXml.replace(
    "</xdr:wsDr>",
    `${generatedStackedBarChartShapesXml(config)}${anchors.join("")}</xdr:wsDr>`,
  );
  zip.file(drawingPath, drawingXml);
  zip.file(drawingRelsPath, drawingRelsXml);
  zip.file(contentTypesPath, contentTypesXml);
  await fs.writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
}

async function renderPreviewFromXlsx(outputPath, previewPath, sheetName) {
  const file = await FileBlob.load(outputPath);
  const workbook = await SpreadsheetFile.importXlsx(file);
  await renderPreview(workbook, previewPath, sheetName);
}

async function verifyWorkbook(workbook, sheetName, config) {
  const sourceStartColIndex = 15;
  const sourceEndCol = columnName(sourceStartColIndex + config.segments.length + 3);
  const sourceEndRow = LAYOUT.sourceTableStartRow + config.banks.length;
  const tableCheck = await workbook.inspect({
    kind: "table",
    range: `${sheetName}!P${LAYOUT.sourceTableStartRow}:${sourceEndCol}${sourceEndRow}`,
    include: "values,formulas",
    tableMaxRows: 12,
    tableMaxCols: 12,
  });
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 50 },
    summary: "final formula error scan",
  });

  console.log(tableCheck.ndjson);
  console.log(errors.ndjson);
}

async function buildRevenueGrowthWorkbook(config, outputPath, previewPath, options = {}) {
  const sheetName = safeSheetName(config);
  config.sheetName = sheetName;
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add(sheetName);

  setColumnWidths(sheet);
  styleSectionHeader(sheet, config);
  addLegend(sheet, config);
  addStackedColumnChart(sheet, config);
  addTopLabels(sheet, config);
  await addBankBranding(sheet, config);
  writeSourceTable(sheet, config);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });

  if (options.verify !== false) {
    await verifyWorkbook(workbook, sheetName, config);
  }
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);
  await postProcessChart(outputPath, config);
  await insertBankLogos(outputPath, config);
  await renderPreviewFromXlsx(outputPath, previewPath, sheetName);
  console.log(outputPath);
}

async function main() {
  const args = parseArgs(process.argv);

  if (args.createInputTemplate) {
    await createInputTemplateWorkbook(path.resolve(args.createInputTemplate));
    return;
  }

  const defaultOutput = path.resolve(
    SCRIPT_DIR,
    "../outputs/revenue-growth-by-segment/revenue_growth_by_segment.xlsx",
  );
  const outputPath = path.resolve(args.output ?? args.createBlankTemplate ?? defaultOutput);
  const previewPath = path.resolve(
    args.preview ?? path.join(path.dirname(outputPath), "preview.png"),
  );
  const config = args.createBlankTemplate ? blankTemplateConfig() : await loadConfig(args.input);
  await buildRevenueGrowthWorkbook(config, outputPath, previewPath, {
    verify: !args.createBlankTemplate,
  });
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  await main();
}

export { loadConfigFromInputWorkbook, prepareConfig, readXlsxSheetCells };
