// generate_excel.js — generates copart_competitive_data.xlsx in the project folder
// Run: node generate_excel.js

const XLSX = require('xlsx');
const path = require('path');

const COMPANIES = ['Copart', 'IAA', 'Manheim', 'OpenLane', 'ACV', 'CarMax', 'Carvana', 'eBay Motors'];

// ── ALL METRICS ──────────────────────────────────────────────────────────────
// Each entry: { name, unit, note, data }
// data order matches COMPANIES above; null = N/A
const METRICS = [
  {
    name: 'Annual Revenue',
    unit: 'USD Billions',
    note: 'Most recent fiscal year (FY2024). CarMax & Carvana include retail sales. eBay Motors is estimated.',
    data: [4.19, 4.06, 4.50, 0.91, 0.63, 26.50, 13.70, 1.80],
  },
  {
    name: 'Market Capitalization',
    unit: 'USD Billions',
    note: 'Public market value, year-end 2024. Manheim (private) and eBay Motors (division) = N/A.',
    data: [54, 8, null, 1.1, 1.4, 11, 49, null],
  },
  {
    name: 'Capital Expenditure',
    unit: 'USD Billions',
    note: 'Annual capex FY2024. Manheim & eBay Motors not separately reported. Copart reflects active land acquisition.',
    data: [1.1, 0.22, null, 0.065, 0.04, 0.38, 0.35, null],
  },
  {
    name: 'Physical Locations',
    unit: 'Number of Sites',
    note: 'Auction yards, storage lots & reconditioning centres, 2024. ACV & eBay are digital-only.',
    data: [270, 200, 75, 30, 0, 246, 36, 0],
  },
  {
    name: 'Land Footprint',
    unit: 'Thousands of Acres',
    note: 'Total owned & operated acreage, 2024. Digital-first platforms show near-zero.',
    data: [120, 30, 14, 4, 0, 5, 2, 0],
  },
  {
    name: 'Countries of Operation',
    unit: 'Countries',
    note: 'Countries with active physical or digital operations, 2024.',
    data: [11, 2, 4, 3, 2, 1, 1, 15],
  },
  {
    name: 'Vehicles Auctioned / Available',
    unit: 'Millions / Year',
    note: 'Vehicles consigned / listed to auction per year, FY2024.',
    data: [3.6, 2.4, 10.0, 1.6, 0.6, 0.95, 0.45, 3.0],
  },
  {
    name: 'Vehicles Sold Annually',
    unit: 'Millions / Year',
    note: 'Vehicles actually sold at auction / retail, FY2024. Reflects sell-through rate.',
    data: [3.5, 2.2, 8.5, 1.3, 0.55, 0.90, 0.42, 2.0],
  },
  {
    name: 'Average Daily Inventory',
    unit: 'Thousands of Vehicles',
    note: 'Vehicles available for bidding or purchase on any given day, 2024. eBay figure = vehicles only.',
    data: [180, 125, 85, 55, 18, 55, 40, 35],
  },
  {
    name: 'Towing & Transporter Network',
    unit: 'Thousands of Contracted Providers',
    note: 'Contracted tow operators & transport providers, 2024. Digital-only & retail platforms = N/A.',
    data: [20, 10, null, null, 1, null, null, null],
  },
  {
    name: 'Title Processing (Insurance / Salvage)',
    unit: 'Millions of Titles / Year',
    note: 'Annual salvage & insurance-seller titles processed, FY2024. Only Copart & IAA operate at scale for this.',
    data: [3.2, 2.1, null, null, null, null, null, null],
  },
  {
    name: 'Total Employees',
    unit: 'Thousands (FTE)',
    note: 'Full-time equivalent headcount, 2024. eBay Motors employees not separately reportable.',
    data: [16, 7, 20, 2.2, 1.7, 28, 15, null],
  },
  {
    name: 'Glassdoor Employee Rating',
    unit: 'Rating (out of 5.0)',
    note: 'Aggregate employee satisfaction score, Glassdoor.com, 2024.',
    data: [3.9, 3.4, 3.7, 3.5, 3.9, 3.9, 3.3, 3.8],
  },
  {
    name: 'Years in Operation',
    unit: 'Years (as of 2024)',
    note: 'Continuous industry operation based on founding year.',
    data: [42, 43, 79, 35, 10, 31, 12, 25],
  },
  {
    name: 'Main App Rating (iOS)',
    unit: 'Rating (out of 5.0)',
    note: 'iOS App Store buyer/seller platform rating, 2024.',
    data: [4.8, 4.3, 3.9, 3.7, 4.6, 4.8, 4.8, 4.4],
  },
  {
    name: 'Main App Downloads',
    unit: 'Millions (cumulative)',
    note: 'Combined iOS + Android total downloads, 2024. eBay Motors app counted separately.',
    data: [5, 2, 1, 0.5, 1, 10, 8, 5],
  },
  {
    name: 'Driver App Rating (iOS)',
    unit: 'Rating (out of 5.0)',
    note: 'Dedicated driver/transport app iOS rating, 2024. N/A where no dedicated driver app exists.',
    data: [4.5, 4.0, 3.8, null, 4.2, 4.3, 4.5, null],
  },
  {
    name: 'Driver App Downloads',
    unit: 'Millions (cumulative)',
    note: 'Driver/transport app combined iOS + Android downloads, 2024. N/A where no dedicated app exists.',
    data: [1.0, 0.5, 0.3, null, 0.4, 0.8, 1.2, null],
  },
  {
    name: 'Registered Global Buyers',
    unit: 'Thousands of Accounts',
    note: 'Total registered buyer/member accounts, 2024. CarMax, Carvana & eBay use B2C model — N/A.',
    data: [750, 500, 70, 85, 25, null, null, null],
  },
  {
    name: 'Buyer Countries',
    unit: 'Countries',
    note: 'Countries from which active buyers placed bids, 2024.',
    data: [170, 40, 25, 20, 3, 1, 1, 65],
  },
  {
    name: 'Monthly Social & Web Mentions',
    unit: 'Thousands / Month',
    note: 'Estimated average monthly brand mentions across social platforms, news & forums, 2024.',
    data: [45, 25, 35, 8, 12, 180, 280, 55],
  },
  {
    name: 'B2B Auction Mentions (Excl. Retail)',
    unit: 'Thousands / Month',
    note: 'Social mentions filtered to B2B/salvage-auction segment only, 2024. Retail platforms = N/A.',
    data: [45, 25, 35, 8, 12, null, null, 55],
  },
  {
    name: 'YouTube Channel Subscribers',
    unit: 'Thousands of Subscribers',
    note: 'Official YouTube channel subscribers, 2024 estimate. B2B platforms have smaller but more targeted audiences.',
    data: [90, 20, 25, 7, 4, 115, 60, 47],
  },
  {
    name: 'Estimated Monthly SEM Spend',
    unit: 'USD Millions / Month',
    note: 'Estimated paid search (Google/Bing) monthly spend, 2024. Not publicly disclosed; industry benchmark estimates.',
    data: [1.1, 0.6, 0.75, 0.35, 0.45, 11.5, 15.0, 5.0],
  },
];

// ── STYLES ───────────────────────────────────────────────────────────────────
const COPART_BLUE = '0047BB';
const HEADER_BG   = '1E3A5F';
const ALT_ROW_BG  = 'EEF4FF';
const COPART_ROW  = 'DBEAFE';
const TITLE_FG    = 'FFFFFF';
const LABEL_FG    = '334155';
const NOTE_FG     = '64748B';
const NA_FG       = 'A0AEC0';

function makeHeaderCell(value, bold = true, bg = HEADER_BG, fg = TITLE_FG) {
  return {
    v: value, t: 's',
    s: {
      font:    { bold, color: { rgb: fg }, name: 'Calibri', sz: 11 },
      fill:    { fgColor: { rgb: bg }, patternType: 'solid' },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top:    { style: 'thin', color: { rgb: 'CBD5E1' } },
        bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
        left:   { style: 'thin', color: { rgb: 'CBD5E1' } },
        right:  { style: 'thin', color: { rgb: 'CBD5E1' } },
      }
    }
  };
}

function makeCell(value, isCopart = false, isAlt = false, align = 'center') {
  const isNA = value === null;
  const bg   = isCopart ? COPART_ROW : (isAlt ? ALT_ROW_BG : 'FFFFFF');
  return {
    v: isNA ? 'N/A' : value,
    t: isNA ? 's' : (typeof value === 'number' ? 'n' : 's'),
    s: {
      font:    { color: { rgb: isNA ? NA_FG : (isCopart ? COPART_BLUE : LABEL_FG) }, name: 'Calibri', sz: 11, bold: isCopart },
      fill:    { fgColor: { rgb: bg }, patternType: 'solid' },
      alignment: { horizontal: align, vertical: 'center' },
      border: {
        top:    { style: 'thin', color: { rgb: 'E2E8F0' } },
        bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
        left:   { style: 'thin', color: { rgb: 'E2E8F0' } },
        right:  { style: 'thin', color: { rgb: 'E2E8F0' } },
      }
    }
  };
}

function makeTitleCell(value) {
  return {
    v: value, t: 's',
    s: {
      font:    { bold: true, sz: 14, color: { rgb: TITLE_FG }, name: 'Calibri' },
      fill:    { fgColor: { rgb: COPART_BLUE }, patternType: 'solid' },
      alignment: { horizontal: 'left', vertical: 'center' },
    }
  };
}

function makeSubCell(value) {
  return {
    v: value, t: 's',
    s: {
      font:    { italic: true, sz: 9, color: { rgb: NOTE_FG }, name: 'Calibri' },
      fill:    { fgColor: { rgb: 'F8FAFC' }, patternType: 'solid' },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    }
  };
}

function makeSectionLabel(value) {
  return {
    v: value, t: 's',
    s: {
      font:    { bold: true, sz: 10, color: { rgb: COPART_BLUE }, name: 'Calibri' },
      fill:    { fgColor: { rgb: 'EFF6FF' }, patternType: 'solid' },
      alignment: { horizontal: 'left', vertical: 'center' },
    }
  };
}

// ── SHEET BUILDER ────────────────────────────────────────────────────────────
function buildSheet(metric) {
  const { name, unit, note, data } = metric;
  const ws = {};
  const ref = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };

  function setCell(r, c, cell) {
    const addr = XLSX.utils.encode_cell({ r, c });
    ws[addr] = cell;
    if (r > ref.e.r) ref.e.r = r;
    if (c > ref.e.c) ref.e.c = c;
  }

  function blank(r, c) {
    setCell(r, c, { v: '', t: 's', s: { fill: { fgColor: { rgb: 'FFFFFF' }, patternType: 'solid' } } });
  }

  const numCols = Math.max(COMPANIES.length + 1, 3);

  // ── ROW 0: Title ──
  setCell(0, 0, makeTitleCell(name));
  for (let c = 1; c < numCols; c++) setCell(0, c, { v: '', t: 's', s: { fill: { fgColor: { rgb: COPART_BLUE }, patternType: 'solid' } } });

  // ── ROW 1: Unit + Note ──
  setCell(1, 0, makeSubCell(`Unit: ${unit}   |   ${note}`));
  for (let c = 1; c < numCols; c++) setCell(1, c, { v: '', t: 's', s: { fill: { fgColor: { rgb: 'F8FAFC' }, patternType: 'solid' } } });

  // ── ROW 2: Blank ──
  for (let c = 0; c < numCols; c++) blank(2, c);

  // ── ROW 3: VERTICAL LAYOUT label ──
  setCell(3, 0, makeSectionLabel('VERTICAL LAYOUT'));
  for (let c = 1; c < numCols; c++) setCell(3, c, { ...makeSectionLabel(''), v: '' });

  // ── ROW 4: Headers (vertical) ──
  setCell(4, 0, makeHeaderCell('Company'));
  setCell(4, 1, makeHeaderCell(`Value (${unit})`));
  setCell(4, 2, makeHeaderCell('Notes'));
  for (let c = 3; c < numCols; c++) blank(4, c);

  // ── ROWS 5–12: Data (vertical) ──
  COMPANIES.forEach((co, i) => {
    const row = 5 + i;
    const isCopart = i === 0;
    const isAlt    = i % 2 === 1;
    const val      = data[i];
    setCell(row, 0, makeCell(co,                isCopart, isAlt, 'left'));
    setCell(row, 1, makeCell(val,               isCopart, isAlt, 'center'));
    setCell(row, 2, makeCell(val === null ? 'Not applicable / not publicly reported' : `${val} ${unit}`, isCopart, isAlt, 'left'));
    for (let c = 3; c < numCols; c++) blank(row, c);
  });

  // ── ROW 13: Blank ──
  for (let c = 0; c < numCols; c++) blank(13, c);

  // ── ROW 14: HORIZONTAL LAYOUT label ──
  const hCols = COMPANIES.length + 1;
  setCell(14, 0, makeSectionLabel('HORIZONTAL LAYOUT'));
  for (let c = 1; c < hCols; c++) setCell(14, c, { ...makeSectionLabel(''), v: '' });

  // ── ROW 15: Company headers (horizontal) ──
  setCell(15, 0, makeHeaderCell('Metric'));
  COMPANIES.forEach((co, i) => setCell(15, i + 1, makeHeaderCell(co, true, i === 0 ? COPART_BLUE : HEADER_BG)));

  // ── ROW 16: Values (horizontal) ──
  setCell(16, 0, makeCell(name, false, false, 'left'));
  COMPANIES.forEach((co, i) => {
    setCell(16, i + 1, makeCell(data[i], i === 0, false, 'center'));
  });

  // ── ROW 17: Unit row ──
  setCell(17, 0, makeCell('Unit', false, true, 'left'));
  COMPANIES.forEach((_co, i) => setCell(17, i + 1, makeCell(unit, false, true, 'center')));

  // ── Column widths ──
  ws['!cols'] = [
    { wch: 28 },  // Company / Metric
    { wch: 18 },  // Value / Copart
    { wch: 42 },  // Notes / IAA
    ...COMPANIES.slice(2).map(() => ({ wch: 14 })),
  ];

  // ── Row heights ──
  ws['!rows'] = [
    { hpx: 32 }, // title
    { hpx: 36 }, // note
    { hpx: 8 },  // blank
    { hpx: 20 }, // section label
    { hpx: 24 }, // header
    ...COMPANIES.map(() => ({ hpx: 22 })), // data rows
    { hpx: 8 },  // blank
    { hpx: 20 }, // section label
    { hpx: 24 }, // header
    { hpx: 22 }, // values
    { hpx: 20 }, // unit
  ];

  // ── Merges ──
  ws['!merges'] = [
    { s: { r: 0,  c: 0 }, e: { r: 0,  c: numCols - 1 } }, // title
    { s: { r: 1,  c: 0 }, e: { r: 1,  c: numCols - 1 } }, // note
    { s: { r: 3,  c: 0 }, e: { r: 3,  c: numCols - 1 } }, // vertical label
    { s: { r: 13, c: 0 }, e: { r: 13, c: hCols - 1 } },   // blank
    { s: { r: 14, c: 0 }, e: { r: 14, c: hCols - 1 } },   // horizontal label
  ];

  ws['!ref'] = XLSX.utils.encode_range(ref);
  return ws;
}

// ── BUILD SUMMARY SHEET ──────────────────────────────────────────────────────
function buildSummarySheet() {
  const ws = {};
  const nCo = COMPANIES.length;
  const nMe = METRICS.length;

  function setCell(r, c, cell) {
    ws[XLSX.utils.encode_cell({ r, c })] = cell;
  }
  function blank(r, c) { setCell(r, c, { v: '', t: 's', s: { fill: { fgColor: { rgb: 'FFFFFF' }, patternType: 'solid' } } }); }

  // Title row
  setCell(0, 0, makeTitleCell('Copart Competitive Dashboard — All Metrics Summary (FY2024)'));
  for (let c = 1; c <= nCo; c++) setCell(0, c, { v: '', t: 's', s: { fill: { fgColor: { rgb: COPART_BLUE }, patternType: 'solid' } } });

  // Sub row
  setCell(1, 0, makeSubCell('Data is directional and approximate — intended for presentation use, not investment decisions.'));
  for (let c = 1; c <= nCo; c++) setCell(1, c, { v: '', t: 's', s: { fill: { fgColor: { rgb: 'F8FAFC' }, patternType: 'solid' } } });

  blank(2, 0); for (let c = 1; c <= nCo; c++) blank(2, c);

  // Header row
  setCell(3, 0, makeHeaderCell('Metric'));
  COMPANIES.forEach((co, i) => setCell(3, i + 1, makeHeaderCell(co, true, i === 0 ? COPART_BLUE : HEADER_BG)));
  setCell(3, nCo + 1, makeHeaderCell('Unit'));

  // Data rows
  METRICS.forEach(({ name, unit, data }, mi) => {
    const row = 4 + mi;
    const isAlt = mi % 2 === 1;
    setCell(row, 0, makeCell(name, false, isAlt, 'left'));
    data.forEach((v, ci) => setCell(row, ci + 1, makeCell(v, ci === 0, isAlt, 'center')));
    setCell(row, nCo + 1, makeCell(unit, false, isAlt, 'left'));
  });

  ws['!cols'] = [
    { wch: 36 },
    ...COMPANIES.map(() => ({ wch: 14 })),
    { wch: 30 },
  ];
  ws['!rows'] = [
    { hpx: 32 }, { hpx: 32 }, { hpx: 8 }, { hpx: 24 },
    ...METRICS.map(() => ({ hpx: 22 })),
  ];
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: nCo } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: nCo } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: nCo } },
  ];
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 4 + nMe, c: nCo + 1 } });
  return ws;
}

// ── MAIN ─────────────────────────────────────────────────────────────────────
const wb = XLSX.utils.book_new();

// Summary sheet first
XLSX.utils.book_append_sheet(wb, buildSummarySheet(), 'Summary');

// One sheet per metric (sanitize: remove invalid chars, truncate to 31)
METRICS.forEach(metric => {
  const safe = metric.name.replace(/[:\\/?\*\[\]]/g, '-');
  const sheetName = safe.length > 31 ? safe.slice(0, 31) : safe;
  XLSX.utils.book_append_sheet(wb, buildSheet(metric), sheetName);
});

const outPath = path.join(__dirname, 'copart_competitive_data.xlsx');
XLSX.writeFile(wb, outPath, { bookType: 'xlsx', cellStyles: true });
console.log(`Excel file written to: ${outPath}`);
console.log(`Sheets: Summary + ${METRICS.length} metric sheets`);
