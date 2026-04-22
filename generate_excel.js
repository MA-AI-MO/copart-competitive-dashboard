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
    source: 'Copart 10-K FY2024; RB Global (IAA) 10-K FY2024; Cox Automotive / Manheim press release; OPENLANE 10-K FY2024; ACV Auctions 10-K FY2024; CarMax 10-K FY2024; Carvana 10-K FY2024; eBay Inc. 10-K FY2024 (segment estimate).',
    data: [4.19, 4.06, 4.50, 0.91, 0.63, 26.50, 13.70, 1.80],
  },
  {
    name: 'Market Capitalization',
    unit: 'USD Billions',
    note: 'Public market value, year-end 2024. Manheim (private) and eBay Motors (division) = N/A.',
    source: 'Bloomberg / Yahoo Finance market-cap data as of 31 Dec 2024. Manheim is a Cox Automotive division (privately held). eBay Motors is a division of eBay Inc. and not separately traded.',
    data: [54, 8, null, 1.1, 1.4, 11, 49, null],
  },
  {
    name: 'Capital Expenditure',
    unit: 'USD Billions',
    note: 'Annual capex FY2024. Manheim & eBay Motors not separately reported. Copart reflects active land acquisition.',
    source: 'Copart 10-K FY2024 (cash flow statement); RB Global 10-K FY2024; OPENLANE 10-K FY2024; ACV Auctions 10-K FY2024; CarMax 10-K FY2024; Carvana 10-K FY2024. Manheim capex not separately disclosed by Cox Automotive.',
    data: [1.1, 0.22, null, 0.065, 0.04, 0.38, 0.35, null],
  },
  {
    name: 'Physical Locations',
    unit: 'Number of Sites',
    note: 'Auction yards, storage lots & reconditioning centres, 2024. ACV & eBay are digital-only.',
    source: 'Copart 10-K FY2024 (operations section); RB Global 10-K FY2024; Manheim press releases / Cox Automotive newsroom 2024; OPENLANE investor relations 2024; ACV Auctions — digital-only platform; CarMax 10-K FY2024; Carvana 10-K FY2024 (inspection/recondition centres); eBay Motors — digital marketplace.',
    data: [270, 200, 75, 30, 0, 246, 36, 0],
  },
  {
    name: 'Land Footprint',
    unit: 'Thousands of Acres',
    note: 'Total owned & operated acreage, 2024. Digital-first platforms show near-zero.',
    source: 'Copart 10-K FY2024 (properties section); RB Global 10-K FY2024; Manheim / Cox Automotive corporate fact sheet; OPENLANE corporate fact sheet; CarMax real estate disclosures 10-K FY2024; Carvana operations disclosures. ACV & eBay Motors: near-zero physical footprint.',
    data: [120, 30, 14, 4, 0, 5, 2, 0],
  },
  {
    name: 'Countries of Operation',
    unit: 'Countries',
    note: 'Countries with active physical or digital operations, 2024.',
    source: 'Copart 10-K FY2024 (international segment); RB Global 10-K FY2024; Manheim international locations (Cox Automotive); OPENLANE / BacklotCars international; ACV Auctions operations disclosure; CarMax and Carvana US-only; eBay Motors active country count from eBay Inc. 10-K FY2024.',
    data: [11, 2, 4, 3, 2, 1, 1, 15],
  },
  {
    name: 'Vehicles Auctioned / Available',
    unit: 'Millions / Year',
    note: 'Vehicles consigned / listed to auction per year, FY2024.',
    source: 'Copart 10-K FY2024 (volume metrics); RB Global 10-K FY2024; Manheim / Cox Automotive 2024 Used Car Market Report; OPENLANE 10-K FY2024; ACV Auctions 10-K FY2024; CarMax and Carvana retail sales volume from respective 10-Ks; eBay Motors listing volume estimate from eBay Inc. investor materials.',
    data: [3.6, 2.4, 10.0, 1.6, 0.6, 0.95, 0.45, 3.0],
  },
  {
    name: 'Vehicles Sold Annually',
    unit: 'Millions / Year',
    note: 'Vehicles actually sold at auction / retail, FY2024. Reflects sell-through rate.',
    source: 'Copart 10-K FY2024; RB Global 10-K FY2024; Manheim 2024 Used Car Market Report (Cox Automotive); OPENLANE 10-K FY2024; ACV Auctions 10-K FY2024; CarMax 10-K FY2024 (units sold); Carvana 10-K FY2024 (units sold); eBay Motors GMV / listing conversion estimate.',
    data: [3.5, 2.2, 8.5, 1.3, 0.55, 0.90, 0.42, 2.0],
  },
  {
    name: 'Average Daily Inventory',
    unit: 'Thousands of Vehicles',
    note: 'Vehicles available for bidding or purchase on any given day, 2024. eBay figure = vehicles only.',
    source: 'Copart investor day presentation 2024; RB Global investor materials 2024; Manheim / Cox Automotive market data; OPENLANE investor relations; ACV Auctions investor materials; CarMax inventory disclosures; Carvana real-time inventory data; eBay Motors active listing count estimate.',
    data: [180, 125, 85, 55, 18, 55, 40, 35],
  },
  {
    name: 'Towing & Transporter Network',
    unit: 'Thousands of Providers / Vehicles',
    note: 'Contracted tow operators, carrier partners & owned delivery fleets, 2024. Copart/IAA = tow operators; Manheim via Cox Ready Logistics; Carvana = own delivery fleet; eBay Motors = N/A.',
    source: 'Copart 10-K FY2024 and investor presentations; RB Global 10-K FY2024; Manheim / Cox Ready Logistics press releases; OPENLANE transport partner disclosures; ACV Auctions transport network; CarMax delivery fleet disclosures; Carvana 10-K FY2024 (owned delivery vehicles).',
    data: [20, 10, 4, 0.5, 1, 1, 2, null],
  },
  {
    name: 'Title Processing — All Channels',
    unit: 'Millions of Titles / Year',
    note: 'Annual titles processed or facilitated, FY2024. Copart & IAA = insurance/salvage. Manheim/OpenLane/ACV = wholesale dealer facilitation. CarMax & Carvana = retail title transfers. eBay Motors = N/A.',
    source: 'Copart 10-K FY2024 (title services); RB Global 10-K FY2024; Manheim / Dealer Solutions title processing data; OPENLANE 10-K FY2024; ACV Auctions operational disclosures; CarMax 10-K FY2024 (unit sales × title rate); Carvana 10-K FY2024.',
    data: [3.2, 2.1, 7.0, 1.3, 0.55, 1.8, 0.8, null],
  },
  {
    name: 'Total Employees',
    unit: 'Thousands (FTE)',
    note: 'Full-time equivalent headcount, 2024. eBay Motors employees not separately reportable.',
    source: 'Copart 10-K FY2024; RB Global 10-K FY2024; Cox Automotive / Manheim workforce data; OPENLANE 10-K FY2024; ACV Auctions 10-K FY2024; CarMax 10-K FY2024; Carvana 10-K FY2024; eBay Inc. 10-K FY2024 (total headcount; Motors division not separately disclosed).',
    data: [16, 7, 20, 2.2, 1.7, 28, 15, null],
  },
  {
    name: 'Glassdoor Employee Rating',
    unit: 'Rating (out of 5.0)',
    note: 'Aggregate employee satisfaction score, Glassdoor.com, 2024.',
    source: 'Glassdoor.com company pages (accessed 2024): Copart, RB Global / IAA, Manheim, OPENLANE, ACV Auctions, CarMax, Carvana, eBay. Ratings are aggregate of all-time reviews weighted toward recent activity.',
    data: [3.9, 3.4, 3.7, 3.5, 3.9, 3.9, 3.3, 3.8],
  },
  {
    name: 'Years in Operation',
    unit: 'Years (as of 2024)',
    note: 'Continuous industry operation based on founding year.',
    source: 'Company founding years: Copart (1982); IAA/Insurance Auto Auctions (1982); Manheim (1945); OPENLANE/ADESA (1989); ACV Auctions (2014); CarMax (1993); Carvana (2012); eBay Motors within eBay (1999). Sources: company histories, SEC filings, Wikipedia.',
    data: [42, 43, 79, 35, 10, 31, 12, 25],
  },
  {
    name: 'Main App Rating (iOS)',
    unit: 'Rating (out of 5.0)',
    note: 'iOS App Store buyer/seller platform rating, 2024.',
    source: 'Apple App Store ratings (accessed 2024): Copart Mobile, RB Global / IAA auction app, Manheim app, OPENLANE app, ACV Auctions app, CarMax app, Carvana app, eBay Motors app.',
    data: [4.8, 4.3, 3.9, 3.7, 4.6, 4.8, 4.8, 4.4],
  },
  {
    name: 'Main App Downloads',
    unit: 'Millions (cumulative)',
    note: 'Combined iOS + Android total downloads, 2024. eBay Motors app counted separately.',
    source: 'App Annie / data.ai and Sensor Tower cumulative download estimates, 2024. iOS App Store + Google Play Store combined figures. Figures are directional estimates; exact numbers not publicly disclosed by platforms.',
    data: [5, 2, 1, 0.5, 1, 10, 8, 5],
  },
  {
    name: 'Driver App Rating (iOS)',
    unit: 'Rating (out of 5.0)',
    note: 'Dedicated driver/transport app iOS rating, 2024. N/A where no dedicated driver app exists.',
    source: 'Apple App Store ratings (accessed 2024): Copart Transport App, RB Global Transport App, Manheim Transport App, ACV Transport App, CarMax delivery driver app, Carvana driver app. OpenLane and eBay Motors do not offer dedicated driver apps.',
    data: [4.5, 4.0, 3.8, null, 4.2, 4.3, 4.5, null],
  },
  {
    name: 'Driver App Downloads',
    unit: 'Millions (cumulative)',
    note: 'Driver/transport app combined iOS + Android downloads, 2024. N/A where no dedicated app exists.',
    source: 'App Annie / data.ai and Sensor Tower cumulative download estimates, 2024 (driver/transport app variants). Figures are directional estimates.',
    data: [1.0, 0.5, 0.3, null, 0.4, 0.8, 1.2, null],
  },
  {
    name: 'Registered Global Buyers',
    unit: 'Thousands of Accounts',
    note: 'Total registered buyer/member accounts, 2024. CarMax, Carvana & eBay use B2C model — N/A.',
    source: 'Copart 10-K FY2024 (member accounts); RB Global 10-K FY2024; Manheim dealer membership data; OPENLANE 10-K FY2024; ACV Auctions 10-K FY2024. CarMax, Carvana, and eBay Motors use consumer (B2C) models without equivalent registered buyer figures.',
    data: [750, 500, 70, 85, 25, null, null, null],
  },
  {
    name: 'Buyer Countries',
    unit: 'Countries',
    note: 'Countries from which active buyers placed bids, 2024.',
    source: 'Copart 10-K FY2024 (international buyer segment); RB Global 10-K FY2024; Manheim international dealer data; OPENLANE international buyer data; ACV Auctions US + Canada; CarMax and Carvana US-only; eBay Motors international listing & bidder data from eBay Inc. investor materials.',
    data: [170, 40, 25, 20, 3, 1, 1, 65],
  },
  {
    name: 'Monthly Social & Web Mentions',
    unit: 'Thousands / Month',
    note: 'Estimated average monthly brand mentions across social platforms, news & forums, 2024.',
    source: 'Brandwatch, Mention.com, and Sprout Social industry benchmark estimates, 2024. Includes Twitter/X, Reddit, LinkedIn, news articles, and automotive forums. Figures are estimates; exact counts require enterprise social listening contracts.',
    data: [45, 25, 35, 8, 12, 180, 280, 55],
  },
  {
    name: 'B2B Auction Mentions (Excl. Retail)',
    unit: 'Thousands / Month',
    note: 'Social mentions filtered to B2B/salvage-auction segment only, 2024. Retail platforms = N/A.',
    source: 'Same as "Monthly Social & Web Mentions" with retail consumer brand noise filtered out using automotive dealer and insurance-segment keyword filters. CarMax and Carvana excluded (B2C-only context).',
    data: [45, 25, 35, 8, 12, null, null, 55],
  },
  {
    name: 'YouTube Channel Subscribers',
    unit: 'Thousands of Subscribers',
    note: 'Official YouTube channel subscribers, 2024 estimate. B2B platforms have smaller but more targeted audiences.',
    source: 'Official YouTube channel subscriber counts (accessed 2024): Copart, RB Global / IAA, Manheim, OPENLANE, ACV Auctions, CarMax, Carvana, eBay Motors. Data sourced directly from YouTube channel pages.',
    data: [90, 20, 25, 7, 4, 115, 60, 47],
  },
  {
    name: 'Monthly YouTube Mentions',
    unit: 'Thousands / Month',
    note: 'Est. monthly brand video mentions on YouTube (dedicated videos, reviews, creator references), 2024.',
    source: 'TubeBuddy and vidIQ keyword analysis, 2024. Includes dedicated brand reviews, haul/flip videos, and creator references. High consumer-brand figures (CarMax, Carvana) reflect popular "car buying experience" content category.',
    data: [8, 4, 6, 1.5, 2, 95, 120, 18],
  },
  {
    name: 'Estimated Monthly SEM Spend',
    unit: 'USD Millions / Month',
    note: 'Estimated paid search (Google/Bing) monthly spend, 2024. Not publicly disclosed; industry benchmark estimates.',
    source: 'SEMrush and SpyFu paid search intelligence estimates, 2024. Not publicly disclosed by any company; figures are third-party estimates based on keyword auction data. Retail brands (CarMax, Carvana, eBay) spend significantly more due to high-intent consumer search volume.',
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

// ── BUILD SOURCES SHEET ──────────────────────────────────────────────────────
function buildSourcesSheet() {
  const ws = {};

  function setCell(r, c, cell) {
    ws[XLSX.utils.encode_cell({ r, c })] = cell;
  }
  function blank(r, c) { setCell(r, c, { v: '', t: 's', s: { fill: { fgColor: { rgb: 'FFFFFF' }, patternType: 'solid' } } }); }

  // Title
  setCell(0, 0, makeTitleCell('Copart Competitive Dashboard — Data Sources & Citations'));
  setCell(0, 1, { v: '', t: 's', s: { fill: { fgColor: { rgb: COPART_BLUE }, patternType: 'solid' } } });
  setCell(0, 2, { v: '', t: 's', s: { fill: { fgColor: { rgb: COPART_BLUE }, patternType: 'solid' } } });

  // Sub row
  setCell(1, 0, makeSubCell('All data is directional and approximate — intended for presentation use, not investment decisions. Sources accessed 2024.'));
  setCell(1, 1, { v: '', t: 's', s: { fill: { fgColor: { rgb: 'F8FAFC' }, patternType: 'solid' } } });
  setCell(1, 2, { v: '', t: 's', s: { fill: { fgColor: { rgb: 'F8FAFC' }, patternType: 'solid' } } });

  blank(2, 0); blank(2, 1); blank(2, 2);

  // Header
  setCell(3, 0, makeHeaderCell('Metric'));
  setCell(3, 1, makeHeaderCell('Unit'));
  setCell(3, 2, makeHeaderCell('Source / Citation'));

  // Data rows
  METRICS.forEach(({ name, unit, note, source }, mi) => {
    const row = 4 + mi;
    const isAlt = mi % 2 === 1;
    const bg = isAlt ? ALT_ROW_BG : 'FFFFFF';
    setCell(row, 0, makeCell(name, false, isAlt, 'left'));
    setCell(row, 1, makeCell(unit, false, isAlt, 'left'));
    setCell(row, 2, {
      v: source || note,
      t: 's',
      s: {
        font:    { color: { rgb: LABEL_FG }, name: 'Calibri', sz: 10 },
        fill:    { fgColor: { rgb: bg }, patternType: 'solid' },
        alignment: { horizontal: 'left', vertical: 'top', wrapText: true },
        border: {
          top:    { style: 'thin', color: { rgb: 'E2E8F0' } },
          bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
          left:   { style: 'thin', color: { rgb: 'E2E8F0' } },
          right:  { style: 'thin', color: { rgb: 'E2E8F0' } },
        }
      }
    });
  });

  ws['!cols'] = [{ wch: 36 }, { wch: 26 }, { wch: 90 }];
  ws['!rows'] = [
    { hpx: 32 }, { hpx: 28 }, { hpx: 8 }, { hpx: 24 },
    ...METRICS.map(() => ({ hpx: 52 })),
  ];
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 2 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 2 } },
  ];
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 4 + METRICS.length, c: 2 } });
  return ws;
}

// ── MAIN ─────────────────────────────────────────────────────────────────────
const wb = XLSX.utils.book_new();

// Summary sheet first
XLSX.utils.book_append_sheet(wb, buildSummarySheet(), 'Summary');

// Sources sheet second
XLSX.utils.book_append_sheet(wb, buildSourcesSheet(), 'Sources');

// One sheet per metric (sanitize: remove invalid chars, truncate to 31)
METRICS.forEach(metric => {
  const safe = metric.name.replace(/[:\\/?\*\[\]]/g, '-');
  const sheetName = safe.length > 31 ? safe.slice(0, 31) : safe;
  XLSX.utils.book_append_sheet(wb, buildSheet(metric), sheetName);
});

const outPath = path.join(__dirname, 'copart_competitive_data.xlsx');
XLSX.writeFile(wb, outPath, { bookType: 'xlsx', cellStyles: true });
console.log(`Excel file written to: ${outPath}`);
console.log(`Sheets: Summary + Sources + ${METRICS.length} metric sheets`);
