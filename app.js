import * as XLSX from 'xlsx';

const ALIAS_MAP = {
  author: ['author full names', 'author names', 'authors', 'author(s)', 'author'],
  title: ['article title', 'publication title', 'paper title', 'title', 'title of the contribution in original language'],
  journal: ['source title', 'journal', 'journal title', 'source publication title'],
  doi: ['doi', 'digital object identifier'],
  year: ['publication year', 'year'],
  affiliations: ['addresses', 'affiliations', 'address', 'institution', 'institutions'],
  ut: ['ut (unique wos id)', 'ut unique id', 'unique wos id', 'wos id', 'ut'],
  pureTitle: ['title of the contribution in original language', 'title'],
  pureSubtitle: ['subtitle of the contribution in original language', 'subtitle'],
  pureJournal: ['journal', 'source title', 'journal title', 'publication title'],
  pureYear: ['publication year', 'year']
};

const tabs = [
  { key: 'missing', label: 'Missing', statusFilter: 'missing' },
  { key: 'review', label: 'Needs review', statusFilter: 'needs_review' },
  { key: 'matched', label: 'Matched', statusFilter: 'matched' },
  { key: 'excluded', label: 'Excluded by affiliation rules', statusFilter: 'excluded' }
];

const state = {
  wosHeaders: [],
  pureHeaders: [],
  wosRows: [],
  pureRows: [],
  matchResults: [],
  selectedTab: 'missing',
  searchText: '',
  sortKey: 'title',
  sortDir: 'asc',
  titleThreshold: 95,
  wosMapping: {},
  pureMapping: {},
  yearWarning: ''
};

const elements = {
  wosFileInput: document.getElementById('wosFileInput'),
  pureFileInput: document.getElementById('pureFileInput'),
  columnMapping: document.getElementById('columnMapping'),
  yearWarning: document.getElementById('yearWarning'),
  exportBtn: document.getElementById('exportBtn'),
  searchInput: document.getElementById('searchInput'),
  resultsTableBody: document.getElementById('resultsTableBody'),
  tabs: document.getElementById('tabs'),
  totalWosCount: document.getElementById('totalWosCount'),
  filteredWosCount: document.getElementById('filteredWosCount'),
  matchedCount: document.getElementById('matchedCount'),
  missingCount: document.getElementById('missingCount'),
  reviewCount: document.getElementById('reviewCount')
};

function normalizeHeader(value) {
  return String(value ?? '')
    .toLowerCase()
    .normalize('NFKC')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function getSectionLabel(key) {
  const map = {
    author: 'Author(s)',
    title: 'Title',
    journal: 'Journal',
    doi: 'DOI',
    year: 'Publication Year',
    affiliations: 'Affiliations',
    ut: 'WoS UT',
    pureTitle: 'PURE Title',
    pureSubtitle: 'PURE Subtitle',
    pureJournal: 'PURE Journal',
    pureYear: 'PURE Year'
  };
  return map[key] || key;
}

function findHeaderRow(rows) {
  let bestIndex = 0;
  let bestScore = -1;

  rows.forEach((row, index) => {
    if (!Array.isArray(row)) return;
    const score = row.reduce((sum, cell) => {
      const normalized = normalizeHeader(cell);
      const aliases = Object.values(ALIAS_MAP).flat();
      return sum + (aliases.includes(normalized) ? 2 : 0);
    }, 0);

    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestScore >= 1 ? bestIndex : 0;
}

function getAcceptedFileType(fileName) {
  const name = String(fileName || '').toLowerCase();
  if (name.endsWith('.bib') || name.endsWith('.bibtex')) return 'bib';
  if (name.endsWith('.csv')) return 'csv';
  if (name.endsWith('.xlsx') || name.endsWith('.xls')) return 'xlsx';
  return 'unsupported';
}

function parseBibText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const text = String(event.target.result || '');
        const entries = text
          .split(/(?=@[A-Za-z0-9_]+)/)
          .map((entry) => entry.trim())
          .filter(Boolean);

        const rows = [[
          'Author Full Names',
          'Article Title',
          'Source Title',
          'DOI',
          'Publication Year',
          'Addresses',
          'UT (Unique WOS ID)'
        ]];

        entries.forEach((entry) => {
          const author = extractBibField(entry, 'author');
          const title = extractBibField(entry, 'title');
          const journal = extractBibField(entry, 'journal') || extractBibField(entry, 'booktitle');
          const doi = extractBibField(entry, 'doi');
          const year = extractBibField(entry, 'year') || extractBibField(entry, 'date');
          const raw = entry.replace(/\s+/g, ' ').trim();

          rows.push([author || '', title || raw || '', journal || '', doi || '', year || '', '', '']);
        });

        resolve(rows.length > 1 ? rows : [['Author Full Names','Article Title','Source Title','DOI','Publication Year','Addresses','UT (Unique WOS ID)'], ['', '', '', '', '', '', '']]);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error('Unable to read BibTeX file.'));
    reader.readAsText(file, 'UTF-8');
  });
}

function extractBibField(entry, fieldName) {
  const regex = new RegExp(`${fieldName}\\s*=\\s*\\{([^}]*)\}`, 'i');
  const match = entry.match(regex);
  if (match) return cleanBibFieldValue(match[1]);

  const quotedRegex = new RegExp(`${fieldName}\\s*=\\s*"([^"]*)"`, 'i');
  const quoted = entry.match(quotedRegex);
  return quoted ? cleanBibFieldValue(quoted[1]) : '';
}

function cleanBibFieldValue(value) {
  return String(value || '')
    .replace(/\s+and\s+/gi, '; ')
    .replace(/\{([^}]+)\}/g, '$1')
    .replace(/\"/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const fileType = getAcceptedFileType(file.name);
        const workbook = XLSX.read(event.target.result, {
          type: fileType === 'csv' ? 'string' : 'array',
          cellDates: true,
          raw: false,
          WTF: false
        });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
        resolve(rows);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error('Unable to read file.'));
    if (getAcceptedFileType(file.name) === 'csv') {
      reader.readAsText(file, 'UTF-8');
    } else {
      reader.readAsArrayBuffer(file);
    }
  });
}

function resolveMapping(headers, aliases) {
  const headerNames = headers.map((header) => normalizeHeader(header));
  const mapping = {};

  Object.entries(aliases).forEach(([key, aliasList]) => {
    const found = headerNames.findIndex((header) => aliasList.some((alias) => header === normalizeHeader(alias)));
    mapping[key] = found >= 0 ? headers[found] : null;
  });

  return mapping;
}

function buildRowsFromSheet(rows, mapping, source) {
  const headerIndex = findHeaderRow(rows);
  const headerCells = rows[headerIndex] || [];
  const semanticHeaders = headerCells.map((value) => String(value ?? '').trim());
  const dataRows = rows.slice(headerIndex + 1);

  const resolved = { ...mapping };

  const finalRows = dataRows
    .filter((row) => Array.isArray(row) && row.some((cell) => String(cell ?? '').trim() !== ''))
    .map((row) => {
      const record = {};
      semanticHeaders.forEach((header, index) => {
        const key = normalizeHeader(header);
        record[key] = row[index] ?? '';
      });

      const canonical = {};
      Object.entries(resolved).forEach(([field, actualHeader]) => {
        if (!actualHeader) return;
        const key = normalizeHeader(actualHeader);
        canonical[field] = record[key] ?? '';
      });

      if (source === 'wos') {
        canonical.author = canonical.author ?? '';
        canonical.title = canonical.title ?? '';
        canonical.journal = canonical.journal ?? '';
        canonical.doi = canonical.doi ?? '';
        canonical.year = canonical.year ?? '';
        canonical.affiliations = canonical.affiliations ?? '';
        canonical.ut = canonical.ut ?? '';
      }

      if (source === 'pure') {
        canonical.title = canonical.pureTitle ?? canonical.title ?? '';
        canonical.subtitle = canonical.pureSubtitle ?? '';
        canonical.journal = canonical.pureJournal ?? canonical.journal ?? '';
        canonical.doi = canonical.doi ?? '';
        canonical.year = canonical.pureYear ?? canonical.year ?? '';
      }

      return canonical;
    });

  return { headerIndex, headerCells: semanticHeaders, rows: finalRows };
}

function normalizeDoi(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\/doi\.org\//i, '')
    .replace(/^doi:\s*/i, '')
    .replace(/\s+/g, '')
    .replace(/\u00a0/g, '');
}

function canonicalizeText(value) {
  if (value === null || value === undefined) return '';
  let text = String(value)
    .replace(/<[^>]+>/g, ' ')
    .replace(/\\[a-zA-Z]+\{([^}]+)\}/g, '$1')
    .replace(/\$\$|\$/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\u2013|\u2014|\u2012/g, ' - ')
    .replace(/[\p{P}\p{S}]+/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return text.toLowerCase();
}

function getTitleKey(value) {
  return canonicalizeText(value);
}

function stripHtmlEntities(value) {
  return String(value ?? '')
    .replace(/&amp;/g, '&')
    .replace(/&nbsp;/g, ' ')
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, '"');
}

function levenshteinDistance(a, b) {
  if (a === b) return 0;
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;

  const matrix = Array.from({ length: a.length + 1 }, () => new Array(b.length + 1).fill(0));

  for (let i = 0; i <= a.length; i += 1) matrix[i][0] = i;
  for (let j = 0; j <= b.length; j += 1) matrix[0][j] = j;

  for (let i = 1; i <= a.length; i += 1) {
    for (let j = 1; j <= b.length; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost
      );
    }
  }

  return matrix[a.length][b.length];
}

function fuzzyTitleScore(a, b) {
  if (!a || !b) return 0;
  const left = canonicalizeText(removeExtraWhitespace(stripHtmlEntities(a)));
  const right = canonicalizeText(removeExtraWhitespace(stripHtmlEntities(b)));
  if (!left || !right) return 0;

  const maxLength = Math.max(left.length, right.length);
  if (maxLength === 0) return 100;
  const distance = levenshteinDistance(left, right);
  return ((maxLength - distance) / maxLength) * 100;
}

function removeExtraWhitespace(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function normalizeAffiliation(value) {
  if (!value) return '';
  return String(value)
    .toLowerCase()
    .normalize('NFKC')
    .replace(/&/g, 'and')
    .replace(/[^a-z0-9\s()]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function affiliationMatchesMainCampus(value) {
  const normalized = normalizeAffiliation(value);
  if (!normalized) return false;

  const hasHongKong = normalized.includes('hong kong') || normalized.includes('hk');
  const hasUniversity = normalized.includes('university') || normalized.includes('univ');
  const hasScience = normalized.includes('science') || normalized.includes('sci');
  const hasTech = normalized.includes('technology') || normalized.includes('technol') || normalized.includes('tech');
  const hasHKUST = normalized.includes('hkust') || normalized.includes('ust');
  const hasGuangzhou = normalized.includes('guangzhou') || normalized.includes('gzu');
  const hasMainCampus = hasHongKong && ((hasUniversity && (hasScience || hasTech)) || hasHKUST || (hasScience && hasTech));

  return hasMainCampus && !hasGuangzhou;
}

function affiliationMatchesGuangzhou(value) {
  const normalized = normalizeAffiliation(value);
  if (!normalized) return false;

  return (normalized.includes('guangzhou') || normalized.includes('gzu')) && (
    normalized.includes('hong kong') || normalized.includes('hkust') || normalized.includes('science') || normalized.includes('technol')
  );
}

function filterWosAffiliations(records) {
  return records.map((record) => {
    const entries = String(record.affiliations ?? '')
      .split(';')
      .map((entry) => entry.trim())
      .filter(Boolean);

    const hasMainCampus = entries.some((entry) => affiliationMatchesMainCampus(entry));
    const hasGuangzhouCampus = entries.some((entry) => affiliationMatchesGuangzhou(entry));
    const hasAnyHKUSTReference = entries.some((entry) => /hong kong|hkust|science|technol|university|univ/i.test(entry));

    if (hasMainCampus) {
      return { ...record, affiliationStatus: 'kept' };
    }

    if (hasGuangzhouCampus && !hasMainCampus) {
      return { ...record, affiliationStatus: 'excluded' };
    }

    if (!hasAnyHKUSTReference) {
      return { ...record, affiliationStatus: 'excluded' };
    }

    return { ...record, affiliationStatus: 'excluded' };
  });
}

function deduplicateRecords(rows, keyNames) {
  const seenKeys = new Set();
  const output = [];

  rows.forEach((row) => {
    const keyParts = keyNames
      .map((field) => {
        if (field === 'doi') return normalizeDoi(row.doi);
        if (field === 'ut') return String(row.ut ?? '').trim().toLowerCase();
        if (field === 'title') return getTitleKey(row.title || row.pureTitle || '');
        return '';
      })
      .filter(Boolean);

    const key = keyParts.join('|');
    if (!key) {
      output.push(row);
      return;
    }

    if (seenKeys.has(key)) {
      return;
    }

    seenKeys.add(key);
    output.push(row);
  });

  return output;
}

function getPureTitle(row) {
  return [row.title, row.subtitle].filter(Boolean).join(' ').trim();
}

function buildComparisonData() {
  const wosRecords = deduplicateRecords(state.wosRows, ['doi', 'ut', 'title']);
  const pureRecords = deduplicateRecords(state.pureRows, ['doi', 'title']);

  const filteredWos = filterWosAffiliations(wosRecords);

  const pureMapByDoi = new Map();
  pureRecords.forEach((row) => {
    const doi = normalizeDoi(row.doi);
    if (doi) pureMapByDoi.set(doi, row);
  });

  const pureTitles = pureRecords.map((row) => ({
    title: getPureTitle(row),
    row
  }));

  const results = filteredWos.map((record) => {
    const doi = normalizeDoi(record.doi || '');
    const wosTitle = record.title || '';
    const normalizedWosTitle = getTitleKey(wosTitle);

    let status = 'missing';
    let reason = 'no DOI match, no title match';

    if (record.affiliationStatus === 'excluded') {
      status = 'excluded';
      reason = 'excluded by affiliation rules';
      return {
        ...record,
        status,
        reason,
        matchStatus: 'Excluded',
        pureMatch: null
      };
    }

    const pureMatchByDoi = doi ? pureMapByDoi.get(doi) : null;
    if (pureMatchByDoi) {
      status = 'matched';
      reason = 'DOI matched';
      return {
        ...record,
        status,
        reason,
        matchStatus: 'Matched',
        pureMatch: pureMatchByDoi
      };
    }

    const exactTitleMatches = pureTitles.filter(({ title }) => title && getTitleKey(title) === normalizedWosTitle);
    if (exactTitleMatches.length > 0) {
      status = 'matched';
      reason = 'title matched exactly';
      return {
        ...record,
        status,
        reason,
        matchStatus: 'Matched',
        pureMatch: exactTitleMatches[0].row
      };
    }

    let bestMatch = null;
    let bestScore = 0;

    pureTitles.forEach(({ title, row }) => {
      if (!title) return;
      const similarity = fuzzyTitleScore(wosTitle, title);
      if (similarity > bestScore) {
        bestScore = similarity;
        bestMatch = { row, similarity };
      }
    });

    if (bestMatch && bestMatch.similarity >= state.titleThreshold) {
      status = 'needs_review';
      reason = `fuzzy title match (${bestMatch.similarity.toFixed(1)}%) needs review`;
      return {
        ...record,
        status,
        reason,
        matchStatus: 'Needs review',
        pureMatch: bestMatch.row,
        score: bestMatch.similarity
      };
    }

    return {
      ...record,
      status,
      reason,
      matchStatus: 'Missing',
      pureMatch: null
    };
  });

  const totalWos = wosRecords.length;
  const filtered = filteredWos.filter((record) => record.affiliationStatus === 'kept').length;
  const matched = results.filter((record) => record.status === 'matched' && record.affiliationStatus !== 'excluded').length;
  const missing = results.filter((record) => record.status === 'missing' && record.affiliationStatus !== 'excluded').length;
  const review = results.filter((record) => record.status === 'needs_review' && record.affiliationStatus !== 'excluded').length;
  const excluded = results.filter((record) => record.status === 'excluded').length;

  state.matchResults = results;
  state.summary = { totalWos, filtered, matched, missing, review, excluded };
  renderSummary();
  renderTabs();
  renderTable();
}

function renderSummary() {
  const { totalWos, filtered, matched, missing, review } = state.summary || { totalWos: 0, filtered: 0, matched: 0, missing: 0, review: 0 };

  elements.totalWosCount.textContent = totalWos;
  elements.filteredWosCount.textContent = filtered;
  elements.matchedCount.textContent = matched;
  elements.missingCount.textContent = missing;
  elements.reviewCount.textContent = review;
}

function renderTabs() {
  elements.tabs.innerHTML = tabs
    .map(
      (tab) => `
        <button type="button" class="tab-btn ${state.selectedTab === tab.key ? 'active' : ''}" data-tab="${tab.key}">
          ${tab.label}
        </button>
      `
    )
    .join('');

  elements.tabs.querySelectorAll('.tab-btn').forEach((button) => {
    button.addEventListener('click', () => {
      state.selectedTab = button.dataset.tab;
      renderTabs();
      renderTable();
    });
  });
}

function renderTable() {
  const activeStatus = tabs.find((tab) => tab.key === state.selectedTab)?.statusFilter || 'missing';
  const search = state.searchText.trim().toLowerCase();

  const missingCount = state.matchResults.filter((record) => record.status === 'missing').length;
  elements.exportBtn.classList.toggle('hidden', missingCount === 0);

  const filteredRows = state.matchResults.filter((record) => {
    const statusMatch = record.status === activeStatus;
    const matchesSearch = !search || `${record.author || ''} ${record.title || ''} ${record.journal || ''}`.toLowerCase().includes(search);
    return statusMatch && matchesSearch;
  });

  const sortedRows = [...filteredRows].sort((a, b) => {
    const left = String(a[state.sortKey] ?? '').toLowerCase();
    const right = String(b[state.sortKey] ?? '').toLowerCase();
    if (left === right) return 0;
    return state.sortDir === 'asc' ? left.localeCompare(right) : right.localeCompare(left);
  });

  if (!sortedRows.length) {
    elements.resultsTableBody.innerHTML = '<tr><td colspan="8" class="empty-state">No records match the active filter.</td></tr>';
    return;
  }

  elements.resultsTableBody.innerHTML = sortedRows
    .map((row) => {
      const doi = row.doi ? `<a class="doi-link" href="${row.doi.startsWith('http') ? row.doi : `https://doi.org/${row.doi}`}" target="_blank" rel="noopener noreferrer">${row.doi}</a>` : '—';
      const statusClass = row.status === 'matched' ? 'status-matched' : row.status === 'missing' ? 'status-missing' : row.status === 'needs_review' ? 'status-review' : 'status-excluded';
      const badgeLabel = row.status === 'matched' ? 'Matched' : row.status === 'missing' ? 'Missing' : row.status === 'needs_review' ? 'Needs review' : 'Excluded';

      return `
        <tr>
          <td>${escapeHtml(row.author || '—')}</td>
          <td>${escapeHtml(row.title || '—')}</td>
          <td>${escapeHtml(row.journal || '—')}</td>
          <td>${escapeHtml(row.year || '—')}</td>
          <td>${doi}</td>
          <td>${escapeHtml(row.ut || '—')}</td>
          <td><span class="status-badge ${statusClass}">${badgeLabel}</span></td>
          <td>${escapeHtml(row.reason || '—')}</td>
        </tr>
      `;
    })
    .join('');

  document.querySelectorAll('thead th[data-sort]').forEach((header) => {
    header.onclick = () => {
      const sortKey = header.dataset.sort;
      if (state.sortKey === sortKey) {
        state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
      } else {
        state.sortKey = sortKey;
        state.sortDir = 'asc';
      }
      renderTable();
    };
  });
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function ensureMappingValue(value, fallback) {
  return value || fallback;
}

function renderMappingControls() {
  if (!state.wosHeaders.length || !state.pureHeaders.length) {
    elements.columnMapping.classList.add('hidden');
    return;
  }

  const wosOptions = state.wosHeaders.map((header) => `<option value="${escapeHtml(header)}">${escapeHtml(header)}</option>`).join('');
  const pureOptions = state.pureHeaders.map((header) => `<option value="${escapeHtml(header)}">${escapeHtml(header)}</option>`).join('');

  const wosMappingFields = [
    { key: 'author', label: 'Author Full Names' },
    { key: 'title', label: 'Article Title' },
    { key: 'journal', label: 'Source Title' },
    { key: 'doi', label: 'DOI' },
    { key: 'year', label: 'Publication Year' },
    { key: 'affiliations', label: 'Affiliations' },
    { key: 'ut', label: 'UT (Unique WOS ID)' }
  ];

  const pureMappingFields = [
    { key: 'pureTitle', label: 'PURE Title' },
    { key: 'pureSubtitle', label: 'PURE Subtitle' },
    { key: 'pureJournal', label: 'PURE Journal' },
    { key: 'doi', label: 'DOI' },
    { key: 'pureYear', label: 'PURE Publication Year' }
  ];

  const renderSelect = (field, label, options, currentValue) => `
    <div class="map-field">
      <label for="${field}">${label}</label>
      <select id="${field}" data-target="${field}">
        <option value="">Auto-detect</option>
        ${options}
      </select>
    </div>
  `;

  elements.columnMapping.innerHTML = `
    <h3>Column mapping</h3>
    <div class="mapping-grid">
      <div>
        <h4>Web of Science</h4>
        <div class="mapping-grid">
          ${wosMappingFields.map((field) => renderSelect(`wos_${field.key}`, field.label, wosOptions, state.wosMapping[field.key] || '')).join('')}
        </div>
      </div>
      <div>
        <h4>PURE</h4>
        <div class="mapping-grid">
          ${pureMappingFields.map((field) => renderSelect(`pure_${field.key}`, field.label, pureOptions, state.pureMapping[field.key] || '')).join('')}
        </div>
      </div>
    </div>
  `;

  elements.columnMapping.classList.remove('hidden');

  elements.columnMapping.querySelectorAll('select').forEach((select) => {
    select.value = select.dataset.target.includes('wos') ? (state.wosMapping[select.dataset.target.replace('wos_', '')] || '') : (state.pureMapping[select.dataset.target.replace('pure_', '')] || '');
    select.addEventListener('change', () => {
      const key = select.dataset.target.replace('wos_', '').replace('pure_', '');
      const source = select.dataset.target.startsWith('wos') ? 'wos' : 'pure';
      if (source === 'wos') {
        state.wosMapping[key] = select.value || null;
      } else {
        state.pureMapping[key] = select.value || null;
      }
      applySelectedMappings();
    });
  });
}

function applySelectedMappings() {
  if (!state.wosRows.length || !state.pureRows.length) return;

  const wosResponse = { author: state.wosMapping.author || null, title: state.wosMapping.title || null, journal: state.wosMapping.journal || null, doi: state.wosMapping.doi || null, year: state.wosMapping.year || null, affiliations: state.wosMapping.affiliations || null, ut: state.wosMapping.ut || null };
  const pureResponse = { pureTitle: state.pureMapping.pureTitle || null, pureSubtitle: state.pureMapping.pureSubtitle || null, pureJournal: state.pureMapping.pureJournal || null, doi: state.pureMapping.doi || null, pureYear: state.pureMapping.pureYear || null };

  const finalWos = buildRowsFromSheet(state.wosRawRows, wosResponse, 'wos').rows;
  const finalPure = buildRowsFromSheet(state.pureRawRows, pureResponse, 'pure').rows;

  state.wosRows = finalWos.map((row) => ({
    author: row.author || '',
    title: row.title || '',
    journal: row.journal || '',
    doi: row.doi || '',
    year: row.year || '',
    affiliations: row.affiliations || '',
    ut: row.ut || ''
  }));

  state.pureRows = finalPure.map((row) => ({
    title: row.title || '',
    subtitle: row.subtitle || '',
    journal: row.journal || '',
    doi: row.doi || '',
    year: row.year || ''
  }));

  compareAndRender();
}

function compareAndRender() {
  if (!state.wosRows.length || !state.pureRows.length) {
    return;
  }

  const wosYears = state.wosRows
    .map((row) => Number.parseInt(String(row.year || '').replace(/[^0-9]/g, ''), 10))
    .filter((year) => Number.isInteger(year));
  const pureYears = state.pureRows
    .map((row) => Number.parseInt(String(row.year || '').replace(/[^0-9]/g, ''), 10))
    .filter((year) => Number.isInteger(year));

  const overlap = [...new Set(wosYears.filter((year) => pureYears.includes(year)))];
  if (!overlap.length) {
    state.yearWarning = 'Warning: the two files do not overlap on publication year. Please confirm they are filtered to the same year range.';
    elements.yearWarning.textContent = state.yearWarning;
    elements.yearWarning.classList.remove('hidden');
  } else {
    state.yearWarning = '';
    elements.yearWarning.textContent = '';
    elements.yearWarning.classList.add('hidden');
  }

  buildComparisonData();
}

function parseUploadedFile(file, source) {
  const type = getAcceptedFileType(file.name);
  if (type === 'unsupported') {
    elements.yearWarning.textContent = 'Unsupported file type. Please upload a BibTeX (.bib), CSV (.csv), or XLSX (.xlsx) file.';
    elements.yearWarning.classList.remove('hidden');
    return Promise.resolve();
  }

  const parser = type === 'bib' ? parseBibText(file) : parseWorkbook(file);
  return parser.then((rows) => {
    const mapping = source === 'wos' ? resolveMapping(rows[findHeaderRow(rows)] || [], ALIAS_MAP) : resolveMapping(rows[findHeaderRow(rows)] || [], ALIAS_MAP);

    if (source === 'wos') {
      state.wosRawRows = rows;
      state.wosHeaders = rows[findHeaderRow(rows)] || [];
      state.wosMapping = mapping; 
    } else {
      state.pureRawRows = rows;
      state.pureHeaders = rows[findHeaderRow(rows)] || [];
      state.pureMapping = mapping;
    }

    const resolved = buildRowsFromSheet(rows, mapping, source);
    const parsed = resolved.rows.map((row) => {
      if (source === 'wos') {
        return {
          author: row.author || '',
          title: row.title || '',
          journal: row.journal || '',
          doi: row.doi || '',
          year: row.year || '',
          affiliations: row.affiliations || '',
          ut: row.ut || ''
        };
      }

      return {
        title: row.title || '',
        subtitle: row.subtitle || '',
        journal: row.journal || '',
        doi: row.doi || '',
        year: row.year || ''
      };
    });

    if (source === 'wos') {
      state.wosRows = parsed;
    } else {
      state.pureRows = parsed;
    }

    renderMappingControls();
    compareAndRender();
  });
}

function exportMissingRows() {
  const missingRows = state.matchResults.filter((record) => record.status === 'missing');
  const exportData = missingRows.map((record) => ({
    Author: record.author || '',
    Title: record.title || '',
    Journal: record.journal || '',
    Year: record.year || '',
    DOI: record.doi || '',
    'WoS UT': record.ut || '',
    Status: 'Missing',
    Reason: record.reason || 'no DOI match, no title match'
  }));

  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Missing records');
  XLSX.writeFile(wb, 'missing_wos_records.xlsx');
}

function wireEvents() {
  elements.wosFileInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;
    parseUploadedFile(file, 'wos');
  });

  elements.pureFileInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;
    parseUploadedFile(file, 'pure');
  });

  elements.searchInput.addEventListener('input', (event) => {
    state.searchText = event.target.value;
    renderTable();
  });

  elements.exportBtn.addEventListener('click', exportMissingRows);
}

wireEvents();

window.addEventListener('DOMContentLoaded', () => {
  renderTabs();
  renderSummary();
});
