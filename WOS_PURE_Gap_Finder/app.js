const XLSX = window.XLSX;

const ALIAS_MAP = {
  author: ['author full names', 'author names', 'authors', 'author(s)', 'author'],
  title: ['article title', 'publication title', 'paper title', 'title', 'title of the contribution in original language'],
  journal: ['source title', 'journal', 'journal title', 'source publication title'],
  doi: ['doi', 'digital object identifier'],
  year: ['publication year', 'year'],
  recordType: ['document type', 'doc type', 'publication type', 'type', 'item type', 'genre'],
  affiliations: ['addresses', 'affiliations', 'address', 'institution', 'institutions'],
  ut: ['ut (unique wos id)', 'ut unique id', 'unique wos id', 'wos id', 'ut'],
  pureTitle: ['title of the contribution in original language', 'title'],
  pureSubtitle: ['subtitle of the contribution in original language', 'subtitle'],
  pureJournal: ['journal', 'source title', 'journal title', 'publication title'],
  pureYear: ['publication year', 'year'],
  pureRecordType: ['type', 'publication type', 'output type', 'document type', 'item type', 'category']
};

const tabs = [
  { key: 'missing', label: 'Missing', statusFilter: 'missing' },
  { key: 'review', label: 'Needs review', statusFilter: 'needs_review' },
  { key: 'matched', label: 'Matched', statusFilter: 'matched' },
  { key: 'excluded', label: 'Excluded by affiliation rules', statusFilter: 'excluded' }
];

const CHUNK_SIZE = 150;
const MAX_FUZZY_CANDIDATES = 140;

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
  compareRunId: 0,
  wosMapping: {},
  pureMapping: {},
  yearWarning: ''
};

const elements = {
  wosFileInput: document.getElementById('wosFileInput'),
  pureFileInput: document.getElementById('pureFileInput'),
  columnMapping: document.getElementById('columnMapping'),
  yearWarning: document.getElementById('yearWarning'),
  debugOutput: document.getElementById('debugOutput'),
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

function renderDebugInfo(info) {
  if (!elements.debugOutput) return;

  const lines = [
    `Parsed WOS rows: ${info.parsedWos ?? 0}`,
    `Parsed PURE rows: ${info.parsedPure ?? 0}`,
    `Deduplicated WOS rows: ${info.dedupWos ?? 0}`,
    `Deduplicated PURE rows: ${info.dedupPure ?? 0}`,
    `WOS years detected: ${info.wosYears ?? 0}`,
    `PURE years detected: ${info.pureYears ?? 0}`,
    `Shared years: ${info.sharedYears ?? 0}`,
    `Rows after year alignment (WOS/PURE): ${info.afterYearWos ?? 0} / ${info.afterYearPure ?? 0}`,
    `WOS types detected: ${info.wosTypes ?? 0}`,
    `PURE types detected: ${info.pureTypes ?? 0}`,
    `Shared types: ${info.sharedTypes ?? 0}`,
    `Rows after type alignment (WOS/PURE): ${info.afterTypeWos ?? 0} / ${info.afterTypePure ?? 0}`,
    `Rows kept after affiliation filter: ${info.affiliationKept ?? 0}`,
    `Excluded by affiliation: ${info.affiliationExcluded ?? 0}`,
    `Matched / Missing / Review: ${info.matched ?? 0} / ${info.missing ?? 0} / ${info.review ?? 0}`
  ];

  elements.debugOutput.textContent = lines.join('\n');
}

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
    recordType: 'Record Type',
    affiliations: 'Affiliations',
    ut: 'WoS UT',
    pureTitle: 'PURE Title',
    pureSubtitle: 'PURE Subtitle',
    pureJournal: 'PURE Journal',
    pureYear: 'PURE Year',
    pureRecordType: 'PURE Type'
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

function showYearWarning(message) {
  state.yearWarning = message;
  elements.yearWarning.textContent = message;
  elements.yearWarning.classList.remove('hidden');
}

function hideYearWarning() {
  state.yearWarning = '';
  elements.yearWarning.textContent = '';
  elements.yearWarning.classList.add('hidden');
}

function parseBibText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const text = decodeBibliographyText(event.target.result);
        const entries = splitBibEntries(text)
          .map((entry) => entry.trim())
          .filter(Boolean);

        const rows = [[
          'Author Full Names',
          'Article Title',
          'Source Title',
          'DOI',
          'Publication Year',
          'Addresses',
          'UT (Unique WOS ID)',
          'Document Type'
        ]];

        if (entries.length) {
          entries.forEach((entry) => {
            const author = extractBibField(entry, 'author');
            const title = extractBibField(entry, 'title');
            const journal = extractBibField(entry, 'journal') || extractBibField(entry, 'booktitle');
            const doi = extractBibField(entry, 'doi');
            const year = extractBibField(entry, 'year') || extractBibField(entry, 'date');
            const affiliations = extractBibField(entry, 'address') || extractBibField(entry, 'affiliation');
            const ut = extractBibField(entry, 'ut') || extractBibField(entry, 'accessionnumber');
            const docType = normalizeBibEntryType(extractBibEntryType(entry));
            const raw = entry.replace(/\s+/g, ' ').trim();

            rows.push([author || '', title || raw || '', journal || '', doi || '', year || '', affiliations || '', ut || '', docType || '']);
          });
        } else {
          // Fallback for tagged exports saved with .bib extension.
          const risRows = parseRisTaggedText(text);
          const taggedRows = risRows.length ? risRows : parseWosTaggedText(text);
          taggedRows.forEach((row) => {
            rows.push([
              row.author || '',
              row.title || '',
              row.journal || '',
              row.doi || '',
              row.year || '',
              row.affiliations || '',
              row.ut || '',
              row.recordType || ''
            ]);
          });
        }

        if (rows.length <= 1) {
          const looseRows = parseBibByLooseSplit(text);
          looseRows.forEach((row) => {
            rows.push([
              row.author || '',
              row.title || '',
              row.journal || '',
              row.doi || '',
              row.year || '',
              row.affiliations || '',
              row.ut || '',
              row.recordType || ''
            ]);
          });
        }

        resolve(rows);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error('Unable to read BibTeX file.'));
    reader.readAsArrayBuffer(file);
  });
}

function parseBibByLooseSplit(text) {
  const source = sanitizeBibliographyText(text);
  const chunks = source
    .split(/(?=@)/g)
    .map((entry) => entry.trim())
    .filter((entry) => /^@/i.test(entry));

  const output = [];
  chunks.forEach((entry) => {
    const title = extractBibField(entry, 'title');
    const author = extractBibField(entry, 'author');
    const journal = extractBibField(entry, 'journal') || extractBibField(entry, 'booktitle');
    const doi = extractBibField(entry, 'doi');
    const year = extractBibField(entry, 'year') || extractBibField(entry, 'date');
    const affiliations = extractBibField(entry, 'address') || extractBibField(entry, 'affiliation');
    const ut = extractBibField(entry, 'ut') || extractBibField(entry, 'accessionnumber');
    const recordType = normalizeBibEntryType(extractBibEntryType(entry));
    const raw = entry.replace(/\s+/g, ' ').trim();

    if (title || author || doi || year || journal || ut) {
      output.push({
        author: author || '',
        title: title || raw || '',
        journal: journal || '',
        doi: doi || '',
        year: year || '',
        affiliations: affiliations || '',
        ut: ut || '',
        recordType: recordType || ''
      });
    }
  });

  return output;
}

function decodeBibliographyText(input) {
  if (typeof input === 'string') {
    return sanitizeBibliographyText(input);
  }

  if (!(input instanceof ArrayBuffer)) {
    return '';
  }

  const bytes = new Uint8Array(input);
  const utf8 = sanitizeBibliographyText(new TextDecoder('utf-8', { fatal: false }).decode(bytes));
  const utf16 = sanitizeBibliographyText(new TextDecoder('utf-16le', { fatal: false }).decode(bytes));

  const scoreText = (text) => {
    const bib = (text.match(/@[a-z0-9_]+\s*[({]/gi) || []).length;
    const ris = (text.match(/\bTY\s*-/g) || []).length;
    const wos = (text.match(/\bPT\s+[A-Z]/g) || []).length;
    return bib * 3 + ris * 2 + wos;
  };

  return scoreText(utf16) > scoreText(utf8) ? utf16 : utf8;
}

function sanitizeBibliographyText(text) {
  return String(text || '')
    .replace(/^\uFEFF/, '')
    .replace(/\u0000/g, '')
    .replace(/\r\n?/g, '\n');
}

function splitBibEntries(text) {
  const source = String(text || '');
  const entries = [];

  let i = 0;
  while (i < source.length) {
    const at = source.indexOf('@', i);
    if (at === -1) break;

    const head = source.slice(at).match(/^@[A-Za-z0-9_]+\s*[({]/);
    if (!head) {
      i = at + 1;
      continue;
    }

    const openPos = at + head[0].length - 1;
    const openChar = source[openPos];
    const closeChar = openChar === '{' ? '}' : ')';

    let depth = 0;
    let inString = false;
    let escaped = false;
    let endPos = -1;

    for (let j = openPos; j < source.length; j += 1) {
      const ch = source[j];

      if (inString) {
        if (!escaped && ch === '"') {
          inString = false;
        }
        escaped = !escaped && ch === '\\';
        continue;
      }

      if (ch === '"') {
        inString = true;
        escaped = false;
        continue;
      }

      if (ch === openChar) {
        depth += 1;
      } else if (ch === closeChar) {
        depth -= 1;
        if (depth === 0) {
          endPos = j;
          break;
        }
      }
    }

    if (endPos !== -1) {
      entries.push(source.slice(at, endPos + 1).trim());
      i = endPos + 1;
    } else {
      // Keep the remainder as a best-effort final entry if braces are imbalanced.
      entries.push(source.slice(at).trim());
      break;
    }
  }

  if (entries.length) {
    return entries;
  }

  return (source.match(/@[A-Za-z0-9_]+\s*[({][\s\S]*?(?=\n\s*@[A-Za-z0-9_]+\s*[({]|$)/g) || [])
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function extractBibEntryType(entry) {
  const match = String(entry || '').match(/^\s*@([A-Za-z0-9_]+)/);
  return match ? match[1] : '';
}

function normalizeBibEntryType(value) {
  const type = String(value || '').trim().toLowerCase();
  if (!type) return '';
  if (type.includes('article')) return 'journal article';
  if (type.includes('inproceedings') || type.includes('proceedings') || type.includes('conference')) return 'conference paper';
  if (type.includes('inbook') || type.includes('incollection') || type.includes('bookchapter')) return 'book chapter';
  if (type.includes('book')) return 'book';
  if (type.includes('phdthesis') || type.includes('mastersthesis') || type.includes('thesis')) return 'thesis';
  return type;
}

function parseRisTaggedText(text) {
  const blocks = String(text || '')
    .split(/\r?\nER\s*-\s*/i)
    .map((block) => block.trim())
    .filter(Boolean);

  if (!blocks.length || !/\bTY\s*-\s*/i.test(text)) {
    return [];
  }

  return blocks.map((block) => {
    const lines = block.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
    const fields = new Map();

    lines.forEach((line) => {
      const match = line.match(/^([A-Z0-9]{2})\s*-\s*(.*)$/i);
      if (!match) return;
      const key = match[1].toUpperCase();
      const value = match[2] || '';
      if (!fields.has(key)) fields.set(key, []);
      fields.get(key).push(value);
    });

    const first = (key) => (fields.get(key)?.[0] || '');
    const many = (key) => (fields.get(key) || []).filter(Boolean);
    const typeRaw = first('TY');

    return {
      author: many('AU').concat(many('A1')).join('; '),
      title: first('TI') || first('T1'),
      journal: first('JO') || first('JF') || first('T2'),
      doi: first('DO'),
      year: first('PY') || first('Y1') || first('DA'),
      affiliations: first('AD') || first('C1'),
      ut: first('AN') || first('ID'),
      recordType: normalizeRisType(typeRaw)
    };
  }).filter((row) => row.title || row.author || row.doi || row.year);
}

function normalizeRisType(typeRaw) {
  const type = String(typeRaw || '').trim().toUpperCase();
  if (!type) return '';
  if (type === 'JOUR' || type === 'JFULL') return 'journal article';
  if (type === 'CPAPER' || type === 'CONF') return 'conference paper';
  if (type === 'CHAP') return 'book chapter';
  if (type === 'BOOK') return 'book';
  if (type === 'THES') return 'thesis';
  return type.toLowerCase();
}

function parseWosTaggedText(text) {
  const raw = String(text || '');
  if (!/\bPT\s+[A-Z]/m.test(raw) && !/\bTI\s/m.test(raw)) {
    return [];
  }

  const lines = raw.split(/\r?\n/);
  const records = [];
  let current = null;
  let lastTag = '';

  lines.forEach((line) => {
    const tagMatch = line.match(/^([A-Z0-9]{2})\s(.*)$/);
    if (tagMatch) {
      const tag = tagMatch[1];
      const value = (tagMatch[2] || '').trim();

      if (tag === 'PT') {
        if (current) {
          records.push(current);
        }
        current = new Map();
      }

      if (!current) {
        return;
      }

      if (!current.has(tag)) current.set(tag, []);
      if (value) current.get(tag).push(value);
      lastTag = tag;

      if (tag === 'ER') {
        records.push(current);
        current = null;
        lastTag = '';
      }
      return;
    }

    // Continuation line for the previous tag in WoS plain text export.
    if (current && lastTag && /^\s{3,}\S/.test(line)) {
      const value = line.trim();
      if (value) {
        const arr = current.get(lastTag) || [];
        arr.push(value);
        current.set(lastTag, arr);
      }
    }
  });

  if (current) {
    records.push(current);
  }

  const first = (map, key) => (map.get(key)?.[0] || '');
  const many = (map, key) => (map.get(key) || []).filter(Boolean);

  return records.map((record) => {
    const dt = first(record, 'DT') || first(record, 'PT');

    return {
      author: many(record, 'AU').join('; ') || many(record, 'AF').join('; '),
      title: first(record, 'TI') || first(record, 'CT'),
      journal: first(record, 'SO') || first(record, 'JI') || first(record, 'SE'),
      doi: first(record, 'DI'),
      year: first(record, 'PY') || first(record, 'YR') || first(record, 'PD'),
      affiliations: many(record, 'C1').join('; ') || many(record, 'RP').join('; '),
      ut: first(record, 'UT'),
      recordType: normalizeRecordType(dt)
    };
  }).filter((row) => row.title || row.author || row.doi || row.year || row.ut);
}

function extractBibField(entry, fieldName) {
  const fieldRegex = new RegExp(`${fieldName}\\s*=\\s*`, 'i');
  const baseMatch = fieldRegex.exec(entry);
  if (!baseMatch) return '';

  let i = baseMatch.index + baseMatch[0].length;
  while (i < entry.length && /\s/.test(entry[i])) i += 1;
  if (i >= entry.length) return '';

  const start = entry[i];

  if (start === '{') {
    let depth = 0;
    let out = '';
    for (let j = i; j < entry.length; j += 1) {
      const ch = entry[j];
      if (ch === '{') {
        depth += 1;
        if (depth > 1) out += ch;
      } else if (ch === '}') {
        depth -= 1;
        if (depth === 0) {
          return cleanBibFieldValue(out);
        }
        out += ch;
      } else if (depth >= 1) {
        out += ch;
      }
    }
    return cleanBibFieldValue(out);
  }

  if (start === '"') {
    let out = '';
    for (let j = i + 1; j < entry.length; j += 1) {
      const ch = entry[j];
      if (ch === '"' && entry[j - 1] !== '\\') {
        return cleanBibFieldValue(out);
      }
      out += ch;
    }
    return cleanBibFieldValue(out);
  }

  const rest = entry.slice(i).split(/[\n,]/)[0] || '';
  return cleanBibFieldValue(rest);
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
        canonical.recordType = canonical.recordType ?? '';
        canonical.affiliations = canonical.affiliations ?? '';
        canonical.ut = canonical.ut ?? '';
      }

      if (source === 'pure') {
        canonical.title = canonical.pureTitle ?? canonical.title ?? '';
        canonical.subtitle = canonical.pureSubtitle ?? '';
        canonical.journal = canonical.pureJournal ?? canonical.journal ?? '';
        canonical.doi = canonical.doi ?? '';
        canonical.year = canonical.pureYear ?? canonical.year ?? '';
        canonical.recordType = canonical.pureRecordType ?? canonical.recordType ?? '';
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

function fuzzyTitleScoreNormalized(left, right) {
  if (!left || !right) return 0;
  const maxLength = Math.max(left.length, right.length);
  if (maxLength === 0) return 100;
  const distance = levenshteinDistance(left, right);
  return ((maxLength - distance) / maxLength) * 100;
}

function nextUiTick() {
  return new Promise((resolve) => setTimeout(resolve, 0));
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

    // BibTeX exports often omit affiliation fields. Keep these records so the comparison can still run.
    if (!entries.length) {
      return { ...record, affiliationStatus: 'kept' };
    }

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

function parseYearValue(value) {
  const match = String(value ?? '').match(/\b(19|20)\d{2}\b/);
  return match ? Number.parseInt(match[0], 10) : null;
}

function getYearSet(rows) {
  return new Set(
    rows
      .map((row) => parseYearValue(row.year))
      .filter((year) => Number.isInteger(year))
  );
}

function normalizeRecordType(value) {
  const raw = canonicalizeText(value || '');
  if (!raw) return '';

  const compact = raw.replace(/\s+/g, ' ').trim();
  if (/(journal|article|review|letter|editorial|news item)/.test(compact)) return 'journal article';
  if (/(conference|proceedings|meeting|symposium|workshop|paper)/.test(compact)) return 'conference paper';
  if (/(book chapter|chapter in book|chapter)/.test(compact)) return 'book chapter';
  if (/(book|monograph)/.test(compact)) return 'book';
  if (/(thesis|dissertation)/.test(compact)) return 'thesis';
  if (/(preprint|working paper)/.test(compact)) return 'preprint';

  return compact;
}

function getSharedRecordTypes(wosRows, pureRows) {
  const wosTypes = new Set(
    wosRows
      .map((row) => normalizeRecordType(row.recordType))
      .filter(Boolean)
  );

  const pureTypes = new Set(
    pureRows
      .map((row) => normalizeRecordType(row.recordType))
      .filter(Boolean)
  );

  if (!wosTypes.size || !pureTypes.size) {
    return { overlap: null, comparable: false };
  }

  const overlap = new Set([...wosTypes].filter((type) => pureTypes.has(type)));
  return { overlap, comparable: true };
}

function getTypeSet(rows) {
  return new Set(
    rows
      .map((row) => normalizeRecordType(row.recordType))
      .filter(Boolean)
  );
}

function autoAlignDatasets(wosRows, pureRows) {
  const wosYears = getYearSet(wosRows);
  const pureYears = getYearSet(pureRows);
  const sharedYears = new Set([...wosYears].filter((year) => pureYears.has(year)));

  let alignedWos = [...wosRows];
  let alignedPure = [...pureRows];
  const notes = [];
  const meta = {
    wosYears: wosYears.size,
    pureYears: pureYears.size,
    sharedYears: sharedYears.size,
    afterYearWos: wosRows.length,
    afterYearPure: pureRows.length,
    wosTypes: 0,
    pureTypes: 0,
    sharedTypes: 0,
    afterTypeWos: wosRows.length,
    afterTypePure: pureRows.length
  };

  if (wosYears.size && pureYears.size) {
    if (sharedYears.size) {
      alignedWos = alignedWos.filter((row) => sharedYears.has(parseYearValue(row.year)));
      alignedPure = alignedPure.filter((row) => sharedYears.has(parseYearValue(row.year)));

      const sortedYears = [...sharedYears].sort((a, b) => a - b);
      notes.push(`Auto-year filter applied: ${sortedYears[0]}-${sortedYears[sortedYears.length - 1]} (${sharedYears.size} shared year(s)).`);
    } else {
      notes.push('No overlapping publication years found between WOS and PURE, so year filtering was skipped.');
    }
  }

  meta.afterYearWos = alignedWos.length;
  meta.afterYearPure = alignedPure.length;

  const wosTypeSet = getTypeSet(alignedWos);
  const pureTypeSet = getTypeSet(alignedPure);
  meta.wosTypes = wosTypeSet.size;
  meta.pureTypes = pureTypeSet.size;

  const sharedTypesResult = getSharedRecordTypes(alignedWos, alignedPure);
  if (sharedTypesResult.comparable && sharedTypesResult.overlap && sharedTypesResult.overlap.size) {
    alignedWos = alignedWos.filter((row) => {
      const type = normalizeRecordType(row.recordType);
      return type && sharedTypesResult.overlap.has(type);
    });

    alignedPure = alignedPure.filter((row) => {
      const type = normalizeRecordType(row.recordType);
      return type && sharedTypesResult.overlap.has(type);
    });

    notes.push(`Auto-type filter applied: ${sharedTypesResult.overlap.size} shared type(s).`);
  } else if (sharedTypesResult.comparable) {
    notes.push('No overlapping record types found between WOS and PURE, so type filtering was skipped.');
  }

  meta.sharedTypes = sharedTypesResult.overlap?.size || 0;
  meta.afterTypeWos = alignedWos.length;
  meta.afterTypePure = alignedPure.length;

  return { wosRows: alignedWos, pureRows: alignedPure, notes, meta };
}

function getPureTitle(row) {
  return [row.title, row.subtitle].filter(Boolean).join(' ').trim();
}

function createTitleBucketKey(titleKey, yearValue) {
  const prefix = String(titleKey || '').slice(0, 18);
  const lengthBucket = Math.floor(String(titleKey || '').length / 10);
  const yearPart = Number.isInteger(yearValue) ? String(yearValue) : 'na';
  return `${yearPart}|${prefix}|${lengthBucket}`;
}

function buildPureLookup(pureRows) {
  const byDoi = new Map();
  const byExactTitle = new Map();
  const byBucket = new Map();
  const allTitleRecords = [];

  pureRows.forEach((row) => {
    const doi = normalizeDoi(row.doi);
    const title = getPureTitle(row);
    const titleKey = getTitleKey(title);
    const yearValue = parseYearValue(row.year);

    if (doi && !byDoi.has(doi)) {
      byDoi.set(doi, row);
    }

    if (titleKey && !byExactTitle.has(titleKey)) {
      byExactTitle.set(titleKey, row);
    }

    if (!titleKey) {
      return;
    }

    const titleRecord = { row, title, titleKey, yearValue };
    allTitleRecords.push(titleRecord);

    const bucketKeyExactYear = createTitleBucketKey(titleKey, yearValue);
    const bucketKeyAnyYear = createTitleBucketKey(titleKey, null);

    if (!byBucket.has(bucketKeyExactYear)) byBucket.set(bucketKeyExactYear, []);
    if (!byBucket.has(bucketKeyAnyYear)) byBucket.set(bucketKeyAnyYear, []);

    byBucket.get(bucketKeyExactYear).push(titleRecord);
    byBucket.get(bucketKeyAnyYear).push(titleRecord);
  });

  return { byDoi, byExactTitle, byBucket, allTitleRecords };
}

function getFuzzyCandidates(lookup, wosTitleKey, wosYear) {
  if (!wosTitleKey) return [];

  const primaryKey = createTitleBucketKey(wosTitleKey, wosYear);
  const fallbackKey = createTitleBucketKey(wosTitleKey, null);
  const primary = lookup.byBucket.get(primaryKey) || [];
  const fallback = lookup.byBucket.get(fallbackKey) || [];

  const merged = [...primary, ...fallback];
  if (!merged.length) {
    return lookup.allTitleRecords.slice(0, Math.min(MAX_FUZZY_CANDIDATES, lookup.allTitleRecords.length));
  }

  const dedup = [];
  const seen = new Set();
  merged.forEach((candidate) => {
    if (seen.has(candidate)) return;
    seen.add(candidate);
    dedup.push(candidate);
  });

  if (dedup.length <= MAX_FUZZY_CANDIDATES) {
    return dedup;
  }

  return dedup
    .sort((a, b) => Math.abs(a.titleKey.length - wosTitleKey.length) - Math.abs(b.titleKey.length - wosTitleKey.length))
    .slice(0, MAX_FUZZY_CANDIDATES);
}

async function buildComparisonData(runId) {
  const wosRecords = deduplicateRecords(state.wosRows, ['doi', 'ut', 'title']);
  const pureRecords = deduplicateRecords(state.pureRows, ['doi', 'title']);
  const aligned = autoAlignDatasets(wosRecords, pureRecords);
  const alignedWos = aligned.wosRows;
  const alignedPure = aligned.pureRows;

  if (aligned.notes.length) {
    showYearWarning(aligned.notes.join(' '));
  } else {
    hideYearWarning();
  }

  const filteredWos = filterWosAffiliations(alignedWos);
  const lookup = buildPureLookup(alignedPure);

  const results = [];
  for (let index = 0; index < filteredWos.length; index += 1) {
    if (runId !== state.compareRunId) {
      return;
    }

    const record = filteredWos[index];
    const doi = normalizeDoi(record.doi || '');
    const wosTitle = record.title || '';
    const normalizedWosTitle = getTitleKey(wosTitle);
    const wosYearValue = parseYearValue(record.year);

    let status = 'missing';
    let reason = 'no DOI match, no title match';

    if (record.affiliationStatus === 'excluded') {
      status = 'excluded';
      reason = 'excluded by affiliation rules';
      results.push({
        ...record,
        status,
        reason,
        matchStatus: 'Excluded',
        pureMatch: null
      });
      continue;
    }

    const pureMatchByDoi = doi ? lookup.byDoi.get(doi) : null;
    if (pureMatchByDoi) {
      status = 'matched';
      reason = 'DOI matched';
      results.push({
        ...record,
        status,
        reason,
        matchStatus: 'Matched',
        pureMatch: pureMatchByDoi
      });
      continue;
    }

    const exactTitleMatch = normalizedWosTitle ? lookup.byExactTitle.get(normalizedWosTitle) : null;
    if (exactTitleMatch) {
      status = 'matched';
      reason = 'title matched exactly';
      results.push({
        ...record,
        status,
        reason,
        matchStatus: 'Matched',
        pureMatch: exactTitleMatch
      });
      continue;
    }

    let bestMatch = null;
    let bestScore = 0;

    const candidates = getFuzzyCandidates(lookup, normalizedWosTitle, wosYearValue);
    candidates.forEach(({ title, titleKey, row }) => {
      if (!titleKey) return;
      const similarity = fuzzyTitleScoreNormalized(normalizedWosTitle, titleKey);
      if (similarity > bestScore) {
        bestScore = similarity;
        bestMatch = { row, similarity };
      }
    });

    if (bestMatch && bestMatch.similarity >= state.titleThreshold) {
      status = 'needs_review';
      reason = `fuzzy title match (${bestMatch.similarity.toFixed(1)}%) needs review`;
      results.push({
        ...record,
        status,
        reason,
        matchStatus: 'Needs review',
        pureMatch: bestMatch.row,
        score: bestMatch.similarity
      });
    } else {
      results.push({
        ...record,
        status,
        reason,
        matchStatus: 'Missing',
        pureMatch: null
      });
    }

    if ((index + 1) % CHUNK_SIZE === 0) {
      showYearWarning(`Comparing records... ${index + 1}/${filteredWos.length}`);
      await nextUiTick();
    }
  }

  if (runId !== state.compareRunId) {
    return;
  }

  const totalWos = alignedWos.length;
  const filtered = filteredWos.filter((record) => record.affiliationStatus === 'kept').length;
  const affiliationExcluded = filteredWos.filter((record) => record.affiliationStatus === 'excluded').length;
  const matched = results.filter((record) => record.status === 'matched' && record.affiliationStatus !== 'excluded').length;
  const missing = results.filter((record) => record.status === 'missing' && record.affiliationStatus !== 'excluded').length;
  const review = results.filter((record) => record.status === 'needs_review' && record.affiliationStatus !== 'excluded').length;
  const excluded = results.filter((record) => record.status === 'excluded').length;

  renderDebugInfo({
    parsedWos: state.wosRows.length,
    parsedPure: state.pureRows.length,
    dedupWos: wosRecords.length,
    dedupPure: pureRecords.length,
    wosYears: aligned.meta.wosYears,
    pureYears: aligned.meta.pureYears,
    sharedYears: aligned.meta.sharedYears,
    afterYearWos: aligned.meta.afterYearWos,
    afterYearPure: aligned.meta.afterYearPure,
    wosTypes: aligned.meta.wosTypes,
    pureTypes: aligned.meta.pureTypes,
    sharedTypes: aligned.meta.sharedTypes,
    afterTypeWos: aligned.meta.afterTypeWos,
    afterTypePure: aligned.meta.afterTypePure,
    affiliationKept: filtered,
    affiliationExcluded,
    matched,
    missing,
    review
  });

  state.matchResults = results;
  state.summary = { totalWos, filtered, matched, missing, review, excluded };
  if (aligned.notes.length) {
    showYearWarning(aligned.notes.join(' '));
  } else {
    hideYearWarning();
  }
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
    const haystack = `${record.author || ''} ${record.title || ''} ${record.journal || ''} ${record.year || ''} ${record.doi || ''} ${record.ut || ''} ${record.matchStatus || ''} ${record.reason || ''}`.toLowerCase();
    const matchesSearch = !search || haystack.includes(search);
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
    { key: 'recordType', label: 'Document Type' },
    { key: 'affiliations', label: 'Affiliations' },
    { key: 'ut', label: 'UT (Unique WOS ID)' }
  ];

  const pureMappingFields = [
    { key: 'pureTitle', label: 'PURE Title' },
    { key: 'pureSubtitle', label: 'PURE Subtitle' },
    { key: 'pureJournal', label: 'PURE Journal' },
    { key: 'doi', label: 'DOI' },
    { key: 'pureYear', label: 'PURE Publication Year' },
    { key: 'pureRecordType', label: 'PURE Type' }
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

  const wosResponse = { author: state.wosMapping.author || null, title: state.wosMapping.title || null, journal: state.wosMapping.journal || null, doi: state.wosMapping.doi || null, year: state.wosMapping.year || null, recordType: state.wosMapping.recordType || null, affiliations: state.wosMapping.affiliations || null, ut: state.wosMapping.ut || null };
  const pureResponse = { pureTitle: state.pureMapping.pureTitle || null, pureSubtitle: state.pureMapping.pureSubtitle || null, pureJournal: state.pureMapping.pureJournal || null, doi: state.pureMapping.doi || null, pureYear: state.pureMapping.pureYear || null, pureRecordType: state.pureMapping.pureRecordType || null };

  const finalWos = buildRowsFromSheet(state.wosRawRows, wosResponse, 'wos').rows;
  const finalPure = buildRowsFromSheet(state.pureRawRows, pureResponse, 'pure').rows;

  state.wosRows = finalWos.map((row) => ({
    author: row.author || '',
    title: row.title || '',
    journal: row.journal || '',
    doi: row.doi || '',
    year: row.year || '',
    recordType: row.recordType || '',
    affiliations: row.affiliations || '',
    ut: row.ut || ''
  }));

  state.pureRows = finalPure.map((row) => ({
    title: row.title || '',
    subtitle: row.subtitle || '',
    journal: row.journal || '',
    doi: row.doi || '',
    year: row.year || '',
    recordType: row.recordType || ''
  }));

  compareAndRender();
}

async function compareAndRender() {
  if (!state.wosRows.length || !state.pureRows.length) {
    const wosCount = state.wosRows.length;
    const pureCount = state.pureRows.length;
    showYearWarning(`Waiting for comparable data. Parsed WOS: ${wosCount}, PURE: ${pureCount}.`);
    renderDebugInfo({ parsedWos: wosCount, parsedPure: pureCount });
    state.summary = { totalWos: wosCount, filtered: 0, matched: 0, missing: 0, review: 0, excluded: 0 };
    renderSummary();
    return;
  }

  const runId = state.compareRunId + 1;
  state.compareRunId = runId;
  showYearWarning('Preparing comparison...');
  await nextUiTick();
  await buildComparisonData(runId);
}

function parseUploadedFile(file, source) {
  const type = getAcceptedFileType(file.name);
  if (type === 'unsupported') {
    elements.yearWarning.textContent = 'Unsupported file type. Please upload a BibTeX (.bib), CSV (.csv), or XLSX (.xlsx) file.';
    elements.yearWarning.classList.remove('hidden');
    return Promise.resolve();
  }

  showYearWarning(`Parsing ${source.toUpperCase()} file: ${file.name} ...`);

  const parser = type === 'bib' ? parseBibText(file) : parseWorkbook(file);
  return parser.then((rows) => {
    if (elements.debugOutput) {
      const existing = elements.debugOutput.textContent || '';
      elements.debugOutput.textContent = `Last parse: ${source.toUpperCase()} rows including header = ${rows.length}\n${existing}`;
    }

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
          recordType: row.recordType || '',
          affiliations: row.affiliations || '',
          ut: row.ut || ''
        };
      }

      return {
        title: row.title || '',
        subtitle: row.subtitle || '',
        journal: row.journal || '',
        doi: row.doi || '',
        year: row.year || '',
        recordType: row.recordType || ''
      };
    });

    if (source === 'wos') {
      state.wosRows = parsed;
    } else {
      state.pureRows = parsed;
    }

    if (!parsed.length) {
      showYearWarning(`No records were parsed from ${file.name}. Please verify the file content and format.`);
    } else if (state.wosRows.length && state.pureRows.length) {
      hideYearWarning();
    }

    renderMappingControls();
    compareAndRender();
  }).catch(() => {
    showYearWarning(`Failed to parse ${file.name}. Please upload a valid BibTeX, CSV, or XLSX file.`);
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
  // Allow selecting the same file repeatedly and still trigger parsing.
  elements.wosFileInput.addEventListener('click', () => {
    elements.wosFileInput.value = '';
  });

  elements.pureFileInput.addEventListener('click', () => {
    elements.pureFileInput.value = '';
  });

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
  if (!XLSX) {
    showYearWarning('Spreadsheet engine failed to load. Please refresh the page and try again.');
  }
  renderTabs();
  renderSummary();
  renderDebugInfo({ parsedWos: 0, parsedPure: 0 });

  // If the browser preserved file selections across reload, parse them immediately.
  const wosFile = elements.wosFileInput.files?.[0];
  const pureFile = elements.pureFileInput.files?.[0];
  if (wosFile) parseUploadedFile(wosFile, 'wos');
  if (pureFile) parseUploadedFile(pureFile, 'pure');
});
