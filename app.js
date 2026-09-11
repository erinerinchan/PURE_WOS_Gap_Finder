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
  pureTitle: ['title of the contribution in original language', 'article title', 'publication title', 'paper title', 'title'],
  pureSubtitle: ['subtitle of the contribution in original language', 'subtitle'],
  pureJournal: ['journal', 'source title', 'journal title', 'publication title'],
  pureYear: ['publication year', 'year'],
  pureRecordType: ['type', 'publication type', 'output type', 'document type', 'item type', 'category']
};

const tabs = [
  { key: 'missing', label: 'Missing from PURE/Scopus', statusFilter: 'missing' }
];

const CHUNK_SIZE = 150;
const MAX_FUZZY_CANDIDATES = 140;
const MATCH_THRESHOLD = 97;
const REVIEW_THRESHOLD = 90;
const YEAR_TOLERANCE = 1;
const AMBIGUITY_SCORE_GAP = 0.75;
const FUZZY_PREFILTER_YEAR_DELTA = 2;
const SCORE_MARGIN_THRESHOLD = 1.0;
const SECOND_PASS_FUZZY_CANDIDATES = 480;
const TITLE_MAPPING_EMPTY_RATIO_THRESHOLD = 0.4;
const TITLE_MAPPING_AVG_LENGTH_THRESHOLD = 12;

const TITLE_STOP_WORDS = new Set(['the', 'a', 'an', 'and', 'of', 'for', 'to', 'in', 'on', 'with', 'by', 'from']);

const TITLE_SCORE_WEIGHTS = {
  tokenJaccard: 0.45,
  trigramJaccard: 0.35,
  editSimilarity: 0.2
};

const METADATA_WEIGHTS = {
  title: 0.78,
  year: 0.08,
  recordType: 0.06,
  journal: 0.05,
  firstAuthor: 0.03
};

const METADATA_PENALTIES = {
  doiConflict: 8,
  yearHardConflict: 14,
  typeHardConflict: 16
};

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
  yearWarning: '',
  selectedYearFilter: 'all',
  availableYears: []
};

const elements = {
  wosFileInput: document.getElementById('wosFileInput'),
  pureFileInput: document.getElementById('pureFileInput'),
  columnMapping: document.getElementById('columnMapping'),
  yearWarning: document.getElementById('yearWarning'),
  debugOutput: document.getElementById('debugOutput'),
  exportBtn: document.getElementById('exportBtn'),
  yearFilterSelect: document.getElementById('yearFilterSelect'),
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
    `Active year filter: ${info.activeYearFilter ?? 'All years'}`,
    `Available years: ${info.availableYearsCount ?? 0}`,
    `Rows after selected-year filter (WOS/PURE): ${info.selectedYearWos ?? 0} / ${info.selectedYearPure ?? 0}`,
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

  if (typeof info.doiMatchedCount === 'number') {
    lines.push(`DOI matches: ${info.doiMatchedCount}`);
  }

  if (typeof info.stableIdMatchedCount === 'number') {
    lines.push(`Stable-ID matches: ${info.stableIdMatchedCount}`);
  }

  if (typeof info.rescuedByFinalStrongSignal === 'number') {
    lines.push(`Final strong-signal rescues: ${info.rescuedByFinalStrongSignal}`);
  }

  if (typeof info.missingBlockedByStrongSignal === 'number') {
    lines.push(`Missing blocked by strong signal: ${info.missingBlockedByStrongSignal}`);
  }

  if (typeof info.doiConflictCount === 'number') {
    lines.push(`DOI conflicts detected: ${info.doiConflictCount}`);
  }

  if (typeof info.fuzzyHighConfidenceCount === 'number') {
    lines.push(`Fuzzy high-confidence matches: ${info.fuzzyHighConfidenceCount}`);
  }

  if (typeof info.fuzzyReviewCount === 'number') {
    lines.push(`Fuzzy review candidates: ${info.fuzzyReviewCount}`);
  }

  if (typeof info.metadataPromotedMatched === 'number') {
    lines.push(`Promoted to matched by metadata: ${info.metadataPromotedMatched}`);
  }

  if (typeof info.metadataDowngradedReview === 'number') {
    lines.push(`Downgraded to review by metadata conflict: ${info.metadataDowngradedReview}`);
  }

  if (typeof info.metadataRejectedHardConflict === 'number') {
    lines.push(`Rejected due to hard conflicts: ${info.metadataRejectedHardConflict}`);
  }

  if (typeof info.ambiguousCandidateCount === 'number') {
    lines.push(`Ambiguous candidate cases: ${info.ambiguousCandidateCount}`);
  }

  if (typeof info.candidatesSkippedByPrefilter === 'number') {
    lines.push(`Candidates skipped by prefilter: ${info.candidatesSkippedByPrefilter}`);
  }

  if (typeof info.forcedReviewByMargin === 'number') {
    lines.push(`Forced to review by score margin: ${info.forcedReviewByMargin}`);
  }

  if (typeof info.forcedReviewByClaimConflict === 'number') {
    lines.push(`Forced to review by claim conflict: ${info.forcedReviewByClaimConflict}`);
  }

  if (typeof info.exactTitleBlockedByMetadata === 'number') {
    lines.push(`Exact-title matches blocked by metadata: ${info.exactTitleBlockedByMetadata}`);
  }

  if (typeof info.secondPassCandidatesChecked === 'number') {
    lines.push(`Second-pass candidates checked: ${info.secondPassCandidatesChecked}`);
  }

  if (typeof info.rescuedToReviewBySecondPass === 'number') {
    lines.push(`Rescued to review by second pass: ${info.rescuedToReviewBySecondPass}`);
  }

  if (typeof info.missingWithHighMetadataAgreement === 'number') {
    lines.push(`Missing with high metadata agreement: ${info.missingWithHighMetadataAgreement}`);
  }

  if (typeof info.mappingQualityWarningTriggered === 'number') {
    lines.push(`Mapping quality warning triggered: ${info.mappingQualityWarningTriggered}`);
  }

  if (typeof info.topCandidateScoreForMissing === 'number') {
    lines.push(`Top candidate score for missing (avg): ${info.topCandidateScoreForMissing.toFixed(1)}%`);
  }

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
            const raw = entry.replace(/\s+/g, ' ').trim();
            const doi = extractBibField(entry, 'doi') || extractDoiCandidates(raw)[0] || '';
            const year = extractBibField(entry, 'year') || extractBibField(entry, 'date');
            const affiliations = extractBibField(entry, 'address') || extractBibField(entry, 'affiliation');
            const stableIdFromUrl = firstStableIdCandidate(
              extractBibField(entry, 'url'),
              extractBibField(entry, 'eid'),
              extractBibField(entry, 'scopusid'),
              extractBibField(entry, 'publicationid')
            );
            const ut = extractBibField(entry, 'ut') || extractBibField(entry, 'accessionnumber') || stableIdFromUrl;
            const docType = normalizeBibEntryType(extractBibEntryType(entry));

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
    const raw = entry.replace(/\s+/g, ' ').trim();
    const doi = extractBibField(entry, 'doi') || extractDoiCandidates(raw)[0] || '';
    const year = extractBibField(entry, 'year') || extractBibField(entry, 'date');
    const affiliations = extractBibField(entry, 'address') || extractBibField(entry, 'affiliation');
    const stableIdFromUrl = firstStableIdCandidate(
      extractBibField(entry, 'url'),
      extractBibField(entry, 'eid'),
      extractBibField(entry, 'scopusid'),
      extractBibField(entry, 'publicationid')
    );
    const ut = extractBibField(entry, 'ut') || extractBibField(entry, 'accessionnumber') || stableIdFromUrl;
    const recordType = normalizeBibEntryType(extractBibEntryType(entry));

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
    const stableId = firstStableIdCandidate(first('UR'), first('L1'), first('L2'), first('ID'));

    return {
      author: many('AU').concat(many('A1')).join('; '),
      title: first('TI') || first('T1'),
      journal: first('JO') || first('JF') || first('T2'),
      doi: first('DO'),
      year: first('PY') || first('Y1') || first('DA'),
      affiliations: first('AD') || first('C1'),
      ut: first('AN') || first('ID') || stableId,
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
        canonical.ut = canonical.ut ?? '';
      }

      return canonical;
    });

  return { headerIndex, headerCells: semanticHeaders, rows: finalRows };
}

function normalizeDoi(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .normalize('NFKC')
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, '')
    .replace(/[\u200E\u200F\u202A-\u202E\u2066-\u2069]/g, '')
    .replace(/[\u2028\u2029]/g, ' ')
    .replace(/\u00a0/g, ' ')
    .trim()
    .replace(/^(?:https?:\/\/)?(?:dx\.)?doi\.org\//i, '')
    .replace(/^doi:\s*/i, '')
    .split('#')[0]
    .split('?')[0]
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/\u00a0/g, '')
    .replace(/[.,;:\)\]\}]+$/g, '');
}

function isValidDoi(value) {
  const doi = normalizeDoi(value);
  return /^10\.\d{4,9}\/[\-._;()/:a-z0-9]+$/i.test(doi);
}

function extractDoiCandidates(text) {
  const source = String(text ?? '');
  if (!source.trim()) return [];

  const pattern = /(?:https?:\/\/(?:dx\.)?doi\.org\/|doi:\s*)?(10\.\d{4,9}\/[\-._;()/:a-z0-9]+)/gi;
  const matches = new Set();
  let match;

  while ((match = pattern.exec(source)) !== null) {
    const normalized = normalizeDoi(match[1]);
    if (isValidDoi(normalized)) {
      matches.add(normalized);
    }
  }

  return [...matches];
}

function firstStableIdCandidate(...values) {
  for (const value of values) {
    const candidates = extractStableIdCandidates(String(value || ''));
    if (candidates.length) {
      return candidates[0];
    }
  }
  return '';
}

function getRecordDoiCandidates(record) {
  const candidates = new Set();
  const directDoi = normalizeDoi(record?.doi);
  if (isValidDoi(directDoi)) {
    candidates.add(directDoi);
  }

  const freeText = [record?.doi, record?.title, record?.subtitle, record?.journal, record?.ut]
    .filter(Boolean)
    .join(' ');

  extractDoiCandidates(freeText).forEach((doi) => candidates.add(doi));
  return candidates;
}

function hasDoiOverlap(leftSet, rightSet) {
  if (!leftSet?.size || !rightSet?.size) return false;
  for (const doi of leftSet) {
    if (rightSet.has(doi)) return true;
  }
  return false;
}

function hasDoiConflict(leftSet, rightSet) {
  if (!leftSet?.size || !rightSet?.size) return false;
  return !hasDoiOverlap(leftSet, rightSet);
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

function getStrictTitleKey(value) {
  return canonicalizeText(value);
}

function getRelaxedTitleKey(value) {
  return canonicalizeText(value)
    .split(/\s+/)
    .filter((token) => token.length >= 3 && !TITLE_STOP_WORDS.has(token))
    .join(' ');
}

function getTitleTokens(value) {
  return new Set(
    getRelaxedTitleKey(value)
      .split(/\s+/)
      .map((token) => token.trim())
      .filter(Boolean)
  );
}

function getTokenBucketKey(value, limit = 4) {
  const tokens = [...getTitleTokens(value)].sort();
  return tokens.slice(0, limit).join('|');
}

function normalizeStableId(value) {
  return String(value ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function extractStableIdCandidates(text) {
  const source = String(text ?? '');
  if (!source.trim()) return [];

  const patterns = [
    /\bUT\s*[:=\-]?\s*([a-z0-9:-]+)/gi,
    /\bAN\s*[:=\-]?\s*([a-z0-9:-]+)/gi,
    /\bEID\s*[:=\-]?\s*([a-z0-9./:-]+)/gi,
    /\bSCOPUS(?:\s+EID)?\s*[:=\-]?\s*([a-z0-9./:-]+)/gi,
    /scopus\.com\/pages\/publications\/([0-9]{8,})/gi,
    /[?&]eid=([^&\s]+)/gi,
    /\bpublication(?:\s|_)?id\s*[:=\-]?\s*([a-z0-9._:-]+)/gi,
    /\bWOS\s*[:=\-]?\s*([a-z0-9./:-]+)/gi,
    /\bPMID\s*[:=\-]?\s*([a-z0-9:-]+)/gi,
    /\bISBN\s*[:=\-]?\s*([a-z0-9-]+)/gi,
    /\b2-s2\.0-\d+(?:\.[0-9]+)?\b/gi,
    /\bWOS:\s*[A-Z0-9]+\b/gi
  ];

  const matches = new Set();

  patterns.forEach((pattern) => {
    let match;
    while ((match = pattern.exec(source)) !== null) {
      const value = normalizeStableId(match[1] || match[0]);
      if (value && value.length >= 6) {
        matches.add(value);
      }
    }
  });

  return [...matches];
}

function getRecordStableIdCandidates(record) {
  const candidates = new Set();
  const directValues = [record?.ut, record?.accessionNumber, record?.accessionnumber, record?.eid, record?.scopusEid];
  directValues.filter(Boolean).forEach((value) => candidates.add(normalizeStableId(value)));

  const freeText = [record?.ut, record?.accessionNumber, record?.accessionnumber, record?.eid, record?.scopusEid, record?.title, record?.subtitle, record?.journal, record?.doi]
    .filter(Boolean)
    .join(' ');

  extractStableIdCandidates(freeText).forEach((id) => candidates.add(id));
  return candidates;
}

function isTitleMappingQualityPoor(rows) {
  const titles = rows
    .map((row) => String(row.title || '').trim())
    .filter(Boolean);

  const total = rows.length || 0;
  const emptyRatio = total ? (total - titles.length) / total : 1;
  const avgLength = titles.length ? titles.reduce((sum, title) => sum + title.length, 0) / titles.length : 0;

  return {
    poor: emptyRatio >= TITLE_MAPPING_EMPTY_RATIO_THRESHOLD || avgLength <= TITLE_MAPPING_AVG_LENGTH_THRESHOLD,
    emptyRatio,
    avgLength
  };
}

function isHighMetadataAgreement(parts) {
  return parts?.year === 'year exact' && parts?.recordType === 'type exact';
}

function shouldRescueLowScoreCandidate(candidate) {
  if (!candidate || candidate.hasHardConflict || candidate.hasDoiConflict) return false;

  const parts = candidate.reasonParts || {};
  const titleScore = Number(parts.titleScore ?? 0);
  const metadataAligned = isHighMetadataAgreement(parts);

  return metadataAligned && titleScore >= 10;
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

function tokenizeTitle(text) {
  const tokens = canonicalizeText(text)
    .split(/\s+/)
    .map((token) => token.trim())
    .filter((token) => token.length >= 3);
  return new Set(tokens);
}

function jaccardSimilarity(setA, setB) {
  if (!setA?.size && !setB?.size) return 1;
  if (!setA?.size || !setB?.size) return 0;

  let intersection = 0;
  const [small, large] = setA.size <= setB.size ? [setA, setB] : [setB, setA];
  for (const item of small) {
    if (large.has(item)) intersection += 1;
  }

  const union = setA.size + setB.size - intersection;
  return union ? intersection / union : 0;
}

function trigramSet(text) {
  const normalized = canonicalizeText(text).replace(/\s+/g, ' ').trim();
  const compact = normalized.replace(/\s/g, '');
  if (compact.length < 3) {
    return new Set(compact ? [compact] : []);
  }

  const set = new Set();
  for (let i = 0; i <= compact.length - 3; i += 1) {
    set.add(compact.slice(i, i + 3));
  }
  return set;
}

function weightedTitleScore(leftKey, rightKey) {
  if (!leftKey || !rightKey) return 0;

  const tokenScore = jaccardSimilarity(tokenizeTitle(leftKey), tokenizeTitle(rightKey)) * 100;
  const trigramScore = jaccardSimilarity(trigramSet(leftKey), trigramSet(rightKey)) * 100;
  const editScore = fuzzyTitleScoreNormalized(leftKey, rightKey);

  return (tokenScore * TITLE_SCORE_WEIGHTS.tokenJaccard)
    + (trigramScore * TITLE_SCORE_WEIGHTS.trigramJaccard)
    + (editScore * TITLE_SCORE_WEIGHTS.editSimilarity);
}

function clampScore(value) {
  if (!Number.isFinite(value)) return 0;
  return Math.max(0, Math.min(100, value));
}

function getFirstAuthorToken(value) {
  const firstAuthor = String(value ?? '')
    .split(';')[0]
    .split(',')[0]
    .trim();

  if (!firstAuthor) return '';

  const normalized = canonicalizeText(firstAuthor)
    .split(/\s+/)
    .filter((token) => token.length >= 2);

  return normalized[0] || '';
}

function scoreFirstAuthorOverlap(wosAuthor, pureAuthor) {
  const left = getFirstAuthorToken(wosAuthor);
  const right = getFirstAuthorToken(pureAuthor);

  if (!left || !right) {
    return { score: 50, label: 'author missing' };
  }

  if (left === right) {
    return { score: 100, label: 'first author match' };
  }

  return { score: 15, label: 'first author mismatch' };
}

function isHardTypeConflict(wosType, pureType) {
  if (!wosType || !pureType || wosType === pureType) return false;

  const archivalTypes = new Set(['thesis', 'book', 'book chapter']);
  const articleTypes = new Set(['journal article', 'conference paper', 'preprint']);

  const leftArchival = archivalTypes.has(wosType);
  const rightArchival = archivalTypes.has(pureType);
  const leftArticle = articleTypes.has(wosType);
  const rightArticle = articleTypes.has(pureType);

  return (leftArchival && rightArticle) || (rightArchival && leftArticle);
}

function getRecordTypeFamily(value) {
  const type = normalizeRecordType(value);
  if (!type) return '';

  if (type === 'journal article' || type === 'conference paper' || type === 'preprint') {
    return 'article_like';
  }

  if (type === 'thesis' || type === 'book' || type === 'book chapter') {
    return 'long_form';
  }

  return 'other';
}

function areTypeFamiliesCompatible(leftType, rightType) {
  const leftFamily = getRecordTypeFamily(leftType);
  const rightFamily = getRecordTypeFamily(rightType);

  if (!leftFamily || !rightFamily) return true;
  if (leftFamily === 'other' || rightFamily === 'other') return true;
  return leftFamily === rightFamily;
}

function passesFuzzyPrefilter(record, candidateRow) {
  const wosYear = parseYearValue(record.year);
  const pureYear = parseYearValue(candidateRow.year);
  const yearCompatible = !Number.isInteger(wosYear)
    || !Number.isInteger(pureYear)
    || Math.abs(wosYear - pureYear) <= FUZZY_PREFILTER_YEAR_DELTA;

  const typeCompatible = areTypeFamiliesCompatible(record.recordType, candidateRow.recordType);

  return {
    pass: yearCompatible && typeCompatible,
    yearCompatible,
    typeCompatible
  };
}

function scoreYearAgreement(wosYear, pureYear) {
  if (!Number.isInteger(wosYear) || !Number.isInteger(pureYear)) {
    return { score: 50, label: 'year missing', hardConflict: false };
  }

  const delta = Math.abs(wosYear - pureYear);
  if (delta === 0) {
    return { score: 100, label: 'year exact', hardConflict: false };
  }

  if (delta <= YEAR_TOLERANCE) {
    return { score: 72, label: `year near (Δ${delta})`, hardConflict: false };
  }

  return { score: 0, label: `year mismatch (${wosYear} vs ${pureYear})`, hardConflict: true };
}

function scoreTypeAgreement(wosType, pureType) {
  if (!wosType || !pureType) {
    return { score: 50, label: 'type missing', hardConflict: false };
  }

  if (wosType === pureType) {
    return { score: 100, label: 'type exact', hardConflict: false };
  }

  if (isHardTypeConflict(wosType, pureType)) {
    return { score: 0, label: `type mismatch (${wosType} vs ${pureType})`, hardConflict: true };
  }

  return { score: 32, label: `type weak (${wosType} vs ${pureType})`, hardConflict: false };
}

function scoreJournalAgreement(wosJournal, pureJournal) {
  const left = getTitleKey(wosJournal || '');
  const right = getTitleKey(pureJournal || '');

  if (!left || !right) {
    return { score: 50, label: 'journal missing' };
  }

  const score = weightedTitleScore(left, right);
  if (score >= 95) return { score, label: 'journal exact' };
  if (score >= 80) return { score, label: 'journal strong' };
  if (score >= 60) return { score, label: 'journal moderate' };
  return { score, label: 'journal weak' };
}

function computeMetadataConfidence({ record, candidateRow, titleScore, hasDoiConflict }) {
  const wosYear = parseYearValue(record.year);
  const pureYear = parseYearValue(candidateRow.year);
  const year = scoreYearAgreement(wosYear, pureYear);

  const wosType = normalizeRecordType(record.recordType);
  const pureType = normalizeRecordType(candidateRow.recordType);
  const recordType = scoreTypeAgreement(wosType, pureType);
  const journal = scoreJournalAgreement(record.journal, candidateRow.journal);
  const firstAuthor = scoreFirstAuthorOverlap(record.author, candidateRow.author);

  let confidence = (titleScore * METADATA_WEIGHTS.title)
    + (year.score * METADATA_WEIGHTS.year)
    + (recordType.score * METADATA_WEIGHTS.recordType)
    + (journal.score * METADATA_WEIGHTS.journal)
    + (firstAuthor.score * METADATA_WEIGHTS.firstAuthor);

  if (hasDoiConflict) {
    confidence -= METADATA_PENALTIES.doiConflict;
  }

  if (year.hardConflict) {
    confidence -= METADATA_PENALTIES.yearHardConflict;
  }

  if (recordType.hardConflict) {
    confidence -= METADATA_PENALTIES.typeHardConflict;
  }

  const hasHardConflict = year.hardConflict || recordType.hardConflict;
  const strongConflict = (year.hardConflict && recordType.hardConflict) || (hasDoiConflict && hasHardConflict);

  return {
    confidence: clampScore(confidence),
    hasHardConflict,
    strongConflict,
    hasDoiConflict,
    parts: {
      titleScore,
      year: year.label,
      recordType: recordType.label,
      journal: journal.label,
      firstAuthor: firstAuthor.label
    }
  };
}

function buildReasonFromMetadata(parts, confidence, suffix = '') {
  const base = `title ${parts.titleScore.toFixed(1)}%, ${parts.year}, ${parts.recordType}, ${parts.journal}, ${parts.firstAuthor}, confidence ${confidence.toFixed(1)}%`;
  return suffix ? `${base}, ${suffix}` : base;
}

function formatScoreMargin(bestScore, secondScore) {
  if (!Number.isFinite(bestScore) || !Number.isFinite(secondScore)) {
    return 'score margin unavailable';
  }

  const delta = bestScore - secondScore;
  return `top ${bestScore.toFixed(1)}%, second ${secondScore.toFixed(1)}%, delta ${delta.toFixed(1)}%`;
}

function computeMissingLikelihood(score, reason) {
  const numericScore = Number.isFinite(score) ? score : 0;

  if (/no DOI match, no title match/i.test(reason || '')) {
    return 98;
  }

  if (/no fuzzy candidates/i.test(reason || '')) {
    return 94;
  }

  if (/hard conflict/i.test(reason || '')) {
    return 90;
  }

  return clampScore(100 - numericScore);
}

function buildMissingAssessment(score, reason) {
  const likelihoodMissing = Math.round(computeMissingLikelihood(score, reason));
  const trustRecommendation = likelihoodMissing >= 80 ? 'Yes' : 'No';

  return { likelihoodMissing, trustRecommendation };
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

function getAvailableYears(wosRows, pureRows) {
  const years = new Set();

  [...wosRows, ...pureRows].forEach((row) => {
    const year = parseYearValue(row?.year);
    if (Number.isInteger(year)) {
      years.add(year);
    }
  });

  return [...years].sort((a, b) => b - a);
}

function refreshYearFilterOptions() {
  if (!elements.yearFilterSelect) return;

  const years = getAvailableYears(state.wosRows, state.pureRows);
  state.availableYears = years;

  if (state.selectedYearFilter !== 'all') {
    const selectedYear = Number.parseInt(state.selectedYearFilter, 10);
    if (!years.includes(selectedYear)) {
      state.selectedYearFilter = 'all';
    }
  }

  const optionsHtml = ['<option value="all">All years</option>']
    .concat(years.map((year) => `<option value="${year}">${year}</option>`))
    .join('');

  elements.yearFilterSelect.innerHTML = optionsHtml;
  elements.yearFilterSelect.value = state.selectedYearFilter;
  elements.yearFilterSelect.disabled = !years.length;
}

function getRowsForSelectedYearFilter() {
  if (state.selectedYearFilter === 'all') {
    return {
      wosRows: [...state.wosRows],
      pureRows: [...state.pureRows],
      label: 'All years'
    };
  }

  const selectedYear = Number.parseInt(state.selectedYearFilter, 10);
  if (!Number.isInteger(selectedYear)) {
    return {
      wosRows: [...state.wosRows],
      pureRows: [...state.pureRows],
      label: 'All years'
    };
  }

  return {
    wosRows: state.wosRows.filter((row) => parseYearValue(row.year) === selectedYear),
    pureRows: state.pureRows.filter((row) => parseYearValue(row.year) === selectedYear),
    label: String(selectedYear)
  };
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
      alignedWos = alignedWos.filter((row) => {
        const year = parseYearValue(row.year);
        return !Number.isInteger(year) || sharedYears.has(year);
      });
      alignedPure = alignedPure.filter((row) => {
        const year = parseYearValue(row.year);
        return !Number.isInteger(year) || sharedYears.has(year);
      });

      const sortedYears = [...sharedYears].sort((a, b) => a - b);
      notes.push(`Auto-year candidate narrowing applied: ${sortedYears[0]}-${sortedYears[sortedYears.length - 1]} (${sharedYears.size} shared year(s)). Comparison still runs across uploaded records.`);
    } else {
      notes.push('No overlapping publication years found between WOS and PURE; year-based candidate narrowing was skipped and comparison still runs across uploaded records.');
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
      return !type || sharedTypesResult.overlap.has(type);
    });

    alignedPure = alignedPure.filter((row) => {
      const type = normalizeRecordType(row.recordType);
      return !type || sharedTypesResult.overlap.has(type);
    });

    notes.push(`Auto-type candidate narrowing applied: ${sharedTypesResult.overlap.size} shared type(s).`);
  } else if (sharedTypesResult.comparable) {
    notes.push('No overlapping record types found between WOS and PURE; type-based candidate narrowing was skipped.');
  }

  meta.sharedTypes = sharedTypesResult.overlap?.size || 0;
  meta.afterTypeWos = alignedWos.length;
  meta.afterTypePure = alignedPure.length;

  return { wosRows: alignedWos, pureRows: alignedPure, notes, meta };
}

function getPureTitle(row) {
  return [row.title, row.subtitle].filter(Boolean).join(' ').trim();
}

function createTitleBucketKey(titleKey, yearValue, prefixLength = 18) {
  const prefix = String(titleKey || '').slice(0, prefixLength);
  const lengthBucket = Math.floor(String(titleKey || '').length / 10);
  const yearPart = Number.isInteger(yearValue) ? String(yearValue) : 'na';
  return `${yearPart}|${prefix}|${lengthBucket}`;
}

function buildPureLookup(pureRows) {
  const byDoi = new Map();
  const byExactTitle = new Map();
  const byRelaxedTitle = new Map();
  const byBucket = new Map();
  const byShortBucket = new Map();
  const byTokenBucket = new Map();
  const byStableId = new Map();
  const allTitleRecords = [];
  const byRowDoiCandidates = new Map();

  pureRows.forEach((row) => {
    const doiCandidates = getRecordDoiCandidates(row);
    const title = getPureTitle(row);
    const strictTitleKey = getStrictTitleKey(title);
    const relaxedTitleKey = getRelaxedTitleKey(title);
    const yearValue = parseYearValue(row.year);
    const shortPrefixKey = createTitleBucketKey(relaxedTitleKey, yearValue, 12);
    const tokenBucketKey = getTokenBucketKey(title);
    const stableIdCandidates = getRecordStableIdCandidates(row);

    byRowDoiCandidates.set(row, doiCandidates);

    stableIdCandidates.forEach((id) => {
      if (!byStableId.has(id)) {
        byStableId.set(id, row);
      }
    });

    doiCandidates.forEach((doi) => {
      if (!byDoi.has(doi)) {
        byDoi.set(doi, row);
      }
    });

    if (strictTitleKey && !byExactTitle.has(strictTitleKey)) {
      byExactTitle.set(strictTitleKey, row);
    }

    if (relaxedTitleKey && !byRelaxedTitle.has(relaxedTitleKey)) {
      byRelaxedTitle.set(relaxedTitleKey, row);
    }

    if (shortPrefixKey) {
      if (!byShortBucket.has(shortPrefixKey)) byShortBucket.set(shortPrefixKey, []);
      byShortBucket.get(shortPrefixKey).push({ row, title, strictTitleKey, relaxedTitleKey, yearValue, doiCandidates, stableIdCandidates, tokenBucketKey, shortPrefixKey });
    }

    if (tokenBucketKey) {
      if (!byTokenBucket.has(tokenBucketKey)) byTokenBucket.set(tokenBucketKey, []);
      byTokenBucket.get(tokenBucketKey).push({ row, title, strictTitleKey, relaxedTitleKey, yearValue, doiCandidates, stableIdCandidates, tokenBucketKey, shortPrefixKey });
    }

    if (!relaxedTitleKey) {
      return;
    }

    const titleRecord = { row, title, strictTitleKey, relaxedTitleKey, yearValue, doiCandidates, stableIdCandidates, tokenBucketKey, shortPrefixKey };
    allTitleRecords.push(titleRecord);

    const bucketKeyExactYear = createTitleBucketKey(relaxedTitleKey, yearValue);
    const bucketKeyAnyYear = createTitleBucketKey(relaxedTitleKey, null);

    if (!byBucket.has(bucketKeyExactYear)) byBucket.set(bucketKeyExactYear, []);
    if (!byBucket.has(bucketKeyAnyYear)) byBucket.set(bucketKeyAnyYear, []);

    byBucket.get(bucketKeyExactYear).push(titleRecord);
    byBucket.get(bucketKeyAnyYear).push(titleRecord);
  });

  return { byDoi, byExactTitle, byRelaxedTitle, byBucket, byShortBucket, byTokenBucket, byStableId, allTitleRecords, byRowDoiCandidates };
}

function getFuzzyCandidates(lookup, wosTitleKey, wosYear, options = {}) {
  if (!wosTitleKey) return [];

  const isSecondPass = !!options.secondPass;
  const maxCandidates = isSecondPass ? SECOND_PASS_FUZZY_CANDIDATES : MAX_FUZZY_CANDIDATES;
  const shortPrefixLength = isSecondPass ? 12 : 18;

  const primaryKey = createTitleBucketKey(wosTitleKey, wosYear);
  const fallbackKey = createTitleBucketKey(wosTitleKey, null);
  const primary = lookup.byBucket.get(primaryKey) || [];
  const fallback = lookup.byBucket.get(fallbackKey) || [];
  const shortPrefixKey = createTitleBucketKey(wosTitleKey, wosYear, shortPrefixLength);
  const shortPrefixFallbackKey = createTitleBucketKey(wosTitleKey, null, shortPrefixLength);
  const shortPrefixCandidates = isSecondPass
    ? [
      ...(lookup.byShortBucket.get(shortPrefixKey) || []),
      ...(lookup.byShortBucket.get(shortPrefixFallbackKey) || [])
    ]
    : [];
  const tokenBucketKey = getTokenBucketKey(wosTitleKey);
  const tokenCandidates = isSecondPass ? (lookup.byTokenBucket.get(tokenBucketKey) || []) : [];
  const relaxedKey = getRelaxedTitleKey(wosTitleKey);
  const relaxedCandidates = isSecondPass ? (lookup.byRelaxedTitle.get(relaxedKey) ? [{ row: lookup.byRelaxedTitle.get(relaxedKey), relaxedTitleKey: relaxedKey }] : []) : [];

  const merged = [...primary, ...fallback, ...shortPrefixCandidates, ...tokenCandidates, ...relaxedCandidates];
  if (!merged.length) {
    return lookup.allTitleRecords.slice(0, Math.min(maxCandidates, lookup.allTitleRecords.length));
  }

  const dedup = [];
  const seen = new Set();
  merged.forEach((candidate) => {
    const key = candidate?.row || candidate;
    if (seen.has(key)) return;
    seen.add(key);
    dedup.push(candidate);
  });

  if (dedup.length <= maxCandidates) {
    return dedup;
  }

  return dedup
    .sort((a, b) => {
      const aKey = a.relaxedTitleKey || a.strictTitleKey || getRelaxedTitleKey(a.title || '');
      const bKey = b.relaxedTitleKey || b.strictTitleKey || getRelaxedTitleKey(b.title || '');
      return Math.abs(aKey.length - wosTitleKey.length) - Math.abs(bKey.length - wosTitleKey.length);
    })
    .slice(0, maxCandidates);
}

function evaluateCandidateMatch(record, candidateData, wosDoiCandidates, options = {}) {
  const row = candidateData.row || candidateData;
  const titleKey = candidateData.relaxedTitleKey || candidateData.strictTitleKey || getRelaxedTitleKey(candidateData.title || row.title || '');
  if (!titleKey) return null;

  const doiCandidates = candidateData.doiCandidates || getRecordDoiCandidates(row);
  const stableIdCandidates = candidateData.stableIdCandidates || getRecordStableIdCandidates(row);
  const hasConflict = hasDoiConflict(wosDoiCandidates, doiCandidates);
  const hasStableIdOverlap = hasDoiOverlap(new Set(Array.from(getRecordStableIdCandidates(record))), new Set(stableIdCandidates));
  const titleScore = weightedTitleScore(getTitleKey(record.title || ''), titleKey);
  const metadata = computeMetadataConfidence({
    record,
    candidateRow: row,
    titleScore,
    hasDoiConflict: hasConflict
  });

  return {
    row,
    titleScore,
    finalScore: metadata.confidence,
    hasDoiConflict: metadata.hasDoiConflict,
    hasHardConflict: metadata.hasHardConflict,
    strongConflict: metadata.strongConflict,
    reasonParts: metadata.parts,
    hasStableIdOverlap,
    secondPassEligible: !!options.secondPassEligible
  };
}

function getSecondPassCandidates(lookup, wosTitleKey, wosYear) {
  const candidates = getFuzzyCandidates(lookup, wosTitleKey, wosYear, { secondPass: true });
  return candidates.slice(0, SECOND_PASS_FUZZY_CANDIDATES);
}

async function buildComparisonData(runId, inputWosRows = state.wosRows, inputPureRows = state.pureRows, yearFilterMeta = {}) {
  const wosRecords = deduplicateRecords(inputWosRows, ['doi', 'ut', 'title']);
  const pureRecords = deduplicateRecords(inputPureRows, ['doi', 'ut', 'title']);
  const aligned = autoAlignDatasets(wosRecords, pureRecords);
  const alignedPure = aligned.pureRows;

  if (aligned.notes.length) {
    showYearWarning(aligned.notes.join(' '));
  } else {
    hideYearWarning();
  }

  const filteredWos = filterWosAffiliations(wosRecords);
  const fullLookup = buildPureLookup(pureRecords);
  const alignedLookup = buildPureLookup(alignedPure);
  const titleMappingQuality = isTitleMappingQualityPoor(alignedPure);
  let doiMatchedCount = 0;
  let stableIdMatchedCount = 0;
  let rescuedByFinalStrongSignal = 0;
  let missingBlockedByStrongSignal = 0;
  let doiConflictCount = 0;
  let fuzzyHighConfidenceCount = 0;
  let fuzzyReviewCount = 0;
  let metadataPromotedMatched = 0;
  let metadataDowngradedReview = 0;
  let metadataRejectedHardConflict = 0;
  let ambiguousCandidateCount = 0;
  let candidatesSkippedByPrefilter = 0;
  let forcedReviewByMargin = 0;
  let forcedReviewByClaimConflict = 0;
  let exactTitleBlockedByMetadata = 0;
  let secondPassCandidatesChecked = 0;
  let rescuedToReviewBySecondPass = 0;
  let missingWithHighMetadataAgreement = 0;
  let mappingQualityWarningTriggered = 0;
  let topCandidateScoreForMissingTotal = 0;
  let topCandidateScoreForMissingCount = 0;
  const pureClaims = new Map();

  if (titleMappingQuality.poor) {
    mappingQualityWarningTriggered = 1;
    showYearWarning(`Mapping quality warning: ${Math.round(titleMappingQuality.emptyRatio * 100)}% of PURE titles are empty or the average title length is only ${titleMappingQuality.avgLength.toFixed(1)} characters.`);
  }

  function finalizeMissing(record, reasonText, score, metadataAgreement = false, wosDoiCandidates = new Set(), wosStableIdCandidates = new Set()) {
    let recordJson = '';
    try {
      recordJson = JSON.stringify(record) || '';
    } catch (error) {
      recordJson = '';
    }

    const emergencyText = [record?.doi, record?.title, record?.journal, record?.ut, recordJson]
      .filter(Boolean)
      .join(' ');

    const emergencyDoiCandidates = new Set(wosDoiCandidates);
    extractDoiCandidates(emergencyText).forEach((doi) => emergencyDoiCandidates.add(doi));

    for (const doi of emergencyDoiCandidates) {
      const emergencyDoiMatch = fullLookup.byDoi.get(doi);
      if (!emergencyDoiMatch) continue;

      rescuedByFinalStrongSignal += 1;
      if (pureClaims.has(emergencyDoiMatch)) {
        forcedReviewByClaimConflict += 1;
        const claim = pureClaims.get(emergencyDoiMatch);
        results.push({
          ...record,
          status: 'needs_review',
          reason: `final strong-signal rescue by DOI (${doi}) found claimed PURE record (${claim.score.toFixed(1)}%), sent to review`,
          matchStatus: 'Needs review',
          pureMatch: emergencyDoiMatch,
          score: claim.score
        });
        return;
      }

      pureClaims.set(emergencyDoiMatch, { score: 100, via: 'final_doi_rescue' });
      doiMatchedCount += 1;
      results.push({
        ...record,
        status: 'matched',
        reason: `final strong-signal rescue by DOI (${doi})`,
        matchStatus: 'Matched',
        pureMatch: emergencyDoiMatch,
        score: 100
      });
      return;
    }

    const emergencyStableIdCandidates = new Set(wosStableIdCandidates);
    extractStableIdCandidates(emergencyText).forEach((id) => emergencyStableIdCandidates.add(id));

    for (const stableId of emergencyStableIdCandidates) {
      const emergencyStableIdMatch = fullLookup.byStableId.get(stableId);
      if (!emergencyStableIdMatch) continue;

      rescuedByFinalStrongSignal += 1;
      const stableDoiConflict = hasDoiConflict(emergencyDoiCandidates, fullLookup.byRowDoiCandidates.get(emergencyStableIdMatch) || new Set());

      if (pureClaims.has(emergencyStableIdMatch)) {
        forcedReviewByClaimConflict += 1;
        const claim = pureClaims.get(emergencyStableIdMatch);
        results.push({
          ...record,
          status: 'needs_review',
          reason: `final strong-signal rescue by stable identifier (${stableId}) found claimed PURE record (${claim.score.toFixed(1)}%), sent to review`,
          matchStatus: 'Needs review',
          pureMatch: emergencyStableIdMatch,
          score: claim.score
        });
        return;
      }

      if (stableDoiConflict) {
        doiConflictCount += 1;
        results.push({
          ...record,
          status: 'needs_review',
          reason: `final strong-signal rescue by stable identifier (${stableId}) with DOI conflict, sent to review`,
          matchStatus: 'Needs review',
          pureMatch: emergencyStableIdMatch,
          score: 99
        });
        return;
      }

      pureClaims.set(emergencyStableIdMatch, { score: 100, via: 'final_stable_id_rescue' });
      stableIdMatchedCount += 1;
      results.push({
        ...record,
        status: 'matched',
        reason: `final strong-signal rescue by stable identifier (${stableId})`,
        matchStatus: 'Matched',
        pureMatch: emergencyStableIdMatch,
        score: 100
      });
      return;
    }

    const missingAssessment = buildMissingAssessment(score, reasonText);
    if (metadataAgreement) {
      missingWithHighMetadataAgreement += 1;
      missingAssessment.trustRecommendation = 'No';
    }
    topCandidateScoreForMissingTotal += Number.isFinite(score) ? score : 0;
    topCandidateScoreForMissingCount += 1;
    results.push({
      ...record,
      status: 'missing',
      reason: reasonText,
      matchStatus: 'Missing',
      pureMatch: null,
      score,
      ...missingAssessment
    });
  }

  const results = [];
  for (let index = 0; index < filteredWos.length; index += 1) {
    if (runId !== state.compareRunId) {
      return;
    }

    const record = filteredWos[index];
    const wosDoiCandidates = getRecordDoiCandidates(record);
    const wosStableIdCandidates = getRecordStableIdCandidates(record);
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

    let pureMatchByDoi = null;
    for (const doi of wosDoiCandidates) {
      const candidateMatch = fullLookup.byDoi.get(doi);
      if (candidateMatch) {
        pureMatchByDoi = candidateMatch;
        break;
      }
    }

    if (pureMatchByDoi) {
      if (pureClaims.has(pureMatchByDoi)) {
        forcedReviewByClaimConflict += 1;
        const claim = pureClaims.get(pureMatchByDoi);
        results.push({
          ...record,
          status: 'needs_review',
          reason: `DOI matched but PURE record already claimed (${claim.score.toFixed(1)}%), sent to review`,
          matchStatus: 'Needs review',
          pureMatch: pureMatchByDoi,
          score: claim.score
        });
        continue;
      }

      pureClaims.set(pureMatchByDoi, { score: 100, via: 'doi' });
      doiMatchedCount += 1;
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

    let pureMatchByStableId = null;
    for (const stableId of wosStableIdCandidates) {
      const candidateMatch = fullLookup.byStableId.get(stableId);
      if (candidateMatch) {
        pureMatchByStableId = candidateMatch;
        break;
      }
    }

    if (pureMatchByStableId) {
      const pureStableIds = getRecordStableIdCandidates(pureMatchByStableId);
      const stableDoiConflict = hasDoiConflict(wosDoiCandidates, fullLookup.byRowDoiCandidates.get(pureMatchByStableId) || new Set());

      if (pureClaims.has(pureMatchByStableId)) {
        forcedReviewByClaimConflict += 1;
        const claim = pureClaims.get(pureMatchByStableId);
        results.push({
          ...record,
          status: 'needs_review',
          reason: `stable identifier matched but PURE record already claimed (${claim.score.toFixed(1)}%), sent to review`,
          matchStatus: 'Needs review',
          pureMatch: pureMatchByStableId,
          score: claim.score
        });
        continue;
      }

      if (stableDoiConflict) {
        doiConflictCount += 1;
        results.push({
          ...record,
          status: 'needs_review',
          reason: 'stable identifier overlap found but DOI conflict detected, sent to review',
          matchStatus: 'Needs review',
          pureMatch: pureMatchByStableId,
          score: 99
        });
        continue;
      }

      if (!hasDoiOverlap(wosStableIdCandidates, pureStableIds)) {
        results.push({
          ...record,
          status: 'needs_review',
          reason: 'stable identifier candidate found, sent to review',
          matchStatus: 'Needs review',
          pureMatch: pureMatchByStableId,
          score: 99
        });
        continue;
      }

      pureClaims.set(pureMatchByStableId, { score: 100, via: 'stable_id' });
      stableIdMatchedCount += 1;
      results.push({
        ...record,
        status: 'matched',
        reason: 'stable identifier matched',
        matchStatus: 'Matched',
        pureMatch: pureMatchByStableId
      });
      continue;
    }

    const exactTitleMatch = normalizedWosTitle ? fullLookup.byExactTitle.get(normalizedWosTitle) : null;
    if (exactTitleMatch) {
      const exactMatchDoiCandidates = fullLookup.byRowDoiCandidates.get(exactTitleMatch) || new Set();
      const exactMatchDoiConflict = hasDoiConflict(wosDoiCandidates, exactMatchDoiCandidates);
      const exactYear = scoreYearAgreement(parseYearValue(record.year), parseYearValue(exactTitleMatch.year));
      const exactType = scoreTypeAgreement(normalizeRecordType(record.recordType), normalizeRecordType(exactTitleMatch.recordType));
      const exactStrongMetadataConflict = exactYear.hardConflict || exactType.hardConflict;

      if (exactMatchDoiConflict) {
        doiConflictCount += 1;
      }

      if (exactStrongMetadataConflict) {
        exactTitleBlockedByMetadata += 1;
        results.push({
          ...record,
          status: 'needs_review',
          reason: `exact title matched but metadata conflict (${exactYear.label}; ${exactType.label}), sent to review`,
          matchStatus: 'Needs review',
          pureMatch: exactTitleMatch
        });
        continue;
      }

      if (!exactMatchDoiConflict) {
        if (pureClaims.has(exactTitleMatch)) {
          forcedReviewByClaimConflict += 1;
          const claim = pureClaims.get(exactTitleMatch);
          results.push({
            ...record,
            status: 'needs_review',
            reason: `exact title matched but PURE record already claimed (${claim.score.toFixed(1)}%), sent to review`,
            matchStatus: 'Needs review',
            pureMatch: exactTitleMatch,
            score: claim.score
          });
          continue;
        }

        pureClaims.set(exactTitleMatch, { score: 99, via: 'exact_title' });
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

      reason = 'exact title matched but DOI conflict found';
    }

    const relaxedTitleKey = getRelaxedTitleKey(wosTitle);
    const relaxedTitleMatch = relaxedTitleKey ? fullLookup.byRelaxedTitle.get(relaxedTitleKey) : null;
    if (relaxedTitleMatch) {
      const relaxedMatchDoiCandidates = fullLookup.byRowDoiCandidates.get(relaxedTitleMatch) || new Set();
      const relaxedMatchDoiConflict = hasDoiConflict(wosDoiCandidates, relaxedMatchDoiCandidates);
      const relaxedYear = scoreYearAgreement(parseYearValue(record.year), parseYearValue(relaxedTitleMatch.year));
      const relaxedType = scoreTypeAgreement(normalizeRecordType(record.recordType), normalizeRecordType(relaxedTitleMatch.recordType));
      const relaxedMetadataConsistent = !relaxedMatchDoiConflict && isHighMetadataAgreement({ year: relaxedYear.label, recordType: relaxedType.label });

      if (relaxedMatchDoiConflict) {
        doiConflictCount += 1;
      }

      if (relaxedMetadataConsistent) {
        if (pureClaims.has(relaxedTitleMatch)) {
          forcedReviewByClaimConflict += 1;
          const claim = pureClaims.get(relaxedTitleMatch);
          results.push({
            ...record,
            status: 'needs_review',
            reason: `normalized title matched and metadata aligned, but PURE record already claimed (${claim.score.toFixed(1)}%), sent to review`,
            matchStatus: 'Needs review',
            pureMatch: relaxedTitleMatch,
            score: claim.score
          });
          continue;
        }

        pureClaims.set(relaxedTitleMatch, { score: 98, via: 'relaxed_title' });
        results.push({
          ...record,
          status: 'needs_review',
          reason: `normalized title matched with aligned metadata (${relaxedYear.label}; ${relaxedType.label}), sent to review`,
          matchStatus: 'Needs review',
          pureMatch: relaxedTitleMatch,
          score: 98
        });
        continue;
      }

    }

    const evaluatedCandidates = [];

    const candidates = getFuzzyCandidates(alignedLookup, normalizedWosTitle, wosYearValue);
    candidates.forEach((candidateData) => {
      const row = candidateData.row || candidateData;
      if (!row) return;

      const prefilter = passesFuzzyPrefilter(record, row);
      if (!prefilter.pass) {
        candidatesSkippedByPrefilter += 1;
        return;
      }

      const evaluated = evaluateCandidateMatch(record, candidateData, wosDoiCandidates);
      if (evaluated) {
        evaluatedCandidates.push(evaluated);
      }
    });

    evaluatedCandidates.sort((a, b) => b.finalScore - a.finalScore);
    const bestMatch = evaluatedCandidates[0] || null;
    const secondBest = evaluatedCandidates[1] || null;
    const scoreDelta = bestMatch && secondBest ? bestMatch.finalScore - secondBest.finalScore : Number.POSITIVE_INFINITY;
    const isAmbiguous = !!(
      bestMatch
      && secondBest
      && bestMatch.finalScore >= REVIEW_THRESHOLD
      && secondBest.finalScore >= REVIEW_THRESHOLD
      && Math.abs(bestMatch.finalScore - secondBest.finalScore) <= AMBIGUITY_SCORE_GAP
    );
    const isLowMargin = !!(
      bestMatch
      && secondBest
      && bestMatch.finalScore >= REVIEW_THRESHOLD
      && scoreDelta < SCORE_MARGIN_THRESHOLD
    );

    if (bestMatch && bestMatch.hasDoiConflict) {
      doiConflictCount += 1;
    }

    if (bestMatch && bestMatch.strongConflict) {
      metadataRejectedHardConflict += 1;
      const reasonText = buildReasonFromMetadata(bestMatch.reasonParts, bestMatch.finalScore, 'hard conflict, rejected');
      finalizeMissing(record, reasonText, bestMatch.finalScore, isHighMetadataAgreement(bestMatch.reasonParts), wosDoiCandidates, wosStableIdCandidates);
    } else if (bestMatch && isLowMargin) {
      forcedReviewByMargin += 1;
      fuzzyReviewCount += 1;
      results.push({
        ...record,
        status: 'needs_review',
        reason: `${buildReasonFromMetadata(bestMatch.reasonParts, bestMatch.finalScore, 'low score margin, sent to review')} (${formatScoreMargin(bestMatch.finalScore, secondBest.finalScore)})`,
        matchStatus: 'Needs review',
        pureMatch: bestMatch.row,
        score: bestMatch.finalScore
      });
    } else if (bestMatch && isAmbiguous) {
      ambiguousCandidateCount += 1;
      fuzzyReviewCount += 1;
      results.push({
        ...record,
        status: 'needs_review',
        reason: buildReasonFromMetadata(bestMatch.reasonParts, bestMatch.finalScore, `ambiguous with close candidate (${secondBest.finalScore.toFixed(1)}%)`),
        matchStatus: 'Needs review',
        pureMatch: bestMatch.row,
        score: bestMatch.finalScore
      });
    } else if (bestMatch && bestMatch.finalScore >= MATCH_THRESHOLD && !bestMatch.hasHardConflict && !bestMatch.hasDoiConflict) {
      if (pureClaims.has(bestMatch.row)) {
        forcedReviewByClaimConflict += 1;
        const claim = pureClaims.get(bestMatch.row);
        results.push({
          ...record,
          status: 'needs_review',
          reason: `${buildReasonFromMetadata(bestMatch.reasonParts, bestMatch.finalScore, 'PURE record already claimed, sent to review')} (${formatScoreMargin(bestMatch.finalScore, secondBest?.finalScore ?? 0)}; claimed ${claim.score.toFixed(1)}%)`,
          matchStatus: 'Needs review',
          pureMatch: bestMatch.row,
          score: bestMatch.finalScore
        });
        continue;
      }

      pureClaims.set(bestMatch.row, { score: bestMatch.finalScore, via: 'fuzzy' });
      fuzzyHighConfidenceCount += 1;
      metadataPromotedMatched += 1;
      status = 'matched';
      reason = buildReasonFromMetadata(bestMatch.reasonParts, bestMatch.finalScore, 'auto-matched');
      results.push({
        ...record,
        status,
        reason,
        matchStatus: 'Matched',
        pureMatch: bestMatch.row,
        score: bestMatch.finalScore
      });
    } else if (bestMatch && bestMatch.finalScore >= REVIEW_THRESHOLD) {
      fuzzyReviewCount += 1;
      if (bestMatch.hasHardConflict || bestMatch.hasDoiConflict) {
        metadataDowngradedReview += 1;
      }

      status = 'needs_review';
      reason = buildReasonFromMetadata(bestMatch.reasonParts, bestMatch.finalScore, bestMatch.hasDoiConflict ? 'DOI conflict, sent to review' : 'sent to review');
      results.push({
        ...record,
        status,
        reason,
        matchStatus: 'Needs review',
        pureMatch: bestMatch.row,
        score: bestMatch.finalScore
      });
    } else if (bestMatch && bestMatch.hasHardConflict) {
      metadataRejectedHardConflict += 1;
      const reasonText = buildReasonFromMetadata(bestMatch.reasonParts, bestMatch.finalScore, 'hard conflict, below threshold');
      finalizeMissing(record, reasonText, bestMatch.finalScore, isHighMetadataAgreement(bestMatch.reasonParts), wosDoiCandidates, wosStableIdCandidates);
    } else {
      let reasonText = bestMatch
        ? buildReasonFromMetadata(bestMatch.reasonParts, bestMatch.finalScore, 'below review threshold')
        : (candidates.length ? reason : 'no fuzzy candidates available');
      let finalScore = bestMatch ? bestMatch.finalScore : 0;

      if (!bestMatch || finalScore < REVIEW_THRESHOLD) {
        const secondPassCandidates = getSecondPassCandidates(alignedLookup, normalizedWosTitle, wosYearValue);
        secondPassCandidatesChecked += secondPassCandidates.length;

        const secondPassEvaluated = [];
        secondPassCandidates.forEach((candidateData) => {
          const row = candidateData.row || candidateData;
          if (!row) return;

          const evaluated = evaluateCandidateMatch(record, candidateData, wosDoiCandidates, { secondPassEligible: true });
          if (!evaluated) return;

          const relaxedPrefilter = passesFuzzyPrefilter(record, row);
          if (!relaxedPrefilter.pass && !evaluated.hasStableIdOverlap) {
            return;
          }

          secondPassEvaluated.push(evaluated);
        });

        secondPassEvaluated.sort((a, b) => b.finalScore - a.finalScore);
        const rescueCandidate = secondPassEvaluated[0] || null;

        if (rescueCandidate && rescueCandidate.finalScore >= REVIEW_THRESHOLD) {
          rescuedToReviewBySecondPass += 1;
          fuzzyReviewCount += 1;
          results.push({
            ...record,
            status: 'needs_review',
            reason: buildReasonFromMetadata(rescueCandidate.reasonParts, rescueCandidate.finalScore, 'second-pass rescue, sent to review'),
            matchStatus: 'Needs review',
            pureMatch: rescueCandidate.row,
            score: rescueCandidate.finalScore
          });
          continue;
        }

        if (rescueCandidate && shouldRescueLowScoreCandidate(rescueCandidate)) {
          rescuedToReviewBySecondPass += 1;
          fuzzyReviewCount += 1;
          results.push({
            ...record,
            status: 'needs_review',
            reason: buildReasonFromMetadata(rescueCandidate.reasonParts, rescueCandidate.finalScore, 'metadata-aligned second-pass rescue, sent to review'),
            matchStatus: 'Needs review',
            pureMatch: rescueCandidate.row,
            score: rescueCandidate.finalScore
          });
          continue;
        }

        if (rescueCandidate) {
          reasonText = buildReasonFromMetadata(rescueCandidate.reasonParts, rescueCandidate.finalScore, 'second-pass candidate found, still below threshold');
          finalScore = rescueCandidate.finalScore;
        }
      }

      if (bestMatch && shouldRescueLowScoreCandidate(bestMatch)) {
        fuzzyReviewCount += 1;
        results.push({
          ...record,
          status: 'needs_review',
          reason: buildReasonFromMetadata(bestMatch.reasonParts, bestMatch.finalScore, 'metadata-aligned low-score rescue, sent to review'),
          matchStatus: 'Needs review',
          pureMatch: bestMatch.row,
          score: bestMatch.finalScore
        });
        continue;
      }

      finalizeMissing(record, reasonText, finalScore, !!bestMatch && isHighMetadataAgreement(bestMatch.reasonParts), wosDoiCandidates, wosStableIdCandidates);
    }

    if ((index + 1) % CHUNK_SIZE === 0) {
      showYearWarning(`Comparing records... ${index + 1}/${filteredWos.length}`);
      await nextUiTick();
    }
  }

  if (runId !== state.compareRunId) {
    return;
  }

  const totalWos = wosRecords.length;
  const filtered = filteredWos.filter((record) => record.affiliationStatus === 'kept').length;
  const affiliationExcluded = filteredWos.filter((record) => record.affiliationStatus === 'excluded').length;
  const matched = results.filter((record) => record.status === 'matched' && record.affiliationStatus !== 'excluded').length;
  const missing = results.filter((record) => record.status === 'missing' && record.affiliationStatus !== 'excluded').length;
  const review = results.filter((record) => record.status === 'needs_review' && record.affiliationStatus !== 'excluded').length;
  const excluded = results.filter((record) => record.status === 'excluded').length;

  renderDebugInfo({
    parsedWos: state.wosRows.length,
    parsedPure: state.pureRows.length,
    activeYearFilter: yearFilterMeta.activeYearFilter || 'All years',
    availableYearsCount: state.availableYears.length,
    selectedYearWos: Number.isInteger(yearFilterMeta.selectedYearWos) ? yearFilterMeta.selectedYearWos : state.wosRows.length,
    selectedYearPure: Number.isInteger(yearFilterMeta.selectedYearPure) ? yearFilterMeta.selectedYearPure : state.pureRows.length,
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
    review,
    doiMatchedCount,
    stableIdMatchedCount,
    rescuedByFinalStrongSignal,
    missingBlockedByStrongSignal,
    doiConflictCount,
    fuzzyHighConfidenceCount,
    fuzzyReviewCount,
    metadataPromotedMatched,
    metadataDowngradedReview,
    metadataRejectedHardConflict,
    ambiguousCandidateCount,
    candidatesSkippedByPrefilter,
    forcedReviewByMargin,
    forcedReviewByClaimConflict,
    exactTitleBlockedByMetadata
    ,secondPassCandidatesChecked,
    rescuedToReviewBySecondPass,
    missingWithHighMetadataAgreement,
    mappingQualityWarningTriggered,
    topCandidateScoreForMissing: topCandidateScoreForMissingCount ? topCandidateScoreForMissingTotal / topCandidateScoreForMissingCount : 0
  });

  // Results table should only show records present in WoS but missing in PURE/Scopus source set.
  state.matchResults = results.filter((record) => record.status === 'missing');
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
  const activeStatus = 'missing';
  const search = state.searchText.trim().toLowerCase();

  const missingCount = state.matchResults.filter((record) => record.status === 'missing').length;
  elements.exportBtn.classList.toggle('hidden', missingCount === 0);

  const filteredRows = state.matchResults.filter((record) => {
    const statusMatch = record.status === activeStatus;
    const haystack = `${record.author || ''} ${record.title || ''} ${record.journal || ''} ${record.year || ''} ${record.doi || ''} ${record.ut || ''} ${record.matchStatus || ''} ${record.reason || ''} ${record.likelihoodMissing ?? ''} ${record.trustRecommendation || ''}`.toLowerCase();
    const matchesSearch = !search || haystack.includes(search);
    return statusMatch && matchesSearch;
  });

  const sortedRows = [...filteredRows].sort((a, b) => {
    if (state.sortKey === 'likelihoodMissing') {
      const left = Number(a.likelihoodMissing ?? 0);
      const right = Number(b.likelihoodMissing ?? 0);
      return state.sortDir === 'asc' ? left - right : right - left;
    }

    const left = String(a[state.sortKey] ?? '').toLowerCase();
    const right = String(b[state.sortKey] ?? '').toLowerCase();
    if (left === right) return 0;
    return state.sortDir === 'asc' ? left.localeCompare(right) : right.localeCompare(left);
  });

  if (!sortedRows.length) {
    elements.resultsTableBody.innerHTML = '<tr><td colspan="10" class="empty-state">No records match the active filter.</td></tr>';
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
          <td>${escapeHtml(String(row.likelihoodMissing ?? 0))}%</td>
          <td>${escapeHtml(row.reason || '—')}</td>
          <td>${escapeHtml(row.trustRecommendation || 'No')}</td>
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
    recordType: row.recordType || '',
    ut: row.ut || ''
  }));

  refreshYearFilterOptions();

  compareAndRender();
}

async function compareAndRender() {
  refreshYearFilterOptions();

  if (!state.wosRows.length || !state.pureRows.length) {
    const wosCount = state.wosRows.length;
    const pureCount = state.pureRows.length;
    showYearWarning(`Waiting for comparable data. Parsed WOS: ${wosCount}, PURE: ${pureCount}.`);
    renderDebugInfo({
      parsedWos: wosCount,
      parsedPure: pureCount,
      activeYearFilter: state.selectedYearFilter === 'all' ? 'All years' : state.selectedYearFilter,
      availableYearsCount: state.availableYears.length,
      selectedYearWos: 0,
      selectedYearPure: 0
    });
    state.summary = { totalWos: wosCount, filtered: 0, matched: 0, missing: 0, review: 0, excluded: 0 };
    renderSummary();
    return;
  }

  const selectedYearRows = getRowsForSelectedYearFilter();

  const runId = state.compareRunId + 1;
  state.compareRunId = runId;
  showYearWarning(`Preparing comparison for ${selectedYearRows.label}...`);
  await nextUiTick();
  await buildComparisonData(runId, selectedYearRows.wosRows, selectedYearRows.pureRows, {
    activeYearFilter: selectedYearRows.label,
    selectedYearWos: selectedYearRows.wosRows.length,
    selectedYearPure: selectedYearRows.pureRows.length
  });
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
        recordType: row.recordType || '',
        ut: row.ut || ''
      };
    });

    if (source === 'wos') {
      state.wosRows = parsed;
    } else {
      state.pureRows = parsed;
    }

    refreshYearFilterOptions();

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

  if (elements.yearFilterSelect) {
    elements.yearFilterSelect.addEventListener('change', (event) => {
      state.selectedYearFilter = event.target.value || 'all';
      compareAndRender();
    });
  }

  elements.exportBtn.addEventListener('click', exportMissingRows);
}

wireEvents();

window.addEventListener('DOMContentLoaded', () => {
  if (!XLSX) {
    showYearWarning('Spreadsheet engine failed to load. Please refresh the page and try again.');
  }
  renderTabs();
  renderSummary();
  refreshYearFilterOptions();
  renderDebugInfo({ parsedWos: 0, parsedPure: 0, activeYearFilter: 'All years', availableYearsCount: 0, selectedYearWos: 0, selectedYearPure: 0 });

  // If the browser preserved file selections across reload, parse them immediately.
  const wosFile = elements.wosFileInput.files?.[0];
  const pureFile = elements.pureFileInput.files?.[0];
  if (wosFile) parseUploadedFile(wosFile, 'wos');
  if (pureFile) parseUploadedFile(pureFile, 'pure');
});
