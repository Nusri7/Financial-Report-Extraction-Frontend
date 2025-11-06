import { Fragment, useEffect, useMemo, useState, useRef, useCallback } from 'react';
import axios from 'axios';
import * as XLSX from 'xlsx';
import './App.css';

const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:5005';

const SOP_METRICS = [
  'Revenues',
  'Gross profit',
  'Operating Profits',
  'Interest Expense',
  'Interest Income',
  'Profit Before Tax',
  'Taxation',
  'Net Profit',
  'Fixed Assets',
  'Inventory',
  'Trade Receivables',
  'Cash',
  'Current Assets',
  'Total Assets',
  'Total Equity',
  'Trade Payables',
  'Current Liabilities',
  'Total Liabilities',
  'Total Debt',
  'Book Value',
  'OCF Qtrly',
  'Depreciation Qtrly',
  'Amortization Qtrly',
  'ICF Qtrly',
  'Capital Exp Qtrly',
  'FCF Qtrly',
  'Net Borrowings Qrtly',
  'Share Price Quaterly',
  'Tot. No. of Shares',
];

const WORKSPACE_TABS = [
  { id: 'overview', label: 'Overview' },
  { id: 'statements', label: 'Statements' },
  { id: 'sop', label: 'SOP Summary' },
  { id: 'exports', label: 'Exports' },
];

const parseNumericValue = (input) => {
  if (input === null || typeof input === 'undefined') {
    return null;
  }
  const str = input.toString().trim();
  if (!str) {
    return null;
  }
  const normalised = str
    .replace(/[\u2212\u2012\u2013\u2014]/g, '-') // map unicode minus variants
    .replace(/[, ]+/g, '')
    .replace(/^\((.*)\)$/, '-$1')
    .replace(/[^0-9.\-]/g, '');
  if (!normalised || normalised === '-' || normalised === '.' || normalised === '-.') {
    return null;
  }
  const value = Number(normalised);
  if (!Number.isFinite(value)) {
    return null;
  }
  return value;
};

const formatNumericValue = (value) => {
  if (!Number.isFinite(value)) {
    return null;
  }
  const formatter = new Intl.NumberFormat('en-US', {
    useGrouping: true,
    maximumFractionDigits: 10,
  });
  const formatted = formatter.format(value);
  return formatted
    .replace(/\.0+$/, '')
    .replace(/(\.\d*?)0+$/, '$1')
    .replace(/\.$/, '');
};

const appendThreeZeros = (input) => {
  if (input === null || typeof input === 'undefined') {
    return null;
  }
  const raw = input.toString();
  if (!raw.trim() || /%|percent|pct/i.test(raw)) {
    return null;
  }
  const numeric = parseNumericValue(raw);
  if (numeric === null) {
    return null;
  }
  return formatNumericValue(numeric * 1000);
};

const convertValueToPositive = (input) => {
  const numeric = parseNumericValue(input);
  if (numeric === null) {
    return null;
  }
  return formatNumericValue(Math.abs(numeric));
};

const normaliseSopValue = (value) => {
  if (value === null || typeof value === 'undefined') {
    return '-';
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value.toString() : '-';
  }
  const trimmed = value.toString().trim();
  return trimmed === '' ? '-' : trimmed;
};

const cleanSopText = (value) => {
  if (value === null || typeof value === 'undefined') {
    return '';
  }
  return value.toString().trim();
};

const toTrimmed = (value) => {
  if (value === null || typeof value === 'undefined') {
    return '';
  }
  return value.toString().trim();
};

const normaliseKey = (value) => toTrimmed(value).toLowerCase();

const buildEmptySopEditDraft = () => ({
  value: '',
  statement: '',
  column: '',
  sourceLine: '',
});

const buildEmptySopSummary = () => SOP_METRICS.map((metric) => ({
  metric,
  value: '-',
  statement: '',
  column: '',
  sourceLine: '',
  manual: false,
}));

const createEmptyCalculationStep = (overrides = {}) => ({
  operator: '+',
  statement: '',
  lineItem: '',
  column: '',
  constant: '',
  ...overrides,
});

const normaliseSopSummaryEntries = (entries) => {
  const populated = new Map();

  if (Array.isArray(entries)) {
    entries.forEach((entry) => {
      if (!entry || typeof entry !== 'object') return;
      const metric = entry.metric || entry.name;
      if (!metric || populated.has(metric)) return;
      populated.set(metric, {
        metric,
        value: normaliseSopValue(entry.value),
        statement: cleanSopText(entry.statement || ''),
        column: cleanSopText(entry.column || entry.sourceColumn || ''),
        sourceLine: cleanSopText(entry.sourceLine || entry.lineItem || ''),
        manual: Boolean(entry.manual),
      });
    });
  }

  return SOP_METRICS.map((metric) => {
    if (populated.has(metric)) {
      return populated.get(metric);
    }
    return {
      metric,
      value: '-',
      statement: '',
      column: '',
      sourceLine: '',
      manual: false,
    };
  });
};

function App() {
  const [pdfName, setPdfName] = useState('');
  const [pdfBase64, setPdfBase64] = useState('');
  const [lineItems, setLineItems] = useState([]);
  const [activeStatement, setActiveStatement] = useState('');
  const [valueColumns, setValueColumns] = useState([]);
  const [sopSummary, setSopSummary] = useState(() => buildEmptySopSummary());
  const [candidateMetrics, setCandidateMetrics] = useState([]);
  const [manualSopEntries, setManualSopEntries] = useState({});
  const [sopMetadata, setSopMetadata] = useState(() => ({ latestColumns: {} }));
  const [expandedSopMetrics, setExpandedSopMetrics] = useState({});
  const [breakdownDrafts, setBreakdownDrafts] = useState({});
  const [qcComplete, setQcComplete] = useState(false);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState(null);
  const [error, setError] = useState('');
  const [showManualEntry, setShowManualEntry] = useState(false);
  const [manualRow, setManualRow] = useState({ statement: '', lineItem: '', values: {} });
  const [verifiedStatements, setVerifiedStatements] = useState({});
  const [activeWorkspaceTab, setActiveWorkspaceTab] = useState('overview');
  const [pdfZoom, setPdfZoom] = useState(1);
  const [statementMultiplierApplied, setStatementMultiplierApplied] = useState({});
  const [editingSopMetric, setEditingSopMetric] = useState(null);
  const [sopEditDraft, setSopEditDraft] = useState(() => buildEmptySopEditDraft());
  const [selectedRowIds, setSelectedRowIds] = useState(() => new Set());
  const [bulkClassificationMetric, setBulkClassificationMetric] = useState('');
  const loadingIntervalRef = useRef(null);
  const selectAllCheckboxRef = useRef(null);

  useEffect(() => {
    if (!loading) {
      if (loadingIntervalRef.current) {
        clearInterval(loadingIntervalRef.current);
        loadingIntervalRef.current = null;
      }
      return;
    }

    if (typeof window === 'undefined') {
      return;
    }

    const messages = [
      'Uploading PDF to the server...',
      'Extracting structured text from the document...',
      'Rebuilding financial tables...',
      'Preparing line items for review...',
    ];

    let index = 0;
    setStatus({
      type: 'info',
      message: `${messages[index]} This can take 2-3 minutes.`,
    });

    const intervalId = window.setInterval(() => {
      index = (index + 1) % messages.length;
      setStatus({
        type: 'info',
        message: `${messages[index]} Still working...`,
      });
    }, 20000);

    loadingIntervalRef.current = intervalId;

    return () => {
      clearInterval(intervalId);
      loadingIntervalRef.current = null;
    };
  }, [loading]);

  const totalRows = lineItems.length;
  const sopMetricOptions = useMemo(() => {
    const merged = new Set(SOP_METRICS);
    if (Array.isArray(candidateMetrics)) {
      candidateMetrics.forEach((metric) => {
        const name = typeof metric === 'string' ? metric.trim() : '';
        if (name) {
          merged.add(name);
        }
      });
    }
    lineItems.forEach((item) => {
      const name = typeof item?.classification === 'string' ? item.classification.trim() : '';
      if (name) {
        merged.add(name);
      }
    });
    return Array.from(merged);
  }, [candidateMetrics, lineItems]);

  const statements = useMemo(() => (
    Array.from(new Set(lineItems.map((item) => item.statement))).filter(Boolean)
  ), [lineItems]);

  const lineItemSuggestionsByStatement = useMemo(() => {
    const map = new Map();
    const addValue = (key, value) => {
      if (!value) {
        return;
      }
      if (!map.has(key)) {
        map.set(key, new Set());
      }
      map.get(key).add(value);
    };
    lineItems.forEach((item) => {
      if (!item || typeof item !== 'object') {
        return;
      }
      const statementName = (item.statement || '').toString().trim();
      const lineItemName = (item.lineItem || item['Line Item'] || '').toString().trim();
      if (!lineItemName) {
        return;
      }
      const statementKey = statementName || '__without_statement__';
      addValue('__all__', lineItemName);
      addValue(statementKey, lineItemName);
    });
    return map;
  }, [lineItems]);

  const getLineItemSuggestions = useCallback((statementName) => {
    if (!statementName) {
      return Array.from(lineItemSuggestionsByStatement.get('__all__') || []);
    }
    const key = statementName.toString().trim() || '__without_statement__';
    const forStatement = lineItemSuggestionsByStatement.get(key);
    if (forStatement && forStatement.size) {
      return Array.from(forStatement);
    }
    return Array.from(lineItemSuggestionsByStatement.get('__all__') || []);
  }, [lineItemSuggestionsByStatement]);

  const statementSuggestions = useMemo(() => {
    const collected = new Set(statements);
    Object.values(manualSopEntries || {}).forEach((entries) => {
      if (!Array.isArray(entries)) {
        return;
      }
      entries.forEach((entry) => {
        if (entry?.statement) {
          collected.add(entry.statement);
        }
        if (Array.isArray(entry?.calculation)) {
          entry.calculation.forEach((step) => {
            if (step?.statement) {
              collected.add(step.statement);
            }
          });
        }
      });
    });
    return Array.from(collected).filter(Boolean);
  }, [statements, manualSopEntries]);

  useEffect(() => {
    if (!statements.length) {
      setActiveStatement('');
      return;
    }
    setActiveStatement((current) => (
      statements.includes(current) ? current : statements[0]
    ));
  }, [statements]);

  useEffect(() => {
    if (!statements.length) {
      setStatementMultiplierApplied({});
      return;
    }
    setStatementMultiplierApplied((prev) => {
      const filtered = {};
      statements.forEach((statement) => {
        if (prev[statement]) {
          filtered[statement] = true;
        }
      });
      const prevKeys = Object.keys(prev).filter((key) => statements.includes(key));
      if (prevKeys.length === Object.keys(filtered).length) {
        const unchanged = statements.every((statement) => !!prev[statement] === !!filtered[statement]);
        if (unchanged) {
          return prev;
        }
      }
      return filtered;
    });
  }, [statements]);

  useEffect(() => {
    if (!statements.length) {
      setVerifiedStatements({});
      return;
    }

    setVerifiedStatements((prev) => {
      const next = {};
      statements.forEach((statement) => {
        next[statement] = Object.prototype.hasOwnProperty.call(prev, statement) ? prev[statement] : false;
      });
      return next;
    });
  }, [statements]);

  useEffect(() => {
    if (!pdfBase64 && activeWorkspaceTab !== 'overview') {
      setActiveWorkspaceTab('overview');
    }
  }, [pdfBase64, activeWorkspaceTab]);

  const visibleItems = useMemo(() => {
    if (!activeStatement) {
      return [];
    }
    return lineItems.filter((item) => item.statement === activeStatement);
  }, [activeStatement, lineItems]);

  useEffect(() => {
    setSelectedRowIds((prev) => {
      if (!prev.size) {
        return prev;
      }
      const visibleSet = new Set(visibleItems.map((item) => item.rowId));
      let changed = false;
      const next = new Set();
      prev.forEach((id) => {
        if (visibleSet.has(id)) {
          next.add(id);
        } else {
          changed = true;
        }
      });
      if (!changed && next.size === prev.size) {
        return prev;
      }
      return next;
    });
  }, [visibleItems]);

  useEffect(() => {
    if (!selectAllCheckboxRef.current) {
      return;
    }
    const checkbox = selectAllCheckboxRef.current;
    const total = visibleItems.length;
    const selected = selectedRowIds.size;
    checkbox.indeterminate = selected > 0 && selected < total;
    checkbox.checked = total > 0 && selected === total;
  }, [selectedRowIds, visibleItems]);

  const statementTotalsMap = useMemo(() => {
    const map = {};
    lineItems.forEach((item) => {
      if (!item.statement) return;
      if (!map[item.statement]) {
        map[item.statement] = 0;
      }
      map[item.statement] += 1;
    });
    return map;
  }, [lineItems]);

  const totalStatements = statements.length;
  const pendingStatements = statements.filter((statement) => !verifiedStatements[statement]).length;
  const activeStatementVerified = activeStatement ? !!verifiedStatements[activeStatement] : false;

  const filterColumnsForStatement = useCallback((columns, statementName) => {
    if (!Array.isArray(columns) || !columns.length) {
      return columns;
    }
    if (!statementName) {
      return columns;
    }
    const lowerStatement = statementName.toString().toLowerCase();
    if (!lowerStatement.includes('profit') && !lowerStatement.includes('loss')) {
      return columns;
    }
    return columns.filter((column) => {
      const text = (column ?? '').toString().toLowerCase();
      if (!text) return true;
      const hasChange = text.includes('change');
      const hasPercent = text.includes('%') || text.includes('percent') || text.includes('pct');
      return !(hasChange && hasPercent);
    });
  }, []);

  const statementValueColumns = useMemo(() => {
    const baseColumns = filterColumnsForStatement(valueColumns, activeStatement);
    if (!visibleItems.length) {
      return baseColumns;
    }
    const populatedColumns = baseColumns.filter((column) => (
      visibleItems.some((item) => {
        const raw = item[column];
        if (typeof raw === 'number') return true;
        return typeof raw === 'string' && raw.trim() !== '';
      })
    ));
    return populatedColumns.length ? populatedColumns : baseColumns;
  }, [valueColumns, visibleItems, activeStatement, filterColumnsForStatement]);

  const manualEntryColumns = useMemo(() => (
    filterColumnsForStatement(valueColumns, manualRow.statement || activeStatement || '')
  ), [valueColumns, manualRow.statement, activeStatement, filterColumnsForStatement]);

  const lineItemBreakdown = useMemo(() => {
    const grouped = {};
    lineItems.forEach((item) => {
      if (!item) {
        return;
      }
      const metric = typeof item.classification === 'string' ? item.classification.trim() : '';
      if (!metric) {
        return;
      }
      const values = {};
      valueColumns.forEach((column) => {
        const raw = item[column];
        if (raw === null || typeof raw === 'undefined') {
          return;
        }
        const text = raw.toString().trim();
        if (text) {
          values[column] = text;
        }
      });
      if (!grouped[metric]) {
        grouped[metric] = [];
      }
      grouped[metric].push({
        type: 'lineItem',
        rowId: item.rowId,
        statement: item.statement || '',
        lineItem: item.lineItem || item['Line Item'] || '',
        values,
      });
    });
    return grouped;
  }, [lineItems, valueColumns]);

  const lineItemLookup = useMemo(() => {
    const lookup = new Map();
    lineItems.forEach((item) => {
      if (!item || typeof item !== 'object') {
        return;
      }
      const statementKey = normaliseKey(item.statement);
      const lineLabelKey = normaliseKey(item.lineItem || item['Line Item']);
      if (!statementKey || !lineLabelKey) {
        return;
      }
      const key = `${statementKey}||${lineLabelKey}`;
      if (!lookup.has(key)) {
        lookup.set(key, []);
      }
      lookup.get(key).push(item);
    });
    return lookup;
  }, [lineItems]);

  const manualSopOverrides = useMemo(() => {
    if (!manualSopEntries || !Object.keys(manualSopEntries).length) {
      return {};
    }

    const latestColumnsInput = (sopMetadata && typeof sopMetadata === 'object')
      ? sopMetadata.latestColumns || {}
      : {};
    const latestColumnByStatement = new Map();
    Object.entries(latestColumnsInput).forEach(([statementName, columnName]) => {
      const normalisedStatement = normaliseKey(statementName);
      const trimmedColumn = toTrimmed(columnName);
      if (!normalisedStatement || !trimmedColumn) {
        return;
      }
      if (!latestColumnByStatement.has(normalisedStatement)) {
        latestColumnByStatement.set(normalisedStatement, trimmedColumn);
      }
    });

    const findMatchingColumnName = (row, columnName) => {
      const trimmed = toTrimmed(columnName);
      if (!trimmed) {
        return '';
      }
      const targetKey = normaliseKey(trimmed);
      const rowKeys = Object.keys(row || {});
      for (let idx = 0; idx < rowKeys.length; idx += 1) {
        const key = rowKeys[idx];
        if (normaliseKey(key) === targetKey) {
          return key;
        }
      }
      return '';
    };

    const resolveLineItemValue = (statementName, lineItemName, columnHint) => {
      const statementKey = normaliseKey(statementName);
      const lineItemKey = normaliseKey(lineItemName);
      if (!statementKey || !lineItemKey) {
        return null;
      }
      const lookupKey = `${statementKey}||${lineItemKey}`;
      const rows = lineItemLookup.get(lookupKey);
      if (!rows || !rows.length) {
        return null;
      }

      const candidates = [];
      const seen = new Set();

      const enqueue = (row, candidateColumn) => {
        if (!row) {
          return;
        }
        const resolvedColumn = findMatchingColumnName(row, candidateColumn);
        if (!resolvedColumn) {
          return;
        }
        const identifier = `${row.rowId || `${normaliseKey(row.statement)}||${normaliseKey(row.lineItem || row['Line Item'])}`}||${resolvedColumn}`;
        if (seen.has(identifier)) {
          return;
        }
        seen.add(identifier);
        candidates.push({ row, column: resolvedColumn });
      };

      rows.forEach((row) => {
        if (columnHint) {
          enqueue(row, columnHint);
        }
      });

      const metaColumn = latestColumnByStatement.get(statementKey);
      rows.forEach((row) => {
        if (metaColumn) {
          enqueue(row, metaColumn);
        }
      });

      valueColumns.forEach((columnName) => {
        rows.forEach((row) => enqueue(row, columnName));
      });

      rows.forEach((row) => {
        Object.keys(row).forEach((key) => {
          if (['rowId', 'statement', 'lineItem', 'Line Item', 'classification', 'aiConfidence'].includes(key)) {
            return;
          }
          enqueue(row, key);
        });
      });

      for (let idx = 0; idx < candidates.length; idx += 1) {
        const candidate = candidates[idx];
        const raw = candidate.row[candidate.column];
        if (raw === null || typeof raw === 'undefined') {
          continue;
        }
        const numeric = parseNumericValue(raw);
        if (numeric === null) {
          continue;
        }
        return {
          numericValue: numeric,
          columnName: candidate.column,
          statementName: candidate.row.statement || statementName,
          lineItemName: candidate.row.lineItem || candidate.row['Line Item'] || lineItemName,
          displayValue: raw,
        };
      }

      return null;
    };

    const evaluateEntry = (rawEntry) => {
      if (!rawEntry || typeof rawEntry !== 'object') {
        return null;
      }
      const entryStatement = toTrimmed(rawEntry.statement);
      const entryLineItem = toTrimmed(rawEntry.lineItem);
      const entryColumn = toTrimmed(rawEntry.column);
      const entryValueText = toTrimmed(rawEntry.value);
      const steps = Array.isArray(rawEntry.calculation) ? rawEntry.calculation : [];

      const columnsUsed = new Set();
      const formulaParts = [];

      const addColumn = (columnName) => {
        if (!columnName) {
          return;
        }
        columnsUsed.add(columnName);
      };

      let runningTotal = null;
      let baseContext = null;

      if (entryValueText) {
        const numeric = parseNumericValue(entryValueText);
        if (numeric === null) {
          return null;
        }
        runningTotal = numeric;
        const formatted = formatNumericValue(numeric) ?? numeric.toString();
        formulaParts.push(`Manual value ${formatted}`);
        addColumn('Manual Input');
      } else if (entryStatement && entryLineItem) {
        const resolved = resolveLineItemValue(entryStatement, entryLineItem, entryColumn);
        if (!resolved) {
          runningTotal = 0;
          baseContext = {
            statementName: entryStatement,
            columnName: '',
            lineItemName: entryLineItem,
          };
          formulaParts.push(`Start at 0 for ${entryLineItem} [${entryStatement}]`);
        } else {
          runningTotal = resolved.numericValue;
          baseContext = resolved;
          const valueDisplay = resolved.displayValue
            ? normaliseSopValue(resolved.displayValue)
            : formatNumericValue(resolved.numericValue) ?? resolved.numericValue.toString();
          const columnDescriptor = resolved.columnName ? ` -> ${resolved.columnName}` : '';
          formulaParts.push(`${resolved.lineItemName || entryLineItem} [${resolved.statementName || entryStatement}${columnDescriptor}] (${valueDisplay})`);
          addColumn(resolved.columnName || '');
        }
      } else {
        return null;
      }

      if (runningTotal === null) {
        return null;
      }

      for (let idx = 0; idx < steps.length; idx += 1) {
        const step = steps[idx];
        if (!step || typeof step !== 'object') {
          continue;
        }
        const operator = ['+', '-', '*', '/'].includes(step.operator) ? step.operator : '+';
        const stepConstant = toTrimmed(step.constant);
        let operandValue = null;
        let operandDescription = '';
        let operandColumn = '';

        if (stepConstant) {
          const numeric = parseNumericValue(stepConstant);
          if (numeric === null) {
            return null;
          }
          operandValue = numeric;
          const formatted = formatNumericValue(numeric) ?? numeric.toString();
          operandDescription = `manual constant ${formatted}`;
          operandColumn = 'Manual Input';
        } else {
          const operandStatement = toTrimmed(step.statement) || baseContext?.statementName || entryStatement;
          const operandLineItem = toTrimmed(step.lineItem);
          if (!operandLineItem) {
            return null;
          }
          const operandColumnHint = toTrimmed(step.column) || baseContext?.columnName || entryColumn;
          const resolved = resolveLineItemValue(operandStatement, operandLineItem, operandColumnHint);
          if (!resolved) {
            return null;
          }
          operandValue = resolved.numericValue;
          operandColumn = resolved.columnName || '';
          const valueDisplay = resolved.displayValue
            ? normaliseSopValue(resolved.displayValue)
            : formatNumericValue(resolved.numericValue) ?? resolved.numericValue.toString();
          const columnDescriptor = resolved.columnName ? ` -> ${resolved.columnName}` : '';
          operandDescription = `${resolved.lineItemName || operandLineItem} [${resolved.statementName || operandStatement}${columnDescriptor}] (${valueDisplay})`;
        }

        switch (operator) {
          case '+':
            runningTotal += operandValue;
            break;
          case '-':
            runningTotal -= operandValue;
            break;
          case '*':
            runningTotal *= operandValue;
            break;
          case '/':
            if (operandValue === 0) {
              return null;
            }
            runningTotal /= operandValue;
            break;
          default:
            runningTotal += operandValue;
            break;
        }

        if (!Number.isFinite(runningTotal)) {
          return null;
        }

        formulaParts.push(`${operator} ${operandDescription}`);
        addColumn(operandColumn);
      }

      const formula = formulaParts
        .map((part) => part.replace(/\s+/g, ' ').trim())
        .filter(Boolean)
        .join(' ');

      return {
        total: runningTotal,
        columns: Array.from(columnsUsed).filter(Boolean),
        formula,
      };
    };

    const overrides = {};
    Object.entries(manualSopEntries).forEach(([metric, entries]) => {
      if (!Array.isArray(entries) || !entries.length) {
        return;
      }

      let aggregate = 0;
      let hasValue = false;
      const formulas = [];
      const columns = new Set();

      entries.forEach((entry) => {
        const result = evaluateEntry(entry);
        if (!result) {
          return;
        }
        aggregate += result.total;
        hasValue = true;
        if (result.formula) {
          formulas.push(result.formula);
        }
        result.columns.forEach((columnName) => columns.add(columnName));
      });

      if (!hasValue) {
        return;
      }

      const formattedValue = formatNumericValue(aggregate);
      const valueText = normaliseSopValue(formattedValue ?? aggregate);
      const columnValues = Array.from(columns).filter(Boolean);
      const columnText = columnValues.length === 0
        ? ''
        : columnValues.length === 1
          ? columnValues[0]
          : 'Multiple';
      const sourceLine = formulas.length
        ? formulas.join('; ')
        : 'Derived from manual calculation';

      overrides[metric] = {
        value: valueText,
        statement: 'Derived',
        column: columnText,
        sourceLine,
      };
    });

    return overrides;
  }, [manualSopEntries, lineItemLookup, sopMetadata, valueColumns]);

  const displayedSopSummary = useMemo(() => (
    sopSummary.map((entry) => {
      const override = manualSopOverrides[entry.metric];
      if (!override) {
        return entry;
      }
      const nextColumn = Object.prototype.hasOwnProperty.call(override, 'column')
        ? override.column
        : entry.column;
      const nextStatement = Object.prototype.hasOwnProperty.call(override, 'statement')
        ? override.statement
        : entry.statement;
      const nextSourceLine = Object.prototype.hasOwnProperty.call(override, 'sourceLine')
        ? override.sourceLine
        : entry.sourceLine;
      return {
        ...entry,
        value: override.value,
        statement: nextStatement,
        column: nextColumn,
        sourceLine: nextSourceLine,
        manual: true,
      };
    })
  ), [sopSummary, manualSopOverrides]);

  const iframeSrc = useMemo(() => {
    if (!pdfBase64) return '';
    return `data:application/pdf;base64,${pdfBase64}`;
  }, [pdfBase64]);

  const handleValueChange = (rowId, columnName, value, options = {}) => {
    const safeValue = value === null || typeof value === 'undefined' ? '' : value.toString();
    let affectedStatement = null;
    let affectedLineItem = '';
    let valueChanged = false;

    setLineItems((items) => items.map((item) => {
      if (item.rowId !== rowId) {
        return item;
      }
      affectedStatement = item.statement;
      affectedLineItem = item.lineItem || item['Line Item'] || '';
      if (item[columnName] === safeValue) {
        return item;
      }
      valueChanged = true;
      return { ...item, [columnName]: safeValue };
    }));

    if (!valueChanged) {
      if (!options?.preserveStatus) {
        setStatus(null);
      }
      return;
    }

    if (affectedStatement) {
      setVerifiedStatements((prev) => ({
        ...prev,
        [affectedStatement]: false,
      }));
    }

    if (affectedStatement && affectedLineItem) {
      const statementKey = affectedStatement.toString().toLowerCase();
      const lineItemKey = affectedLineItem.toString().toLowerCase();
      const columnKey = columnName.toString().toLowerCase();

      setSopSummary((current) => current.map((entry) => {
        if (!entry.statement || !entry.sourceLine || !entry.column) {
          return entry;
        }
        if (entry.manual) {
          return entry;
        }
        const entryStatement = entry.statement.toString().toLowerCase();
        const entryLine = entry.sourceLine.toString().toLowerCase();
        const entryColumn = entry.column.toString().toLowerCase();

        if (entryStatement === statementKey
          && entryLine === lineItemKey
          && entryColumn === columnKey) {
          return {
            ...entry,
            value: normaliseSopValue(safeValue),
          };
        }
        return entry;
      }));
    }

    setQcComplete(false);
    if (!options?.preserveStatus) {
      setStatus(null);
    }
  };

  const handleSopEditStart = (metric) => {
    const currentEntry = displayedSopSummary.find((entry) => entry.metric === metric);
    if (!currentEntry) {
      return;
    }
    setEditingSopMetric(metric);
    setSopEditDraft({
      value: currentEntry.value === '-' ? '' : currentEntry.value || '',
      statement: currentEntry.statement || '',
      column: currentEntry.column || '',
      sourceLine: currentEntry.sourceLine || '',
    });
  };

  const handleSopEditCancel = () => {
    setEditingSopMetric(null);
    setSopEditDraft(buildEmptySopEditDraft());
  };

  const handleSopEditFieldChange = (field, value) => {
    setSopEditDraft((prev) => ({
      ...prev,
      [field]: value,
    }));
  };

  const handleSopEditSave = () => {
    if (!editingSopMetric) {
      return;
    }
    setSopSummary((current) => current.map((entry) => {
      if (entry.metric !== editingSopMetric) {
        return entry;
      }
      return {
        ...entry,
        value: normaliseSopValue(sopEditDraft.value),
        statement: cleanSopText(sopEditDraft.statement),
        column: cleanSopText(sopEditDraft.column),
        sourceLine: cleanSopText(sopEditDraft.sourceLine),
        manual: true,
      };
    }));
    setStatus({ type: 'success', message: `${editingSopMetric} updated in SOP summary.` });
    setQcComplete(false);
    handleSopEditCancel();
  };

  const buildEmptyValueMap = () => valueColumns.reduce((acc, column) => {
    acc[column] = '';
    return acc;
  }, {});

  const openManualEntry = () => {
    const defaultStatement = activeStatement || statements[0] || '';
    setManualRow({
      statement: defaultStatement,
      lineItem: '',
      values: buildEmptyValueMap(),
    });
    setShowManualEntry(true);
    setStatus(null);
  };

  const handleManualFieldChange = (field, value) => {
    setManualRow((prev) => ({
      ...prev,
      [field]: value,
    }));
  };

  const handleManualValueInput = (column, value) => {
    setManualRow((prev) => ({
      ...prev,
      values: {
        ...(prev.values || {}),
        [column]: value,
      },
    }));
  };

  const handleManualRowCancel = () => {
    setShowManualEntry(false);
    setManualRow({ statement: activeStatement || '', lineItem: '', values: buildEmptyValueMap() });
    setStatus(null);
  };

  const handleManualRowSubmit = () => {
    const statement = (manualRow.statement || '').trim();
    const lineItem = (manualRow.lineItem || '').trim();

    if (!statement || !lineItem) {
      setStatus({ type: 'error', message: 'Enter both a statement name and line item before adding a row.' });
      return;
    }

    const newRowId = `manual-${Date.now()}`;
    const newRow = {
      rowId: newRowId,
      statement,
      lineItem,
    };

    valueColumns.forEach((column) => {
      newRow[column] = (manualRow.values?.[column] || '').trim();
    });

    setLineItems((prev) => [...prev, newRow]);
    setVerifiedStatements((prev) => ({
      ...prev,
      [statement]: false,
    }));
    setActiveStatement(statement);
    setShowManualEntry(false);
    setManualRow({
      statement,
      lineItem: '',
      values: buildEmptyValueMap(),
    });
    setQcComplete(false);
    setActiveWorkspaceTab('statements');
    setStatus({ type: 'info', message: 'Manual line item added. Review the values and mark the statement when ready.' });
  };

  const handleClassificationChange = (rowIds, metricName, options = {}) => {
    const trimmedMetric = typeof metricName === 'string' ? metricName.trim() : '';
    const targets = Array.isArray(rowIds) ? rowIds : [rowIds];
    const targetSet = new Set(targets.filter(Boolean));

    if (!targetSet.size) {
      if (!options?.preserveStatus) {
        setStatus({ type: 'warning', message: 'Select at least one row before updating the SOP metric.' });
      }
      return;
    }

    const affectedStatements = new Set();
    let changedCount = 0;

    setLineItems((items) => items.map((item) => {
      if (!targetSet.has(item.rowId)) {
        return item;
      }
      if (item.statement) {
        affectedStatements.add(item.statement);
      }
      const current = typeof item.classification === 'string' ? item.classification.trim() : '';
      if (current === trimmedMetric) {
        return item;
      }
      changedCount += 1;
      if (!trimmedMetric) {
        const next = { ...item };
        delete next.classification;
        return next;
      }
      return { ...item, classification: trimmedMetric };
    }));

    if (!changedCount) {
      if (!options?.preserveStatus) {
        setStatus({
          type: 'info',
          message: trimmedMetric
            ? `Selected row${targetSet.size === 1 ? '' : 's'} already linked to "${trimmedMetric}".`
            : 'Selected row(s) already had no SOP metric assigned.',
        });
      }
      return;
    }

    if (affectedStatements.size) {
      setVerifiedStatements((prev) => {
        const next = { ...prev };
        affectedStatements.forEach((statement) => {
          if (statement) {
            next[statement] = false;
          }
        });
        return next;
      });
    }
    setQcComplete(false);
    if (!options?.preserveStatus) {
      setStatus({
        type: 'success',
        message: trimmedMetric
          ? `Linked ${changedCount} row${changedCount === 1 ? '' : 's'} to "${trimmedMetric}".`
          : `Cleared SOP metric for ${changedCount} row${changedCount === 1 ? '' : 's'}.`,
      });
    }
  };

  const selectedRowCount = selectedRowIds.size;

  const handleRowSelectionToggle = (rowId, checked) => {
    setSelectedRowIds((prev) => {
      const next = new Set(prev);
      if (checked) {
        next.add(rowId);
      } else {
        next.delete(rowId);
      }
      return next;
    });
  };

  const handleSelectAllVisibleRows = (checked) => {
    if (checked) {
      setSelectedRowIds(new Set(visibleItems.map((item) => item.rowId)));
    } else {
      setSelectedRowIds(new Set());
    }
  };

  const handleBulkClassificationApply = () => {
    const trimmedMetric = typeof bulkClassificationMetric === 'string'
      ? bulkClassificationMetric.trim()
      : '';
    if (!selectedRowCount) {
      setStatus({ type: 'warning', message: 'Select at least one row before applying a bulk classification.' });
      return;
    }
    if (!trimmedMetric) {
      setStatus({ type: 'warning', message: 'Choose a SOP metric before applying the bulk classification.' });
      return;
    }
    handleClassificationChange(Array.from(selectedRowIds), trimmedMetric);
    setSelectedRowIds(new Set());
  };

  const handleClearRowSelection = () => {
    setSelectedRowIds(new Set());
  };

  const toggleSopMetricExpansion = (metric) => {
    setExpandedSopMetrics((prev) => ({
      ...prev,
      [metric]: !prev?.[metric],
    }));
  };

  const handleManualBreakdownDraftChange = (metric, field, value) => {
    setBreakdownDrafts((prev) => {
      const current = prev?.[metric] || {};
      const next = {
        statement: current.statement || '',
        lineItem: current.lineItem || '',
        calculation: Array.isArray(current.calculation)
          ? current.calculation
          : [],
      };
      if (field === 'statement' || field === 'lineItem') {
        next[field] = value;
      }
      return {
        ...prev,
        [metric]: next,
      };
    });
  };

  const handleManualBreakdownDraftCalculationChange = (metric, index, field, value) => {
    setBreakdownDrafts((prev) => {
      const current = prev?.[metric] || {};
      const steps = Array.isArray(current.calculation)
        ? current.calculation.map((step) => ({ ...createEmptyCalculationStep(), ...step }))
        : [];
      while (steps.length <= index) {
        steps.push(createEmptyCalculationStep());
      }
      steps[index] = {
        ...steps[index],
        [field]: value,
      };
      return {
        ...prev,
        [metric]: {
          statement: current.statement || '',
          lineItem: current.lineItem || '',
          column: current.column || '',
          value: current.value || '',
          calculation: steps,
        },
      };
    });
  };

  const handleAddManualBreakdownDraftStep = (metric) => {
    setBreakdownDrafts((prev) => {
      const current = prev?.[metric] || {};
      const steps = Array.isArray(current.calculation)
        ? [...current.calculation.map((step) => ({ ...createEmptyCalculationStep(), ...step }))]
        : [];
      steps.push(createEmptyCalculationStep({ operator: steps.length ? '+' : '+' }));
      return {
        ...prev,
        [metric]: {
          statement: current.statement || '',
          lineItem: current.lineItem || '',
          column: current.column || '',
          value: current.value || '',
          calculation: steps,
        },
      };
    });
  };

  const handleRemoveManualBreakdownDraftStep = (metric, index) => {
    setBreakdownDrafts((prev) => {
      const current = prev?.[metric];
      if (!current) {
        return prev;
      }
      const steps = Array.isArray(current.calculation)
        ? current.calculation.filter((_, stepIndex) => stepIndex !== index)
        : [];
      return {
        ...prev,
        [metric]: {
          statement: current.statement || '',
          lineItem: current.lineItem || '',
          column: current.column || '',
          value: current.value || '',
          calculation: steps,
        },
      };
    });
  };

  const handleAddManualBreakdownEntry = (metric) => {
    const draft = breakdownDrafts?.[metric] || {};
    const statement = (draft.statement || '').trim();
    const lineItem = (draft.lineItem || '').trim();
    const calculationSteps = Array.isArray(draft.calculation)
      ? draft.calculation
        .map((step) => ({
          ...createEmptyCalculationStep(),
          ...step,
          operator: step?.operator || '+',
          statement: (step?.statement || '').trim(),
          lineItem: (step?.lineItem || '').trim(),
          column: (step?.column || '').trim(),
          constant: (step?.constant || '').trim(),
        }))
        .filter((step) => (
          step.statement
          || step.lineItem
          || step.column
          || step.constant
        ))
      : [];

    if (!statement || !lineItem) {
      setStatus({ type: 'error', message: 'Provide both a statement and line item before adding a breakdown row.' });
      return;
    }

    if (!calculationSteps.length) {
      setStatus({ type: 'error', message: 'Add at least one calculation step before adding a breakdown row.' });
      return;
    }

    const newEntry = {
      id: `manual-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      statement,
      lineItem,
      column: '',
      value: '',
      calculation: calculationSteps,
    };

    setManualSopEntries((prev) => {
      const next = { ...(prev || {}) };
      const existing = Array.isArray(next[metric]) ? next[metric] : [];
      next[metric] = [...existing, newEntry];
      return next;
    });

    setBreakdownDrafts((prev) => ({
      ...prev,
      [metric]: {
        statement: '',
        lineItem: '',
        calculation: [],
      },
    }));

    setQcComplete(false);
    setStatus({ type: 'success', message: `Added manual breakdown entry to "${metric}".` });
  };

  const handleManualBreakdownValueChange = (metric, entryId, field, value) => {
    setManualSopEntries((prev) => {
      const existing = Array.isArray(prev?.[metric]) ? prev[metric] : [];
      const updated = existing.map((entry) => {
        if (entry.id !== entryId) {
          return entry;
        }
        return {
          ...entry,
          [field]: value,
        };
      });
      return {
        ...(prev || {}),
        [metric]: updated,
      };
    });
    setQcComplete(false);
  };

  const handleManualBreakdownCalculationChange = (metric, entryId, index, field, value) => {
    setManualSopEntries((prev) => {
      const existing = Array.isArray(prev?.[metric]) ? prev[metric] : [];
      const updated = existing.map((entry) => {
        if (entry.id !== entryId) {
          return entry;
        }
        const steps = Array.isArray(entry.calculation)
          ? entry.calculation.map((step) => ({ ...createEmptyCalculationStep(), ...step }))
          : [];
        while (steps.length <= index) {
          steps.push(createEmptyCalculationStep());
        }
        steps[index] = {
          ...steps[index],
          [field]: value,
        };
        return {
          ...entry,
          calculation: steps,
        };
      });
      return {
        ...(prev || {}),
        [metric]: updated,
      };
    });
    setQcComplete(false);
  };

  const handleAddManualBreakdownCalculationStep = (metric, entryId) => {
    setManualSopEntries((prev) => {
      const existing = Array.isArray(prev?.[metric]) ? prev[metric] : [];
      const updated = existing.map((entry) => {
        if (entry.id !== entryId) {
          return entry;
        }
        const steps = Array.isArray(entry.calculation)
          ? entry.calculation.map((step) => ({ ...createEmptyCalculationStep(), ...step }))
          : [];
        steps.push(createEmptyCalculationStep({ operator: steps.length ? '+' : '+' }));
        return {
          ...entry,
          calculation: steps,
        };
      });
      return {
        ...(prev || {}),
        [metric]: updated,
      };
    });
    setQcComplete(false);
  };

  const handleRemoveManualBreakdownCalculationStep = (metric, entryId, index) => {
    setManualSopEntries((prev) => {
      const existing = Array.isArray(prev?.[metric]) ? prev[metric] : [];
      const updated = existing.map((entry) => {
        if (entry.id !== entryId) {
          return entry;
        }
        const steps = Array.isArray(entry.calculation)
          ? entry.calculation.filter((_, stepIndex) => stepIndex !== index)
          : [];
        return {
          ...entry,
          calculation: steps,
        };
      });
      return {
        ...(prev || {}),
        [metric]: updated,
      };
    });
    setQcComplete(false);
  };

  const handleRemoveManualBreakdownEntry = (metric, entryId) => {
    setManualSopEntries((prev) => {
      const existing = Array.isArray(prev?.[metric]) ? prev[metric] : [];
      const filtered = existing.filter((entry) => entry.id !== entryId);
      const next = { ...(prev || {}) };
      if (filtered.length) {
        next[metric] = filtered;
      } else {
        delete next[metric];
      }
      return next;
    });
    setQcComplete(false);
    setStatus({ type: 'info', message: `Removed a manual breakdown entry from "${metric}".` });
  };

  const handleStatementAddZeros = () => {
    if (!activeStatement) {
      setStatus({ type: 'warning', message: 'Select a statement before applying the multiplier.' });
      return;
    }
    if (statementMultiplierApplied[activeStatement]) {
      setStatus({
        type: 'info',
        message: `The x1,000 multiplier has already been applied to ${activeStatement}.`,
      });
      return;
    }
    if (!statementValueColumns.length) {
      setStatus({ type: 'info', message: 'No numeric columns available in this statement.' });
      return;
    }
    let updatedCells = 0;
    lineItems.forEach((row) => {
      if (row.statement !== activeStatement) {
        return;
      }
      statementValueColumns.forEach((column) => {
        const nextValue = appendThreeZeros(row[column]);
        if (nextValue === null) {
          return;
        }
        const original = row[column] === null || typeof row[column] === 'undefined'
          ? ''
          : row[column].toString().trim();
        if (nextValue !== original) {
          updatedCells += 1;
          handleValueChange(row.rowId, column, nextValue, { preserveStatus: true });
        }
      });
    });
    if (updatedCells) {
      setStatementMultiplierApplied((prev) => ({
        ...prev,
        [activeStatement]: true,
      }));
      setStatus({
        type: 'success',
        message: `Added three zeros to ${updatedCells} cell${updatedCells === 1 ? '' : 's'} in ${activeStatement}.`,
      });
    } else {
      setStatus({ type: 'info', message: 'No numeric values were updated.' });
    }
  };

  const handleStatementMakePositive = () => {
    if (!activeStatement) {
      setStatus({ type: 'warning', message: 'Select a statement before converting values.' });
      return;
    }
    if (!statementValueColumns.length) {
      setStatus({ type: 'info', message: 'No numeric columns available in this statement.' });
      return;
    }
    let updatedCells = 0;
    lineItems.forEach((row) => {
      if (row.statement !== activeStatement) {
        return;
      }
      statementValueColumns.forEach((column) => {
        const nextValue = convertValueToPositive(row[column]);
        if (nextValue === null) {
          return;
        }
        const original = row[column] === null || typeof row[column] === 'undefined'
          ? ''
          : row[column].toString().trim();
        if (nextValue !== original) {
          updatedCells += 1;
          handleValueChange(row.rowId, column, nextValue, { preserveStatus: true });
        }
      });
    });
    if (updatedCells) {
      setStatus({
        type: 'success',
        message: `Converted ${updatedCells} cell${updatedCells === 1 ? '' : 's'} to positive in ${activeStatement}.`,
      });
    } else {
      setStatus({ type: 'info', message: 'No negative values were found to update.' });
    }
  };

  const handleDeleteRow = (rowId) => {
    const targetRow = lineItems.find((item) => item.rowId === rowId);
    if (!targetRow) {
      return;
    }
    const statementName = targetRow.statement || '';
    const lineItemName = targetRow.lineItem || targetRow['Line Item'] || '';
    const statementKey = statementName.toString().toLowerCase();
    const lineItemKey = lineItemName.toString().toLowerCase();

    setLineItems((prev) => prev.filter((item) => item.rowId !== rowId));

    setSelectedRowIds((prev) => {
      if (!prev.has(rowId)) {
        return prev;
      }
      const next = new Set(prev);
      next.delete(rowId);
      return next;
    });

    if (statementName) {
      setVerifiedStatements((prev) => ({
        ...prev,
        [statementName]: false,
      }));
    }

    if (statementKey && lineItemKey) {
      setSopSummary((current) => current.map((entry) => {
        if (!entry || !entry.statement || !entry.sourceLine || !entry.column) {
          return entry;
        }
        if (entry.manual) {
          return entry;
        }
        const entryStatement = entry.statement.toString().toLowerCase();
        const entryLine = entry.sourceLine.toString().toLowerCase();
        if (entryStatement === statementKey && entryLine === lineItemKey) {
          return {
            ...entry,
            value: '-',
          };
        }
        return entry;
      }));
    }

    setQcComplete(false);
    setStatus({ type: 'info', message: 'Line item removed from the dataset.' });
  };

  const handleRemoveColumn = (columnName) => {
    if (!columnName || !valueColumns.includes(columnName)) {
      return;
    }
    const columnKey = columnName.toString().toLowerCase();

    setValueColumns((cols) => cols.filter((column) => column !== columnName));

    setLineItems((prev) => prev.map((item) => {
      if (!Object.prototype.hasOwnProperty.call(item, columnName)) {
        return item;
      }
      const next = { ...item };
      delete next[columnName];
      return next;
    }));

    setManualRow((prev) => {
      const nextValues = Object.entries(prev.values || {}).reduce((acc, [name, val]) => {
        if (name !== columnName) {
          acc[name] = val;
        }
        return acc;
      }, {});
      return {
        ...prev,
        values: nextValues,
      };
    });

    setManualSopEntries((prev) => {
      if (!prev || !Object.keys(prev).length) {
        return prev;
      }
      let changed = false;
      const next = {};
      Object.entries(prev).forEach(([metric, entries]) => {
        if (!Array.isArray(entries)) {
          next[metric] = entries;
          return;
        }
        const updated = entries.map((entry) => {
          if (!entry || typeof entry !== 'object') {
            return entry;
          }
          const entryColumn = entry.column ? entry.column.toString().toLowerCase() : '';
          if (entryColumn === columnKey) {
            changed = true;
            return {
              ...entry,
              column: '',
              value: '',
            };
          }
          return entry;
        });
        next[metric] = updated;
      });
      return changed ? next : prev;
    });

    setSopSummary((current) => current.map((entry) => {
      if (!entry || !entry.column) {
        return entry;
      }
      if (entry.column.toString().toLowerCase() !== columnKey) {
        return entry;
      }
      return {
        ...entry,
        column: '',
        value: '-',
      };
    }));

    setVerifiedStatements(() => {
      const next = {};
      statements.forEach((statement) => {
        next[statement] = false;
      });
      return next;
    });

    setQcComplete(false);
    setStatus({ type: 'info', message: `Column "${columnName}" removed.` });
  };

  const handleFileChange = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError('');
    setStatus({
      type: 'info',
      message: 'Uploading PDF to the server... This can take 2-3 minutes.',
    });
    setQcComplete(false);
    setSopSummary(buildEmptySopSummary());
    setSopMetadata({ latestColumns: {} });
    setEditingSopMetric(null);
    setSopEditDraft(buildEmptySopEditDraft());
    setActiveWorkspaceTab('overview');

    let successMessage = null;

    try {
      const formData = new FormData();
      formData.append('file', file);

      const response = await axios.post(
        `${API_BASE_URL}/api/process`,
        formData,
        {
          headers: { 'Content-Type': 'multipart/form-data' },
        },
      );

      const nextValueColumns = response.data.valueColumns || [];
      const sopEntries = normaliseSopSummaryEntries(response.data.sopSummary);
      const candidateMetricList = Array.isArray(response.data.candidateMetrics)
        ? response.data.candidateMetrics
        : [];

      const incomingLineItems = response.data.lineItems || [];
      const sanitizedLineItems = incomingLineItems.map((item) => {
        if (!item || typeof item !== 'object') return item;
        const { verified: _discardVerified, Verified: _discardVerifiedUpper, ...rest } = item;
        return { ...rest };
      });

      const lineLookup = new Map();
      sanitizedLineItems.forEach((item) => {
        if (!item || typeof item !== 'object') {
          return;
        }
        const statementKey = (item.statement || '').toString().toLowerCase().trim();
        const labelKey = (item.lineItem || item['Line Item'] || '').toString().toLowerCase().trim();
        if (!statementKey || !labelKey) {
          return;
        }
        const key = `${statementKey}||${labelKey}`;
        if (!lineLookup.has(key)) {
          lineLookup.set(key, item);
        }
      });

      const initialManualEntries = {};
      sopEntries.forEach((entry, index) => {
        if (!entry || typeof entry !== 'object') {
          return;
        }
        const statementKey = entry.statement ? entry.statement.toString().toLowerCase().trim() : '';
        const lineKey = entry.sourceLine ? entry.sourceLine.toString().toLowerCase().trim() : '';
        const lookupKey = statementKey && lineKey ? `${statementKey}||${lineKey}` : '';
        if (lookupKey && lineLookup.has(lookupKey)) {
          const targetRow = lineLookup.get(lookupKey);
          if (targetRow && !targetRow.classification && entry.metric) {
            targetRow.classification = entry.metric;
          }
          return;
        }
        const hasDetails = [entry.statement, entry.column, entry.sourceLine, entry.value].some((field) => {
          if (field === null || typeof field === 'undefined') {
            return false;
          }
          const text = field.toString().trim();
          return Boolean(text && text !== '-');
        });
        if (!hasDetails) {
          return;
        }
        if (!initialManualEntries[entry.metric]) {
          initialManualEntries[entry.metric] = [];
        }
        initialManualEntries[entry.metric].push({
          id: `seed-${index}`,
          statement: entry.statement || '',
          lineItem: entry.sourceLine || '',
          column: entry.column || '',
          value: entry.value === '-' ? '' : (entry.value || ''),
          calculation: Array.isArray(entry.calculation)
            ? entry.calculation.map((step) => ({
              ...createEmptyCalculationStep(),
              ...step,
              operator: step?.operator || '+',
              statement: (step?.statement || '').trim(),
              lineItem: (step?.lineItem || '').trim(),
              column: (step?.column || '').trim(),
              constant: (step?.constant || '').trim(),
            }))
            : [],
        });
      });

      const statementsFromResponse = Array.from(new Set(sanitizedLineItems.map((item) => item?.statement))).filter(Boolean);
      const emptyValues = nextValueColumns.reduce((acc, column) => ({ ...acc, [column]: '' }), {});

      setPdfName(response.data.pdfName);
      setPdfBase64(response.data.pdfBase64);
      setLineItems(sanitizedLineItems);
      setCandidateMetrics(candidateMetricList);
      setValueColumns(nextValueColumns);
      setStatementMultiplierApplied({});
      setSopSummary(sopEntries);
      setSopMetadata(response.data.sopMetadata || { latestColumns: {} });
      setManualSopEntries(initialManualEntries);
      setExpandedSopMetrics({});
      setBreakdownDrafts({});
      setEditingSopMetric(null);
      setSopEditDraft(buildEmptySopEditDraft());
      setVerifiedStatements(() => {
        const map = {};
        statementsFromResponse.forEach((statement) => {
          map[statement] = false;
        });
        return map;
      });
      setShowManualEntry(false);
      setManualRow({
        statement: statementsFromResponse[0] || '',
        lineItem: '',
        values: emptyValues,
      });
      setPdfZoom(1);
      setActiveWorkspaceTab('statements');

      successMessage = 'Extraction complete. Review each statement and mark it as reviewed when done.';
    } catch (err) {
      console.error(err);
      const errorMessage = err?.response?.data?.error || 'Failed to process PDF. Please try again.';
      if (loadingIntervalRef.current) {
        clearInterval(loadingIntervalRef.current);
        loadingIntervalRef.current = null;
      }
      setError(errorMessage);
      setStatus({ type: 'error', message: errorMessage });
      setSopSummary(buildEmptySopSummary());
      setSopMetadata({ latestColumns: {} });
      setEditingSopMetric(null);
      setSopEditDraft(buildEmptySopEditDraft());
    } finally {
      setLoading(false);
      if (successMessage) {
        setStatus({ type: 'success', message: successMessage });
      }
      if (event.target) {
        event.target.value = '';
      }
    }
  };

  const handleFinalize = () => {
    if (!totalRows) {
      setStatus({ type: 'warning', message: 'No line items available to finalise.' });
      return;
    }
    if (pendingStatements > 0) {
      setStatus({ type: 'error', message: 'Please mark every statement as reviewed before finalising.' });
      return;
    }
    setQcComplete(true);
    setStatus({ type: 'success', message: 'QC complete. You may now export the verified dataset.' });
  };

  const handleStatementVerify = () => {
    if (!activeStatement) {
      return;
    }
    setVerifiedStatements((prev) => ({
      ...prev,
      [activeStatement]: true,
    }));
    setStatus({ type: 'success', message: `${activeStatement} marked as reviewed.` });
    setQcComplete(false);
  };

  const handleDownload = () => {
    if (!qcComplete) return;

    const workbook = XLSX.utils.book_new();
    const preferredOrder = [
      'Profit or Loss',
      'Comprehensive Income',
      'Financial Position',
      'Changes in Equity',
      'Cash Flows',
    ];

    const grouped = lineItems.reduce((acc, item) => {
      const key = item.statement || 'Unassigned Statement';
      if (!acc[key]) {
        acc[key] = [];
      }
      acc[key].push(item);
      return acc;
    }, {});

    const orderedStatements = [
      ...preferredOrder,
      ...Object.keys(grouped).filter((name) => !preferredOrder.includes(name)),
    ].filter((name, index, self) => name && self.indexOf(name) === index);

            const sanitizeSheetName = (name) => {
      const base = (name || 'Sheet');
      const withoutCommonInvalids = base.replace(/[\/?*]/g, '');
      const cleaned = withoutCommonInvalids.replace(/\[|\]/g, '').slice(0, 31);
      return cleaned || 'Sheet';
    };
    if (!orderedStatements.length) {
    const emptySheet = XLSX.utils.json_to_sheet([{
        'Line Item': '',
      }]);
      emptySheet['!cols'] = [{ wch: 35 }];
      const headerCell = XLSX.utils.encode_cell({ c: 0, r: 0 });
      if (emptySheet[headerCell]) {
        emptySheet[headerCell].s = {
          fill: { patternType: 'solid', fgColor: { rgb: 'FFFF00' } },
          font: { bold: true },
        };
      }
      XLSX.utils.book_append_sheet(workbook, emptySheet, 'Line_Items');
    } else {
      orderedStatements.forEach((statementName) => {
        const rows = grouped[statementName] || [];
        const relevantColumns = valueColumns.filter((column) => (
          rows.some((item) => {
            const raw = item[column];
            if (typeof raw === 'number') return true;
            return typeof raw === 'string' && raw.trim() !== '';
          })
        ));

        const headerOrder = [
          'Line Item',
          ...relevantColumns,
        ];

        const sheetRows = rows.map((item) => {
          const record = {
            'Line Item': item.lineItem,
          };

          relevantColumns.forEach((column) => {
            record[column] = item[column] ?? '';
          });

          return record;
        });

        const headerTemplate = headerOrder.reduce((acc, key) => {
          acc[key] = '';
          return acc;
        }, {});

        const worksheet = XLSX.utils.json_to_sheet(
          sheetRows.length ? sheetRows : [headerTemplate],
          { header: headerOrder },
        );

        if (!worksheet['!cols']) {
          worksheet['!cols'] = [];
        }
        headerOrder.forEach((_, idx) => {
          worksheet['!cols'][idx] = { wch: idx === 0 ? 35 : 18 };
          const cellAddress = XLSX.utils.encode_cell({ c: idx, r: 0 });
          if (!worksheet[cellAddress]) {
            worksheet[cellAddress] = { v: headerOrder[idx], t: 's' };
          }
          worksheet[cellAddress].s = {
            fill: { patternType: 'solid', fgColor: { rgb: 'FFFF00' } },
            font: { bold: true },
          };
        });

        XLSX.utils.book_append_sheet(
          workbook,
          worksheet,
          sanitizeSheetName(statementName || 'Statement'),
        );
      });
    }

    const sopSheetRows = displayedSopSummary.map((row) => ({
      Metric: row.metric,
      'Latest Quarter': row.value ?? '-',
      Statement: row.statement || '',
      'Source Column': row.column || '',
      'Source Line Item': row.sourceLine || '',
    }));
    const sopSheet = XLSX.utils.json_to_sheet(
      sopSheetRows.length ? sopSheetRows : [{ Metric: 'No SOP metrics available', 'Latest Quarter': '-' }],
    );
    sopSheet['!cols'] = [
      { wch: 35 },
      { wch: 18 },
      { wch: 24 },
      { wch: 24 },
      { wch: 28 },
    ];
    sopSheet['!cols'].forEach((_, idx) => {
      const cellAddress = XLSX.utils.encode_cell({ c: idx, r: 0 });
      if (sopSheet[cellAddress]) {
        sopSheet[cellAddress].s = {
          fill: { patternType: 'solid', fgColor: { rgb: 'FFFF00' } },
          font: { bold: true },
        };
      }
    });
    XLSX.utils.book_append_sheet(workbook, sopSheet, 'SOP_Summary');

  const suggestedName = pdfName
    ? `${pdfName.replace(/\.pdf$/i, '')}_verified.xlsx`
    : 'verified_line_items.xlsx';

  XLSX.writeFile(workbook, suggestedName, { cellStyles: true });
};

  const hasPdf = Boolean(pdfBase64);
  const reviewedCount = Object.values(verifiedStatements).filter(Boolean).length;
  const reviewProgressPercent = totalStatements ? Math.round((reviewedCount / totalStatements) * 100) : 0;
  const workflowSteps = [
    {
      title: 'Upload PDF',
      description: 'Select a quarterly Group/Consolidated report to begin.',
      status: hasPdf ? 'Complete' : loading ? 'In Progress' : 'Pending',
    },
    {
      title: 'Review Statements',
      description: 'Compare each line item with the source PDF and mark statements as reviewed.',
      status: reviewedCount === totalStatements && totalStatements > 0 ? 'Complete'
        : hasPdf ? 'In Progress' : 'Locked',
    },
    {
      title: 'Finalize & Export',
      description: 'Finalize QC to unlock the export options and download results.',
      status: qcComplete ? 'Complete'
        : reviewedCount === totalStatements && totalStatements > 0 ? 'Ready'
        : 'Locked',
    },
  ];

  const tabAvailability = {
    overview: true,
    statements: hasPdf,
    sop: lineItems.length > 0,
    exports: hasPdf,
  };

  const renderOverviewTab = () => (
    <div className="tab-panel-body overview-tab">
      <div className="overview-tab-grid">
        <div className="overview-card">
          <h3>Workflow Overview</h3>
          <ol>
            <li>Upload a financial PDF and run AI extraction.</li>
            <li>Review every line item against the source PDF.</li>
            <li>Adjust values where needed and mark each statement as reviewed.</li>
            <li>Finalize QC to unlock the export button.</li>
          </ol>
        </div>
        <div className="overview-card metrics">
          <h3>QC Snapshot</h3>
          <div className="overview-metrics">
            <div>
              <span className="metric-label">Line Items</span>
              <span className="metric-value">{totalRows}</span>
            </div>
            <div>
              <span className="metric-label">Statements</span>
              <span className="metric-value">{totalStatements}</span>
            </div>
            <div>
              <span className="metric-label">Reviewed</span>
              <span className="metric-value">{reviewedCount}</span>
            </div>
            <div>
              <span className="metric-label">Pending</span>
              <span className="metric-value">{pendingStatements}</span>
            </div>
          </div>
          <div className="progress-bar">
            <div
              className="progress-bar-fill"
              style={{ width: `${reviewProgressPercent}%` }}
            />
          </div>
          <p className="progress-caption">
            {totalStatements
              ? `${reviewProgressPercent}% of statements reviewed`
              : 'Upload a PDF to begin the workflow.'}
          </p>
        </div>
      </div>
    </div>
  );

  const renderStatementsTab = () => {
    if (!lineItems.length) {
    return (
      <div className="tab-placeholder">
        Upload a PDF to populate the statements workspace.
      </div>
    );
    }

    const multiplierApplied = Boolean(statementMultiplierApplied?.[activeStatement]);
    const bulkMetricTrimmed = typeof bulkClassificationMetric === 'string'
      ? bulkClassificationMetric.trim()
      : '';
    const canApplyBulkClassification = selectedRowCount > 0 && Boolean(bulkMetricTrimmed);
    const hasBulkSelection = selectedRowCount > 0;

    return (
      <div className="tab-panel-body statements-tab">
        <div className="tab-header">
          <div>
            <h3>Statement Review</h3>
            <p>
              Compare each line item with the PDF on the left, adjust values where needed, and mark the statement as reviewed.
            </p>
          </div>
          <div className="tab-actions">
            <button
              type="button"
              className="secondary-button"
              onClick={openManualEntry}
              disabled={showManualEntry}
            >
              Add Manual Line Item
            </button>
          </div>
        </div>
        <div className="panel-card statement-hint">
          <p>
            If a row is missing from the extraction, use <strong>Add Manual Line Item</strong> to capture it here.
          </p>
          <p>
            Remember to review each statement and click <strong>Mark Statement Reviewed</strong> once the values match the source PDF.
          </p>
          <p>
            Use the <strong>SOP Metric</strong> column to map each row to your SOP categories so the summary breakdown stays in sync.
          </p>
        </div>
        {showManualEntry && (
          <div className="manual-entry-card panel-card">
            <div className="manual-entry-grid">
              <label>
                <span>Statement</span>
                <input
                  type="text"
                  value={manualRow.statement}
                  onChange={(event) => handleManualFieldChange('statement', event.target.value)}
                  placeholder="e.g. Statement of Profit or Loss"
                />
              </label>
              <label>
                <span>Line Item</span>
                <input
                  type="text"
                  value={manualRow.lineItem}
                  onChange={(event) => handleManualFieldChange('lineItem', event.target.value)}
                  placeholder="Enter line item description"
                />
              </label>
            </div>
            <div className="manual-entry-values">
              {manualEntryColumns.length ? (
                manualEntryColumns.map((column) => (
                  <label key={column}>
                    <span>{column}</span>
                    <input
                      type="text"
                      value={manualRow.values?.[column] ?? ''}
                      onChange={(event) => handleManualValueInput(column, event.target.value)}
                      placeholder={`Value for ${column}`}
                    />
                  </label>
                ))
              ) : (
                <p className="manual-entry-empty">No numeric columns detected. You can still add the line item now and edit values later.</p>
              )}
            </div>
            <div className="manual-entry-actions">
              <button type="button" className="secondary-button" onClick={handleManualRowCancel}>Cancel</button>
              <button type="button" className="finalize-button" onClick={handleManualRowSubmit}>Insert Row</button>
            </div>
          </div>
        )}
        <div className="statement-nav panel-card">
          {statements.length > 0 ? (
            <>
              <div className="statement-tabs">
                {statements.map((statement) => {
                  const totalForStatement = statementTotalsMap[statement] || 0;
                  const statementReviewed = !!verifiedStatements[statement];
                  return (
                    <button
                      key={statement}
                      type="button"
                      className={`statement-tab${statement === activeStatement ? ' active' : ''}${statementReviewed ? ' reviewed' : ''}`}
                      onClick={() => setActiveStatement(statement)}
                    >
                      <span>{statement}</span>
                      <span className="statement-tab-count">
                        {totalForStatement}{statementReviewed ? ' \u2713' : ''}
                      </span>
                    </button>
                  );
                })}
              </div>
              <div className="statement-summary">
                {activeStatement ? (
                  <>
                    <span className="summary-label">Active Statement:</span>
                    <span className="summary-value">{activeStatement}</span>
                  </>
                ) : (
                  <span className="summary-label">Select a statement tab to review line items.</span>
                )}
              </div>
              {activeStatement && (
                <div className="statement-controls">
                  <button
                    type="button"
                    className={activeStatementVerified ? 'statement-verify-button verified' : 'statement-verify-button'}
                    onClick={handleStatementVerify}
                    disabled={activeStatementVerified}
                  >
                    {activeStatementVerified ? 'Statement Reviewed' : 'Mark Statement Reviewed'}
                  </button>
                  {activeStatementVerified && <span className="statement-status-badge">Reviewed</span>}
                </div>
              )}
            </>
          ) : (
            <div className="statement-empty">
              Upload a PDF or add a manual line item to create your first statement.
            </div>
          )}
        </div>
        {activeStatement && (
          <div className="statement-tools panel-card">
            <div className="statement-tools-buttons">
              <button
                type="button"
                className="secondary-button"
                onClick={handleStatementAddZeros}
                disabled={!statementValueColumns.length}
              >
                {multiplierApplied ? 'x1,000 Applied' : 'Apply x1,000 to Values'}
              </button>
              <button
                type="button"
                className="secondary-button"
                onClick={handleStatementMakePositive}
                disabled={!statementValueColumns.length}
              >
                Convert Negatives to Positive
              </button>
            </div>
            <span className="statement-tools-hint">
              These actions affect only the active statement.
              {multiplierApplied ? ' The x1,000 multiplier has already been applied.' : ''}
            </span>
            <div className="statement-tools-bulk">
              <div className="statement-tools-bulk-header">
                <span>Bulk classify selected rows</span>
                <span className="statement-tools-bulk-count">
                  {selectedRowCount} selected
                </span>
              </div>
              <div className="statement-tools-bulk-controls">
                <select
                  value={bulkClassificationMetric}
                  onChange={(event) => setBulkClassificationMetric(event.target.value)}
                  aria-label="Bulk SOP metric selection"
                >
                  <option value="">Select SOP metric</option>
                  {sopMetricOptions.map((metric) => (
                    <option key={metric} value={metric}>{metric}</option>
                  ))}
                </select>
                <button
                  type="button"
                  className="secondary-button"
                  onClick={handleBulkClassificationApply}
                  disabled={!canApplyBulkClassification}
                >
                  Apply to Selected
                </button>
                <button
                  type="button"
                  className="text-button"
                  onClick={handleClearRowSelection}
                  disabled={!hasBulkSelection}
                >
                  Clear Selection
                </button>
              </div>
            </div>
          </div>
        )}
        <div className="statement-table panel-card">
          {visibleItems.length ? (
            <div className="table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th className="select-column">
                      <input
                        type="checkbox"
                        ref={selectAllCheckboxRef}
                        onChange={(event) => handleSelectAllVisibleRows(event.target.checked)}
                        aria-label="Select all rows for this statement"
                        disabled={!visibleItems.length}
                      />
                    </th>
                    <th>Line Item</th>
                    <th className="metric-column-header">SOP Metric</th>
                    {statementValueColumns.map((column) => (
                      <th key={column}>
                        <div className="column-header">
                          <span>{column}</span>
                          <button
                            type="button"
                            className="column-remove-button"
                            onClick={() => handleRemoveColumn(column)}
                            title={`Remove column ${column}`}
                          >
                            Remove
                          </button>
                        </div>
                      </th>
                    ))}
                    <th className="actions-header">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleItems.map((row) => {
                    const isSelected = selectedRowIds.has(row.rowId);
                    const rowLabel = row.lineItem || row['Line Item'] || 'Row';
                    return (
                      <tr key={row.rowId} className={isSelected ? 'selected-row' : ''}>
                        <td className="select-cell">
                          <input
                            type="checkbox"
                            checked={isSelected}
                            onChange={(event) => handleRowSelectionToggle(row.rowId, event.target.checked)}
                            aria-label={`Select ${rowLabel}`}
                          />
                        </td>
                        <td>{rowLabel}</td>
                        <td className="classification-cell">
                          <select
                            value={row.classification ?? ''}
                            onChange={(event) => handleClassificationChange(row.rowId, event.target.value)}
                          >
                            <option value="">Unassigned</option>
                            {sopMetricOptions.map((metric) => (
                              <option key={metric} value={metric}>{metric}</option>
                            ))}
                          </select>
                        </td>
                        {statementValueColumns.map((column) => (
                          <td key={column} className="value-cell">
                            <input
                              type="text"
                              value={row[column] ?? ''}
                              onChange={(event) => {
                                const { value } = event.target;
                                handleValueChange(row.rowId, column, value);
                              }}
                            />
                          </td>
                        ))}
                        <td className="row-actions">
                          <button
                            type="button"
                            className="row-action-button"
                            onClick={() => handleDeleteRow(row.rowId)}
                          >
                            Delete Row
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="placeholder">No line items available for the selected statement.</div>
          )}
        </div>
      </div>
    );
  };

  const renderSopTab = () => {
    if (!lineItems.length) {
      return (
        <div className="tab-placeholder">
        Upload a PDF to populate the SOP summary metrics.
        </div>
      );
    }

    return (
      <div className="tab-panel-body sop-tab">
        <div className="panel-card sop-summary-card">
          <div className="sop-card-header">
            <h3>SOP Summary</h3>
            <p>Review or edit the derived metrics for the Statement of Performance, and open the breakdown to inspect linked rows or add manual adjustments.</p>
          </div>
          <div className="sop-summary-table-wrapper">
            <table className="sop-summary-table">
              <thead>
                <tr>
                  <th>Metric</th>
                  <th>Latest Quarter</th>
                  <th>Source</th>
                  <th className="sop-actions-column">Actions</th>
                </tr>
              </thead>
              <tbody>
                {displayedSopSummary.map((row) => {
                  const sourceParts = [row.statement, row.column, row.sourceLine]
                    .map((part) => (part || '').trim())
                    .filter((part) => part);
                  const sourceText = sourceParts.length ? sourceParts.join(' - ') : '-';
                  const isEditing = editingSopMetric === row.metric;
                  const isExpanded = Boolean(expandedSopMetrics?.[row.metric]);
                  const linkedRows = lineItemBreakdown[row.metric] || [];
                  const manualEntriesForMetric = manualSopEntries[row.metric] || [];
                  const breakdownDraft = breakdownDrafts[row.metric] || {
                    statement: '',
                    lineItem: '',
                    calculation: [],
                  };
                  const metricSlug = (row.metric || 'metric').toString().replace(/[^a-zA-Z0-9]+/g, '-').toLowerCase();

                  return (
                    <Fragment key={row.metric}>
                      <tr className={`sop-summary-row${row.manual ? ' manual' : ''}`}>
                        <td>{row.metric}</td>
                        <td>
                          <div className="sop-value-display">
                            {row.value ?? '-'}
                            {row.manual && (
                              <span className="sop-manual-indicator">Manual</span>
                            )}
                          </div>
                        </td>
                        <td>{sourceText}</td>
                        <td className="sop-actions-cell">
                          <div className="sop-action-buttons">
                            <button
                              type="button"
                              className={`sop-action-button${isExpanded ? ' active' : ''}`}
                              onClick={() => toggleSopMetricExpansion(row.metric)}
                            >
                              {isExpanded ? 'Hide' : 'Breakdown'}
                            </button>
                            {isEditing ? (
                              <>
                                <button
                                  type="button"
                                  className="sop-action-button primary"
                                  onClick={handleSopEditSave}
                                >
                                  Save
                                </button>
                                <button
                                  type="button"
                                  className="sop-action-button"
                                  onClick={handleSopEditCancel}
                                >
                                  Cancel
                                </button>
                              </>
                            ) : (
                              <button
                                type="button"
                                className="sop-action-button"
                                onClick={() => handleSopEditStart(row.metric)}
                                disabled={Boolean(editingSopMetric) && !isEditing}
                              >
                                Edit
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                      {isEditing && (
                        <tr className="sop-edit-row">
                          <td colSpan={4}>
                            <div className="sop-edit-form">
                              <label>
                                <span>Value</span>
                                <input
                                  type="text"
                                  value={sopEditDraft.value}
                                  onChange={(event) => handleSopEditFieldChange('value', event.target.value)}
                                  placeholder="Leave blank for '-'"
                                />
                              </label>
                              <label>
                                <span>Statement</span>
                                <input
                                  type="text"
                                  value={sopEditDraft.statement}
                                  onChange={(event) => handleSopEditFieldChange('statement', event.target.value)}
                                />
                              </label>
                              <label>
                                <span>Column</span>
                                <input
                                  type="text"
                                  value={sopEditDraft.column}
                                  onChange={(event) => handleSopEditFieldChange('column', event.target.value)}
                                />
                              </label>
                              <label>
                                <span>Source Line</span>
                                <input
                                  type="text"
                                  value={sopEditDraft.sourceLine}
                                  onChange={(event) => handleSopEditFieldChange('sourceLine', event.target.value)}
                                />
                              </label>
                            </div>
                          </td>
                        </tr>
                      )}
                      {isExpanded && (
                        <tr className="sop-breakdown-row">
                          <td colSpan={4}>
                            <div className="sop-breakdown">
                              <div className="sop-breakdown-section">
                                <div className="sop-breakdown-section-header">
                                  <h4>Linked line items</h4>
                                  <span>{linkedRows.length} linked</span>
                                </div>
                                {linkedRows.length ? (
                                  <ul className="sop-breakdown-list">
                                    {linkedRows.map((item) => {
                                      const valueEntries = Object.entries(item.values || {});
                                      return (
                                        <li key={item.rowId} className="sop-breakdown-list-item">
                                          <div className="sop-breakdown-item-header">
                                            <span className="sop-breakdown-title">
                                              {(item.statement || 'Statement unknown')}
                                              {' · '}
                                              {(item.lineItem || 'Unnamed line item')}
                                            </span>
                                            <span className="sop-breakdown-tag">Linked</span>
                                          </div>
                                          <div className="sop-breakdown-values">
                                            {valueEntries.length ? (
                                              valueEntries.map(([column, value]) => (
                                                <span key={column}>{column}: {value}</span>
                                              ))
                                            ) : (
                                              <span className="sop-breakdown-empty">No numeric values captured.</span>
                                            )}
                                          </div>
                                        </li>
                                      );
                                    })}
                                  </ul>
                                ) : (
                                  <p className="sop-breakdown-empty">
                                    No linked rows yet. Use the SOP Metric dropdown in the statements tab to link line items.
                                  </p>
                                )}
                              </div>
                              <div className="sop-breakdown-section">
                                <div className="sop-breakdown-section-header">
                                  <h4>Manual adjustments</h4>
                                  <span>{manualEntriesForMetric.length} manual</span>
                                </div>
                                {manualEntriesForMetric.length ? (
                                  <div className="sop-breakdown-manual-list">
                                    {manualEntriesForMetric.map((entry, entryIndex) => {
                                      const metricSlug = (row.metric || 'metric').toString().replace(/[^a-zA-Z0-9]+/g, '-').toLowerCase();
                                      const entrySuffix = `${metricSlug}-${entry.id || entryIndex}`;
                                      const entryStatementListId = `manual-statement-${entrySuffix}`;
                                      const entryLineItemListId = `manual-line-item-${entrySuffix}`;
                                      const calculationSteps = Array.isArray(entry.calculation) ? entry.calculation : [];
                                      const lineItemOptions = getLineItemSuggestions(entry.statement);
                                      return (
                                        <div key={entry.id} className="sop-breakdown-manual-item">
                                          <div className="sop-breakdown-manual-grid">
                                            <label>
                                              <span>Statement</span>
                                              <input
                                                type="text"
                                                list={entryStatementListId}
                                                value={entry.statement || ''}
                                                placeholder="Select statement"
                                                onChange={(event) => handleManualBreakdownValueChange(
                                                  row.metric,
                                                  entry.id,
                                                  'statement',
                                                  event.target.value,
                                                )}
                                              />
                                              <datalist id={entryStatementListId}>
                                                {statementSuggestions.map((statementName) => (
                                                  <option key={statementName} value={statementName} />
                                                ))}
                                              </datalist>
                                            </label>
                                            <label>
                                              <span>Line Item</span>
                                              <input
                                                type="text"
                                                list={entryLineItemListId}
                                                value={entry.lineItem || ''}
                                                placeholder="Select line item"
                                                onChange={(event) => handleManualBreakdownValueChange(
                                                  row.metric,
                                                  entry.id,
                                                  'lineItem',
                                                  event.target.value,
                                                )}
                                              />
                                              <datalist id={entryLineItemListId}>
                                                {lineItemOptions.map((itemName) => (
                                                  <option key={itemName} value={itemName} />
                                                ))}
                                              </datalist>
                                            </label>
                                          </div>
                                          <div className="sop-breakdown-calculation">
                                            <div className="sop-breakdown-calculation-header">
                                              <span>Calculation Steps</span>
                                              <button
                                                type="button"
                                                className="sop-breakdown-add-step"
                                                onClick={() => handleAddManualBreakdownCalculationStep(row.metric, entry.id)}
                                              >
                                                Add Step
                                              </button>
                                            </div>
                                            {calculationSteps.length ? (
                                              <div className="sop-breakdown-calculation-list">
                                                {calculationSteps.map((step, stepIndex) => {
                                                  const stepSuffix = `${entrySuffix}-${stepIndex}`;
                                                  const stepStatementListId = `manual-step-statement-${stepSuffix}`;
                                                  const stepLineItemListId = `manual-step-line-item-${stepSuffix}`;
                                                  const stepLineItemOptions = getLineItemSuggestions(step.statement || entry.statement);
                                                  return (
                                                    <div key={stepSuffix} className="sop-breakdown-calculation-row">
                                                      <label>
                                                        <span>Operator</span>
                                                        <select
                                                          value={step.operator || '+'}
                                                          onChange={(event) => handleManualBreakdownCalculationChange(
                                                            row.metric,
                                                            entry.id,
                                                            stepIndex,
                                                            'operator',
                                                            event.target.value,
                                                          )}
                                                        >
                                                          <option value="+">+</option>
                                                          <option value="-">-</option>
                                                          <option value="*">*</option>
                                                          <option value="/">/</option>
                                                        </select>
                                                      </label>
                                                      <label>
                                                        <span>Statement</span>
                                                        <input
                                                          type="text"
                                                          list={stepStatementListId}
                                                          value={step.statement || ''}
                                                          placeholder="Statement"
                                                          onChange={(event) => handleManualBreakdownCalculationChange(
                                                            row.metric,
                                                            entry.id,
                                                            stepIndex,
                                                            'statement',
                                                            event.target.value,
                                                          )}
                                                        />
                                                        <datalist id={stepStatementListId}>
                                                          {statementSuggestions.map((statementName) => (
                                                            <option key={statementName} value={statementName} />
                                                          ))}
                                                        </datalist>
                                                      </label>
                                                      <label>
                                                        <span>Line Item</span>
                                                        <input
                                                          type="text"
                                                          list={stepLineItemListId}
                                                          value={step.lineItem || ''}
                                                          placeholder="Line item"
                                                          onChange={(event) => handleManualBreakdownCalculationChange(
                                                            row.metric,
                                                            entry.id,
                                                            stepIndex,
                                                            'lineItem',
                                                            event.target.value,
                                                          )}
                                                        />
                                                        <datalist id={stepLineItemListId}>
                                                          {stepLineItemOptions.map((itemName) => (
                                                            <option key={itemName} value={itemName} />
                                                          ))}
                                                        </datalist>
                                                      </label>
                                                      <label>
                                                        <span>Constant/Value</span>
                                                        <input
                                                          type="text"
                                                          value={step.constant || ''}
                                                          placeholder="Optional number"
                                                          onChange={(event) => handleManualBreakdownCalculationChange(
                                                            row.metric,
                                                            entry.id,
                                                            stepIndex,
                                                            'constant',
                                                            event.target.value,
                                                          )}
                                                        />
                                                      </label>
                                                      <button
                                                        type="button"
                                                        className="sop-breakdown-remove-step"
                                                        onClick={() => handleRemoveManualBreakdownCalculationStep(row.metric, entry.id, stepIndex)}
                                                      >
                                                        Remove
                                                      </button>
                                                    </div>
                                                  );
                                                })}
                                              </div>
                                            ) : (
                                              <p className="sop-breakdown-empty muted">No calculation steps configured.</p>
                                            )}
                                          </div>
                                          <button
                                            type="button"
                                            className="sop-action-button danger"
                                            onClick={() => handleRemoveManualBreakdownEntry(row.metric, entry.id)}
                                          >
                                            Remove Entry
                                          </button>
                                        </div>
                                      );
                                    })}
                                </div>
                              ) : (
                                  <p className="sop-breakdown-empty">No manual adjustments recorded.</p>
                                )}
                                <div className="sop-breakdown-form">
                                  <div className="sop-breakdown-manual-grid">
                                    <label>
                                      <span>Statement</span>
                                      <input
                                        type="text"
                                        list={`draft-statement-${metricSlug}`}
                                        value={breakdownDraft.statement}
                                        placeholder="Select statement"
                                        onChange={(event) => handleManualBreakdownDraftChange(
                                          row.metric,
                                          'statement',
                                          event.target.value,
                                        )}
                                      />
                                      <datalist id={`draft-statement-${metricSlug}`}>
                                        {statementSuggestions.map((statementName) => (
                                          <option key={statementName} value={statementName} />
                                        ))}
                                      </datalist>
                                    </label>
                                    <label>
                                      <span>Line Item</span>
                                      <input
                                        type="text"
                                        list={`draft-line-item-${metricSlug}`}
                                        value={breakdownDraft.lineItem}
                                        placeholder="Select line item"
                                        onChange={(event) => handleManualBreakdownDraftChange(
                                          row.metric,
                                          'lineItem',
                                          event.target.value,
                                        )}
                                      />
                                      <datalist id={`draft-line-item-${metricSlug}`}>
                                        {getLineItemSuggestions(breakdownDraft.statement).map((itemName) => (
                                          <option key={itemName} value={itemName} />
                                        ))}
                                      </datalist>
                                    </label>
                                  </div>
                                  <div className="sop-breakdown-calculation">
                                    <div className="sop-breakdown-calculation-header">
                                      <span>Calculation Steps</span>
                                      <button
                                        type="button"
                                        className="sop-breakdown-add-step"
                                        onClick={() => handleAddManualBreakdownDraftStep(row.metric)}
                                      >
                                        Add Step
                                      </button>
                                    </div>
                                    {Array.isArray(breakdownDraft.calculation) && breakdownDraft.calculation.length ? (
                                      <div className="sop-breakdown-calculation-list">
                                        {breakdownDraft.calculation.map((step, stepIndex) => {
                                          const stepSuffix = `${metricSlug}-draft-${stepIndex}`;
                                          const stepStatementListId = `draft-step-statement-${stepSuffix}`;
                                          const stepLineItemListId = `draft-step-line-item-${stepSuffix}`;
                                          const stepLineItemOptions = getLineItemSuggestions(step.statement || breakdownDraft.statement);
                                          return (
                                            <div key={stepSuffix} className="sop-breakdown-calculation-row">
                                              <label>
                                                <span>Operator</span>
                                                <select
                                                  value={step.operator || '+'}
                                                  onChange={(event) => handleManualBreakdownDraftCalculationChange(
                                                    row.metric,
                                                    stepIndex,
                                                    'operator',
                                                    event.target.value,
                                                  )}
                                                >
                                                  <option value="+">+</option>
                                                  <option value="-">-</option>
                                                  <option value="*">*</option>
                                                  <option value="/">/</option>
                                                </select>
                                              </label>
                                              <label>
                                                <span>Statement</span>
                                                <input
                                                  type="text"
                                                  list={stepStatementListId}
                                                  value={step.statement || ''}
                                                  placeholder="Statement"
                                                  onChange={(event) => handleManualBreakdownDraftCalculationChange(
                                                    row.metric,
                                                    stepIndex,
                                                    'statement',
                                                    event.target.value,
                                                  )}
                                                />
                                                <datalist id={stepStatementListId}>
                                                  {statementSuggestions.map((statementName) => (
                                                    <option key={statementName} value={statementName} />
                                                  ))}
                                                </datalist>
                                              </label>
                                              <label>
                                                <span>Line Item</span>
                                                <input
                                                  type="text"
                                                  list={stepLineItemListId}
                                                  value={step.lineItem || ''}
                                                  placeholder="Line item"
                                                  onChange={(event) => handleManualBreakdownDraftCalculationChange(
                                                    row.metric,
                                                    stepIndex,
                                                    'lineItem',
                                                    event.target.value,
                                                  )}
                                                />
                                                <datalist id={stepLineItemListId}>
                                                  {stepLineItemOptions.map((itemName) => (
                                                    <option key={itemName} value={itemName} />
                                                  ))}
                                                </datalist>
                                              </label>
                                              <label>
                                                <span>Constant/Value</span>
                                                <input
                                                  type="text"
                                                  value={step.constant || ''}
                                                  placeholder="Optional number"
                                                  onChange={(event) => handleManualBreakdownDraftCalculationChange(
                                                    row.metric,
                                                    stepIndex,
                                                    'constant',
                                                    event.target.value,
                                                  )}
                                                />
                                              </label>
                                              <button
                                                type="button"
                                                className="sop-breakdown-remove-step"
                                                onClick={() => handleRemoveManualBreakdownDraftStep(row.metric, stepIndex)}
                                              >
                                                Remove
                                              </button>
                                            </div>
                                          );
                                        })}
                                      </div>
                                    ) : (
                                      <p className="sop-breakdown-empty muted">No calculation steps configured.</p>
                                    )}
                                  </div>
                                  <div className="sop-breakdown-form-actions">
                                    <button
                                      type="button"
                                      className="sop-action-button primary"
                                      onClick={() => handleAddManualBreakdownEntry(row.metric)}
                                    >
                                      Add Entry
                                    </button>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  };

  const renderExportsTab = () => {
    if (!lineItems.length) {
      return (
        <div className="tab-placeholder">
          Upload a PDF to unlock export tools.
        </div>
      );
    }

    const allStatementsReviewed = reviewedCount === totalStatements && totalStatements > 0;

    return (
      <div className="tab-panel-body exports-tab">
        <div className="panel-card export-card">
          <h3>Finalize & Export</h3>
          <p>Finalize QC once every statement is reviewed to enable the download.</p>
          <div className="export-actions">
            <button
              type="button"
              className="finalize-button"
              onClick={handleFinalize}
              disabled={loading || !totalRows}
            >
              Finalize Quality Control
            </button>
            <button
              type="button"
              className="primary-download-button"
              onClick={handleDownload}
              disabled={!qcComplete}
            >
              Download Verified Dataset (.xlsx)
            </button>
          </div>
          <div className="export-status">
            {qcComplete ? (
              <p className="export-ready">Quality control is complete. You can download the dataset now.</p>
            ) : (
              <p className="export-hint">
                {allStatementsReviewed
                  ? 'Finalize QC to enable the download.'
                  : 'Review every statement before finalizing.'}
              </p>
            )}
          </div>
        </div>
      </div>
    );
  };

  const renderActiveTab = () => {
    switch (activeWorkspaceTab) {
      case 'statements':
        return renderStatementsTab();
      case 'sop':
        return renderSopTab();
      case 'exports':
        return renderExportsTab();
      case 'overview':
      default:
        return renderOverviewTab();
    }
  };

  return (
    <div className="app-shell">
      {loading && (
        <div className="loading-overlay">
          <div className="loading-content">
            <div className="loading-spinner" />
            <p>{status?.message || 'Processing PDF...'}</p>
            <p className="loading-subtext">Large reports can take a couple of minutes. Please keep this tab open.</p>
          </div>
        </div>
      )}
      <header className="shell-header">
        <div className="shell-brand">
          <img
            src="https://www.valueit.io/wp-content/uploads/2024/06/Valueit.png"
            alt="Valueit logo"
            className="shell-logo"
          />
          <h1>Financial Data QC Workbench</h1>
        </div>
        <label className="upload-button">
          <svg
            className="upload-button-icon"
            viewBox="0 0 24 24"
            aria-hidden="true"
            focusable="false"
          >
            <path
              fill="currentColor"
              d="M12 3a1 1 0 0 1 .78.37l4 5a1 1 0 1 1-1.56 1.26L13 6.54V15a1 1 0 0 1-2 0V6.54L8.78 9.63a1 1 0 0 1-1.56-1.26l4-5A1 1 0 0 1 12 3zm-7 12a1 1 0 0 1 1 1v3h12v-3a1 1 0 1 1 2 0v3a3 3 0 0 1-3 3H8a3 3 0 0 1-3-3v-3a1 1 0 0 1 1-1z"
            />
          </svg>
          <span>Upload PDF</span>
          <input
            type="file"
            accept="application/pdf"
            onChange={handleFileChange}
            disabled={loading}
          />
        </label>
      </header>

      <section className="workflow-strip">
        {workflowSteps.map((step) => {
          const statusClass = step.status.toLowerCase().replace(/\s+/g, '-');
          return (
            <div key={step.title} className={`workflow-step ${statusClass}`}>
              <div className="workflow-step-header">
                <span className="workflow-status">{step.status}</span>
                <h3>{step.title}</h3>
              </div>
              <p>{step.description}</p>
            </div>
          );
        })}
      </section>

      {error && (
        <div className="alert error">
          {error}
        </div>
      )}

      {status && (
        <div className={`alert ${status.type}`}>
          {status.message}
        </div>
      )}

      <div className="workspace-grid">
        <section className="workspace-left">
          <div className="workspace-card pdf-card">
            <div className="pdf-card-header">
              <div>
                <h2>Source PDF</h2>
                <span className="file-name">{pdfName || 'No file uploaded'}</span>
              </div>
              <div className="pdf-controls">
                <span>Zoom</span>
                <div className="pdf-controls-buttons">
                  <button
                    type="button"
                    onClick={() => setPdfZoom((value) => Math.max(0.5, Number((value - 0.1).toFixed(2))))}
                    disabled={!iframeSrc}
                  >
                    -
                  </button>
                  <span className="zoom-value">{Math.round(pdfZoom * 100)}%</span>
                  <button
                    type="button"
                    onClick={() => setPdfZoom((value) => Math.min(3, Number((value + 0.1).toFixed(2))))}
                    disabled={!iframeSrc}
                  >
                    +
                  </button>
                  <button
                    type="button"
                    onClick={() => setPdfZoom(1)}
                    disabled={!iframeSrc || pdfZoom === 1}
                  >
                    Reset
                  </button>
                </div>
              </div>
            </div>
            {iframeSrc ? (
              <div className="pdf-frame-scroll">
                <div
                  className="pdf-frame-wrapper"
                  style={{
                    transform: `scale(${pdfZoom})`,
                    transformOrigin: 'top left',
                  }}
                >
                  <iframe title="Uploaded PDF" src={iframeSrc} className="pdf-frame" />
                </div>
              </div>
            ) : (
              <div className="placeholder">Upload a financial PDF to begin.</div>
            )}
          </div>
        </section>
        <section className="workspace-right">
          <div className="right-tab-bar">
            {WORKSPACE_TABS.map((tab) => {
              const disabled = !tabAvailability[tab.id];
              return (
                <button
                  key={tab.id}
                  type="button"
                  className={`right-tab${activeWorkspaceTab === tab.id ? ' active' : ''}`}
                  onClick={() => setActiveWorkspaceTab(tab.id)}
                  disabled={disabled}
                >
                  {tab.label}
                </button>
              );
            })}
          </div>
          <div className="right-tab-panel">
            {renderActiveTab()}
          </div>
        </section>
      </div>
    </div>
  );
}

export default App;



