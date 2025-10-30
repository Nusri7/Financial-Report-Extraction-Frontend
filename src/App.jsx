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
  const [, setCandidateMetrics] = useState([]);
  const [qcComplete, setQcComplete] = useState(false);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState(null);
  const [error, setError] = useState('');
  const [showManualEntry, setShowManualEntry] = useState(false);
  const [manualRow, setManualRow] = useState({ statement: '', lineItem: '', values: {} });
  const [verifiedStatements, setVerifiedStatements] = useState({});
  const [activeWorkspaceTab, setActiveWorkspaceTab] = useState('overview');
  const [pdfZoom, setPdfZoom] = useState(1);
  const [editingSopMetric, setEditingSopMetric] = useState(null);
  const [sopEditDraft, setSopEditDraft] = useState(() => buildEmptySopEditDraft());
  const loadingIntervalRef = useRef(null);

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
  
  const statements = useMemo(() => (
    Array.from(new Set(lineItems.map((item) => item.statement))).filter(Boolean)
  ), [lineItems]);

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

  const iframeSrc = useMemo(() => {
    if (!pdfBase64) return '';
    return `data:application/pdf;base64,${pdfBase64}`;
  }, [pdfBase64]);

  const handleValueChange = (rowId, columnName, value) => {
    let affectedStatement = null;
    let affectedLineItem = '';

    setLineItems((items) => items.map((item) => {
      if (item.rowId !== rowId) {
        return item;
      }
      affectedStatement = item.statement;
      affectedLineItem = item.lineItem || item['Line Item'] || '';
      return { ...item, [columnName]: value };
    }));

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
            value: normaliseSopValue(value),
          };
        }
        return entry;
      }));
    }

    setQcComplete(false);
    setStatus(null);
  };

  const handleSopEditStart = (metric) => {
    const currentEntry = sopSummary.find((entry) => entry.metric === metric);
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

      const incomingLineItems = response.data.lineItems || [];
      const sanitizedLineItems = incomingLineItems.map((item) => {
        if (!item || typeof item !== 'object') return item;
        const { verified: _discardVerified, Verified: _discardVerifiedUpper, ...rest } = item;
        return { ...rest };
      });
      const statementsFromResponse = Array.from(new Set(sanitizedLineItems.map((item) => item?.statement))).filter(Boolean);
      const emptyValues = nextValueColumns.reduce((acc, column) => ({ ...acc, [column]: '' }), {});

      setPdfName(response.data.pdfName);
      setPdfBase64(response.data.pdfBase64);
      setLineItems(sanitizedLineItems);
      setCandidateMetrics(response.data.candidateMetrics || []);
      setValueColumns(nextValueColumns);
      setSopSummary(sopEntries);
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

    const sopSheetRows = sopSummary.map((row) => ({
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
        <div className="statement-table panel-card">
          {visibleItems.length ? (
            <div className="table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th>Line Item</th>
                    {statementValueColumns.map((column) => (
                      <th key={column}>{column}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {visibleItems.map((row) => (
                    <tr key={row.rowId}>
                      <td>{row.lineItem}</td>
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
                    </tr>
                  ))}
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
            <p>Review or edit the derived metrics for the Statement of Performance.</p>
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
                {sopSummary.map((row) => {
                  const sourceParts = [row.statement, row.column, row.sourceLine]
                    .map((part) => (part || '').trim())
                    .filter((part) => part);
                  const sourceText = sourceParts.length ? sourceParts.join(' - ') : '-';
                  const isEditing = editingSopMetric === row.metric;

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
