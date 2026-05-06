import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { analyzeWorkbook, downloadCleanedWorkbook, getPreview, getSheets } from './api.js';
import './styles.css';

const columnFields = [
  { key: 'entity_column', label: 'Primary column', required: true },
  { key: 'year_column', label: 'Year/date column', required: false },
  { key: 'type_column', label: 'Type/category column', required: false },
  { key: 'amount_column', label: 'Amount column', required: false },
  { key: 'unit_column', label: 'Unit column', required: false },
  { key: 'notes_column_1', label: 'Notes column 1', required: false },
  { key: 'notes_column_2', label: 'Notes column 2', required: false },
];

function App() {
  const [file, setFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [columns, setColumns] = useState([]);
  const [previewRows, setPreviewRows] = useState([]);
  const [columnConfig, setColumnConfig] = useState({});
  const [fuzzyThreshold, setFuzzyThreshold] = useState(88);
  const [minManualScore, setMinManualScore] = useState(0.75);
  const [analysis, setAnalysis] = useState(null);
  const [statusMessage, setStatusMessage] = useState('');
  const [busyAction, setBusyAction] = useState('');

  const previewHeaders = useMemo(() => (previewRows.length > 0 ? Object.keys(previewRows[0]) : columns), [previewRows, columns]);

  const setError = (error) => {
    const detail = typeof error?.message === 'string' ? error.message : 'An unexpected error occurred.';
    setStatusMessage(detail);
  };

  const handleFileChange = async (event) => {
    const nextFile = event.target.files?.[0] || null;
    setFile(nextFile);
    setSheetNames([]);
    setSelectedSheet('');
    setColumns([]);
    setPreviewRows([]);
    setColumnConfig({});
    setAnalysis(null);
    setStatusMessage('');

    if (!nextFile) return;

    try {
      setBusyAction('Loading sheets...');
      const data = await getSheets(nextFile);
      setSheetNames(data.sheet_names || []);
      setStatusMessage('Workbook loaded. Choose a sheet to continue.');
    } catch (error) {
      setError(error);
    } finally {
      setBusyAction('');
    }
  };

  const handleSheetSelect = async (event) => {
    const nextSheet = event.target.value;
    setSelectedSheet(nextSheet);
    setAnalysis(null);

    if (!nextSheet || !file) return;

    try {
      setBusyAction('Loading sheet preview...');
      const data = await getPreview(file, nextSheet);
      setColumns(data.columns || []);
      setPreviewRows(data.preview_rows || []);
      setColumnConfig(data.recommended_column_config || {});
      setStatusMessage('Preview loaded. Confirm configuration to run analysis.');
    } catch (error) {
      setError(error);
    } finally {
      setBusyAction('');
    }
  };

  const handleConfigChange = (key, value) => {
    setColumnConfig((prev) => ({ ...prev, [key]: value || null }));
  };

  const handleAnalyze = async () => {
    if (!file || !selectedSheet || !columnConfig.entity_column) {
      setStatusMessage('Upload a workbook, select a sheet, and set Primary column before running analysis.');
      return;
    }

    try {
      setBusyAction('Running analysis...');
      const data = await analyzeWorkbook(file, selectedSheet, columnConfig, fuzzyThreshold, minManualScore);
      setAnalysis(data);
      setStatusMessage('Analysis complete. Review Step 3 and export when ready.');
    } catch (error) {
      setError(error);
    } finally {
      setBusyAction('');
    }
  };

  const handleExport = async () => {
    if (!file || !selectedSheet || !columnConfig.entity_column) {
      setStatusMessage('Complete steps 1 and 2 before exporting.');
      return;
    }

    try {
      setBusyAction('Preparing cleaned workbook...');
      const response = await downloadCleanedWorkbook(file, selectedSheet, columnConfig, fuzzyThreshold, minManualScore);
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `cleaned_${selectedSheet}.xlsx`;
      anchor.click();
      URL.revokeObjectURL(url);
      setStatusMessage('Cleaned workbook downloaded successfully.');
    } catch (error) {
      setError(error);
    } finally {
      setBusyAction('');
    }
  };

  return (
    <div className="page-shell">
      <main className="workbench">
        <header className="workbench-header">
          <h1>Duplicate Review Workbench</h1>
          <p>Upload a workbook, configure column mappings, run duplicate analysis, and export a cleaned workbook.</p>
        </header>

        <section className="card">
          <h2>Step 1: Upload</h2>
          <label className="field-label" htmlFor="workbook-file">Excel workbook</label>
          <input id="workbook-file" type="file" accept=".xlsx,.xls" onChange={handleFileChange} />
          {sheetNames.length > 0 && (
            <div className="field-group">
              <label className="field-label" htmlFor="sheet-select">Select sheet</label>
              <select id="sheet-select" value={selectedSheet} onChange={handleSheetSelect}>
                <option value="">Choose a sheet</option>
                {sheetNames.map((sheet) => (
                  <option key={sheet} value={sheet}>{sheet}</option>
                ))}
              </select>
            </div>
          )}
        </section>

        <section className="card">
          <h2>Step 2: Configure</h2>
          <div className="mapping-grid">
            {columnFields.map((field) => (
              <label key={field.key}>
                <span>{field.label}</span>
                <select
                  value={columnConfig[field.key] || ''}
                  onChange={(event) => handleConfigChange(field.key, event.target.value)}
                >
                  {!field.required && <option value="">None</option>}
                  {columns.map((column) => <option key={column} value={column}>{column}</option>)}
                </select>
              </label>
            ))}
          </div>
          {previewRows.length > 0 && (
            <div className="table-wrapper">
              <table>
                <thead>
                  <tr>{previewHeaders.map((header) => <th key={header}>{header}</th>)}</tr>
                </thead>
                <tbody>
                  {previewRows.map((row, index) => (
                    <tr key={index}>{previewHeaders.map((header) => <td key={header}>{String(row[header] ?? '')}</td>)}</tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </section>

        <section className="card">
          <h2>Step 3: Analyze</h2>
          <div className="threshold-row">
            <label>
              <span>Name-match strictness</span>
              <input type="number" min="0" max="100" value={fuzzyThreshold} onChange={(event) => setFuzzyThreshold(Number(event.target.value))} />
            </label>
            <label>
              <span>Overall evidence threshold</span>
              <input type="number" min="0" max="1" step="0.01" value={minManualScore} onChange={(event) => setMinManualScore(Number(event.target.value))} />
            </label>
          </div>
          <button className="primary-btn" onClick={handleAnalyze} disabled={Boolean(busyAction)}>Run analysis</button>

          {analysis?.summary && (
            <div className="summary-grid">
              <article><h3>Unique raw primary values</h3><p>{analysis.summary.unique_raw_primary_values}</p></article>
              <article><h3>Safe auto-groups</h3><p>{analysis.summary.safe_auto_groups}</p></article>
              <article><h3>Manual queue</h3><p>{analysis.summary.manual_queue}</p></article>
              <article><h3>Merged names now</h3><p>{analysis.summary.merged_names_now}</p></article>
            </div>
          )}

          {analysis?.auto_groups?.length > 0 && <DataTable title="Auto groups" rows={analysis.auto_groups} />}
          {analysis?.manual_queue?.length > 0 && <DataTable title="Manual queue" rows={analysis.manual_queue} />}
        </section>

        <section className="card export-card">
          <h2>Step 4: Export</h2>
          <p>Download the cleaned workbook generated with the current sheet, mappings, and thresholds.</p>
          <button className="primary-btn large" onClick={handleExport} disabled={Boolean(busyAction)}>Download cleaned workbook</button>
        </section>

        <footer className="status-line">{busyAction || statusMessage}</footer>
      </main>
    </div>
  );
}

function DataTable({ title, rows }) {
  const headers = Object.keys(rows[0] || {});
  return (
    <div className="table-block">
      <h3>{title}</h3>
      <div className="table-wrapper">
        <table>
          <thead><tr>{headers.map((h) => <th key={h}>{h}</th>)}</tr></thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr key={idx}>{headers.map((h) => <td key={h}>{String(row[h] ?? '')}</td>)}</tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

createRoot(document.getElementById('root')).render(<App />);
