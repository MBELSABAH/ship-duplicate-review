import { useMemo, useState } from 'react';
import { exportCleanWorkbook, fetchWorkbookPreview, fetchWorkbookSheets, runAnalyze } from './api';

const steps = ['Upload', 'Configure', 'Auto review', 'Manual review', 'Export'];
const friendlyMap = {
  entity_column: 'Primary value column',
  year_column: 'Year/date evidence column',
  type_column: 'Type/category evidence column',
  amount_column: 'Amount evidence column',
  unit_column: 'Unit evidence column',
  notes_column_1: 'Notes evidence column 1',
  notes_column_2: 'Notes evidence column 2',
};

export default function App() {
  const [step, setStep] = useState(0);
  const [file, setFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [sheetName, setSheetName] = useState('');
  const [columns, setColumns] = useState([]);
  const [previewRows, setPreviewRows] = useState([]);
  const [columnConfig, setColumnConfig] = useState({});
  const [analyzeResult, setAnalyzeResult] = useState(null);
  const [autoStatus, setAutoStatus] = useState({});
  const [manualDecisions, setManualDecisions] = useState({});
  const [manualIndex, setManualIndex] = useState(0);

  const manualQueue = analyzeResult?.manual_queue || [];
  const activeCandidate = manualQueue[manualIndex];

  const summary = useMemo(() => {
    const autos = Object.values(autoStatus);
    const manuals = Object.values(manualDecisions);
    return {
      acceptedAuto: autos.filter((v) => v === 'accepted').length,
      rejectedAuto: autos.filter((v) => v === 'rejected').length,
      merge: manuals.filter((v) => v === 'merge').length,
      keep: manuals.filter((v) => v === 'keep_separate').length,
      unsure: manuals.filter((v) => v === 'unsure').length,
    };
  }, [autoStatus, manualDecisions]);

  async function handleWorkbook(fileInput) {
    setFile(fileInput);
    const sheets = await fetchWorkbookSheets(fileInput);
    setSheetNames(sheets.sheet_names || []);
    setSheetName('');
    setColumns([]);
    setPreviewRows([]);
    setAnalyzeResult(null);
  }

  async function handleLoadSheet(nextSheet) {
    setSheetName(nextSheet);
    const preview = await fetchWorkbookPreview(file, nextSheet);
    setColumns(preview.columns || []);
    setPreviewRows((preview.preview_rows || []).slice(0, 6));
    setColumnConfig(preview.recommended_column_config || {});
  }

  async function handleAnalyze() {
    const data = await runAnalyze({ file, sheet_name: sheetName, column_config: columnConfig, auto_status: autoStatus, manual_decisions: manualDecisions });
    setAnalyzeResult(data);
    setManualIndex(0);
  }

  async function handleExport() {
    const blob = await exportCleanWorkbook({ file, sheet_name: sheetName, column_config: columnConfig, auto_status: autoStatus, manual_decisions: manualDecisions });
    const href = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = href;
    a.download = `cleaned_${sheetName}.xlsx`;
    a.click();
    URL.revokeObjectURL(href);
  }

  return (
    <div className="container">
      <h1>Duplicate Review Workflow</h1>
      <ol className="stepper">{steps.map((s, i) => <li key={s} className={i === step ? 'active' : i < step ? 'done' : ''}>{s}</li>)}</ol>

      {step === 0 && <section className="card"><h2>Upload</h2><input type="file" accept=".xlsx,.xls" onChange={(e) => e.target.files?.[0] && handleWorkbook(e.target.files[0])} />
        {file && <p>Workbook loaded: {file.name}</p>}
        <p>Sheet count: {sheetNames.length}</p>
        {sheetNames.length > 0 && <select value={sheetName} onChange={(e) => handleLoadSheet(e.target.value)}><option value="">Select sheet</option>{sheetNames.map((s) => <option key={s}>{s}</option>)}</select>}
      </section>}

      {step === 1 && <section className="card"><h2>Configure</h2>
        {Object.keys(friendlyMap).map((key) => <label key={key}>{friendlyMap[key]}<select value={columnConfig[key] || ''} onChange={(e) => setColumnConfig({ ...columnConfig, [key]: e.target.value || null })}><option value="">None</option>{columns.map((c) => <option key={c} value={c}>{c}</option>)}</select></label>)}
        <h3>Preview</h3><div className="table-wrap"><table><thead><tr>{columns.map((c) => <th key={c}>{c}</th>)}</tr></thead><tbody>{previewRows.map((row, i) => <tr key={i}>{columns.map((c) => <td key={c}>{String(row[c] ?? '')}</td>)}</tr>)}</tbody></table></div>
      </section>}

      {step === 2 && <section className="card"><h2>Auto review</h2><button onClick={handleAnalyze}>Run analyze</button>
        <div className="cards">{(analyzeResult?.auto_groups || []).map((g, idx) => {
          const key = g.group_id || `${g.canonical_name}-${idx}`;
          return <article className="mini-card" key={key}><h3>{g.canonical_name || 'Suggested canonical'}</h3><p>Member names: {(g.member_names || []).join(', ')}</p><p>Total rows: {g.total_rows}</p><p>Years: {(g.years || []).join(', ')}</p><p>Reason: {g.reason || 'Safe auto-group'}</p><p>Status: {autoStatus[key] || 'pending'}</p><div className="actions"><button onClick={() => setAutoStatus({ ...autoStatus, [key]: 'accepted' })}>Accept</button><button onClick={() => setAutoStatus({ ...autoStatus, [key]: 'rejected' })}>Reject</button><button onClick={() => { const n = { ...autoStatus }; delete n[key]; setAutoStatus(n); }}>Undo</button></div></article>;
        })}</div>
      </section>}

      {step === 3 && <section className="card"><h2>Manual review</h2>{activeCandidate ? <>
        <p>Candidate {manualIndex + 1} of {manualQueue.length}</p><p>Score: {activeCandidate.score}</p><p>Suggested canonical: {activeCandidate.suggested_canonical || 'N/A'}</p><p>Reasons: {(activeCandidate.reasons || []).join(', ') || 'N/A'}</p>
        <div className="compare"><div><h3>name_a</h3><p>{activeCandidate.name_a}</p></div><div><h3>name_b</h3><p>{activeCandidate.name_b}</p></div></div>
        <ul><li>Name similarity: {activeCandidate.score}</li><li>Years: {activeCandidate.year_overlap ?? 'N/A'}</li><li>Unit: {activeCandidate.unit_match ?? 'N/A'}</li><li>Type/category: {activeCandidate.type_match ?? 'N/A'}</li><li>Amount: {activeCandidate.amount_overlap ?? 'N/A'}</li></ul>
        <p>Status: {manualDecisions[activeCandidate.pair_key] || 'pending'}</p>
        <div className="actions"><button onClick={() => setManualDecisions({ ...manualDecisions, [activeCandidate.pair_key]: 'merge' })}>Merge</button><button onClick={() => setManualDecisions({ ...manualDecisions, [activeCandidate.pair_key]: 'keep_separate' })}>Keep separate</button><button onClick={() => setManualDecisions({ ...manualDecisions, [activeCandidate.pair_key]: 'unsure' })}>Unsure</button><button onClick={() => { const n = { ...manualDecisions }; delete n[activeCandidate.pair_key]; setManualDecisions(n); }}>Undo</button><button onClick={() => setManualIndex(Math.max(0, manualIndex - 1))}>Previous candidate</button><button onClick={() => setManualIndex(Math.min(manualQueue.length - 1, manualIndex + 1))}>Next candidate</button></div>
      </> : <p>No candidates available.</p>}</section>}

      {step === 4 && <section className="card"><h2>Export</h2><button onClick={handleExport}>Download cleaned workbook</button>
        <ul><li>accepted auto groups: {summary.acceptedAuto}</li><li>rejected auto groups: {summary.rejectedAuto}</li><li>manual merges: {summary.merge}</li><li>kept separate: {summary.keep}</li><li>unsure: {summary.unsure}</li></ul>
      </section>}

      <div className="footer-nav"><button disabled={step === 0} onClick={() => setStep(step - 1)}>Previous</button><button disabled={step === steps.length - 1 || (step===0 && (!file || !sheetName))} onClick={() => setStep(step + 1)}>Next</button></div>
    </div>
  );
}
