const API_BASE = 'http://127.0.0.1:8000';

function toFormData(payload) {
  const form = new FormData();
  Object.entries(payload).forEach(([key, value]) => {
    if (value === undefined || value === null) return;
    if (typeof value === 'object' && !(value instanceof Blob) && !(value instanceof File)) {
      form.append(key, JSON.stringify(value));
      return;
    }
    form.append(key, value);
  });
  return form;
}

export async function fetchWorkbookSheets(file) {
  const form = new FormData();
  form.append('file', file);
  const res = await fetch(`${API_BASE}/workbook/sheets`, { method: 'POST', body: form });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

export async function fetchWorkbookPreview(file, sheetName) {
  const form = new FormData();
  form.append('file', file);
  form.append('sheet_name', sheetName);
  const res = await fetch(`${API_BASE}/workbook/preview`, { method: 'POST', body: form });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

export async function runAnalyze(payload) {
  const form = toFormData(payload);
  const res = await fetch(`${API_BASE}/dedupe/analyze`, { method: 'POST', body: form });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

export async function exportCleanWorkbook(payload) {
  const form = toFormData(payload);
  const res = await fetch(`${API_BASE}/export/cleaned-workbook`, { method: 'POST', body: form });
  if (!res.ok) throw new Error(await res.text());
  return res.blob();
}
