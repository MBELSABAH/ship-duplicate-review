const API_BASE_URL = 'http://127.0.0.1:8000';

async function handleResponse(response) {
  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `Request failed with status ${response.status}`);
  }
  return response;
}

export async function getSheets(file) {
  const formData = new FormData();
  formData.append('file', file);

  const response = await fetch(`${API_BASE_URL}/workbook/sheets`, {
    method: 'POST',
    body: formData,
  });

  const safeResponse = await handleResponse(response);
  return safeResponse.json();
}

export async function getPreview(file, sheetName) {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('sheet_name', sheetName);

  const response = await fetch(`${API_BASE_URL}/workbook/preview`, {
    method: 'POST',
    body: formData,
  });

  const safeResponse = await handleResponse(response);
  return safeResponse.json();
}

export async function analyzeWorkbook(file, sheetName, columnConfig, fuzzyThreshold, minManualScore) {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('sheet_name', sheetName);
  formData.append('column_config', JSON.stringify(columnConfig));
  formData.append('fuzzy_threshold', String(fuzzyThreshold));
  formData.append('min_manual_score', String(minManualScore));
  formData.append('auto_status', JSON.stringify({}));
  formData.append('manual_decisions', JSON.stringify({}));

  const response = await fetch(`${API_BASE_URL}/dedupe/analyze`, {
    method: 'POST',
    body: formData,
  });

  const safeResponse = await handleResponse(response);
  return safeResponse.json();
}

export async function downloadCleanedWorkbook(file, sheetName, columnConfig, fuzzyThreshold, minManualScore) {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('sheet_name', sheetName);
  formData.append('column_config', JSON.stringify(columnConfig));
  formData.append('fuzzy_threshold', String(fuzzyThreshold));
  formData.append('min_manual_score', String(minManualScore));
  formData.append('auto_status', JSON.stringify({}));
  formData.append('manual_decisions', JSON.stringify({}));

  const response = await fetch(`${API_BASE_URL}/export/cleaned-workbook`, {
    method: 'POST',
    body: formData,
  });

  return handleResponse(response);
}
