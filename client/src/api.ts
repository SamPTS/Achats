import type { ConditionsVersion, GenerateTemplateOption, SearchMatch, Template } from './types';

async function handleJson<T>(res: Response): Promise<T> {
  if (!res.ok) {
    let message = `Erreur ${res.status}`;
    try {
      const body = await res.json();
      if (body?.error) message = body.error;
      if (body?.details?.length) message += ` — ${body.details.join(' ; ')}`;
    } catch {
      // ignore parse error
    }
    throw new Error(message);
  }
  return res.json() as Promise<T>;
}

// --- Module 1 : Conditions commerciales ---

export function listConditions(): Promise<ConditionsVersion[]> {
  return fetch('/api/conditions').then((r) => handleJson(r));
}

export function uploadConditions(file: File, deposePar: string): Promise<ConditionsVersion> {
  const form = new FormData();
  form.append('file', file);
  if (deposePar) form.append('deposePar', deposePar);
  return fetch('/api/conditions', { method: 'POST', body: form }).then((r) => handleJson(r));
}

export function activateConditions(id: string): Promise<{ ok: true }> {
  return fetch(`/api/conditions/${id}/activate`, { method: 'POST' }).then((r) => handleJson(r));
}

export function archiveConditions(id: string): Promise<{ ok: true }> {
  return fetch(`/api/conditions/${id}/archive`, { method: 'POST' }).then((r) => handleJson(r));
}

export function setConditionsCodeColumn(id: string, colonneCodeSousSegment: string): Promise<ConditionsVersion> {
  return fetch(`/api/conditions/${id}`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ colonneCodeSousSegment }),
  }).then((r) => handleJson(r));
}

export function downloadConditionsUrl(id: string): string {
  return `/api/conditions/${id}/download`;
}

// --- Module 2 : Templates de contrats ---

export function listTemplates(): Promise<Template[]> {
  return fetch('/api/templates').then((r) => handleJson(r));
}

export function getTemplateVersions(groupId: string): Promise<Template[]> {
  return fetch(`/api/templates/group/${groupId}/versions`).then((r) => handleJson(r));
}

export function uploadTemplate(opts: {
  file: File;
  libelle: string;
  departement: string;
  groupId?: string;
  deposePar: string;
}): Promise<Template> {
  const form = new FormData();
  form.append('file', opts.file);
  form.append('libelle', opts.libelle);
  if (opts.departement) form.append('departement', opts.departement);
  if (opts.groupId) form.append('groupId', opts.groupId);
  if (opts.deposePar) form.append('deposePar', opts.deposePar);
  return fetch('/api/templates', { method: 'POST', body: form }).then((r) => handleJson(r));
}

export function downloadTemplateUrl(id: string): string {
  return `/api/templates/${id}/download`;
}

export function downloadBlankMappingUrl(id: string, conditionsVersionId?: string): string {
  return conditionsVersionId
    ? `/api/templates/${id}/mapping/blank?conditionsVersionId=${conditionsVersionId}`
    : `/api/templates/${id}/mapping/blank`;
}

export function uploadMapping(id: string, file: File, conditionsVersionId?: string): Promise<Template> {
  const form = new FormData();
  form.append('file', file);
  if (conditionsVersionId) form.append('conditionsVersionId', conditionsVersionId);
  return fetch(`/api/templates/${id}/mapping`, { method: 'POST', body: form }).then((r) => handleJson(r));
}

export function setMappingLine(
  id: string,
  variable: string,
  body: { colonneCorrespondante?: string | null; statut?: 'mappee' | 'libre' | 'manquante' }
): Promise<Template> {
  return fetch(`/api/templates/${id}/mapping/${encodeURIComponent(variable)}`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  }).then((r) => handleJson(r));
}

// --- Module 3 : Génération ---

export function listGenerableTemplates(): Promise<GenerateTemplateOption[]> {
  return fetch('/api/generate/templates').then((r) => handleJson(r));
}

export async function searchCode(
  conditionsVersionId: string,
  code: string
): Promise<{ matches: SearchMatch[] } | { error: string }> {
  const res = await fetch(
    `/api/generate/search?conditionsVersionId=${encodeURIComponent(conditionsVersionId)}&code=${encodeURIComponent(code)}`
  );
  if (res.status === 404) {
    const body = await res.json();
    return { error: body.error };
  }
  return handleJson(res);
}

export function getMappedValues(
  conditionsVersionId: string,
  templateId: string,
  rowIndex: number
): Promise<{ values: Record<string, string> }> {
  return fetch(
    `/api/generate/mapped-values?conditionsVersionId=${conditionsVersionId}&templateId=${templateId}&rowIndex=${rowIndex}`
  ).then((r) => handleJson(r));
}

export async function generateContract(payload: {
  templateId: string;
  conditionsVersionId: string;
  codeSousSegment: string;
  values: Record<string, string>;
  traitePar: string;
}): Promise<{ blob: Blob; filename: string }> {
  const res = await fetch('/api/generate/download', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body.error || `Erreur ${res.status}`);
  }
  const disposition = res.headers.get('Content-Disposition') || '';
  const match = /filename="([^"]+)"/.exec(disposition);
  const filename = match ? match[1] : 'contrat.docx';
  const blob = await res.blob();
  return { blob, filename };
}
