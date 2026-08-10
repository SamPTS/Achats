import { useEffect, useRef, useState } from 'react';
import {
  downloadBlankMappingUrl,
  downloadTemplateUrl,
  listConditions,
  listTemplates,
  setMappingLine,
  uploadMapping,
  uploadTemplate,
} from '../api';
import type { ConditionsVersion, Template } from '../types';

function fmtDate(iso: string): string {
  return new Date(iso).toLocaleString('fr-FR');
}

function statusBadge(statut: Template['statutMapping']) {
  if (statut === 'complet') return <span className="badge ok">Mapping complet</span>;
  if (statut === 'incomplet') return <span className="badge warn">Mapping incomplet</span>;
  return <span className="badge danger">Mapping absent</span>;
}

export default function TemplatesPage() {
  const [templates, setTemplates] = useState<Template[]>([]);
  const [conditions, setConditions] = useState<ConditionsVersion[]>([]);
  const [selected, setSelected] = useState<Template | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [showUpload, setShowUpload] = useState(false);

  async function refresh() {
    const [t, c] = await Promise.all([listTemplates(), listConditions()]);
    setTemplates(t);
    setConditions(c);
    if (selected) {
      const fresh = t.find((x) => x.id === selected.id);
      setSelected(fresh ?? null);
    }
  }

  useEffect(() => {
    refresh().catch((e) => setError(e.message));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const activeConditions = conditions.find((c) => c.estActive) ?? null;

  return (
    <div>
      <h1>Templates de contrats</h1>
      <p className="subtitle">
        Dépôt et historisation des templates Word (variables au format <code>{'{{Variable}}'}</code>) et
        correspondance avec les colonnes du fichier de conditions commerciales.
      </p>

      {error && <div className="alert error">{error}</div>}

      <div className="panel">
        <button onClick={() => setShowUpload((v) => !v)}>
          {showUpload ? 'Annuler' : '+ Déposer un template'}
        </button>
        {showUpload && (
          <UploadTemplateForm
            onDone={async () => {
              setShowUpload(false);
              await refresh();
            }}
          />
        )}
      </div>

      <div className="panel">
        <table>
          <thead>
            <tr>
              <th>Libellé</th>
              <th>Département</th>
              <th>Version</th>
              <th>Dépôt</th>
              <th>Variables</th>
              <th>Statut mapping</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {templates.map((t) => (
              <tr key={t.id}>
                <td>
                  <a href={downloadTemplateUrl(t.id)}>{t.libelle}</a>
                </td>
                <td>{t.departement || <span className="muted">—</span>}</td>
                <td>v{t.version}</td>
                <td>{fmtDate(t.dateDepot)}</td>
                <td>{t.variables.length}</td>
                <td>{statusBadge(t.statutMapping)}</td>
                <td>
                  <button className="secondary" onClick={() => setSelected(t)}>
                    Gérer le mapping
                  </button>
                </td>
              </tr>
            ))}
            {templates.length === 0 && (
              <tr>
                <td colSpan={7} className="muted">
                  Aucun template déposé pour l'instant.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {selected && (
        <MappingPanel
          template={selected}
          activeConditions={activeConditions}
          conditions={conditions}
          onClose={() => setSelected(null)}
          onChanged={refresh}
        />
      )}
    </div>
  );
}

function UploadTemplateForm({ onDone }: { onDone: () => Promise<void> }) {
  const [libelle, setLibelle] = useState('');
  const [departement, setDepartement] = useState('');
  const [deposePar, setDeposePar] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);

  async function submit() {
    const file = fileInput.current?.files?.[0];
    if (!file) {
      setError('Merci de sélectionner un fichier .docx.');
      return;
    }
    if (!libelle.trim()) {
      setError('Le libellé est requis.');
      return;
    }
    setBusy(true);
    setError(null);
    try {
      await uploadTemplate({ file, libelle, departement, deposePar });
      await onDone();
    } catch (e: any) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="mt1">
      {error && <div className="alert error">{error}</div>}
      <div className="form-row">
        <div>
          <label>Libellé *</label>
          <input type="text" value={libelle} onChange={(e) => setLibelle(e.target.value)} placeholder="Contrat Beauty — Capillaire" />
        </div>
        <div>
          <label>Département (optionnel)</label>
          <input type="text" value={departement} onChange={(e) => setDepartement(e.target.value)} />
        </div>
        <div>
          <label>Traité par (optionnel)</label>
          <input type="text" value={deposePar} onChange={(e) => setDeposePar(e.target.value)} />
        </div>
        <div>
          <label>Fichier .docx *</label>
          <input ref={fileInput} type="file" accept=".docx" />
        </div>
      </div>
      <button disabled={busy} onClick={submit}>
        {busy ? 'Dépôt en cours…' : 'Déposer'}
      </button>
    </div>
  );
}

function MappingPanel({
  template,
  activeConditions,
  conditions,
  onClose,
  onChanged,
}: {
  template: Template;
  activeConditions: ConditionsVersion | null;
  conditions: ConditionsVersion[];
  onClose: () => void;
  onChanged: () => Promise<void>;
}) {
  const [conditionsVersionId, setConditionsVersionId] = useState<string>(activeConditions?.id ?? '');
  const [error, setError] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);

  const refColumns = conditions.find((c) => c.id === conditionsVersionId)?.colonnes ?? [];

  async function importMapping() {
    const file = fileInput.current?.files?.[0];
    if (!file) {
      setError('Merci de sélectionner le fichier de mapping rempli.');
      return;
    }
    setBusy(true);
    setError(null);
    setInfo(null);
    try {
      await uploadMapping(template.id, file, conditionsVersionId || undefined);
      setInfo('Mapping importé avec succès.');
      await onChanged();
    } catch (e: any) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  }

  async function markFree(variable: string) {
    await setMappingLine(template.id, variable, { statut: 'libre', colonneCorrespondante: null });
    await onChanged();
  }

  async function assignColumn(variable: string, colonne: string) {
    await setMappingLine(template.id, variable, {
      statut: colonne ? 'mappee' : 'manquante',
      colonneCorrespondante: colonne || null,
    });
    await onChanged();
  }

  return (
    <div className="panel">
      <div className="form-row" style={{ justifyContent: 'space-between' }}>
        <h2 style={{ margin: 0 }}>
          Mapping — {template.libelle} (v{template.version})
        </h2>
        <button className="secondary" onClick={onClose}>
          Fermer
        </button>
      </div>

      <div className="form-row">
        <div>
          <label>Fichier de conditions de référence</label>
          <select value={conditionsVersionId} onChange={(e) => setConditionsVersionId(e.target.value)}>
            <option value="">Version active par défaut</option>
            {conditions.map((c) => (
              <option key={c.id} value={c.id}>
                {c.nomFichier} {c.estActive ? '(active)' : ''}
              </option>
            ))}
          </select>
        </div>
        <a href={downloadBlankMappingUrl(template.id, conditionsVersionId || undefined)}>
          <button className="secondary" type="button">
            Télécharger le mapping vierge
          </button>
        </a>
      </div>

      {error && <div className="alert error">{error}</div>}
      {info && <div className="alert success">{info}</div>}

      <div className="form-row">
        <input ref={fileInput} type="file" accept=".xlsx,.csv" />
        <button disabled={busy} onClick={importMapping}>
          {busy ? 'Import…' : 'Importer le mapping rempli'}
        </button>
      </div>

      <table className="mt1">
        <thead>
          <tr>
            <th>Variable</th>
            <th>Statut</th>
            <th>Colonne correspondante</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          {template.mappings.map((m) => (
            <tr key={m.variable}>
              <td>
                <code>{`{{${m.variable}}}`}</code>
              </td>
              <td>
                {m.statut === 'mappee' && <span className="badge ok">Mappée</span>}
                {m.statut === 'libre' && <span className="badge muted">Saisie libre</span>}
                {m.statut === 'manquante' && <span className="badge danger">Manquante</span>}
              </td>
              <td>
                <select
                  value={m.colonneCorrespondante ?? ''}
                  disabled={m.statut === 'libre'}
                  onChange={(e) => assignColumn(m.variable, e.target.value)}
                >
                  <option value="">—</option>
                  {refColumns.map((c) => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
              </td>
              <td>
                {m.statut !== 'libre' && (
                  <button className="secondary small" onClick={() => markFree(m.variable)}>
                    Marquer comme saisie libre
                  </button>
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
