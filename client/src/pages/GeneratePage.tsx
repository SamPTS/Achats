import { useEffect, useState } from 'react';
import {
  generateContract,
  getMappedValues,
  listConditions,
  listGenerableTemplates,
  listGenerationLog,
  searchCode,
} from '../api';
import type { ConditionsVersion, GenerateTemplateOption, GenerationLogEntry, SearchMatch } from '../types';

function fmtDate(iso: string): string {
  return new Date(iso).toLocaleString('fr-FR');
}

export default function GeneratePage() {
  const [templates, setTemplates] = useState<GenerateTemplateOption[]>([]);
  const [conditions, setConditions] = useState<ConditionsVersion[]>([]);
  const [templateId, setTemplateId] = useState('');
  const [conditionsVersionId, setConditionsVersionId] = useState('');
  const [code, setCode] = useState('');
  const [traitePar, setTraitePar] = useState('');

  const [matches, setMatches] = useState<SearchMatch[] | null>(null);
  const [rowIndex, setRowIndex] = useState<number | null>(null);
  const [values, setValues] = useState<Record<string, string> | null>(null);

  const [error, setError] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const [log, setLog] = useState<GenerationLogEntry[]>([]);

  async function refreshLog() {
    setLog(await listGenerationLog());
  }

  useEffect(() => {
    Promise.all([listGenerableTemplates(), listConditions()]).then(([t, c]) => {
      setTemplates(t);
      setConditions(c);
      if (t.length > 0) setTemplateId(t[0].id);
      const active = c.find((x) => x.estActive);
      if (active) setConditionsVersionId(active.id);
    });
    refreshLog().catch((e) => setError(e.message));
  }, []);

  function resetSearch() {
    setMatches(null);
    setRowIndex(null);
    setValues(null);
    setInfo(null);
  }

  async function runSearch() {
    setError(null);
    resetSearch();
    if (!conditionsVersionId || !code.trim()) return;
    setBusy(true);
    try {
      const result = await searchCode(conditionsVersionId, code.trim());
      if ('error' in result) {
        setError(result.error);
        return;
      }
      setMatches(result.matches);
      if (result.matches.length === 1) {
        await resolveRow(result.matches[0].rowIndex);
      }
    } catch (e: any) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  }

  async function resolveRow(idx: number) {
    setRowIndex(idx);
    setError(null);
    setBusy(true);
    try {
      const { values: v } = await getMappedValues(conditionsVersionId, templateId, idx);
      setValues(v);
    } catch (e: any) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  }

  async function download() {
    if (!values || rowIndex === null) return;
    setBusy(true);
    setError(null);
    setInfo(null);
    try {
      const { blob, filename } = await generateContract({
        templateId,
        conditionsVersionId,
        codeSousSegment: code.trim(),
        values,
        traitePar,
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setInfo('Contrat généré et téléchargé.');
      await refreshLog();
    } catch (e: any) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div>
      <h1>Générer un contrat</h1>
      <p className="subtitle">
        Sélectionnez un template et un fichier de conditions, saisissez le code sous-segment, relisez le
        formulaire pré-rempli puis téléchargez le contrat.
      </p>

      {error && <div className="alert error">{error}</div>}
      {info && <div className="alert success">{info}</div>}

      <div className="panel">
        <div className="form-row">
          <div>
            <label>Template de contrat</label>
            <select
              value={templateId}
              onChange={(e) => {
                setTemplateId(e.target.value);
                resetSearch();
              }}
            >
              {templates.map((t) => (
                <option key={t.id} value={t.id}>
                  {t.libelle} {t.departement ? `(${t.departement})` : ''} — v{t.version}
                </option>
              ))}
            </select>
            {templates.length === 0 && (
              <div className="muted small mt1">
                Aucun template avec un mapping complet n'est disponible pour l'instant.
              </div>
            )}
          </div>
          <div>
            <label>Fichier de conditions commerciales</label>
            <select
              value={conditionsVersionId}
              onChange={(e) => {
                setConditionsVersionId(e.target.value);
                resetSearch();
              }}
            >
              <option value="">—</option>
              {conditions.map((c) => (
                <option key={c.id} value={c.id}>
                  {c.nomFichier} {c.estActive ? '(active)' : ''}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label>Traité par (optionnel)</label>
            <input type="text" value={traitePar} onChange={(e) => setTraitePar(e.target.value)} />
          </div>
        </div>

        <div className="form-row">
          <div>
            <label>Code sous-segment</label>
            <input
              type="search"
              value={code}
              onChange={(e) => setCode(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && runSearch()}
              placeholder="ex : SS001"
            />
          </div>
          <button disabled={busy || !templateId || !conditionsVersionId || !code.trim()} onClick={runSearch}>
            Rechercher
          </button>
        </div>

        {matches && matches.length > 1 && rowIndex === null && (
          <div className="mt1">
            <p className="muted small">
              Plusieurs lignes correspondent à ce code — sélectionnez la ligne concernée :
            </p>
            <table>
              <thead>
                <tr>
                  {Object.keys(matches[0].preview).map((k) => (
                    <th key={k}>{k}</th>
                  ))}
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {matches.map((m) => (
                  <tr key={m.rowIndex}>
                    {Object.values(m.preview).map((v, i) => (
                      <td key={i}>{v}</td>
                    ))}
                    <td>
                      <button className="secondary" onClick={() => resolveRow(m.rowIndex)}>
                        Choisir
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {values && (
        <div className="panel">
          <h2 className="mb0">Formulaire — relecture et correction</h2>
          <p className="muted small">
            Les champs sont pré-remplis à partir du fichier de conditions ; vous pouvez les corriger avant
            génération. Un champ vide signifie que la colonne mappée est vide — complétez-le si besoin.
          </p>
          <div className="field-grid mt1">
            {Object.entries(values).map(([variable, value]) => (
              <div className="field" key={variable}>
                <label>{variable}</label>
                <input
                  type="text"
                  value={value}
                  onChange={(e) => setValues({ ...values, [variable]: e.target.value })}
                />
              </div>
            ))}
          </div>
          <button className="mt1" disabled={busy} onClick={download}>
            {busy ? 'Génération…' : 'Télécharger le contrat'}
          </button>
        </div>
      )}

      <div className="panel">
        <h2 className="mb0">Historique des générations</h2>
        <p className="muted small">
          Traçabilité de chaque contrat généré : template, fichier de conditions et code sous-segment
          utilisés.
        </p>
        <table className="mt1">
          <thead>
            <tr>
              <th>Date</th>
              <th>Template</th>
              <th>Fichier de conditions</th>
              <th>Code sous-segment</th>
              <th>Traité par</th>
            </tr>
          </thead>
          <tbody>
            {log.map((g) => (
              <tr key={g.id}>
                <td>{fmtDate(g.dateGeneration)}</td>
                <td>{g.templateLibelle}</td>
                <td>{g.conditionsNomFichier}</td>
                <td>{g.codeSousSegment}</td>
                <td>{g.traitePar || <span className="muted">—</span>}</td>
              </tr>
            ))}
            {log.length === 0 && (
              <tr>
                <td colSpan={5} className="muted">
                  Aucun contrat généré pour l'instant.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
