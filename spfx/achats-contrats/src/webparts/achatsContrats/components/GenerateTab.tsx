import * as React from 'react';
import { useEffect, useState } from 'react';
import * as generateService from '../services/generateService';
import * as conditionsService from '../services/conditionsService';
import { logAndGetMessage } from '../services/errorLog';
import type { ConditionsVersion, GenerateTemplateOption, SearchMatch } from '../model/types';

export default function GenerateTab({ onGoToTemplates }: { onGoToTemplates: () => void }): JSX.Element {
  const [templates, setTemplates] = useState<GenerateTemplateOption[]>([]);
  const [conditions, setConditions] = useState<ConditionsVersion[]>([]);
  const [templateId, setTemplateId] = useState('');
  const [conditionsVersionId, setConditionsVersionId] = useState('');
  const [code, setCode] = useState('');
  const [traitePar, setTraitePar] = useState('');

  const [matches, setMatches] = useState<SearchMatch[] | null>(null);
  const [rowIndex, setRowIndex] = useState<number | null>(null);
  const [values, setValues] = useState<Record<string, string> | null>(null);
  const [colonnesManquantes, setColonnesManquantes] = useState<string[]>([]);

  const [error, setError] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    Promise.all([generateService.listGenerableTemplates(), conditionsService.listConditions()]).then(([t, c]) => {
      setTemplates(t);
      setConditions(c);
      // Préférer un template déjà utilisable ; à défaut, laisser le premier sélectionné pour que
      // l'utilisateur voie tout de suite pourquoi il bloque (mapping incomplet).
      const premierUtilisable = t.find((x) => x.mappingComplet);
      if (premierUtilisable) setTemplateId(premierUtilisable.id);
      else if (t.length > 0) setTemplateId(t[0].id);
      const active = c.find((x) => x.estActive);
      if (active) setConditionsVersionId(active.id);
    }, (e) => setError(logAndGetMessage(e, 'GenerateTab (chargement initial)')));
  }, []);

  const selectedTemplate = templates.find((t) => t.id === templateId) ?? null;

  function resetSearch(): void {
    setMatches(null);
    setRowIndex(null);
    setValues(null);
    setColonnesManquantes([]);
    setInfo(null);
  }

  async function runSearch(): Promise<void> {
    setError(null);
    resetSearch();
    if (!conditionsVersionId || !code.trim() || !selectedTemplate?.mappingComplet) return;
    setBusy(true);
    try {
      const result = await generateService.searchCode(conditionsVersionId, code.trim());
      setMatches(result);
      if (result.length === 1) await resolveRow(result[0].rowIndex);
    } catch (e) {
      setError(logAndGetMessage(e, 'GenerateTab.runSearch'));
    } finally {
      setBusy(false);
    }
  }

  async function resolveRow(idx: number): Promise<void> {
    setRowIndex(idx);
    setError(null);
    setBusy(true);
    try {
      const { values: v, colonnesManquantes: cm } = await generateService.getMappedValues(conditionsVersionId, templateId, idx);
      setValues(v);
      setColonnesManquantes(cm);
    } catch (e) {
      setError(logAndGetMessage(e, 'GenerateTab.resolveRow'));
    } finally {
      setBusy(false);
    }
  }

  async function download(): Promise<void> {
    if (!values || rowIndex === null) return;
    setBusy(true);
    setError(null);
    setInfo(null);
    try {
      const { blob, filename } = await generateService.generateContract({
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
    } catch (e) {
      setError(logAndGetMessage(e, 'GenerateTab.download'));
    } finally {
      setBusy(false);
    }
  }

  return (
    <div>
      <h1>Générer un contrat</h1>
      <p className="subtitle">
        Sélectionnez un template et un fichier de conditions, saisissez le code sous-segment, relisez le formulaire
        pré-rempli puis téléchargez le contrat.
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
                  {!t.mappingComplet ? ' — mapping incomplet' : ''}
                </option>
              ))}
            </select>
            {templates.length === 0 && <div className="muted small mt1">Aucun template déposé pour l&apos;instant.</div>}
            {selectedTemplate && !selectedTemplate.mappingComplet && (
              <div className="alert error mt1">
                Le mapping de ce template n&apos;est pas complet (ou une colonne mappée n&apos;existe plus dans le
                fichier de conditions actif) : terminez-le dans l&apos;onglet{' '}
                <a
                  href="#"
                  onClick={(e) => {
                    e.preventDefault();
                    onGoToTemplates();
                  }}
                >
                  Templates de contrats
                </a>{' '}
                avant de pouvoir générer un contrat avec.
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
          <button
            disabled={busy || !templateId || !conditionsVersionId || !code.trim() || !selectedTemplate?.mappingComplet}
            onClick={runSearch}
          >
            Rechercher
          </button>
        </div>

        {matches && matches.length > 1 && rowIndex === null && (
          <div className="mt1">
            <p className="muted small">Plusieurs lignes correspondent à ce code — sélectionnez la ligne concernée :</p>
            <table>
              <thead>
                <tr>
                  {Object.keys(matches[0].preview).map((k) => (
                    <th key={k}>{k}</th>
                  ))}
                  <th />
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
            Les champs sont pré-remplis à partir du fichier de conditions ; vous pouvez les corriger avant génération.
            Un champ vide signifie que la colonne mappée est vide — complétez-le si besoin.
          </p>
          {colonnesManquantes.length > 0 && (
            <div className="alert error">
              La colonne mappée pour {colonnesManquantes.length > 1 ? "ces variables n'existe" : "cette variable n'existe"}{' '}
              pas dans le fichier de conditions sélectionné ({colonnesManquantes.join(', ')}) — vérifiez le mapping du
              template ou complétez la valeur manuellement avant de générer.
            </div>
          )}
          <div className="field-grid mt1">
            {Object.entries(values).map(([variable, value]) => (
              <div className="field" key={variable}>
                <label>{variable}</label>
                <input type="text" value={value} onChange={(e) => setValues({ ...values, [variable]: e.target.value })} />
              </div>
            ))}
          </div>
          <button className="mt1" disabled={busy} onClick={download}>
            {busy ? 'Génération…' : 'Télécharger le contrat'}
          </button>
        </div>
      )}
    </div>
  );
}
