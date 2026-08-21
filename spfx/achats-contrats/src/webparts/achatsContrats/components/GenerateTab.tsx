import * as React from 'react';
import { useEffect, useState } from 'react';
import * as generateService from '../services/generateService';
import * as conditionsService from '../services/conditionsService';
import { logAndGetMessage } from '../services/errorLog';
import type { ConditionsVersion, GenerateTemplateOption, SearchMatch } from '../model/types';

interface BatchItem {
  code: string;
  statut: 'resolue' | 'ambigue' | 'introuvable' | 'erreur';
  rowIndex?: number;
  matches?: SearchMatch[];
  values?: Record<string, string>;
  colonnesManquantes?: string[];
  message?: string;
  expanded?: boolean;
}

function parseCodes(text: string): string[] {
  const seen = new Set<string>();
  const codes: string[] = [];
  for (const line of text.split('\n')) {
    const code = line.trim();
    if (!code || seen.has(code)) continue;
    seen.add(code);
    codes.push(code);
  }
  return codes;
}

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

  // --- Mode "plusieurs contrats à la fois" ---
  const [batchMode, setBatchMode] = useState(false);
  const [batchCodesText, setBatchCodesText] = useState('');
  const [batchItems, setBatchItems] = useState<BatchItem[]>([]);
  const [batchBusy, setBatchBusy] = useState(false);

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
  const canSearch = !!templateId && !!conditionsVersionId && !!selectedTemplate?.mappingComplet;

  function resetSearch(): void {
    setMatches(null);
    setRowIndex(null);
    setValues(null);
    setColonnesManquantes([]);
    setInfo(null);
    setBatchItems([]);
  }

  async function runSearch(): Promise<void> {
    setError(null);
    resetSearch();
    if (!canSearch || !code.trim()) return;
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

  // --- Mode lot ---

  async function runBatchSearch(): Promise<void> {
    setError(null);
    setInfo(null);
    const codes = parseCodes(batchCodesText);
    if (!canSearch || codes.length === 0) return;
    setBatchBusy(true);
    try {
      const items: BatchItem[] = [];
      for (const c of codes) {
        try {
          const result = await generateService.searchCode(conditionsVersionId, c);
          if (result.length === 1) {
            const { values: v, colonnesManquantes: cm } = await generateService.getMappedValues(
              conditionsVersionId,
              templateId,
              result[0].rowIndex,
            );
            items.push({ code: c, statut: 'resolue', rowIndex: result[0].rowIndex, values: v, colonnesManquantes: cm });
          } else {
            items.push({ code: c, statut: 'ambigue', matches: result });
          }
        } catch (e) {
          items.push({ code: c, statut: 'introuvable', message: e instanceof Error ? e.message : String(e) });
        }
      }
      setBatchItems(items);
    } catch (e) {
      setError(logAndGetMessage(e, 'GenerateTab.runBatchSearch'));
    } finally {
      setBatchBusy(false);
    }
  }

  async function resolveBatchAmbiguous(code: string, idx: number): Promise<void> {
    setError(null);
    try {
      const { values: v, colonnesManquantes: cm } = await generateService.getMappedValues(conditionsVersionId, templateId, idx);
      setBatchItems((cur) =>
        cur.map((it) => (it.code === code ? { ...it, statut: 'resolue', rowIndex: idx, values: v, colonnesManquantes: cm } : it)),
      );
    } catch (e) {
      setError(logAndGetMessage(e, 'GenerateTab.resolveBatchAmbiguous'));
    }
  }

  function updateBatchValue(code: string, variable: string, value: string): void {
    setBatchItems((cur) =>
      cur.map((it) => (it.code === code && it.values ? { ...it, values: { ...it.values, [variable]: value } } : it)),
    );
  }

  function toggleExpand(code: string): void {
    setBatchItems((cur) => cur.map((it) => (it.code === code ? { ...it, expanded: !it.expanded } : it)));
  }

  const resoluCount = batchItems.filter((it) => it.statut === 'resolue').length;

  async function downloadBatch(): Promise<void> {
    const aTraiter = batchItems.filter((it) => it.statut === 'resolue' && it.values);
    if (aTraiter.length === 0) return;
    setBatchBusy(true);
    setError(null);
    setInfo(null);
    try {
      const { blob, filename, erreurs } = await generateService.generateContractsBatch(
        aTraiter.map((it) => ({
          templateId,
          conditionsVersionId,
          codeSousSegment: it.code,
          values: it.values!,
          traitePar,
        })),
      );
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      const ignores = batchItems.length - aTraiter.length;
      const parts = [`${aTraiter.length - erreurs.length} contrat(s) généré(s) et regroupé(s) dans ${filename}.`];
      if (erreurs.length > 0) parts.push(`${erreurs.length} échec(s) : ${erreurs.map((e) => `${e.codeSousSegment} (${e.message})`).join(' · ')}.`);
      if (ignores > 0) parts.push(`${ignores} code(s) ignoré(s) (non résolus).`);
      setInfo(parts.join(' '));
    } catch (e) {
      setError(logAndGetMessage(e, 'GenerateTab.downloadBatch'));
    } finally {
      setBatchBusy(false);
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

        <div className="form-row" style={{ alignItems: 'center' }}>
          <label style={{ display: 'flex', alignItems: 'center', gap: '0.4rem', fontWeight: 'normal' }}>
            <input
              type="checkbox"
              checked={batchMode}
              onChange={(e) => {
                setBatchMode(e.target.checked);
                resetSearch();
                setCode('');
                setBatchCodesText('');
              }}
            />
            Générer plusieurs contrats à la fois
          </label>
        </div>

        {!batchMode ? (
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
            <button disabled={busy || !canSearch || !code.trim()} onClick={runSearch}>
              Rechercher
            </button>
          </div>
        ) : (
          <div className="form-row">
            <div style={{ flex: 1 }}>
              <label>Codes sous-segment (un par ligne)</label>
              <textarea
                value={batchCodesText}
                onChange={(e) => setBatchCodesText(e.target.value)}
                rows={5}
                placeholder={'ex :\nSS001\nSS002\nSS003'}
                style={{ width: '100%', fontFamily: 'inherit' }}
              />
            </div>
            <button disabled={batchBusy || !canSearch || parseCodes(batchCodesText).length === 0} onClick={runBatchSearch}>
              {batchBusy ? 'Recherche…' : 'Rechercher tout'}
            </button>
          </div>
        )}

        {!batchMode && matches && matches.length > 1 && rowIndex === null && (
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

      {!batchMode && values && (
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

      {batchMode && batchItems.length > 0 && (
        <div className="panel">
          <h2 className="mb0">Relecture du lot ({batchItems.length} code(s))</h2>
          <p className="muted small">
            Dépliez une ligne pour relire ou corriger ses valeurs avant génération, comme pour un contrat unique. Un code
            introuvable ou resté ambigu (plusieurs lignes correspondantes non tranchées) sera ignoré au téléchargement.
          </p>
          <table className="mt1">
            <thead>
              <tr>
                <th>Code</th>
                <th>Statut</th>
                <th>Alertes</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {batchItems.map((it) => (
                <React.Fragment key={it.code}>
                  <tr>
                    <td>
                      <code>{it.code}</code>
                    </td>
                    <td>
                      {it.statut === 'resolue' && <span className="badge ok">Résolu</span>}
                      {it.statut === 'ambigue' && <span className="badge warn">Plusieurs lignes — à choisir</span>}
                      {it.statut === 'introuvable' && <span className="badge danger">Introuvable</span>}
                      {it.statut === 'erreur' && <span className="badge danger">Erreur</span>}
                    </td>
                    <td>
                      {it.statut === 'resolue' && it.colonnesManquantes && it.colonnesManquantes.length > 0 && (
                        <span className="badge warn">{it.colonnesManquantes.length} colonne(s) manquante(s)</span>
                      )}
                      {it.statut === 'introuvable' && <span className="muted small">{it.message}</span>}
                    </td>
                    <td>
                      {it.statut === 'resolue' && (
                        <button className="secondary small" onClick={() => toggleExpand(it.code)}>
                          {it.expanded ? 'Masquer' : 'Voir / modifier'}
                        </button>
                      )}
                    </td>
                  </tr>
                  {it.statut === 'ambigue' && it.matches && (
                    <tr>
                      <td colSpan={4}>
                        <p className="muted small">Plusieurs lignes correspondent à ce code — choisissez la bonne :</p>
                        <table>
                          <thead>
                            <tr>
                              {Object.keys(it.matches[0].preview).map((k) => (
                                <th key={k}>{k}</th>
                              ))}
                              <th />
                            </tr>
                          </thead>
                          <tbody>
                            {it.matches.map((m) => (
                              <tr key={m.rowIndex}>
                                {Object.values(m.preview).map((v, i) => (
                                  <td key={i}>{v}</td>
                                ))}
                                <td>
                                  <button className="secondary small" onClick={() => resolveBatchAmbiguous(it.code, m.rowIndex)}>
                                    Choisir
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </td>
                    </tr>
                  )}
                  {it.statut === 'resolue' && it.expanded && it.values && (
                    <tr>
                      <td colSpan={4}>
                        <div className="field-grid">
                          {Object.entries(it.values).map(([variable, value]) => (
                            <div className="field" key={variable}>
                              <label>{variable}</label>
                              <input
                                type="text"
                                value={value}
                                onChange={(e) => updateBatchValue(it.code, variable, e.target.value)}
                              />
                            </div>
                          ))}
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              ))}
            </tbody>
          </table>
          <button className="mt1" disabled={batchBusy || resoluCount === 0} onClick={downloadBatch}>
            {batchBusy ? 'Génération…' : `Télécharger ${resoluCount} contrat(s) (ZIP)`}
          </button>
        </div>
      )}
    </div>
  );
}
