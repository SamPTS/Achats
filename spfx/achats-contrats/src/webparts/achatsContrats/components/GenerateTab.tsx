import * as React from 'react';
import { useEffect, useState } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
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

function parseCodes(lines: string[]): string[] {
  const seen = new Set<string>();
  const codes: string[] = [];
  for (const line of lines) {
    const code = line.trim();
    if (!code || seen.has(code)) continue;
    seen.add(code);
    codes.push(code);
  }
  return codes;
}

export default function GenerateTab({
  onGoToTemplates,
}: {
  /** Absent si l'utilisateur courant n'a pas accès à l'onglet Templates de contrats (droits
   * insuffisants — voir permissions.ts) : le lien devient un simple texte informatif plutôt qu'une
   * action impossible à mener à bien. */
  onGoToTemplates?: () => void;
}): JSX.Element {
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
  // Une ligne de saisie par code, plutôt qu'une zone de texte multi-lignes : on démarre avec une
  // seule ligne et "Ajouter une ligne" en ajoute une nouvelle vide à la suite.
  const [batchCodeLines, setBatchCodeLines] = useState<string[]>(['']);
  const [batchItems, setBatchItems] = useState<BatchItem[]>([]);
  const [batchBusy, setBatchBusy] = useState(false);
  // Progression affichée pendant runBatchSearch/downloadBatch (voir leurs boucles) — sans elle,
  // seul le libellé du bouton changeait ("Recherche…"/"Génération…") pendant tout le traitement
  // du lot, sans aucun repère de progression, ce qui peut donner l'impression que l'application
  // est figée sur un lot volumineux avec une latence réseau non négligeable.
  const [batchProgress, setBatchProgress] = useState<{ done: number; total: number } | null>(null);

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

  function updateBatchCodeLine(index: number, value: string): void {
    setBatchCodeLines((cur) => cur.map((c, i) => (i === index ? value : c)));
  }

  function addBatchCodeLine(): void {
    setBatchCodeLines((cur) => [...cur, '']);
  }

  function removeBatchCodeLine(index: number): void {
    setBatchCodeLines((cur) => (cur.length <= 1 ? [''] : cur.filter((_, i) => i !== index)));
  }

  async function runBatchSearch(): Promise<void> {
    setError(null);
    setInfo(null);
    const codes = parseCodes(batchCodeLines);
    if (!canSearch || codes.length === 0) return;
    setBatchBusy(true);
    setBatchItems([]);
    setBatchProgress({ done: 0, total: codes.length });
    try {
      let done = 0;
      for (const c of codes) {
        let item: BatchItem;
        try {
          const result = await generateService.searchCode(conditionsVersionId, c);
          if (result.length === 1) {
            const { values: v, colonnesManquantes: cm } = await generateService.getMappedValues(
              conditionsVersionId,
              templateId,
              result[0].rowIndex,
            );
            item = { code: c, statut: 'resolue', rowIndex: result[0].rowIndex, values: v, colonnesManquantes: cm };
          } else {
            item = { code: c, statut: 'ambigue', matches: result };
          }
        } catch (e) {
          item = { code: c, statut: 'introuvable', message: e instanceof Error ? e.message : String(e) };
        }
        // Mise à jour incrémentale (un code traité à la fois) plutôt qu'un seul setBatchItems en
        // fin de boucle : la table de relecture se remplit au fur et à mesure, visible pendant le
        // traitement plutôt qu'en un seul bloc à la fin.
        setBatchItems((cur) => [...cur, item]);
        done++;
        setBatchProgress({ done, total: codes.length });
      }
    } catch (e) {
      setError(logAndGetMessage(e, 'GenerateTab.runBatchSearch'));
    } finally {
      setBatchBusy(false);
      setBatchProgress(null);
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
    setBatchProgress({ done: 0, total: aTraiter.length });
    try {
      const { blob, filename, erreurs } = await generateService.generateContractsBatch(
        aTraiter.map((it) => ({
          templateId,
          conditionsVersionId,
          codeSousSegment: it.code,
          values: it.values!,
          traitePar,
        })),
        (done, total) => setBatchProgress({ done, total }),
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
      setBatchProgress(null);
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
                fichier de conditions actif) : {onGoToTemplates ? (
                  <>
                    terminez-le dans l&apos;onglet{' '}
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
                  </>
                ) : (
                  "contactez un propriétaire du site pour le compléter avant de pouvoir générer un contrat avec."
                )}
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
                setBatchCodeLines(['']);
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
          <div className="mt1">
            <label>Codes sous-segment</label>
            {batchCodeLines.map((line, i) => (
              <div key={i} className="form-row" style={{ marginTop: i === 0 ? 0 : '0.4rem', alignItems: 'center' }}>
                <input
                  type="search"
                  value={line}
                  onChange={(e) => updateBatchCodeLine(i, e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key !== 'Enter') return;
                    if (i === batchCodeLines.length - 1) addBatchCodeLine();
                  }}
                  placeholder="ex : SS001"
                  style={{ flex: 1 }}
                />
                <button
                  type="button"
                  className="icon-btn danger"
                  title="Retirer cette ligne"
                  disabled={batchCodeLines.length === 1 && !line}
                  onClick={() => removeBatchCodeLine(i)}
                >
                  <Icon iconName="Delete" />
                </button>
              </div>
            ))}
            <div className="form-row mt1">
              <button type="button" className="secondary" onClick={addBatchCodeLine}>
                <Icon iconName="Add" /> Ajouter une ligne
              </button>
              <button disabled={batchBusy || !canSearch || parseCodes(batchCodeLines).length === 0} onClick={runBatchSearch}>
                {batchProgress ? `Recherche… (${batchProgress.done}/${batchProgress.total})` : 'Rechercher tout'}
              </button>
            </div>
          </div>
        )}

        {!templateId && templates.length > 0 && (
          <div className="alert error mt1">Sélectionnez un template de contrat pour pouvoir rechercher un code.</div>
        )}
        {templateId && !conditionsVersionId && (
          <div className="alert error mt1">
            Sélectionnez un fichier de conditions commerciales pour pouvoir rechercher un code — ou déposez-en un dans
            l&apos;onglet Conditions commerciales si aucun n&apos;est encore disponible.
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
            {batchProgress ? `Génération… (${batchProgress.done}/${batchProgress.total})` : `Télécharger ${resoluCount} contrat(s) (ZIP)`}
          </button>
        </div>
      )}
    </div>
  );
}
