import * as React from 'react';
import styles from './AchatsContrats.module.scss';
import { GenerateService } from '../services/generateService';
import { ConditionsService } from '../services/conditionsService';
import { IConditionsVersion, IGenerateTemplateOption, ISearchMatch } from '../models/types';

export interface IGenerateTabProps {
  generateService: GenerateService;
  conditionsService: ConditionsService;
}

export default function GenerateTab(props: IGenerateTabProps): JSX.Element {
  const { generateService, conditionsService } = props;
  const [templates, setTemplates] = React.useState<IGenerateTemplateOption[]>([]);
  const [conditions, setConditions] = React.useState<IConditionsVersion[]>([]);
  const [templateId, setTemplateId] = React.useState('');
  const [conditionsVersionId, setConditionsVersionId] = React.useState('');
  const [code, setCode] = React.useState('');
  const [traitePar, setTraitePar] = React.useState('');

  const [matches, setMatches] = React.useState<ISearchMatch[] | null>(null);
  const [rowIndex, setRowIndex] = React.useState<number | null>(null);
  const [values, setValues] = React.useState<Record<string, string> | null>(null);

  const [error, setError] = React.useState<string | null>(null);
  const [info, setInfo] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState(false);

  React.useEffect(() => {
    Promise.all([generateService.listGenerableTemplates(), conditionsService.listVersions()]).then(
      ([t, c]) => {
        setTemplates(t);
        setConditions(c);
        if (t.length > 0) setTemplateId(String(t[0].id));
        const active = c.find((x) => x.estActive);
        if (active) setConditionsVersionId(String(active.id));
      }
    );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function resetSearch(): void {
    setMatches(null);
    setRowIndex(null);
    setValues(null);
    setInfo(null);
  }

  function runSearch(): void {
    setError(null);
    resetSearch();
    if (!conditionsVersionId || !code.trim()) return;
    setBusy(true);
    generateService
      .search(Number(conditionsVersionId), code.trim())
      .then((found) => {
        setMatches(found);
        if (found.length === 1) return resolveRow(found[0].rowIndex);
        return undefined;
      })
      .catch((e: any) => setError(e.message))
      .finally(() => setBusy(false));
  }

  function resolveRow(idx: number): Promise<void> {
    setRowIndex(idx);
    setError(null);
    setBusy(true);
    return generateService
      .getMappedValues(Number(conditionsVersionId), Number(templateId), idx)
      .then((v) => setValues(v))
      .catch((e: any) => setError(e.message))
      .finally(() => setBusy(false));
  }

  function download(): void {
    if (!values || rowIndex === null) return;
    setBusy(true);
    setError(null);
    setInfo(null);
    generateService
      .generate({
        templateId: Number(templateId),
        conditionsVersionId: Number(conditionsVersionId),
        codeSousSegment: code.trim(),
        values,
        traitePar,
      })
      .then(({ blob, filename }) => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        setInfo('Contrat généré et téléchargé.');
      })
      .catch((e: any) => setError(e.message))
      .finally(() => setBusy(false));
  }

  return (
    <div>
      <h2>Générer un contrat</h2>
      <p className={styles.muted}>
        Sélectionnez un template et un fichier de conditions, saisissez le code sous-segment, relisez le
        formulaire pré-rempli puis téléchargez le contrat.
      </p>

      {error && <div className={`${styles.alert} ${styles.alertError}`}>{error}</div>}
      {info && <div className={`${styles.alert} ${styles.alertSuccess}`}>{info}</div>}

      <div className={styles.panel}>
        <div className={styles.formRow}>
          <div className={styles.field}>
            <label>Template de contrat</label>
            <select
              value={templateId}
              onChange={(e) => {
                setTemplateId(e.target.value);
                resetSearch();
              }}
            >
              {templates.map((t) => (
                <option key={t.id} value={String(t.id)}>
                  {t.libelle} {t.departement ? `(${t.departement})` : ''} — v{t.version}
                </option>
              ))}
            </select>
            {templates.length === 0 && (
              <div className={`${styles.muted} ${styles.small}`}>
                Aucun template avec un mapping complet n'est disponible pour l'instant.
              </div>
            )}
          </div>
          <div className={styles.field}>
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
                <option key={c.id} value={String(c.id)}>
                  {c.nomFichier} {c.estActive ? '(active)' : ''}
                </option>
              ))}
            </select>
          </div>
          <div className={styles.field}>
            <label>Traité par (optionnel)</label>
            <input type="text" value={traitePar} onChange={(e) => setTraitePar(e.target.value)} />
          </div>
        </div>

        <div className={styles.formRow}>
          <div className={styles.field}>
            <label>Code sous-segment</label>
            <input
              type="search"
              value={code}
              onChange={(e) => setCode(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && runSearch()}
              placeholder="ex : SS001"
            />
          </div>
          <button className={styles.btn} disabled={busy || !templateId || !conditionsVersionId || !code.trim()} onClick={runSearch}>
            Rechercher
          </button>
        </div>

        {matches && matches.length > 1 && rowIndex === null && (
          <div style={{ marginTop: '1rem' }}>
            <p className={`${styles.muted} ${styles.small}`}>
              Plusieurs lignes correspondent à ce code — sélectionnez la ligne concernée :
            </p>
            <table className={styles.table}>
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
                      <button className={styles.btnSecondary} onClick={() => resolveRow(m.rowIndex)}>
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
        <div className={styles.panel}>
          <h3 style={{ marginBottom: 0 }}>Formulaire — relecture et correction</h3>
          <p className={`${styles.muted} ${styles.small}`}>
            Les champs sont pré-remplis à partir du fichier de conditions ; vous pouvez les corriger avant
            génération. Un champ vide signifie que la colonne mappée est vide — complétez-le si besoin.
          </p>
          <div className={styles.fieldGrid} style={{ marginTop: '1rem' }}>
            {Object.entries(values).map(([variable, value]) => (
              <div className={styles.field} key={variable}>
                <label>{variable}</label>
                <input type="text" value={value} onChange={(e) => setValues({ ...values, [variable]: e.target.value })} />
              </div>
            ))}
          </div>
          <button className={styles.btn} style={{ marginTop: '1rem' }} disabled={busy} onClick={download}>
            {busy ? 'Génération…' : 'Télécharger le contrat'}
          </button>
        </div>
      )}
    </div>
  );
}
