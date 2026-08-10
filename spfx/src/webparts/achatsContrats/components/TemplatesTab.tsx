import * as React from 'react';
import styles from './AchatsContrats.module.scss';
import { TemplatesService } from '../services/templatesService';
import { ConditionsService } from '../services/conditionsService';
import { IConditionsVersion, ITemplate } from '../models/types';

export interface ITemplatesTabProps {
  templatesService: TemplatesService;
  conditionsService: ConditionsService;
}

function fmtDate(iso: string): string {
  return new Date(iso).toLocaleString('fr-FR');
}

function statusBadge(statut: ITemplate['statutMapping']): JSX.Element {
  if (statut === 'complet') return <span className={`${styles.badge} ${styles.badgeOk}`}>Mapping complet</span>;
  if (statut === 'incomplet') return <span className={`${styles.badge} ${styles.badgeWarn}`}>Mapping incomplet</span>;
  return <span className={`${styles.badge} ${styles.badgeDanger}`}>Mapping absent</span>;
}

export default function TemplatesTab(props: ITemplatesTabProps): JSX.Element {
  const { templatesService, conditionsService } = props;
  const [templates, setTemplates] = React.useState<ITemplate[]>([]);
  const [conditions, setConditions] = React.useState<IConditionsVersion[]>([]);
  const [selected, setSelected] = React.useState<ITemplate | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [showUpload, setShowUpload] = React.useState(false);

  const refresh = React.useCallback(async () => {
    const [t, c] = await Promise.all([templatesService.listLatestPerGroup(), conditionsService.listVersions()]);
    setTemplates(t);
    setConditions(c);
    if (selected) {
      const fresh = t.find((x) => x.id === selected.id);
      setSelected(fresh ?? null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [templatesService, conditionsService]);

  React.useEffect(() => {
    refresh().catch((e) => setError(e.message));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div>
      <h2>Templates de contrats</h2>
      <p className={styles.muted}>
        Dépôt et historisation des templates Word (variables au format <code>{'{{Variable}}'}</code>) et
        correspondance avec les colonnes du fichier de conditions commerciales.
      </p>

      {error && <div className={`${styles.alert} ${styles.alertError}`}>{error}</div>}

      <div className={styles.panel}>
        <button className={styles.btn} onClick={() => setShowUpload((v) => !v)}>
          {showUpload ? 'Annuler' : '+ Déposer un template'}
        </button>
        {showUpload && (
          <UploadTemplateForm
            templatesService={templatesService}
            onDone={async () => {
              setShowUpload(false);
              await refresh();
            }}
          />
        )}
      </div>

      <div className={styles.panel}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>Libellé</th>
              <th>Département</th>
              <th>Version</th>
              <th>Dépôt</th>
              <th>Variables</th>
              <th>Statut mapping</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {templates.map((t) => (
              <tr key={t.id}>
                <td>
                  <a href={templatesService.downloadUrl(t)}>{t.libelle}</a>
                </td>
                <td>{t.departement || <span className={styles.muted}>—</span>}</td>
                <td>v{t.version}</td>
                <td>{fmtDate(t.dateDepot)}</td>
                <td>{t.variables.length}</td>
                <td>{statusBadge(t.statutMapping)}</td>
                <td>
                  <button className={styles.btnSecondary} onClick={() => setSelected(t)}>
                    Gérer le mapping
                  </button>
                </td>
              </tr>
            ))}
            {templates.length === 0 && (
              <tr>
                <td colSpan={7} className={styles.muted}>
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
          conditions={conditions}
          templatesService={templatesService}
          onClose={() => setSelected(null)}
          onChanged={refresh}
        />
      )}
    </div>
  );
}

function UploadTemplateForm(props: {
  templatesService: TemplatesService;
  onDone: () => Promise<void>;
}): JSX.Element {
  const { templatesService, onDone } = props;
  const [libelle, setLibelle] = React.useState('');
  const [departement, setDepartement] = React.useState('');
  const [deposePar, setDeposePar] = React.useState('');
  const [error, setError] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState(false);
  const fileInput = React.useRef<HTMLInputElement>(null);

  function submit(): void {
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
    templatesService
      .upload({ file, libelle, departement, deposePar })
      .then(onDone)
      .catch((e: any) => setError(e.message))
      .finally(() => setBusy(false));
  }

  return (
    <div style={{ marginTop: '1rem' }}>
      {error && <div className={`${styles.alert} ${styles.alertError}`}>{error}</div>}
      <div className={styles.formRow}>
        <div className={styles.field}>
          <label>Libellé *</label>
          <input type="text" value={libelle} onChange={(e) => setLibelle(e.target.value)} placeholder="Contrat Beauty — Capillaire" />
        </div>
        <div className={styles.field}>
          <label>Département (optionnel)</label>
          <input type="text" value={departement} onChange={(e) => setDepartement(e.target.value)} />
        </div>
        <div className={styles.field}>
          <label>Traité par (optionnel)</label>
          <input type="text" value={deposePar} onChange={(e) => setDeposePar(e.target.value)} />
        </div>
        <div className={styles.field}>
          <label>Fichier .docx *</label>
          <input ref={fileInput} type="file" accept=".docx" />
        </div>
      </div>
      <button className={styles.btn} disabled={busy} onClick={submit}>
        {busy ? 'Dépôt en cours…' : 'Déposer'}
      </button>
    </div>
  );
}

function MappingPanel(props: {
  template: ITemplate;
  conditions: IConditionsVersion[];
  templatesService: TemplatesService;
  onClose: () => void;
  onChanged: () => Promise<void>;
}): JSX.Element {
  const { template, conditions, templatesService, onClose, onChanged } = props;
  const activeConditions = conditions.find((c) => c.estActive) ?? null;
  const [conditionsVersionId, setConditionsVersionId] = React.useState<string>(
    activeConditions ? String(activeConditions.id) : ''
  );
  const [error, setError] = React.useState<string | null>(null);
  const [info, setInfo] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState(false);
  const fileInput = React.useRef<HTMLInputElement>(null);

  const refColumns = conditions.find((c) => String(c.id) === conditionsVersionId)?.colonnes ?? [];

  function downloadBlank(): void {
    templatesService.buildBlankMapping(template, refColumns).then((blob) => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `mapping_${template.libelle}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    });
  }

  function importMapping(): void {
    const file = fileInput.current?.files?.[0];
    if (!file) {
      setError('Merci de sélectionner le fichier de mapping rempli.');
      return;
    }
    setBusy(true);
    setError(null);
    setInfo(null);
    templatesService
      .importMapping(template.id, file, refColumns)
      .then(() => {
        setInfo('Mapping importé avec succès.');
        return onChanged();
      })
      .catch((e: any) => setError(e.message))
      .finally(() => setBusy(false));
  }

  function markFree(variable: string): void {
    templatesService
      .setMappingLine(template.id, variable, { statut: 'libre', colonneCorrespondante: null })
      .then(onChanged)
      .catch((e: any) => setError(e.message));
  }

  function assignColumn(variable: string, colonne: string): void {
    templatesService
      .setMappingLine(template.id, variable, { statut: colonne ? 'mappee' : 'manquante', colonneCorrespondante: colonne || null })
      .then(onChanged)
      .catch((e: any) => setError(e.message));
  }

  return (
    <div className={styles.panel}>
      <div className={styles.formRow} style={{ justifyContent: 'space-between' }}>
        <h3 style={{ margin: 0 }}>
          Mapping — {template.libelle} (v{template.version})
        </h3>
        <button className={styles.btnSecondary} onClick={onClose}>
          Fermer
        </button>
      </div>

      <div className={styles.formRow}>
        <div className={styles.field}>
          <label>Fichier de conditions de référence</label>
          <select value={conditionsVersionId} onChange={(e) => setConditionsVersionId(e.target.value)}>
            {conditions.map((c) => (
              <option key={c.id} value={String(c.id)}>
                {c.nomFichier} {c.estActive ? '(active)' : ''}
              </option>
            ))}
          </select>
        </div>
        <button className={styles.btnSecondary} onClick={downloadBlank}>
          Télécharger le mapping vierge
        </button>
      </div>

      {error && <div className={`${styles.alert} ${styles.alertError}`}>{error}</div>}
      {info && <div className={`${styles.alert} ${styles.alertSuccess}`}>{info}</div>}

      <div className={styles.formRow}>
        <input ref={fileInput} type="file" accept=".xlsx,.csv" />
        <button className={styles.btn} disabled={busy} onClick={importMapping}>
          {busy ? 'Import…' : 'Importer le mapping rempli'}
        </button>
      </div>

      <table className={styles.table} style={{ marginTop: '1rem' }}>
        <thead>
          <tr>
            <th>Variable</th>
            <th>Statut</th>
            <th>Colonne correspondante</th>
            <th />
          </tr>
        </thead>
        <tbody>
          {template.mappings.map((m) => (
            <tr key={m.variable}>
              <td>
                <code>{`{{${m.variable}}}`}</code>
              </td>
              <td>
                {m.statut === 'mappee' && <span className={`${styles.badge} ${styles.badgeOk}`}>Mappée</span>}
                {m.statut === 'libre' && <span className={`${styles.badge} ${styles.badgeMuted}`}>Saisie libre</span>}
                {m.statut === 'manquante' && <span className={`${styles.badge} ${styles.badgeDanger}`}>Manquante</span>}
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
                  <button className={styles.btnSecondary} onClick={() => markFree(m.variable)}>
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
