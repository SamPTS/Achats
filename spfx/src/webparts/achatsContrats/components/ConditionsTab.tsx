import * as React from 'react';
import styles from './AchatsContrats.module.scss';
import { ConditionsService } from '../services/conditionsService';
import { IConditionsVersion } from '../models/types';

export interface IConditionsTabProps {
  service: ConditionsService;
}

function fmtDate(iso: string): string {
  return new Date(iso).toLocaleString('fr-FR');
}

export default function ConditionsTab(props: IConditionsTabProps): JSX.Element {
  const { service } = props;
  const [versions, setVersions] = React.useState<IConditionsVersion[]>([]);
  const [deposePar, setDeposePar] = React.useState('');
  const [dragOver, setDragOver] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [info, setInfo] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState(false);
  const fileInput = React.useRef<HTMLInputElement>(null);

  const refresh = React.useCallback(() => {
    return service.listVersions().then(setVersions);
  }, [service]);

  React.useEffect(() => {
    refresh().catch((e) => setError(e.message));
  }, [refresh]);

  function handleFile(file: File): void {
    setError(null);
    setInfo(null);
    setBusy(true);
    service
      .upload(file, deposePar)
      .then((created) => {
        const warn =
          created.doublonsDetectes > 0
            ? ` ⚠ ${created.doublonsDetectes} code(s) sous-segment en doublon détecté(s).`
            : '';
        setInfo(
          `Fichier déposé : ${created.nomFichier} — ${created.nbLignes} lignes, ${created.nbColonnes} colonnes.${warn}`
        );
        return refresh();
      })
      .catch((e: any) => setError(e.message))
      .finally(() => setBusy(false));
  }

  return (
    <div>
      <h2>Conditions commerciales</h2>
      <p className={styles.muted}>
        Dépôt et historisation des fichiers Excel de conditions commerciales. Chaque dépôt crée une nouvelle
        version — aucune version n'est jamais écrasée.
      </p>

      {error && <div className={`${styles.alert} ${styles.alertError}`}>{error}</div>}
      {info && <div className={`${styles.alert} ${styles.alertSuccess}`}>{info}</div>}

      <div className={styles.panel}>
        <div className={styles.formRow}>
          <div className={styles.field}>
            <label>Traité par (optionnel)</label>
            <input type="text" value={deposePar} onChange={(e) => setDeposePar(e.target.value)} />
          </div>
        </div>
        <div
          className={`${styles.dropzone} ${dragOver ? styles.dropzoneOver : ''}`}
          onDragOver={(e) => {
            e.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragOver(false);
            const file = e.dataTransfer.files?.[0];
            if (file) handleFile(file);
          }}
          onClick={() => fileInput.current?.click()}
        >
          {busy ? 'Dépôt en cours…' : 'Glissez-déposez un fichier .xlsx ici, ou cliquez pour le sélectionner'}
          <input
            ref={fileInput}
            type="file"
            accept=".xlsx"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) handleFile(file);
              e.target.value = '';
            }}
          />
        </div>
      </div>

      <div className={styles.panel}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>Statut</th>
              <th>Fichier</th>
              <th>Dépôt</th>
              <th>Lignes</th>
              <th>Colonnes</th>
              <th>Code sous-segment</th>
              <th>Alertes</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {versions.map((v) => (
              <VersionRow key={v.id} version={v} service={service} onChanged={refresh} />
            ))}
            {versions.length === 0 && (
              <tr>
                <td colSpan={8} className={styles.muted}>
                  Aucun fichier déposé pour l'instant.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function VersionRow(props: {
  version: IConditionsVersion;
  service: ConditionsService;
  onChanged: () => Promise<void>;
}): JSX.Element {
  const { version, service, onChanged } = props;
  const [busy, setBusy] = React.useState(false);
  const [choosingCol, setChoosingCol] = React.useState(false);
  const [col, setCol] = React.useState(version.colonneCodeSousSegment ?? '');

  function activate(): void {
    setBusy(true);
    service.activate(version.id).then(onChanged).finally(() => setBusy(false));
  }

  function archive(): void {
    // eslint-disable-next-line no-alert
    if (!confirm(`Archiver la version "${version.nomFichier}" (${fmtDate(version.dateDepot)}) ?`)) return;
    setBusy(true);
    service.archive(version.id).then(onChanged).finally(() => setBusy(false));
  }

  function saveCol(): void {
    setBusy(true);
    service
      .setCodeColumn(version.id, col)
      .then(() => setChoosingCol(false))
      .then(onChanged)
      .finally(() => setBusy(false));
  }

  return (
    <tr>
      <td>
        {version.estActive ? (
          <span className={`${styles.badge} ${styles.badgeOk}`}>Active</span>
        ) : (
          <button className={styles.btnSecondary} disabled={busy} onClick={activate}>
            Activer
          </button>
        )}
      </td>
      <td>
        <a href={service.downloadUrl(version)}>{version.nomFichier}</a>
        {version.deposePar && <div className={`${styles.muted} ${styles.small}`}>par {version.deposePar}</div>}
      </td>
      <td>{fmtDate(version.dateDepot)}</td>
      <td>
        {version.nbLignes}
        {version.nbLignesVides > 0 && (
          <div className={`${styles.muted} ${styles.small}`}>{version.nbLignesVides} ligne(s) vide(s)</div>
        )}
      </td>
      <td>{version.nbColonnes}</td>
      <td>
        {choosingCol ? (
          <div style={{ display: 'flex', gap: '0.4rem' }}>
            <select value={col} onChange={(e) => setCol(e.target.value)}>
              <option value="">—</option>
              {version.colonnes.map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
            </select>
            <button className={styles.btn} disabled={busy || !col} onClick={saveCol}>
              OK
            </button>
          </div>
        ) : version.colonneCodeSousSegment ? (
          <span>
            {version.colonneCodeSousSegment}{' '}
            <button className={styles.btnSecondary} onClick={() => setChoosingCol(true)}>
              modifier
            </button>
          </span>
        ) : (
          <button className={styles.btnDanger} onClick={() => setChoosingCol(true)}>
            À désigner
          </button>
        )}
      </td>
      <td>
        {version.doublonsDetectes > 0 && (
          <span className={`${styles.badge} ${styles.badgeWarn}`}>{version.doublonsDetectes} doublon(s)</span>
        )}
      </td>
      <td>
        <button className={styles.btnDanger} disabled={busy || version.estActive} onClick={archive}>
          Archiver
        </button>
      </td>
    </tr>
  );
}
