import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import * as conditionsService from '../services/conditionsService';
import { logAndGetMessage } from '../services/errorLog';
import type { ConditionsVersion } from '../model/types';

function fmtDate(iso: string): string {
  return new Date(iso).toLocaleString('fr-FR');
}

export default function ConditionsTab(): JSX.Element {
  const [versions, setVersions] = useState<ConditionsVersion[]>([]);
  const [deposePar, setDeposePar] = useState('');
  const [dragOver, setDragOver] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);

  // --- Lien vers un fichier existant ailleurs sur le site (alternative au dépôt classique) ---
  const [showLink, setShowLink] = useState(false);
  const [linkRef, setLinkRef] = useState('');
  const [linkBusy, setLinkBusy] = useState(false);

  async function refresh(): Promise<void> {
    const list = await conditionsService.listConditions();
    // Efface une éventuelle erreur d'un rechargement précédent : sans ça, une erreur transitoire
    // (ex. lecture juste après provisionnement à froid) reste affichée indéfiniment même une fois
    // qu'un rechargement suivant a réussi et que les données s'affichent normalement — trompeur
    // pour l'utilisateur, qui voit un message d'erreur sur un écran qui fonctionne en réalité.
    setError(null);
    setVersions(list);
  }

  useEffect(() => {
    refresh().catch((e) => setError(logAndGetMessage(e, 'ConditionsTab.refresh (chargement initial)')));
  }, []);

  async function handleFile(file: File): Promise<void> {
    setError(null);
    setInfo(null);
    setBusy(true);
    try {
      const created = await conditionsService.uploadConditions(file, deposePar);
      const warn =
        created.doublonsDetectes > 0 ? ` ⚠ ${created.doublonsDetectes} code(s) sous-segment en doublon détecté(s).` : '';
      setInfo(`Fichier déposé : ${created.nomFichier} — ${created.nbLignes} lignes, ${created.nbColonnes} colonnes.${warn}`);
      await refresh();
    } catch (e) {
      setError(logAndGetMessage(e, 'ConditionsTab.handleFile (dépôt d\'un fichier)'));
    } finally {
      setBusy(false);
    }
  }

  async function handleLink(): Promise<void> {
    if (!linkRef.trim()) return;
    setError(null);
    setInfo(null);
    setLinkBusy(true);
    try {
      const created = await conditionsService.linkExternalConditions(linkRef, deposePar);
      const warn =
        created.doublonsDetectes > 0 ? ` ⚠ ${created.doublonsDetectes} code(s) sous-segment en doublon détecté(s).` : '';
      setInfo(
        `Fichier lié : ${created.nomFichier} — ${created.nbLignes} lignes, ${created.nbColonnes} colonnes.${warn} ` +
          'Modifiez-le directement à son emplacement d\'origine : les changements seront pris en compte automatiquement.',
      );
      setLinkRef('');
      setShowLink(false);
      await refresh();
    } catch (e) {
      setError(logAndGetMessage(e, 'ConditionsTab.handleLink (lien vers un fichier existant)'));
    } finally {
      setLinkBusy(false);
    }
  }

  return (
    <div>
      <h1>Conditions commerciales</h1>
      <p className="subtitle">
        Dépôt et historisation des fichiers Excel de conditions commerciales fournisseurs. Chaque dépôt crée une
        nouvelle version — aucune version n&apos;est jamais écrasée.
      </p>

      {error && <div className="alert error">{error}</div>}
      {info && <div className="alert success">{info}</div>}

      <div className="panel">
        <div className="form-row">
          <div>
            <label>Traité par (optionnel)</label>
            <input
              type="text"
              value={deposePar}
              onChange={(e) => setDeposePar(e.target.value)}
              placeholder="Nom / initiales"
            />
          </div>
        </div>
        <div
          className={`dropzone${dragOver ? ' dragover' : ''}`}
          role="button"
          tabIndex={0}
          aria-label="Déposer ou sélectionner un fichier .xlsx de conditions commerciales"
          onDragOver={(e) => {
            e.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragOver(false);
            const file = e.dataTransfer.files?.[0];
            if (file) handleFile(file).catch((err) => setError(logAndGetMessage(err, 'ConditionsTab.handleFile (glisser-déposer)')));
          }}
          onClick={() => fileInput.current?.click()}
          onKeyDown={(e) => {
            // Zone rendue accessible au clavier (Entrée/Espace) : un simple <div onClick> ne peut
            // pas être activé sans souris, alors que c'est le principal point d'entrée de cet onglet.
            if (e.key === 'Enter' || e.key === ' ') {
              e.preventDefault();
              fileInput.current?.click();
            }
          }}
        >
          {busy ? 'Dépôt en cours…' : 'Glissez-déposez un fichier .xlsx ici, ou cliquez pour le sélectionner'}
          <input
            ref={fileInput}
            type="file"
            accept=".xlsx"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) handleFile(file).catch((err) => setError(logAndGetMessage(err, 'ConditionsTab.handleFile (glisser-déposer)')));
              e.target.value = '';
            }}
          />
        </div>
        <div className="mt1">
          <button type="button" className="secondary" onClick={() => setShowLink((v) => !v)}>
            {showLink ? 'Annuler' : "Lier un fichier existant plutôt qu'en déposer un"}
          </button>
          {showLink && (
            <div className="mt1">
              <p className="muted small">
                Plutôt qu&apos;une copie déposée ici, référence un fichier Excel déjà présent ailleurs sur le site :
                vous pourrez continuer à le modifier directement à son emplacement d&apos;origine, l&apos;application
                relira automatiquement son contenu à jour à chaque recherche/génération — sans jamais avoir besoin
                de le redéposer.
              </p>
              <div className="form-row">
                <div style={{ flex: 1 }}>
                  <label>Lien du fichier (menu SharePoint « Copier le lien », ou adresse de la page du fichier)</label>
                  <input
                    type="text"
                    value={linkRef}
                    onChange={(e) => setLinkRef(e.target.value)}
                    placeholder="https://…/Documents partagés/Conditions.xlsx"
                    style={{ width: '100%' }}
                  />
                </div>
                <button disabled={linkBusy || !linkRef.trim()} onClick={handleLink}>
                  {linkBusy ? 'Liaison…' : 'Lier ce fichier'}
                </button>
              </div>
            </div>
          )}
        </div>
      </div>

      <div className="panel">
        <table>
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
              <VersionRow key={v.id} version={v} onChanged={refresh} />
            ))}
            {versions.length === 0 && (
              <tr>
                <td colSpan={8} className="muted">
                  Aucun fichier déposé pour l&apos;instant.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function VersionRow({
  version,
  onChanged,
}: {
  version: ConditionsVersion;
  onChanged: () => Promise<void>;
}): JSX.Element {
  const [busy, setBusy] = useState(false);
  const [choosingCol, setChoosingCol] = useState(false);
  const [col, setCol] = useState(version.colonneCodeSousSegment ?? '');
  const [error, setError] = useState<string | null>(null);

  async function activate(): Promise<void> {
    setBusy(true);
    setError(null);
    try {
      await conditionsService.activateConditions(version.id);
      await onChanged();
    } catch (e) {
      setError(logAndGetMessage(e, 'ConditionsTab.VersionRow.activate'));
    } finally {
      setBusy(false);
    }
  }

  async function archive(): Promise<void> {
    if (!confirm(`Archiver la version "${version.nomFichier}" (${fmtDate(version.dateDepot)}) ?`)) return;
    setBusy(true);
    setError(null);
    try {
      await conditionsService.archiveConditions(version.id);
      await onChanged();
    } catch (e) {
      setError(logAndGetMessage(e, 'ConditionsTab.VersionRow.archive'));
    } finally {
      setBusy(false);
    }
  }

  async function saveCol(): Promise<void> {
    setBusy(true);
    setError(null);
    try {
      await conditionsService.setConditionsCodeColumn(version.id, col);
      setChoosingCol(false);
      await onChanged();
    } catch (e) {
      setError(logAndGetMessage(e, 'ConditionsTab.VersionRow.saveCol'));
    } finally {
      setBusy(false);
    }
  }

  async function remove(): Promise<void> {
    const consequenceFichier = version.externe
      ? "seul le lien sera retiré : le fichier d'origine, ailleurs sur le site, ne sera pas supprimé."
      : 'le fichier déposé sera également supprimé.';
    if (
      !confirm(
        `Supprimer définitivement la version "${version.nomFichier}" (${fmtDate(version.dateDepot)}) ? Cette action est irréversible : ${consequenceFichier}`,
      )
    )
      return;
    setBusy(true);
    setError(null);
    try {
      await conditionsService.deleteConditions(version.id);
      await onChanged();
    } catch (e) {
      setError(logAndGetMessage(e, 'ConditionsTab.VersionRow.remove'));
    } finally {
      setBusy(false);
    }
  }

  return (
    <tr>
      <td>
        {version.estActive ? (
          <span className="badge ok">Active</span>
        ) : (
          <button className="secondary" disabled={busy} onClick={activate}>
            Activer
          </button>
        )}
      </td>
      <td>
        <a href={version.cheminStockage} target="_blank" rel="noreferrer">
          {version.nomFichier}
        </a>{' '}
        {version.externe && (
          <span className="badge muted" title="Fichier référencé à son emplacement d'origine, jamais copié — modifiable directement là-bas.">
            Lien externe
          </span>
        )}
        {version.deposePar && <div className="muted small">par {version.deposePar}</div>}
      </td>
      <td>{fmtDate(version.dateDepot)}</td>
      <td>
        {version.nbLignes}
        {version.nbLignesVides > 0 && <div className="muted small">{version.nbLignesVides} ligne(s) vide(s)</div>}
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
            <button disabled={busy || !col} onClick={saveCol}>
              OK
            </button>
          </div>
        ) : version.colonneCodeSousSegment ? (
          <span>
            {version.colonneCodeSousSegment}{' '}
            <button className="icon-btn" title="Modifier la colonne" aria-label="Modifier la colonne" onClick={() => setChoosingCol(true)}>
              <Icon iconName="Edit" />
            </button>
          </span>
        ) : (
          <button className="danger" onClick={() => setChoosingCol(true)}>
            À désigner
          </button>
        )}
      </td>
      <td>{version.doublonsDetectes > 0 && <span className="badge warn">{version.doublonsDetectes} doublon(s)</span>}</td>
      <td>
        {error && <div className="alert error small">{error}</div>}
        <div style={{ display: 'flex', gap: '0.4rem' }}>
          <button
            className="icon-btn danger"
            title="Archiver"
            aria-label="Archiver"
            disabled={busy || version.estActive}
            onClick={archive}
          >
            <Icon iconName="Archive" />
          </button>
          <button
            className="icon-btn danger"
            title="Supprimer"
            aria-label="Supprimer"
            disabled={busy || version.estActive}
            onClick={remove}
          >
            <Icon iconName="Delete" />
          </button>
        </div>
      </td>
    </tr>
  );
}
