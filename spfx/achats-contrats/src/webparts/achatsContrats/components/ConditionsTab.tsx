import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import * as conditionsService from '../services/conditionsService';
import { logAndGetMessage } from '../services/errorLog';
import type { ConditionsVersion } from '../model/types';

function fmtDate(iso: string): string {
  return new Date(iso).toLocaleString('fr-FR');
}

/**
 * Plus d'historique/versions ici (changement demandé par l'utilisateur) : cet onglet n'affiche
 * qu'un seul fichier de conditions "courant". Changer de fichier (nouveau dépôt ou nouveau lien)
 * supprime définitivement l'ancien — aucune trace ne subsiste, conformément au choix explicite de
 * ne plus garder d'historique. Le modèle de données sous-jacent (ConditionsVersions, EstActive...)
 * reste inchangé pour ne pas perturber le reste de l'application (génération, mapping des
 * templates) : seule cette interface se comporte désormais comme "un seul fichier à la fois".
 */
export default function ConditionsTab(): JSX.Element {
  const [version, setVersion] = useState<ConditionsVersion | null | undefined>(undefined); // undefined = chargement
  const [deposePar, setDeposePar] = useState('');
  const [dragOver, setDragOver] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);

  // --- Remplacement du fichier courant (dépôt, ou lien vers un fichier existant ailleurs) ---
  const [showReplace, setShowReplace] = useState(false);
  const [showLink, setShowLink] = useState(false);
  const [linkRef, setLinkRef] = useState('');
  const [linkBusy, setLinkBusy] = useState(false);

  // --- Modification de la colonne "code sous-segment" du fichier courant ---
  const [choosingCol, setChoosingCol] = useState(false);
  const [col, setCol] = useState('');
  const [colBusy, setColBusy] = useState(false);

  async function refresh(): Promise<void> {
    const list = await conditionsService.listConditions();
    // Efface une éventuelle erreur d'un rechargement précédent : sans ça, une erreur transitoire
    // (ex. lecture juste après provisionnement à froid) reste affichée indéfiniment même une fois
    // qu'un rechargement suivant a réussi et que les données s'affichent normalement — trompeur
    // pour l'utilisateur, qui voit un message d'erreur sur un écran qui fonctionne en réalité.
    setError(null);
    // S'il devait rester plus d'un enregistrement (ancien historique jamais nettoyé, ou état
    // transitoire), on affiche le plus récent — jamais une liste, conformément au nouveau
    // fonctionnement "un seul fichier courant".
    const current = list.length === 0 ? null : list.reduce((a, b) => (a.dateDepot > b.dateDepot ? a : b));
    setVersion(current);
    setCol(current?.colonneCodeSousSegment ?? '');
  }

  useEffect(() => {
    refresh().catch((e) => setError(logAndGetMessage(e, 'ConditionsTab.refresh (chargement initial)')));
  }, []);

  async function handleFile(file: File): Promise<void> {
    setError(null);
    setInfo(null);
    setBusy(true);
    try {
      const created = await conditionsService.replaceConditionsWithUpload(file, deposePar);
      const warn =
        created.doublonsDetectes > 0 ? ` ⚠ ${created.doublonsDetectes} code(s) sous-segment en doublon détecté(s).` : '';
      setInfo(`Fichier déposé : ${created.nomFichier} — ${created.nbLignes} lignes, ${created.nbColonnes} colonnes.${warn}`);
      setShowReplace(false);
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
      const created = await conditionsService.replaceConditionsWithLink(linkRef, deposePar);
      const warn =
        created.doublonsDetectes > 0 ? ` ⚠ ${created.doublonsDetectes} code(s) sous-segment en doublon détecté(s).` : '';
      setInfo(
        `Fichier lié : ${created.nomFichier} — ${created.nbLignes} lignes, ${created.nbColonnes} colonnes.${warn} ` +
          'Modifiez-le directement à son emplacement d\'origine : les changements seront pris en compte automatiquement.',
      );
      setLinkRef('');
      setShowLink(false);
      setShowReplace(false);
      await refresh();
    } catch (e) {
      setError(logAndGetMessage(e, 'ConditionsTab.handleLink (lien vers un fichier existant)'));
    } finally {
      setLinkBusy(false);
    }
  }

  async function saveCol(): Promise<void> {
    if (!version) return;
    setColBusy(true);
    setError(null);
    try {
      await conditionsService.setConditionsCodeColumn(version.id, col);
      setChoosingCol(false);
      await refresh();
    } catch (e) {
      setError(logAndGetMessage(e, 'ConditionsTab.saveCol'));
    } finally {
      setColBusy(false);
    }
  }

  function openReplace(): void {
    if (
      version &&
      !confirm(
        `Remplacer "${version.nomFichier}" ? Cette action est irréversible : le fichier actuel ${
          version.externe ? "sera dé-référencé (le fichier d'origine ne sera pas supprimé)" : 'sera définitivement supprimé'
        }, sans conserver d'historique.`,
      )
    ) {
      return;
    }
    setError(null);
    setInfo(null);
    setShowReplace(true);
  }

  const replaceForm = (
    <div className="panel">
      <div className="form-row">
        <div>
          <label>Traité par (optionnel)</label>
          <input type="text" value={deposePar} onChange={(e) => setDeposePar(e.target.value)} placeholder="Nom / initiales" />
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
            if (file) handleFile(file).catch((err) => setError(logAndGetMessage(err, 'ConditionsTab.handleFile (sélection)')));
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
              Plutôt qu&apos;une copie déposée ici, référence un fichier Excel déjà présent ailleurs sur le site : vous
              pourrez continuer à le modifier directement à son emplacement d&apos;origine, l&apos;application relira
              automatiquement son contenu à jour à chaque recherche/génération — sans jamais avoir besoin de le
              redéposer.
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
      <button className="secondary mt1" type="button" onClick={() => setShowReplace(false)}>
        Annuler
      </button>
    </div>
  );

  return (
    <div>
      <h1>Conditions commerciales</h1>
      <p className="subtitle">
        Un seul fichier de conditions commerciales à la fois — le remplacer (nouveau dépôt ou nouveau lien) efface
        définitivement le précédent, sans historique conservé.
      </p>

      {error && <div className="alert error">{error}</div>}
      {info && <div className="alert success">{info}</div>}

      {version === undefined ? (
        <div className="panel muted">Chargement…</div>
      ) : version === null ? (
        <>
          <div className="panel muted">Aucun fichier de conditions configuré pour l&apos;instant.</div>
          {replaceForm}
        </>
      ) : showReplace ? (
        replaceForm
      ) : (
        <div className="panel">
          <div className="form-row" style={{ justifyContent: 'space-between', alignItems: 'flex-start' }}>
            <div>
              <a href={version.cheminStockage} target="_blank" rel="noreferrer">
                {version.nomFichier}
              </a>{' '}
              {version.externe && (
                <span
                  className="badge muted"
                  title="Fichier référencé à son emplacement d'origine, jamais copié — modifiable directement là-bas."
                >
                  Lien externe
                </span>
              )}
              {version.deposePar && (
                <div className="muted small">
                  {version.externe ? 'lié' : 'déposé'} par {version.deposePar} le {fmtDate(version.dateDepot)}
                </div>
              )}
            </div>
            <button type="button" className="secondary" onClick={openReplace}>
              Changer de fichier
            </button>
          </div>

          <div className="form-row mt1">
            <div>
              <label>Lignes</label>
              <div>
                {version.nbLignes}
                {version.nbLignesVides > 0 && <span className="muted small"> ({version.nbLignesVides} vide(s))</span>}
              </div>
            </div>
            <div>
              <label>Colonnes</label>
              <div>{version.nbColonnes}</div>
            </div>
            <div>
              <label>Code sous-segment</label>
              {choosingCol ? (
                <div style={{ display: 'flex', gap: '0.4rem', flexWrap: 'wrap', alignItems: 'center' }}>
                  <select
                    value={col}
                    onChange={(e) => setCol(e.target.value)}
                    title={col}
                    style={{ maxWidth: '320px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                  >
                    <option value="">—</option>
                    {version.colonnes.map((c) => (
                      <option key={c} value={c} title={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                  <button disabled={colBusy || !col} onClick={saveCol}>
                    OK
                  </button>
                </div>
              ) : version.colonneCodeSousSegment ? (
                <span style={{ display: 'inline-flex', alignItems: 'center', gap: '0.3rem', maxWidth: '100%' }}>
                  <span
                    title={version.colonneCodeSousSegment}
                    style={{ maxWidth: '360px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                  >
                    {version.colonneCodeSousSegment}
                  </span>{' '}
                  <button
                    className="icon-btn"
                    title="Modifier la colonne"
                    aria-label="Modifier la colonne"
                    onClick={() => setChoosingCol(true)}
                  >
                    <Icon iconName="Edit" />
                  </button>
                </span>
              ) : (
                <button className="danger" onClick={() => setChoosingCol(true)}>
                  À désigner
                </button>
              )}
            </div>
          </div>

          {version.doublonsDetectes > 0 && (
            <div className="mt1">
              <span className="badge warn">{version.doublonsDetectes} code(s) sous-segment en doublon détecté(s)</span>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
