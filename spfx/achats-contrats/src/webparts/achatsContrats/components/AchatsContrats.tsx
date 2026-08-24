import * as React from 'react';
import { useEffect, useState } from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import styles from './AchatsContrats.module.scss';
import type { IAchatsContratsProps } from './IAchatsContratsProps';
import { getSP } from '../services/spClient';
import { ensureProvisioned } from '../services/provisioning';
import { getAppPermissions, AppPermissions } from '../services/permissions';
import { logAndGetMessage } from '../services/errorLog';
import ConditionsTab from './ConditionsTab';
import TemplatesTab from './TemplatesTab';
import GenerateTab from './GenerateTab';

// Enregistre la police d'icônes Fluent UI une seule fois au chargement du bundle (pas dans le
// composant, pour ne pas la répéter à chaque rendu). Nécessaire pour que les boutons à
// pictogramme (Edit/Archive/Delete) affichent réellement un glyphe : sans cet appel, l'icône
// reste invisible (aucune erreur, juste un bouton vide) tant qu'aucune autre partie de la page
// n'a déjà enregistré ce jeu d'icônes — observé dans une simulation locale sans le chrome
// SharePoint habituel. Un double enregistrement (si SharePoint l'a déjà fait) ne produit qu'un
// avertissement console bénin ("icon was re-registered"), jamais une erreur.
initializeIcons();

type TabKey = 'conditions' | 'templates' | 'generer';

export default function AchatsContrats(props: IAchatsContratsProps): JSX.Element {
  const [tab, setTab] = useState<TabKey | null>(null);
  const [ready, setReady] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [perms, setPerms] = useState<AppPermissions | null>(null);
  // Force le remontage des onglets Templates/Générer après un aller-retour sur Templates,
  // pour qu'ils rechargent des données fraîches (pas de store partagé entre onglets ici).
  const [refreshKey, setRefreshKey] = useState(0);

  useEffect(() => {
    getSP(props.context);
    ensureProvisioned()
      .then(async () => {
        const p = await getAppPermissions();
        setPerms(p);
        // Onglet de départ : le premier auquel l'utilisateur a droit — un Membre sans accès aux
        // deux premiers onglets ne doit jamais atterrir sur un onglet Conditions vide/en erreur.
        setTab(p.isOwner ? 'conditions' : p.canGenerate ? 'generer' : null);
        setReady(true);
      })
      .catch((e) =>
        setError(
          `Impossible de préparer les listes SharePoint nécessaires : ${logAndGetMessage(e, 'AchatsContrats.ensureProvisioned')}. ` +
            "Vérifiez que vous disposez des droits de gestion des listes sur ce site (nécessaire uniquement pour l'initialisation).",
        ),
      );
  }, []);

  if (error) {
    return (
      <section className={styles.achatsContrats}>
        <div className="alert error">{error}</div>
      </section>
    );
  }

  if (!ready || !perms) {
    return (
      <section className={styles.achatsContrats}>
        <Spinner size={SpinnerSize.medium} label="Préparation de l'application…" />
      </section>
    );
  }

  if (!perms.isOwner && !perms.canGenerate) {
    return (
      <section className={styles.achatsContrats}>
        <div className={styles.brandHeader}>Achats — Contrats fournisseurs</div>
        <div className="alert error">
          Vous n&apos;avez pas les autorisations nécessaires pour utiliser cette application. Contactez un
          propriétaire de ce site pour obtenir l&apos;accès.
        </div>
      </section>
    );
  }

  return (
    <section className={styles.achatsContrats}>
      <div className={styles.brandHeader}>Achats — Contrats fournisseurs</div>
      <div className={styles.pivotWrap}>
        <Pivot
          selectedKey={tab}
          onLinkClick={(item) => {
            const key = (item?.props.itemKey as TabKey) ?? tab;
            setTab(key);
            setRefreshKey((k) => k + 1);
          }}
        >
          {/* Conditions commerciales et Templates de contrats réservés aux propriétaires du site
              (droits de gestion complets sur les fichiers/mappings) ; Générer un contrat ouvert
              aux membres, qui n'ont besoin que d'ajouter des éléments (journal des générations). */}
          {perms.isOwner && <PivotItem headerText="1. Conditions commerciales" itemKey="conditions" />}
          {perms.isOwner && <PivotItem headerText="2. Templates de contrats" itemKey="templates" />}
          {perms.canGenerate && <PivotItem headerText="3. Générer un contrat" itemKey="generer" />}
        </Pivot>
      </div>

      {tab === 'conditions' && perms.isOwner && <ConditionsTab key={`conditions-${refreshKey}`} />}
      {tab === 'templates' && perms.isOwner && <TemplatesTab key={`templates-${refreshKey}`} />}
      {tab === 'generer' && perms.canGenerate && (
        <GenerateTab
          key={`generer-${refreshKey}`}
          onGoToTemplates={perms.isOwner ? () => setTab('templates') : undefined}
        />
      )}
    </section>
  );
}
