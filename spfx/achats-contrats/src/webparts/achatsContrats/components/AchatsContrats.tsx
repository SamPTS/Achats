import * as React from 'react';
import { useEffect, useState } from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from './AchatsContrats.module.scss';
import type { IAchatsContratsProps } from './IAchatsContratsProps';
import { getSP } from '../services/spClient';
import { ensureProvisioned } from '../services/provisioning';
import { logAndGetMessage } from '../services/errorLog';
import ConditionsTab from './ConditionsTab';
import TemplatesTab from './TemplatesTab';
import GenerateTab from './GenerateTab';

type TabKey = 'conditions' | 'templates' | 'generer';

export default function AchatsContrats(props: IAchatsContratsProps): JSX.Element {
  const [tab, setTab] = useState<TabKey>('conditions');
  const [ready, setReady] = useState(false);
  const [error, setError] = useState<string | null>(null);
  // Force le remontage des onglets Templates/Générer après un aller-retour sur Templates,
  // pour qu'ils rechargent des données fraîches (pas de store partagé entre onglets ici).
  const [refreshKey, setRefreshKey] = useState(0);

  useEffect(() => {
    getSP(props.context);
    ensureProvisioned()
      .then(() => setReady(true))
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

  if (!ready) {
    return (
      <section className={styles.achatsContrats}>
        <Spinner size={SpinnerSize.medium} label="Préparation de l'application…" />
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
            const key = (item?.props.itemKey as TabKey) ?? 'conditions';
            setTab(key);
            setRefreshKey((k) => k + 1);
          }}
        >
          <PivotItem headerText="1. Conditions commerciales" itemKey="conditions" />
          <PivotItem headerText="2. Templates de contrats" itemKey="templates" />
          <PivotItem headerText="3. Générer un contrat" itemKey="generer" />
        </Pivot>
      </div>

      {tab === 'conditions' && <ConditionsTab key={`conditions-${refreshKey}`} />}
      {tab === 'templates' && <TemplatesTab key={`templates-${refreshKey}`} />}
      {tab === 'generer' && <GenerateTab key={`generer-${refreshKey}`} onGoToTemplates={() => setTab('templates')} />}
    </section>
  );
}
