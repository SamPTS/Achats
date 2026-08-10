import * as React from 'react';
import { SPFI } from '@pnp/sp';
import styles from './AchatsContrats.module.scss';
import { IListsConfig } from '../services/spClient';
import { ensureSharePointStructure } from '../services/provisioning';
import { ConditionsService } from '../services/conditionsService';
import { TemplatesService } from '../services/templatesService';
import { GenerateService } from '../services/generateService';
import ConditionsTab from './ConditionsTab';
import TemplatesTab from './TemplatesTab';
import GenerateTab from './GenerateTab';

export interface IAppProps {
  sp: SPFI;
  config: IListsConfig;
}

type TabKey = 'conditions' | 'templates' | 'generer';

export default function App(props: IAppProps): JSX.Element {
  const { sp, config } = props;
  const [tab, setTab] = React.useState<TabKey>('conditions');
  const [ready, setReady] = React.useState(false);
  const [initError, setInitError] = React.useState<string | null>(null);

  const conditionsService = React.useMemo(() => new ConditionsService(sp, config), [sp, config]);
  const templatesService = React.useMemo(() => new TemplatesService(sp, config), [sp, config]);
  const generateService = React.useMemo(() => new GenerateService(sp, config), [sp, config]);

  const initialize = React.useCallback(() => {
    setInitError(null);
    ensureSharePointStructure(sp, config)
      .then(() => setReady(true))
      .catch((e: any) =>
        setInitError(
          `Impossible d'initialiser les listes/bibliothèque SharePoint (droits insuffisants ?) : ${e.message}`
        )
      );
  }, [sp, config]);

  React.useEffect(() => {
    initialize();
  }, [initialize]);

  if (!ready) {
    return (
      <div className={styles.app}>
        {initError ? (
          <div className={styles.alert + ' ' + styles.alertError}>
            {initError}
            <div className="mt1">
              <button className={styles.btn} onClick={initialize}>
                Réessayer
              </button>
            </div>
          </div>
        ) : (
          <div className={styles.alert + ' ' + styles.alertInfo}>Initialisation du stockage SharePoint…</div>
        )}
      </div>
    );
  }

  return (
    <div className={styles.app}>
      <div className={styles.tabs}>
        <button className={`${styles.tab} ${tab === 'conditions' ? styles.tabActive : ''}`} onClick={() => setTab('conditions')}>
          1. Conditions commerciales
        </button>
        <button className={`${styles.tab} ${tab === 'templates' ? styles.tabActive : ''}`} onClick={() => setTab('templates')}>
          2. Templates de contrats
        </button>
        <button className={`${styles.tab} ${tab === 'generer' ? styles.tabActive : ''}`} onClick={() => setTab('generer')}>
          3. Générer un contrat
        </button>
      </div>

      {tab === 'conditions' && <ConditionsTab service={conditionsService} />}
      {tab === 'templates' && (
        <TemplatesTab templatesService={templatesService} conditionsService={conditionsService} />
      )}
      {tab === 'generer' && (
        <GenerateTab generateService={generateService} conditionsService={conditionsService} />
      )}
    </div>
  );
}
