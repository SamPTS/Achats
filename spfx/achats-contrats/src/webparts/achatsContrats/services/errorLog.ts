/**
 * Journalise une erreur avant de l'afficher à l'écran, puis renvoie son message.
 *
 * Sans ça, une erreur interceptée par un try/catch React (comme celles affichées dans les
 * bandeaux rouges de l'interface) n'apparaît jamais comme une exception "Uncaught" dans la
 * console du navigateur : seul le message final (souvent générique, ex. "Cannot read properties
 * of undefined (reading 'Id')") est visible à l'écran, sans pile d'appels ni contexte permettant
 * de savoir quelle action l'a déclenchée — ce qui a rendu certains bugs très difficiles à
 * diagnostiquer à distance en conditions réelles. `context` doit identifier sans ambiguïté
 * l'action en cours (ex. "ConditionsTab.refresh").
 */
export function logAndGetMessage(e: unknown, context: string): string {
  // eslint-disable-next-line no-console
  console.error(`[achats-contrats] Erreur dans ${context} :`, e);
  return e instanceof Error ? e.message : String(e);
}
