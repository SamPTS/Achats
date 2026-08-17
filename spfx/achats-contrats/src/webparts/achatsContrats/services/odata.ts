/**
 * Échappe une valeur pour l'insérer dans un filtre OData construit par interpolation de chaîne
 * (`.filter(\`Champ eq '${valeur}'\`)`). Convention OData standard : un guillemet simple s'échappe
 * en le doublant.
 *
 * Nécessaire même quand la valeur vient aujourd'hui toujours d'un identifiant interne sûr (un ID
 * SharePoint numérique converti en chaîne, ou un uuid généré côté client) : ce code s'exécute
 * entièrement dans le navigateur, sans frontière serveur — n'importe quel utilisateur connecté
 * peut appeler ces fonctions de service directement depuis la console du navigateur avec
 * n'importe quel argument. Sans cet échappement, une valeur comme `x' or 1 eq 1 or TemplateId eq 'x`
 * élargirait la requête au-delà de ce que la fonction est censée filtrer.
 */
export function odataEscape(value: string): string {
  return value.replace(/'/g, "''");
}
