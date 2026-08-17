import express from 'express';
import path from 'path';
import fs from 'fs';
import http from 'http';
import https from 'https';
import { exec } from 'child_process';
import './db'; // initialise le schéma au démarrage
import conditionsRouter from './routes/conditions';
import templatesRouter from './routes/templates';
import generateRouter from './routes/generate';
import { getAppDir } from './runtimePaths';

const app = express();
// Pas de middleware CORS : le frontend est toujours servi depuis la même origine que l'API
// (même port en production ; en développement, le proxy Vite fait que le navigateur ne voit
// jamais de requête cross-origin — voir client/vite.config.ts). Un CORS ouvert à toute origine
// n'apporterait donc rien d'utile, seulement une surface d'attaque inutile pour une application
// sans authentification qui manipule des conditions commerciales confidentielles.
app.use(express.json({ limit: '5mb' }));

app.use('/api/conditions', conditionsRouter);
app.use('/api/templates', templatesRouter);
app.use('/api/generate', generateRouter);

app.get('/api/health', (_req, res) => res.json({ ok: true }));

// Sert le frontend buildé sur le même port, pour un lancement en un seul exécutable/script :
// évite d'avoir à démarrer un second serveur (Vite) pour tester l'appli. Le routeur front
// utilise le mode "hash" (#/...), donc aucune route "catch-all" n'est requise.
//
// Deux emplacements possibles selon le mode d'exécution :
// - "public" à côté de l'exécutable : c'est le mode de distribution de l'exe compilé avec
//   `bun build --compile` (voir server/package.json, build:exe) — Bun n'embarque pas les
//   fichiers statiques dans le binaire, il faut donc les livrer à côté (voir LISEZ-MOI).
// - client/dist, résolu depuis __dirname : en développement, ou pour un éventuel paquet
//   pkg/Node qui embarque ces fichiers dans son système de fichiers virtuel.
const clientDistCandidates = [
  path.join(getAppDir(), 'public'),
  path.join(__dirname, '..', '..', 'client', 'dist'),
];
const clientDist = clientDistCandidates.find((p) => fs.existsSync(p));
if (clientDist) {
  app.use(express.static(clientDist));
}

/**
 * Ouvre l'URL dans le navigateur par défaut du système. Utilisé pour que
 * l'exécutable packagé (.exe) se comporte comme le script
 * lancer-application.bat, qui ouvrait la page via `start` : sans cela,
 * double-cliquer sur l'exe démarre bien le serveur mais rien ne s'affiche
 * tant que l'utilisateur ne va pas lui-même sur localhost:4000.
 */
function openBrowser(url: string) {
  const cmd =
    process.platform === 'win32'
      ? `start "" "${url}"`
      : process.platform === 'darwin'
        ? `open "${url}"`
        : `xdg-open "${url}"`;
  exec(cmd, (err) => {
    if (err) {
      console.warn(
        `Impossible d'ouvrir automatiquement le navigateur (${err.message}). Ouvrez ${url} manuellement.`,
      );
    }
  });
}

const PORT = process.env.PORT ? Number(process.env.PORT) : 4000;
// Écoute sur 127.0.0.1 uniquement par défaut : sans cela, Node écoute sur toutes les
// interfaces réseau (0.0.0.0), ce qui rendrait l'application — sans authentification,
// manipulant des conditions commerciales confidentielles — accessible à quiconque sur le
// même réseau (Wi-Fi bureau, VLAN partagé), pas seulement depuis le poste qui l'exécute.
// HOST reste réglable explicitement pour un usage assumé de partage réseau (ex : instance
// unique hébergée pour toute l'équipe, ouverte depuis une page SharePoint).
const HOST = process.env.HOST || '127.0.0.1';

/**
 * Support HTTPS optionnel (SSL_CERT_PATH / SSL_KEY_PATH) : nécessaire pour intégrer
 * l'application dans une page SharePoint Online, qui est toujours servie en HTTPS — un
 * navigateur bloque l'affichage d'un contenu http:// dans une page https:// ("contenu
 * mixte"), sauf exception pour http://localhost. Pour une instance partagée hébergée sur
 * un poste/serveur du réseau (HOST différent de 127.0.0.1), un certificat est donc requis
 * pour que la page SharePoint puisse l'afficher pour tout le monde, pas seulement pour la
 * personne qui héberge. En local (HOST par défaut), http:// suffit.
 */
function loadHttpsOptions(): https.ServerOptions | null {
  const certPath = process.env.SSL_CERT_PATH;
  const keyPath = process.env.SSL_KEY_PATH;
  if (!certPath || !keyPath) return null;
  try {
    return { cert: fs.readFileSync(certPath), key: fs.readFileSync(keyPath) };
  } catch (e: any) {
    console.error(`Impossible de lire le certificat HTTPS (SSL_CERT_PATH/SSL_KEY_PATH) : ${e.message}`);
    return null;
  }
}

const httpsOptions = loadHttpsOptions();
const scheme = httpsOptions ? 'https' : 'http';
const server = httpsOptions ? https.createServer(httpsOptions, app) : http.createServer(app);

server.listen(PORT, HOST, () => {
  const url = `${scheme}://localhost:${PORT}`;
  console.log(`Achats — application disponible sur ${url}`);
  if (!httpsOptions && HOST !== '127.0.0.1' && HOST !== 'localhost') {
    console.warn(
      "Attention : HOST écoute au-delà de localhost mais aucun certificat n'est configuré " +
        '(SSL_CERT_PATH / SSL_KEY_PATH). Une page SharePoint Online (toujours en HTTPS) ne pourra ' +
        "afficher cette application que pour la personne qui l'exécute (exception locale du " +
        "navigateur pour localhost), pas pour le reste de l'équipe.",
    );
  }
  openBrowser(url);
});
