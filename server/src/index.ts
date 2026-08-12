import express from 'express';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import { exec } from 'child_process';
import './db'; // initialise le schéma au démarrage
import conditionsRouter from './routes/conditions';
import templatesRouter from './routes/templates';
import generateRouter from './routes/generate';
import { getAppDir } from './runtimePaths';

const app = express();
app.use(cors());
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
app.listen(PORT, () => {
  const url = `http://localhost:${PORT}`;
  console.log(`Achats — application disponible sur ${url}`);
  openBrowser(url);
});
