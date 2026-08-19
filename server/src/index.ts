import express from 'express';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import { exec } from 'child_process';
import './db'; // initialise le schéma au démarrage
import conditionsRouter from './routes/conditions';
import templatesRouter from './routes/templates';
import generateRouter from './routes/generate';

const app = express();
app.use(cors());
app.use(express.json({ limit: '5mb' }));

app.use('/api/conditions', conditionsRouter);
app.use('/api/templates', templatesRouter);
app.use('/api/generate', generateRouter);

app.get('/api/health', (_req, res) => res.json({ ok: true }));

// Sert le frontend buildé (client/dist) sur le même port, pour un lancement en un seul
// exécutable/script : évite d'avoir à démarrer un second serveur (Vite) pour tester l'appli.
// Le routeur front utilise le mode "hash" (#/...), donc aucune route "catch-all" n'est requise.
const clientDist = path.join(__dirname, '..', '..', 'client', 'dist');
if (fs.existsSync(clientDist)) {
  app.use(express.static(clientDist));
}

/**
 * Ouvre le navigateur par défaut sur l'URL donnée. Utilisé uniquement quand l'application
 * tourne en exécutable autonome (pkg) : dans ce mode il n'y a pas de terminal "utilisateur"
 * qui affiche une URL cliquable, donc sans cette ouverture automatique la fenêtre console qui
 * apparaît au double-clic ne donne visuellement l'impression que "l'application ne s'ouvre pas".
 */
function openBrowser(url: string) {
  const command =
    process.platform === 'win32'
      ? `start "" "${url}"`
      : process.platform === 'darwin'
        ? `open "${url}"`
        : `xdg-open "${url}"`;
  exec(command, (err) => {
    if (err) {
      console.warn(`Impossible d'ouvrir automatiquement le navigateur : ${err.message}`);
      console.warn(`Ouvrez-le manuellement sur ${url}`);
    }
  });
}

const PORT = process.env.PORT ? Number(process.env.PORT) : 4000;
const server = app.listen(PORT, () => {
  const url = `http://localhost:${PORT}`;
  console.log(`Achats — application disponible sur ${url}`);

  const pkgProcess = process as unknown as { pkg?: unknown };
  if (pkgProcess.pkg) {
    // Exécutable autonome (.exe) : ouverture automatique, comme le fait lancer-application.bat
    // pour le mode développement.
    openBrowser(url);
  }
});

// Sans ce handler, une erreur de démarrage (ex: port déjà utilisé par une autre instance déjà
// lancée) fait planter le process silencieusement — la fenêtre console de l'exécutable se
// referme aussitôt, donnant l'impression que "l'application ne s'ouvre pas" sans aucun message
// lisible pour l'utilisateur.
server.on('error', (err: NodeJS.ErrnoException) => {
  if (err.code === 'EADDRINUSE') {
    console.error(
      `\n[ERREUR] Le port ${PORT} est déjà utilisé — l'application est peut-être déjà ouverte ` +
        `dans une autre fenêtre ou un autre onglet. Vérifiez http://localhost:${PORT}, ou fermez ` +
        `l'instance déjà lancée avant de relancer.\n`
    );
  } else {
    console.error(`\n[ERREUR] Impossible de démarrer l'application : ${err.message}\n`);
  }
  console.error('Appuyez sur Entrée pour fermer cette fenêtre...');
  process.stdin.resume();
  process.stdin.once('data', () => process.exit(1));
});
