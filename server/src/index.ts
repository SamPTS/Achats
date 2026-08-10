import express from 'express';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
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

const PORT = process.env.PORT ? Number(process.env.PORT) : 4000;
app.listen(PORT, () => {
  console.log(`Achats — application disponible sur http://localhost:${PORT}`);
});
