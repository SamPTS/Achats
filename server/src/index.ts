import express from 'express';
import cors from 'cors';
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

const PORT = process.env.PORT ? Number(process.env.PORT) : 4000;
app.listen(PORT, () => {
  console.log(`Achats API en écoute sur http://localhost:${PORT}`);
});
