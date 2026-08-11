import { NavLink, Route, HashRouter, Routes, Navigate } from 'react-router-dom';
import ConditionsPage from './pages/ConditionsPage';
import TemplatesPage from './pages/TemplatesPage';
import GeneratePage from './pages/GeneratePage';

export default function App() {
  return (
    <HashRouter>
      <div className="layout">
        <aside className="sidebar">
          <div className="brand">Achats — Contrats fournisseurs</div>
          <nav>
            <NavLink to="/conditions" className={({ isActive }) => (isActive ? 'active' : '')}>
              1. Conditions commerciales
            </NavLink>
            <NavLink to="/templates" className={({ isActive }) => (isActive ? 'active' : '')}>
              2. Templates de contrats
            </NavLink>
            <NavLink to="/generer" className={({ isActive }) => (isActive ? 'active' : '')}>
              3. Générer un contrat
            </NavLink>
          </nav>
        </aside>
        <main className="content">
          <Routes>
            <Route path="/" element={<Navigate to="/conditions" replace />} />
            <Route path="/conditions" element={<ConditionsPage />} />
            <Route path="/templates" element={<TemplatesPage />} />
            <Route path="/generer" element={<GeneratePage />} />
          </Routes>
        </main>
      </div>
    </HashRouter>
  );
}
