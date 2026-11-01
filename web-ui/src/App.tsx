import React, { Suspense, lazy } from 'react';
import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import Navbar from './components/layout/Navbar';
import Footer from './components/layout/Footer';
import Toast from './components/common/Toast';
import SettingsModal from './components/common/SettingsModal';
import LoadingSpinner from './components/ui/LoadingSpinner';

// Lazy-loaded pages for better performance
const HomePage = lazy(() => import('./pages/HomePage'));
const CategoryInputPage = lazy(() => import('./pages/CategoryInputPage'));
const ResultPage = lazy(() => import('./pages/ResultPage'));
const RulesConfigPage = lazy(() => import('./pages/RulesConfigPage'));

const PageLoading: React.FC = () => (
  <div className="flex items-center justify-center min-h-[60vh]">
    <LoadingSpinner size="lg" text="加载页面..." />
  </div>
);

const App: React.FC = () => {
  return (
    <BrowserRouter>
      <div className="min-h-screen flex flex-col bg-gray-50">
        <Navbar />
        <main className="flex-1">
          <Suspense fallback={<PageLoading />}>
            <Routes>
              <Route path="/" element={<HomePage />} />
              <Route path="/home" element={<Navigate to="/" replace />} />
              <Route path="/input/:module" element={<CategoryInputPage />} />
              <Route path="/result" element={<ResultPage />} />
              <Route path="/rules-config" element={<RulesConfigPage />} />
              <Route path="*" element={<Navigate to="/" replace />} />
            </Routes>
          </Suspense>
        </main>
        <Footer />
        <Toast />
        <SettingsModal />
      </div>
    </BrowserRouter>
  );
};

export default App;
