/**
 * DEX Event Experience Platform - Main App
 *
 * Routing und grundlegendes Layout.
 * EventProvider umschliesst alles, damit der Event-State
 * in allen Seiten verfuegbar ist.
 */

import { BrowserRouter, Routes, Route } from 'react-router-dom';
import Header from './components/Header';
import LandingPage from './pages/LandingPage';
import StartPage from './pages/StartPage';
import EventListPage from './pages/EventListPage';
import RegistrationPage from './pages/RegistrationPage';
import MyEventsPage from './pages/MyEventsPage';
import EventCreationPage from './pages/EventCreationPage';
import SettingsPage from './pages/SettingsPage';
import { EventProvider } from './context/EventContext';
import './App.css';
import './pages/pages.css';

function App() {
  return (
    <BrowserRouter>
      <EventProvider>
        <div className="app-layout">
          <Header />
          <main className="main-content">
            <Routes>
              <Route path="/" element={<LandingPage />} />
              <Route path="/start" element={<StartPage />} />
              <Route path="/register" element={<EventListPage />} />
              <Route path="/register/:eventId" element={<RegistrationPage />} />
              <Route path="/my-events" element={<MyEventsPage />} />
              <Route path="/create-event" element={<EventCreationPage />} />
              <Route path="/settings" element={<SettingsPage />} />
            </Routes>
          </main>
        </div>
      </EventProvider>
    </BrowserRouter>
  );
}

export default App;
