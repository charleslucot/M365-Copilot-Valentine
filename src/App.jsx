import React from 'react';
import { MsalProvider } from '@azure/msal-react';
import { PublicClientApplication } from '@azure/msal-browser';
import { msalConfig } from './services/authConfig';
import ChatInterface from './components/ChatInterface';
import Hearts from './components/Hearts';

// Initialize MSAL
const msalInstance = new PublicClientApplication(msalConfig);

function App() {
  return (
    <MsalProvider instance={msalInstance}>
      <div className="app">
        <Hearts />
        <div className="app-container">
          <header className="app-header">
            <h1 className="app-title">
              <span className="heart-icon">💕</span>
              M365 Copilot
              <span className="heart-icon">💕</span>
            </h1>
            <p className="app-subtitle">Valentine Edition</p>
          </header>
          <ChatInterface />
        </div>
      </div>
    </MsalProvider>
  );
}

export default App;