// src/components/LoginButton.jsx

import React from 'react';
import { useMsal } from '@azure/msal-react';
import { loginRequest } from '../services/authConfig';

function LoginButton() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = accounts.length > 0;

  const handleLogin = async () => {
    try {
      await instance.loginPopup(loginRequest);
    } catch (error) {
      console.error('Login failed:', error);
    }
  };

  const handleLogout = async () => {
    try {
      await instance.logoutPopup();
    } catch (error) {
      console.error('Logout failed:', error);
    }
  };

  return (
    <div className="login-section">
      {!isAuthenticated ? (
        <button onClick={handleLogin} className="valentine-button login-button">
          <span className="button-icon">🔐</span>
          Sign in with Microsoft
        </button>
      ) : (
        <div className="user-info">
          <span className="user-name">
            ❤️ {accounts[0]?.name || 'User'}
          </span>
          <button onClick={handleLogout} className="valentine-button logout-button">
            Sign Out
          </button>
        </div>
      )}
    </div>
  );
}

export default LoginButton;