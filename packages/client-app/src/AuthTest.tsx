import React, { useState, useEffect } from 'react';
import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';

const msalConfig = {
  auth: {
    clientId: process.env.REACT_APP_CLIENT_ID!,
    authority: `https://login.microsoftonline.com/${process.env.REACT_APP_TENANT_ID}`,
    redirectUri: 'http://localhost:3000'
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  }
};

const msalInstance = new PublicClientApplication(msalConfig);

// Initialize MSAL
msalInstance.initialize().then(() => {
  console.log('MSAL initialized');
}).catch((error) => {
  console.error('MSAL initialization error:', error);
});

export const AuthTest: React.FC = () => {
  const [user, setUser] = useState<string>('');
  const [token, setToken] = useState<string>('');
  const [error, setError] = useState<string>('');
  const [isInitialized, setIsInitialized] = useState(false);

  useEffect(() => {
    msalInstance.initialize().then(() => {
      setIsInitialized(true);
      setError('');
    }).catch((err) => {
      setError(`MSAL initialization failed: ${err.message}`);
    });
  }, []);

  const testLogin = async () => {
    if (!isInitialized) {
      setError('MSAL not initialized yet');
      return;
    }
    
    try {
      const loginResponse = await msalInstance.loginPopup({
        scopes: ['user.read']
      });
      setUser(loginResponse.account.username);
      setError('');
      console.log('Login successful:', loginResponse);
    } catch (err: any) {
      setError(`Login failed: ${err.message}`);
      console.error('Login error:', err);
    }
  };

  const testGetToken = async () => {
    if (!isInitialized) {
      setError('MSAL not initialized yet');
      return;
    }

    try {
      const tokenResponse = await msalInstance.acquireTokenSilent({
        scopes: [`api://${process.env.REACT_APP_CLIENT_ID}/Container.Manage`],
        account: msalInstance.getAllAccounts()[0]
      });
      setToken(tokenResponse.accessToken.substring(0, 50) + '...');
      setError('');
      console.log('Token acquired:', tokenResponse);
    } catch (err: any) {
      if (err instanceof InteractionRequiredAuthError) {
        try {
          const tokenResponse = await msalInstance.acquireTokenPopup({
            scopes: [`api://${process.env.REACT_APP_CLIENT_ID}/Container.Manage`]
          });
          setToken(tokenResponse.accessToken.substring(0, 50) + '...');
          setError('');
        } catch (popupErr: any) {
          setError(`Token popup failed: ${popupErr.message}`);
        }
      } else {
        setError(`Token failed: ${err.message}`);
      }
    }
  };

  const testCallFunction = async () => {
    if (!isInitialized) {
      setError('MSAL not initialized yet');
      return;
    }

    try {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length === 0) {
        setError('No account - please login first');
        return;
      }

      const tokenResponse = await msalInstance.acquireTokenSilent({
        scopes: [`api://${process.env.REACT_APP_CLIENT_ID}/Container.Manage`],
        account: accounts[0]
      });

      const response = await fetch('http://localhost:7071/api/healthCheck', {
        headers: {
          'Authorization': `Bearer ${tokenResponse.accessToken}`
        }
      });

      const data = await response.json();
      setError('');
      console.log('Function call success:', data);
      alert(`Function call success: ${JSON.stringify(data)}`);
    } catch (err: any) {
      setError(`Function call failed: ${err.message}`);
    }
  };

  return (
    <div style={{ padding: '20px' }}>
      <h1>Authentication Test</h1>
      
      <div style={{ marginBottom: '20px' }}>
        <button onClick={testLogin} disabled={!isInitialized} style={{ marginRight: '10px' }}>
          Step 1: Login
        </button>
        <button onClick={testGetToken} disabled={!isInitialized} style={{ marginRight: '10px' }}>
          Step 2: Get Token
        </button>
        <button onClick={testCallFunction} disabled={!isInitialized}>
          Step 3: Call Function
        </button>
      </div>

      {!isInitialized && <p>Initializing MSAL...</p>}

      <div style={{ marginTop: '20px' }}>
        {user && <p><strong>User:</strong> {user}</p>}
        {token && <p><strong>Token:</strong> {token}</p>}
        {error && <p style={{ color: 'red' }}><strong>Error:</strong> {error}</p>}
      </div>

      <div style={{ marginTop: '40px', fontSize: '12px', color: '#666' }}>
        <p>Client ID: {process.env.REACT_APP_CLIENT_ID}</p>
        <p>Tenant ID: {process.env.REACT_APP_TENANT_ID}</p>
        <p>MSAL Initialized: {isInitialized ? 'Yes' : 'No'}</p>
      </div>
    </div>
  );
};