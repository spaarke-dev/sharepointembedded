import React, { useState, useEffect } from 'react';
import { PublicClientApplication } from '@azure/msal-browser';

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

export const SPETest: React.FC = () => {
  const [results, setResults] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [initialized, setInitialized] = useState(false);
  const [containerId, setContainerId] = useState<string>('');

  useEffect(() => {
    msalInstance.initialize().then(() => {
      setInitialized(true);
    });
  }, []);

  const fetchWithAuth = async (url: string, options: RequestInit = {}) => {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      throw new Error('No user logged in - please login first');
    }

    // Using Graph scopes
    const tokenResponse = await msalInstance.acquireTokenSilent({
      scopes: ['https://graph.microsoft.com/Files.ReadWrite.All'],
      account: accounts[0]
    }).catch(async (error) => {
      // If silent fails, try popup
      return await msalInstance.acquireTokenPopup({
        scopes: ['https://graph.microsoft.com/Files.ReadWrite.All']
      });
    });

    return fetch(url, {
      ...options,
      headers: {
        ...options.headers,
        'Authorization': `Bearer ${tokenResponse.accessToken}`
      }
    });
  };

  const runTest = async (endpoint: string, method: string = 'GET', body?: any) => {
    setLoading(true);
    try {
      const options: RequestInit = { method };
      if (body) {
        options.body = JSON.stringify(body);
        options.headers = { 'Content-Type': 'application/json' };
      }
      
      const response = await fetchWithAuth(
        `http://localhost:7071/api/${endpoint}`,
        options
      );
      const data = await response.json();
      
      // Extract container ID if present
      if (data.data?.id) {
        setContainerId(data.data.id);
      }
      
      setResults(prev => [...prev, { 
        endpoint,
        timestamp: new Date().toISOString(),
        ...data 
      }]);
    } catch (error: any) {
      setResults(prev => [...prev, { 
        endpoint,
        timestamp: new Date().toISOString(),
        success: false, 
        error: error.message 
      }]);
    }
    setLoading(false);
  };

  return (
    <div style={{ padding: '20px' }}>
      <h1>SPE API Tests</h1>
      
      {!initialized && <p>Initializing...</p>}
      
      <div style={{ marginBottom: '20px' }}>
        <h3>Container Type Verification</h3>
        <button 
          onClick={() => runTest('getTokenClaims')} 
          disabled={loading || !initialized}
          style={{ marginRight: '10px', padding: '10px', backgroundColor: '#28a745', color: 'white', border: 'none', borderRadius: '4px' }}
        >
          🔑 Check Token Permissions
        </button>
        <button 
          onClick={() => runTest('verifyContainerType')} 
          disabled={loading || !initialized}
          style={{ marginRight: '10px', padding: '10px', backgroundColor: '#0078d4', color: 'white', border: 'none', borderRadius: '4px' }}
        >
          🔍 Verify Container Type Registration
        </button>
        <button 
          onClick={() => runTest('testRegistrationStatus')} 
          disabled={loading || !initialized}
          style={{ marginRight: '10px', padding: '10px', backgroundColor: '#6f42c1', color: 'white', border: 'none', borderRadius: '4px' }}
        >
          🔐 Test Registration Status
        </button>
        
        <h3 style={{ marginTop: '20px' }}>Basic Tests</h3>
        <button onClick={() => runTest('testListContainerTypes')} disabled={loading || !initialized}>
          Test Auth Flow
        </button>
        <button onClick={() => runTest('debugGraph')} disabled={loading || !initialized} style={{ marginLeft: '10px' }}>
          Debug Graph Access
        </button>
        
        <h3 style={{ marginTop: '20px' }}>SPE Operations</h3>
        <button onClick={() => runTest('listContainers')} disabled={loading || !initialized}>
          List Containers
        </button>
        <button onClick={() => runTest('createContainer', 'POST')} disabled={loading || !initialized} style={{ marginLeft: '10px' }}>
          Create Container
        </button>
        <button 
          onClick={() => runTest('uploadFile', 'POST', { containerId })} 
          disabled={loading || !initialized || !containerId}
          style={{ marginLeft: '10px' }}
        >
          Upload File (needs container ID)
        </button>
      </div>

      {containerId && (
        <div style={{ padding: '10px', background: '#e0e0e0', marginBottom: '20px', borderRadius: '4px' }}>
          <strong>Container ID:</strong> {containerId}
        </div>
      )}

      <div style={{ marginTop: '20px' }}>
        <h2>Results:</h2>
        <button onClick={() => setResults([])} style={{ marginBottom: '10px' }}>Clear Results</button>
        <pre style={{ background: '#f0f0f0', padding: '10px', overflow: 'auto', maxHeight: '400px', borderRadius: '4px' }}>
          {JSON.stringify(results, null, 2)}
        </pre>
      </div>

      <div style={{ marginTop: '40px', fontSize: '12px', color: '#666' }}>
        <p><strong>Environment Variables:</strong></p>
        <p>Container Type ID: {process.env.REACT_APP_CONTAINER_TYPE_ID || 'Not set'}</p>
        <p>Client ID: {process.env.REACT_APP_CLIENT_ID}</p>
        <p>Tenant ID: {process.env.REACT_APP_TENANT_ID}</p>
      </div>
    </div>
  );
};