import React, { useState, useEffect } from 'react';
import './App.css';

function App() {
  const [iframeUrl, setIframeUrl] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [fileInfo, setFileInfo] = useState(null);
  const [serverStatus, setServerStatus] = useState('checking');

  useEffect(() => {
    checkServerStatus();
    loadDocument();
  }, []);

  const checkServerStatus = async () => {
    try {
      // Use relative URL - Vite proxy will handle this
      const response = await fetch('/api/health');
      if (response.ok) {
        setServerStatus('online');
      } else {
        setServerStatus('error');
      }
    } catch (err) {
      setServerStatus('error');
    }
  };

  const loadDocument = async () => {
    try {
      setLoading(true);
      setError('');
      
      console.log('🔄 Loading document from server...');
      
      // Use relative URL - Vite proxy will handle this
      const response = await fetch('/api/generate-wopi-url');
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Server returned ${response.status}: ${errorText}`);
      }
      
      const { wopiSrc, accessToken, fileId } = await response.json();
      
      console.log('✅ Received WOPI data:', { wopiSrc, fileId });
      
      // Construct the Word Online editor URL
      const wordOnlineUrl = new URL('https://word-edit.officeapps.live.com/we/wordeditorframe.aspx');
      
      // Add WOPI parameters
      wordOnlineUrl.searchParams.append('WOPISrc', wopiSrc);
      wordOnlineUrl.searchParams.append('access_token', accessToken);
      wordOnlineUrl.searchParams.append('ui', 'en-US');
      wordOnlineUrl.searchParams.append('rs', 'en-US');
      wordOnlineUrl.searchParams.append('wdorigin', '1');
      wordOnlineUrl.searchParams.append('hid', 'ngrok-' + Date.now());
      
      const finalUrl = wordOnlineUrl.toString();
      console.log('🔗 Final Word Online URL:', finalUrl);
      
      setIframeUrl(finalUrl);
      
      // Also get file info to display (this will go through Vite proxy)
      try {
        const fileInfoResponse = await fetch(`/wopi/files/${fileId}?access_token=${accessToken}`);
        if (fileInfoResponse.ok) {
          const fileInfoData = await fileInfoResponse.json();
          setFileInfo(fileInfoData);
          console.log('📄 File info loaded:', fileInfoData);
        }
      } catch (infoError) {
        console.warn('Could not load file info:', infoError);
      }
      
    } catch (err) {
      console.error('❌ Error loading document:', err);
      setError(`Failed to load document: ${err.message}. Make sure the backend server is running on port 8080.`);
    } finally {
      setLoading(false);
    }
  };

  const handleReload = () => {
    setIframeUrl(null);
    setFileInfo(null);
    loadDocument();
  };

  const testServerConnection = async () => {
    try {
      setLoading(true);
      const response = await fetch('/api/health');
      if (response.ok) {
        const health = await response.json();
        alert(`✅ Server is healthy!\nFiles: ${health.files.join(', ')}\nTime: ${health.serverTime}`);
        setServerStatus('online');
      } else {
        throw new Error(`Server health check failed: ${response.status}`);
      }
    } catch (err) {
      alert(`❌ Server connection failed: ${err.message}`);
      setServerStatus('error');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="App">
      <header className="App-header">
        <h1>WOPI Document Editor Demo</h1>
        <p>Edit your Word document directly in the browser using Microsoft Word Online</p>
        
        <div className="server-status">
          <span className={`status-indicator ${serverStatus}`}>
            Server: {serverStatus === 'online' ? '✅ Online' : serverStatus === 'checking' ? '🔄 Checking...' : '❌ Offline'}
          </span>
        </div>
        
        {fileInfo && (
          <div className="file-info">
            <span><strong>File:</strong> {fileInfo.BaseFileName}</span>
            <span><strong>Size:</strong> {Math.round(fileInfo.Size / 1024)} KB</span>
            <span><strong>Editable:</strong> {fileInfo.UserCanWrite ? '✅ Yes' : '❌ No'}</span>
          </div>
        )}
        
        <div className="controls">
          <button onClick={handleReload} disabled={loading}>
            {loading ? '🔄 Loading...' : '📄 Reload Document'}
          </button>
          <button onClick={testServerConnection} disabled={loading}>
            🔧 Test Server
          </button>
        </div>
      </header>

      <main className="App-main">
        {error && (
          <div className="error">
            <h3>❌ Error Loading Document</h3>
            <p>{error}</p>
            <div className="error-actions">
              <button onClick={handleReload}>🔄 Try Again</button>
              <button onClick={testServerConnection}>🔧 Test Connection</button>
            </div>
            <div className="debug-info">
              <p><strong>Troubleshooting:</strong></p>
              <p>1. Make sure backend is running: <code>node server.js</code></p>
              <p>2. Backend should be on http://localhost:8080</p>
              <p>3. Check that sample.docx exists in server/files/</p>
              <p>4. Restart both frontend and backend if needed</p>
            </div>
          </div>
        )}
        
        {loading ? (
          <div className="loading">
            <div className="spinner"></div>
            <p>Loading document editor...</p>
            <p className="loading-sub">Connecting to Microsoft Word Online</p>
          </div>
        ) : iframeUrl ? (
          <div className="editor-container">
            <iframe
              key={iframeUrl}
              src={iframeUrl}
              title="Word Online Editor"
              className="word-iframe"
              allow="autoplay; fullscreen; clipboard-read; clipboard-write"
              sandbox="allow-scripts allow-same-origin allow-forms allow-popups"
              onLoad={() => console.log('✅ Iframe loaded successfully')}
              onError={(e) => console.error('❌ Iframe error:', e)}
            />
          </div>
        ) : null}
      </main>
      
      <footer className="App-footer">
        <p>
          Demo Application - WOPI Protocol Integration
          <br />
          Using Vite Proxy to avoid CORS issues
          <br />
          Changes are automatically saved back to the server when you click Save in Word Online
        </p>
      </footer>
    </div>
  );
}

export default App;