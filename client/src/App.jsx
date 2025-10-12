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
      const response = await fetch('/api/health');
      if (response.ok) {
        const data = await response.json();
        setServerStatus('online');
        console.log('✅ Server health:', data);
      } else {
        setServerStatus('error');
      }
    } catch (err) {
      setServerStatus('error');
      console.error('❌ Server health check failed:', err);
    }
  };

  const loadDocument = async () => {
    try {
      setLoading(true);
      setError('');
      
      console.log('🔄 Loading document for editing...');
      
      const response = await fetch('/api/generate-wopi-url');
      
      if (!response.ok) {
        throw new Error(`Server returned ${response.status}`);
      }
      
      const wopiData = await response.json();
      
      console.log('✅ Received WOPI data:', wopiData);
      
      // Method 1: Direct edit URL (recommended)
      // const editUrl = `https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=${encodeURIComponent(wopiData.wopiSrc)}&access_token=${wopiData.accessToken}`;
      // New: Add &action=edit for explicit instruction
      const editUrl = `https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=${encodeURIComponent(wopiData.wopiSrc)}&access_token=${wopiData.accessToken}&action=edit`;
      
      console.log('🔗 Edit URL:', editUrl);
      setIframeUrl(editUrl);
      
      // Verify file permissions
      await verifyFilePermissions(wopiData);
      
    } catch (err) {
      console.error('❌ Error loading document:', err);
      setError(`Failed to load document: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const verifyFilePermissions = async (wopiData) => {
    try {
      const fileInfoResponse = await fetch(`/wopi/files/${wopiData.fileId}?access_token=${wopiData.accessToken}`);
      if (fileInfoResponse.ok) {
        const fileInfoData = await fileInfoResponse.json();
        setFileInfo(fileInfoData);
        
        console.log('📄 File Info Response:', fileInfoData);
        console.log('🔑 Key Permissions:', {
          UserCanWrite: fileInfoData.UserCanWrite,
          ReadOnly: fileInfoData.ReadOnly,
          SupportsUpdate: fileInfoData.SupportsUpdate,
          SupportsLocks: fileInfoData.SupportsLocks,
          SupportsCobalt: fileInfoData.SupportsCobalt
        });
        
        if (!fileInfoData.UserCanWrite || fileInfoData.ReadOnly) {
          console.error('❌ FILE OPENING IN READ-ONLY MODE DUE TO PERMISSIONS');
        } else {
          console.log('✅ FILE SHOULD OPEN IN EDIT MODE');
        }
      }
    } catch (infoError) {
      console.warn('Could not load file info:', infoError);
    }
  };

  const handleReload = () => {
    setIframeUrl(null);
    setFileInfo(null);
    loadDocument();
  };

  const tryDifferentUrls = () => {
    const methods = [
      {
        name: 'Method 1: Direct Edit',
        url: `https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=${encodeURIComponent(fileInfo.HostEditUrl)}&access_token=${fileInfo.access_token}`
      },
      {
        name: 'Method 2: With Action',
        url: `https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=${encodeURIComponent(fileInfo.HostEditUrl)}&access_token=${fileInfo.access_token}&action=edit`
      },
      {
        name: 'Method 3: Office Online',
        url: `https://office.live.com/we/wordeditorframe.aspx?WOPISrc=${encodeURIComponent(fileInfo.HostEditUrl)}&access_token=${fileInfo.access_token}`
      }
    ];
    
    methods.forEach((method, index) => {
      console.log(`${index + 1}. ${method.name}: ${method.url}`);
    });
    
    // Try the first method
    setIframeUrl(methods[0].url);
  };

  return (
    <div className="App">
      <header className="App-header">
        <h1>WOPI Document Editor - EDIT MODE</h1>
        <p>Forcing edit mode with complete WOPI implementation</p>
        
        <div className="server-status">
          <span className={`status-indicator ${serverStatus}`}>
            Server: {serverStatus === 'online' ? '✅ Online' : serverStatus === 'checking' ? '🔄 Checking...' : '❌ Offline'}
          </span>
        </div>
        
        {fileInfo && (
          <div className="file-info">
            <h3>File Permissions:</h3>
            <div className="permissions-grid">
              <span className={fileInfo.UserCanWrite ? 'success' : 'error'}>
                UserCanWrite: {fileInfo.UserCanWrite ? '✅ true' : '❌ false'}
              </span>
              <span className={!fileInfo.ReadOnly ? 'success' : 'error'}>
                ReadOnly: {fileInfo.ReadOnly ? '❌ true' : '✅ false'}
              </span>
              <span className={fileInfo.SupportsUpdate ? 'success' : 'error'}>
                SupportsUpdate: {fileInfo.SupportsUpdate ? '✅ true' : '❌ false'}
              </span>
              <span className={fileInfo.SupportsLocks ? 'success' : 'error'}>
                SupportsLocks: {fileInfo.SupportsLocks ? '✅ true' : '❌ false'}
              </span>
            </div>
          </div>
        )}
        
        <div className="controls">
          <button onClick={handleReload} disabled={loading}>
            {loading ? '🔄 Loading...' : '📄 Reload Document'}
          </button>
          <button onClick={tryDifferentUrls} disabled={loading || !fileInfo}>
            Try Different URL
          </button>
        </div>
      </header>

      <main className="App-main">
        {error && (
          <div className="error">
            <h3>❌ Error</h3>
            <p>{error}</p>
            <button onClick={handleReload}>🔄 Try Again</button>
          </div>
        )}
        
        {loading ? (
          <div className="loading">
            <div className="spinner"></div>
            <p>Loading Word Online Editor...</p>
            <p className="loading-sub">Edit mode forced in CheckFileInfo</p>
          </div>
        ) : iframeUrl ? (
          <div className="editor-container">
            <iframe
              key={iframeUrl}
              src={iframeUrl}
              title="Word Online Editor"
              className="word-iframe"
              allow="autoplay; fullscreen; clipboard-read; clipboard-write"
              sandbox="allow-scripts allow-same-origin allow-forms allow-popups allow-top-navigation"
              onLoad={() => {
                console.log('✅ Editor iframe loaded');
                console.log('📝 Check browser console for Office Online messages');
              }}
              onError={(e) => console.error('❌ Iframe error:', e)}
            />
          </div>
        ) : null}
      </main>
      
      <footer className="App-footer">
        <div className="troubleshooting">
          <h4>🔧 Troubleshooting Steps:</h4>
          <ol>
            <li>Check browser console for detailed logs</li>
            <li>Clear browser cache completely</li>
            <li>Verify file permissions above show ✅ true for edit mode</li>
            <li>Try "Try Different URL" button</li>
            <li>Check server logs for WOPI requests</li>
          </ol>
        </div>
      </footer>
    </div>
  );
}

export default App;