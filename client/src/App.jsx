import React, { useState, useEffect } from 'react';
import './App.css';

function App() {
  const [iframeUrl, setIframeUrl] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const loadDocument = async () => {
    try {
      setLoading(true);
      setError('');
      
      console.log('🔄 Loading document...');
      
      const response = await fetch('/api/generate-wopi-url');
      
      if (!response.ok) {
        throw new Error('Failed to generate WOPI URL');
      }
      
      const wopiData = await response.json();
      console.log('✅ WOPI data received');
      
      // Use the basic edit URL (Method 4)
      const finalUrl = wopiData.editUrl;
      console.log('🔗 Using basic edit URL:', finalUrl);
      
      setIframeUrl(finalUrl);
      
    } catch (err) {
      console.error('❌ Error:', err);
      setError(`Failed to load document: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const handleReload = () => {
    setIframeUrl(null);
    loadDocument();
  };

  // Load document on component mount
  useEffect(() => {
    loadDocument();
  }, []);

  return (
    <div className="App">
      {/* Minimal Header - Only shows when not in iframe view */}
      {!iframeUrl && (
        <header className="App-header">
          <h1>Word Online Editor</h1>
          <p>Integrated with WOPI Host</p>
          
          <div className="controls">
            <button onClick={loadDocument} disabled={loading}>
              {loading ? '🔄 Loading Editor...' : '📝 Open Word Editor'}
            </button>
          </div>

          {error && (
            <div className="error">
              <h3>Error</h3>
              <p>{error}</p>
              <button onClick={loadDocument}>Try Again</button>
            </div>
          )}
        </header>
      )}

      {/* Full-screen iframe when loaded */}
      {iframeUrl && (
        <div className="fullscreen-container">
          {/* Minimal toolbar that auto-hides */}
          <div className="toolbar">
            <button onClick={handleReload} title="Reload">
              🔄
            </button>
            <button onClick={() => setIframeUrl(null)} title="Close">
              ❌
            </button>
            <span>Word Online Editor - Edit Mode</span>
          </div>
          
          <iframe
            key={iframeUrl}
            src={iframeUrl}
            title="Word Online Editor"
            className="fullscreen-iframe"
            allow="autoplay; fullscreen; clipboard-read; clipboard-write"
            sandbox="allow-scripts allow-same-origin allow-forms allow-popups"
          />
        </div>
      )}

      {/* Loading screen */}
      {loading && (
        <div className="loading-screen">
          <div className="loading-spinner"></div>
          <p>Loading Word Online Editor...</p>
        </div>
      )}
    </div>
  );
}

export default App;