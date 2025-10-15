// client/src/App.jsx — minimal fix (token is now URL-encoded on server, no other changes required)
import React, { useState, useEffect } from 'react';
import './App.css';

export default function App() {
  const [iframeUrl, setIframeUrl] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const loadDocument = async () => {
    try {
      setLoading(true);
      setError('');
      const resp = await fetch('https://6e51ccd22f2a.ngrok-free.app/api/generate-wopi-url');
      if (!resp.ok) throw new Error('Failed to generate WOPI URL');
      const { editUrl } = await resp.json();
      setIframeUrl(editUrl);
    } catch (e) {
      setError(e.message || 'Failed to load document');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { loadDocument(); }, []);

  return (
    <div className="App">
      {!iframeUrl && (
        <header className="App-header">
          <h1>Word Online Editor</h1>
          <button onClick={loadDocument} disabled={loading}>
            {loading ? 'Loading…' : 'Open Word Editor'}
          </button>
          {error && <p className="error">{error}</p>}
        </header>
      )}

      {iframeUrl && (
        <div className="fullscreen-container">
          <div className="toolbar">
            <button onClick={() => setIframeUrl(null)}>Close</button>
            <button onClick={loadDocument}>Reload</button>
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

      {loading && (
        <div className="loading-screen">
          <div className="loading-spinner" />
          <p>Loading Word Online Editor…</p>
        </div>
      )}
    </div>
  );
}
