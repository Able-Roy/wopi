const express = require('express');
const app = express();
const PORT = 8080;

// Simple test endpoints
app.get('/api/health', (req, res) => {
  console.log('Health check called!');
  res.json({ 
    status: 'OK', 
    message: 'Test server is working',
    timestamp: new Date().toISOString()
  });
});

app.get('/', (req, res) => {
  res.json({ message: 'Server is running' });
});

// Start server
app.listen(PORT, () => {
  console.log(`🔧 Test server running on http://localhost:${PORT}`);
  console.log(`📋 Test: http://localhost:${PORT}/api/health`);
});