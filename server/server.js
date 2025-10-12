const express = require('express');
const fs = require('fs');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 8080;

// Enable CORS for all routes
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  next();
});

app.use(express.raw({ 
  type: ['*/*', 'application/octet-stream'],
  limit: '50mb' 
}));
app.use(express.json());

const files = new Map();

// Initialize sample file
const initializeSampleFile = () => {
  const filePath = path.join(__dirname, 'files', 'sample.docx');
  if (fs.existsSync(filePath)) {
    const fileBuffer = fs.readFileSync(filePath);
    const stats = fs.statSync(filePath);
    const sha256 = crypto.createHash('sha256').update(fileBuffer).digest('base64');
    
    files.set('sample.docx', {
      name: 'sample.docx',
      size: stats.size,
      lastModified: stats.mtime.toISOString(),
      version: stats.mtime.getTime().toString(),
      sha256: sha256,
      contents: fileBuffer
    });
    console.log(`✅ Sample file loaded: ${stats.size} bytes`);
  } else {
    console.error('❌ Sample file not found at:', filePath);
  }
};

initializeSampleFile();

// Generate access token
const generateAccessToken = (fileId) => {
  const tokenData = {
    fileId: fileId,
    timestamp: Date.now()
  };
  return Buffer.from(JSON.stringify(tokenData)).toString('base64');
};

// Validate access token
const validateAccessToken = (token, fileId) => {
  try {
    const decoded = JSON.parse(Buffer.from(token, 'base64').toString());
    return decoded.fileId === fileId;
  } catch (error) {
    return false;
  }
};

// WOPI Discovery endpoint (optional but recommended)
app.get('/hosting/discovery', (req, res) => {
  const discoveryXml = `<?xml version="1.0" encoding="utf-8"?>
    <wopi-discovery>
      <net-zone name="external-https">
        <app name="Word" favIconUrl="https://localhost:${PORT}/favicon.ico">
          <action name="view" ext="docx" urlsrc="https://word-view.officeapps.live.com/wv/wordviewerframe.aspx?"/>
          <action name="edit" ext="docx" urlsrc="https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?"/>
          <action name="editnew" ext="docx" urlsrc="https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?"/>
        </app>
      </net-zone>
    </wopi-discovery>`;
  
  res.set('Content-Type', 'application/xml');
  res.send(discoveryXml);
});

// CheckFileInfo - Minimal working version
app.get('/wopi/files/:fileId', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  
  console.log('📋 CheckFileInfo for:', fileId);
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }

  const file = files.get(fileId);
  if (!file) {
    return res.status(404).json({ error: 'File not found' });
  }

  // Minimal required properties
  const fileInfo = {
    BaseFileName: file.name,
    Size: file.size,
    OwnerId: 'user@example.com',
    UserId: 'user@example.com',
    Version: file.version,
    UserFriendlyName: 'Test User',
    
    // Core capabilities
    UserCanWrite: true,
    SupportsUpdate: true,
    ReadOnly: false,
    
    // URLs
    HostEditUrl: `https://c67feb255965.ngrok-free.app/`,
    CloseUrl: `https://c67feb255965.ngrok-free.app/`,
    
    // File validation
    SHA256: file.sha256,
    LastModifiedTime: file.lastModified,
    
    // Breadcrumbs
    BreadcrumbBrandName: 'WOPI Demo',
    BreadcrumbBrandUrl: 'https://c67feb255965.ngrok-free.app/',
    BreadcrumbDocName: file.name
  };

  res.json(fileInfo);
});

// GetFile
app.get('/wopi/files/:fileId/contents', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  
  console.log('📥 GetFile for:', fileId);
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }

  const file = files.get(fileId);
  if (!file) {
    return res.status(404).json({ error: 'File not found' });
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.send(file.contents);
});

// PutFile
app.post('/wopi/files/:fileId/contents', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  
  console.log('💾 PutFile for:', fileId, 'Size:', req.body.length);
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }

  try {
    const sha256 = crypto.createHash('sha256').update(req.body).digest('base64');
    
    files.set(fileId, {
      name: fileId,
      size: req.body.length,
      lastModified: new Date().toISOString(),
      version: Date.now().toString(),
      sha256: sha256,
      contents: req.body
    });

    console.log(`✅ File saved: ${req.body.length} bytes`);
    res.setHeader('X-WOPI-ItemVersion', files.get(fileId).version);
    res.status(200).send();
  } catch (error) {
    console.error('❌ Save error:', error);
    res.status(500).json({ error: 'Save failed' });
  }
});

// API endpoints
app.get('/api/health', (req, res) => {
  res.json({ 
    status: 'OK', 
    files: Array.from(files.keys()),
    timestamp: new Date().toISOString()
  });
});

app.get('/api/generate-wopi-url', (req, res) => {
  const fileId = 'sample.docx';
  
  if (!files.has(fileId)) {
    return res.status(404).json({ error: 'File not found' });
  }

  const accessToken = generateAccessToken(fileId);
  const wopiSrc = `https://c67feb255965.ngrok-free.app/wopi/files/${fileId}`;
  
  res.json({
    wopiSrc,
    accessToken,
    fileId
  });
});

app.get('/', (req, res) => {
  res.json({ 
    message: 'WOPI Server Running',
    endpoints: {
      health: '/api/health',
      wopiUrl: '/api/generate-wopi-url',
      discovery: '/hosting/discovery'
    }
  });
});

app.listen(PORT, () => {
  console.log(`\n🚀 WOPI Server: http://localhost:${PORT}`);
  console.log(`🌐 Ngrok: https://c67feb255965.ngrok-free.app`);
  console.log(`\n✅ Ready for WOPI integration\n`);
});