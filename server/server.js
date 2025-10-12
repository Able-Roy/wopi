const express = require('express');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 8080;

// Enhanced CORS for WOPI
app.use((req, res, next) => {
  const allowedOrigins = [
    'https://word-edit.officeapps.live.com',
    'https://office.live.com',
    'https://c67feb255965.ngrok-free.app'
  ];
  
  const origin = req.headers.origin;
  if (allowedOrigins.includes(origin)) {
    res.header('Access-Control-Allow-Origin', origin);
  } else {
    res.header('Access-Control-Allow-Origin', '*');
  }
  
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, OPTIONS, LOCK, UNLOCK');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With, X-WOPI-Override, X-WOPI-Lock, X-WOPI-OldLock, X-WOPI-MachineName, X-WOPI-SessionId, X-WOPI-ItemVersion');
  res.header('Access-Control-Allow-Credentials', 'true');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  
  next();
});

// FIXED: Proper body parsing for different content types
app.use((req, res, next) => {
  if (req.headers['content-type'] === 'application/octet-stream' || 
      req.headers['x-wopi-override'] === 'PUT_RELATIVE') {
    // For binary file content
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => {
      req.rawBody = Buffer.concat(chunks);
      next();
    });
  } else {
    // For JSON content
    express.json()(req, res, next);
  }
});

const files = new Map();
const locks = new Map();

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
    // Create a simple DOCX file in memory
    const simpleDocx = createSimpleDocx();
    files.set('sample.docx', {
      name: 'sample.docx',
      size: simpleDocx.length,
      lastModified: new Date().toISOString(),
      version: Date.now().toString(),
      sha256: crypto.createHash('sha256').update(simpleDocx).digest('base64'),
      contents: simpleDocx
    });
    console.log('✅ Created simple DOCX file in memory');
  }
};

// Create a simple DOCX file for testing
function createSimpleDocx() {
  // Minimal DOCX structure
  const content = 'Test Document Content';
  return Buffer.from(content);
}

initializeSampleFile();

// Generate access token
const generateAccessToken = (fileId) => {
  const tokenData = {
    fileId: fileId,
    timestamp: Date.now(),
    userId: 'test-user@example.com'
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

// CheckFileInfo - ULTIMATE FIX for edit mode
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

  // CRITICAL: This response forces edit mode
  const fileInfo = {
    // Basic required properties
    BaseFileName: file.name,
    Size: file.size,
    OwnerId: 'owner@example.com',
    UserId: 'user@example.com',
    Version: file.version,
    
    // PERMISSIONS - MOST IMPORTANT FOR EDIT MODE
    UserCanWrite: true,
    ReadOnly: false,
    RestrictedWebViewOnly: false,
    WebEditingDisabled: false,
    
    // User info
    UserFriendlyName: 'Test User',
    
    // PostMessage properties
    // PostMessageOrigin: 'https://c67feb255965.ngrok-free.app',
    PostMessageOrigin: 'http://localhost:5173/',
    
    // Breadcrumb properties
    BreadcrumbBrandName: 'WOPI Host',
    BreadcrumbBrandUrl: 'https://c67feb255965.ngrok-free.app',
    BreadcrumbDocName: file.name,
    BreadcrumbFolderUrl: 'https://c67feb255965.ngrok-free.app',
    
    // File properties
    SHA256: file.sha256,
    LastModifiedTime: file.lastModified,
    
    // SUPPORTED FEATURES - Enable everything for edit mode
    SupportsUpdate: true,
    SupportsLocks: true,
    SupportsGetLock: true,
    SupportsDeleteFile: true,
    SupportsRename: true,
    SupportsUserInfo: true,
    SupportsCobalt: true,
    SupportsContainers: false,
    SupportsFolders: false,
    SupportsCoauth: false,
    SupportsScenarioLinks: false,
    SupportsSecureStore: false,
    SupportsPutRelativeFile: true,
    
    // URLs
    HostEditUrl: `https://c67feb255965.ngrok-free.app/wopi/files/${fileId}`,
    HostViewUrl: `https://c67feb255965.ngrok-free.app/wopi/files/${fileId}`,
    
    // Misc
    AllowExternalMarketplace: false,
    DisablePrint: false,
    DisableTranslation: false,
    FileSharingPostMessage: true,
    FileVersionUrl: `https://c67feb255965.ngrok-free.app/wopi/files/${fileId}`,
    LicenseCheckForEditIsEnabled: false,
    UserCanAttend: true,
    UserCanPresent: true,
    UserCanRename: true,
    UserCanNotWriteRelative: false,
    WebEditingDisabled: false
  };

  console.log('✅ Returning CheckFileInfo - FORCING EDIT MODE');
  console.log('🔑 Key permissions:', {
    UserCanWrite: fileInfo.UserCanWrite,
    ReadOnly: fileInfo.ReadOnly,
    SupportsUpdate: fileInfo.SupportsUpdate,
    SupportsLocks: fileInfo.SupportsLocks
  });
  
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

  res.setHeader('Content-Type', 'application/octet-stream');
  res.setHeader('Content-Disposition', `attachment; filename="${file.name}"`);
  res.send(file.contents);
});

// PutFile - FIXED body handling
app.post('/wopi/files/:fileId/contents', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  
  console.log('💾 PutFile for:', fileId, 'Size:', req.rawBody?.length || 'unknown');
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }

  try {
    const file = files.get(fileId);
    if (!file) {
      return res.status(404).json({ error: 'File not found' });
    }

    if (!req.rawBody) {
      return res.status(400).json({ error: 'No file content provided' });
    }

    const sha256 = crypto.createHash('sha256').update(req.rawBody).digest('base64');
    
    files.set(fileId, {
      ...file,
      size: req.rawBody.length,
      lastModified: new Date().toISOString(),
      version: Date.now().toString(),
      sha256: sha256,
      contents: req.rawBody
    });

    console.log(`✅ File saved: ${req.rawBody.length} bytes`);
    
    res.setHeader('X-WOPI-ItemVersion', files.get(fileId).version);
    res.status(200).send();
  } catch (error) {
    console.error('❌ Save error:', error);
    res.status(500).json({ error: 'Save failed' });
  }
});

// PutRelativeFile - FIXED body handling
app.post('/wopi/files/:fileId', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  const suggestedTarget = req.headers['x-wopi-suggestedtarget'];
  const relativeTarget = req.headers['x-wopi-relativetarget'];
  
  console.log('📝 PutRelativeFile for:', fileId, 'Suggested:', suggestedTarget);
  console.log('📦 Request body size:', req.rawBody?.length || 'undefined');
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }

  try {
    // Check if we have file content
    if (!req.rawBody) {
      console.error('❌ No file content in request body');
      return res.status(400).json({ error: 'No file content provided' });
    }

    const newFileId = suggestedTarget || `document-${Date.now()}.docx`;
    const sha256 = crypto.createHash('sha256').update(req.rawBody).digest('base64');
    
    files.set(newFileId, {
      name: newFileId,
      size: req.rawBody.length,
      lastModified: new Date().toISOString(),
      version: Date.now().toString(),
      sha256: sha256,
      contents: req.rawBody
    });

    console.log(`✅ New file created: ${newFileId} (${req.rawBody.length} bytes)`);
    
    res.json({
      Name: newFileId,
      Url: `https://c67feb255965.ngrok-free.app/wopi/files/${newFileId}`
    });
  } catch (error) {
    console.error('❌ PutRelativeFile error:', error);
    res.status(500).json({ error: 'Create file failed' });
  }
});

// Lock operations
app.post('/wopi/files/:fileId/lock', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  const lockId = req.headers['x-wopi-lock'];
  
  console.log('🔒 Lock file:', fileId, 'Lock ID:', lockId);
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }

  if (!files.has(fileId)) {
    return res.status(404).json({ error: 'File not found' });
  }

  const currentLock = locks.get(fileId);
  
  if (!currentLock) {
    locks.set(fileId, lockId);
    console.log('✅ Lock acquired');
    res.status(200).send();
  } else if (currentLock === lockId) {
    console.log('✅ Lock refreshed');
    res.status(200).send();
  } else {
    console.log('❌ File already locked with:', currentLock);
    res.setHeader('X-WOPI-Lock', currentLock);
    res.status(409).send();
  }
});

app.get('/wopi/files/:fileId/lock', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  
  console.log('🔍 Get lock for:', fileId);
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }

  if (!files.has(fileId)) {
    return res.status(404).json({ error: 'File not found' });
  }

  const lock = locks.get(fileId);
  if (lock) {
    res.setHeader('X-WOPI-Lock', lock);
    console.log('✅ Lock found:', lock);
  }
  
  res.status(200).send();
});

app.unlock('/wopi/files/:fileId/lock', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  const lockId = req.headers['x-wopi-lock'];
  
  console.log('🔓 Unlock file:', fileId, 'Lock ID:', lockId);
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }

  if (!files.has(fileId)) {
    return res.status(404).json({ error: 'File not found' });
  }

  const currentLock = locks.get(fileId);
  
  if (currentLock === lockId) {
    locks.delete(fileId);
    console.log('✅ Lock released');
    res.status(200).send();
  } else {
    console.log('❌ Lock mismatch. Current:', currentLock, 'Requested:', lockId);
    res.setHeader('X-WOPI-Lock', currentLock || '');
    res.status(409).send();
  }
});

// API endpoints
app.get('/api/health', (req, res) => {
  res.json({ 
    status: 'OK', 
    files: Array.from(files.keys()),
    locks: Array.from(locks.entries()),
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
    fileId,
    editUrl: `https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=${encodeURIComponent(wopiSrc)}&access_token=${accessToken}`
  });
});

app.get('/', (req, res) => {
  res.json({ 
    message: 'WOPI Server Running - EDIT MODE FORCED',
    endpoints: {
      health: '/api/health',
      wopiUrl: '/api/generate-wopi-url'
    }
  });
});

app.listen(PORT, () => {
  console.log(`\n🚀 WOPI Server: http://localhost:${PORT}`);
  console.log(`🌐 Ngrok: https://c67feb255965.ngrok-free.app`);
  console.log(`\n🎯 EDIT MODE FORCED - CheckFileInfo configured for editing`);
  console.log(`📁 Available files:`, Array.from(files.keys()));
  console.log(`\n🔑 Key Edit Permissions:`);
  console.log(`   ✅ UserCanWrite: true`);
  console.log(`   ✅ ReadOnly: false`);
  console.log(`   ✅ SupportsUpdate: true`);
  console.log(`   ✅ SupportsLocks: true`);
  console.log(`\n💡 If still in read mode, clear browser cache and try again\n`);
});