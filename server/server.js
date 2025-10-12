const express = require('express');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 8080;

// CORS setup
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, OPTIONS, LOCK, UNLOCK');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With, X-WOPI-Override, X-WOPI-Lock, X-WOPI-OldLock, X-WOPI-MachineName, X-WOPI-SessionId, X-WOPI-ItemVersion');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  
  next();
});

// Body parsing
app.use((req, res, next) => {
  if (req.headers['content-type'] === 'application/octet-stream' || 
      req.headers['x-wopi-override']) {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => {
      req.rawBody = Buffer.concat(chunks);
      next();
    });
  } else {
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
    console.log('⚠️  Creating minimal DOCX file...');
    const minimalDocx = createMinimalDocx();
    files.set('sample.docx', {
      name: 'sample.docx',
      size: minimalDocx.length,
      lastModified: new Date().toISOString(),
      version: Date.now().toString(),
      sha256: crypto.createHash('sha256').update(minimalDocx).digest('base64'),
      contents: minimalDocx
    });
  }
};

function createMinimalDocx() {
  const base64Docx = 'UEsDBBQAAAAIAJySrVHHW+cgBAAAAAQAAAAIAAAAbWltZXR5cGVhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtb2ZmaWNlZG9jdW1lbnQud29yZHByb2Nlc3NpbmdtbC5kb2N1bWVudFBLBQYAAAAAAQABAD4AAAA0AAAAAAA=';
  return Buffer.from(base64Docx, 'base64');
}

initializeSampleFile();

// Generate access token - UPDATED to work with any file
const generateAccessToken = (fileId) => {
  const tokenData = {
    fileId: fileId,
    timestamp: Date.now(),
    userId: 'test-user'
  };
  return Buffer.from(JSON.stringify(tokenData)).toString('base64');
};

// Validate access token - UPDATED to work with any file
const validateAccessToken = (token, fileId) => {
  try {
    const decoded = JSON.parse(Buffer.from(token, 'base64').toString());
    // Check if file exists in our storage
    return files.has(decoded.fileId) && decoded.fileId === fileId;
  } catch (error) {
    return false;
  }
};

// CheckFileInfo - UPDATED to handle any file
app.get('/wopi/files/:fileId', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  
  console.log('📋 CheckFileInfo for:', fileId);
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    console.log('❌ Invalid access token for file:', fileId);
    return res.status(401).json({ error: 'Invalid token' });
  }

  const file = files.get(fileId);
  if (!file) {
    console.log('❌ File not found:', fileId);
    return res.status(404).json({ error: 'File not found' });
  }

  const fileInfo = {
    // Required properties
    BaseFileName: file.name,
    Size: file.size,
    OwnerId: 'test-owner',
    UserId: 'test-user',
    Version: file.version,
    
    // Edit permissions
    UserCanWrite: true,
    ReadOnly: false,
    UserCanNotWriteRelative: false,
    
    // User information
    UserFriendlyName: 'Test User',
    
    // File properties
    SHA256: file.sha256,
    LastModifiedTime: file.lastModified,
    
    // Supported features
    SupportsUpdate: true,
    SupportsLocks: true,
    SupportsGetLock: true,
    SupportsCobalt: true,
    SupportsPutRelativeFile: true,
    SupportsRename: true,
    SupportsDeleteFile: true,
    
    // ACTION URLs - Critical for edit mode
    HostEditUrl: `https://c67feb255965.ngrok-free.app/wopi/files/${fileId}`,
    HostViewUrl: `https://c67feb255965.ngrok-free.app/wopi/files/${fileId}`,
    
    // Additional properties
    AllowExternalMarketplace: false,
    DisablePrint: false,
    DisableTranslation: false,
    LicenseCheckForEditIsEnabled: false,
    UserCanAttend: true,
    UserCanPresent: true,
    WebEditingDisabled: false
  };

  console.log('✅ CheckFileInfo - EDIT MODE ENABLED for:', fileId);
  res.json(fileInfo);
});

// GetFile - UPDATED to handle any file
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

// PutFile - UPDATED to handle any file
app.post('/wopi/files/:fileId/contents', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  
  console.log('💾 PutFile for:', fileId, 'Size:', req.rawBody?.length || 'unknown');
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    console.log('❌ Invalid access token for PutFile:', fileId);
    return res.status(401).json({ error: 'Invalid token' });
  }

  try {
    // Check if we have content
    if (!req.rawBody || req.rawBody.length === 0) {
      console.log('❌ Empty content in PutFile request for:', fileId);
      return res.status(400).json({ error: 'No file content provided' });
    }

    const sha256 = crypto.createHash('sha256').update(req.rawBody).digest('base64');
    
    // Create or update the file
    const existingFile = files.get(fileId);
    const updatedFile = {
      name: fileId,
      size: req.rawBody.length,
      lastModified: new Date().toISOString(),
      version: Date.now().toString(),
      sha256: sha256,
      contents: req.rawBody
    };
    
    files.set(fileId, updatedFile);

    console.log(`✅ File saved: ${fileId} - ${req.rawBody.length} bytes`);
    
    // Return proper WOPI response headers
    res.setHeader('X-WOPI-ItemVersion', updatedFile.version);
    res.setHeader('X-WOPI-Lock', locks.get(fileId) || '');
    
    res.status(200).send();
    
  } catch (error) {
    console.error('❌ PutFile save error for:', fileId, error);
    res.status(500).json({ error: 'Save failed' });
  }
});

// Lock operations - UPDATED to handle any file
app.post('/wopi/files/:fileId/lock', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  const lockId = req.headers['x-wopi-lock'];
  
  console.log('🔒 Lock file:', fileId, 'Lock:', lockId);
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }

  // File doesn't need to exist for locking - it might be created soon
  const currentLock = locks.get(fileId);
  
  if (!currentLock) {
    locks.set(fileId, lockId);
    console.log('✅ Lock acquired for:', fileId);
    res.status(200).send();
  } else if (currentLock === lockId) {
    console.log('✅ Lock refreshed for:', fileId);
    res.status(200).send();
  } else {
    console.log('❌ File already locked:', fileId, 'with:', currentLock);
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

  const lock = locks.get(fileId);
  if (lock) {
    res.setHeader('X-WOPI-Lock', lock);
    console.log('✅ Current lock for', fileId, ':', lock);
  }
  
  res.status(200).send();
});

app.unlock('/wopi/files/:fileId/lock', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  const lockId = req.headers['x-wopi-lock'];
  
  console.log('🔓 Unlock file:', fileId, 'Lock:', lockId);
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }

  const currentLock = locks.get(fileId);
  
  if (currentLock === lockId) {
    locks.delete(fileId);
    console.log('✅ Lock released for:', fileId);
    res.status(200).send();
  } else {
    console.log('❌ Lock mismatch for:', fileId, 'Current:', currentLock, 'Requested:', lockId);
    res.setHeader('X-WOPI-Lock', currentLock || '');
    res.status(409).send();
  }
});

// // PutRelativeFile - FIXED to generate proper access tokens for new files
app.post('/wopi/files/:fileId', (req, res) => {
  const { fileId } = req.params;
  const accessToken = req.query.access_token;
  const suggestedTarget = req.headers['x-wopi-suggestedtarget'];
  
  console.log('📝 PutRelativeFile for:', fileId, 'Suggested:', suggestedTarget);
  
  if (!accessToken || !validateAccessToken(accessToken, fileId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }

  try {
    if (!req.rawBody) {
      return res.status(400).json({ error: 'No file content provided' });
    }

    const newFileId = suggestedTarget || `document-${Date.now()}.docx`;
    const sha256 = crypto.createHash('sha256').update(req.rawBody).digest('base64');
    
    // Create the new file
    files.set(newFileId, {
      name: newFileId,
      size: req.rawBody.length,
      lastModified: new Date().toISOString(),
      version: Date.now().toString(),
      sha256: sha256,
      contents: req.rawBody
    });

    console.log(`✅ New file created: ${newFileId} (${req.rawBody.length} bytes)`);
    
    // Generate access token for the new file
    const newFileAccessToken = generateAccessToken(newFileId);
    const newFileWopiSrc = `https://c67feb255965.ngrok-free.app/wopi/files/${newFileId}`;
    
    res.json({
      Name: newFileId,
      Url: newFileWopiSrc
    });
    
  } catch (error) {
    console.error('❌ PutRelativeFile error:', error);
    res.status(500).json({ error: 'Create file failed' });
  }
});

// PutRelativeFile - DISABLED to prevent new file creation
// app.post('/wopi/files/:fileId', (req, res) => {
//   console.log('❌ PutRelativeFile called - returning error to force save to sample.docx');
//   res.status(501).json({ error: 'Save As not supported - use Save to update original file' });
// });

// NEW: Endpoint to list all files (for debugging)
app.get('/api/files', (req, res) => {
  const fileList = Array.from(files.entries()).map(([fileId, file]) => ({
    fileId,
    name: file.name,
    size: file.size,
    version: file.version,
    lastModified: file.lastModified
  }));
  
  res.json({
    files: fileList,
    total: fileList.length
  });
});

// Debug endpoint to check file status
app.get('/api/debug-file/:fileId', (req, res) => {
  const { fileId } = req.params;
  const file = files.get(fileId);
  
  if (!file) {
    return res.status(404).json({ error: 'File not found' });
  }
  
  res.json({
    fileId,
    exists: true,
    size: file.size,
    version: file.version,
    lastModified: file.lastModified,
    sha256: file.sha256.substring(0, 16) + '...',
    hasLock: !!locks.get(fileId)
  });
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
  
  const editUrl = `https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=${encodeURIComponent(wopiSrc)}&access_token=${accessToken}&ui=en-US&rs=en-US`;
  
  res.json({
    wopiSrc,
    accessToken,
    fileId,
    editUrl
  });
});

app.get('/', (req, res) => {
  res.json({ 
    message: 'WOPI Server Running - COMPLETE FILE MANAGEMENT',
    endpoints: {
      health: '/api/health',
      wopiUrl: '/api/generate-wopi-url',
      files: '/api/files',
      debugFile: '/api/debug-file/:fileId'
    }
  });
});

app.listen(PORT, () => {
  console.log(`\n🚀 WOPI Server: http://localhost:${PORT}`);
  console.log(`🌐 Ngrok: https://c67feb255965.ngrok-free.app`);
  console.log(`\n✨ COMPLETE FILE MANAGEMENT - All files supported`);
  console.log(`📁 Initial files:`, Array.from(files.keys()));
  console.log(`\n💡 Visit /api/files to see all created files`);
});