// server.js (FIXED LOCK COMPARISON)
const express = require('express');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 8080;
const PUBLIC_BASE = process.env.PUBLIC_BASE || 'https://6e51ccd22f2a.ngrok-free.app';
const FILES_DIR = path.join(__dirname, 'files');
const SECRET = process.env.WOPI_SECRET || 'demo-secret';

// Enable detailed logging
const DEBUG = true;

// --- CORS ---
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, PUT');
  res.header('Access-Control-Allow-Headers', [
    'Content-Type',
    'Authorization',
    'X-Requested-With',
    'X-WOPI-Override',
    'X-WOPI-Lock',
    'X-WOPI-OldLock',
    'X-WOPI-MachineName',
    'X-WOPI-SessionId',
    'X-WOPI-ItemVersion',
    'Prefer'
  ].join(', '));
  if (req.method === 'OPTIONS') return res.status(200).end();
  next();
});

// --- Raw body for binary saves ---
app.use((req, res, next) => {
  const isBinary = req.headers['content-type'] === 'application/octet-stream';
  const hasOverride = !!req.headers['x-wopi-override'];
  if (isBinary || hasOverride) {
    const chunks = [];
    req.on('data', c => chunks.push(c));
    req.on('end', () => {
      req.rawBody = Buffer.concat(chunks);
      next();
    });
  } else {
    express.json()(req, res, next);
  }
});

// --- In-memory indices + lock table ---
const locks = new Map();        // fileId -> lockString (store as string)
const versions = new Map();     // fileId -> version string

// --- Helpers ---
const ensureFilesDir = () => { 
  if (!fs.existsSync(FILES_DIR)) fs.mkdirSync(FILES_DIR, { recursive: true }); 
};
ensureFilesDir();

const filePathOf = (fileId) => path.join(FILES_DIR, fileId);

const readFileOrNull = (fileId) => {
  const p = filePathOf(fileId);
  if (!fs.existsSync(p)) return null;
  return fs.readFileSync(p);
};

const writeFile = (fileId, buf) => {
  ensureFilesDir();
  fs.writeFileSync(filePathOf(fileId), buf);
  const v = Date.now().toString();
  versions.set(fileId, v);
  if (DEBUG) console.log(`📁 [WRITE] ${fileId} - ${buf.length} bytes - version: ${v}`);
  return v;
};

const computeSha256 = (buf) => crypto.createHash('sha256').update(buf).digest('base64');

const generateToken = (fileId, userId = 'test-user') => {
  const payload = `${fileId}:${userId}`;
  return crypto.createHmac('sha256', SECRET).update(payload).digest('base64url');
};

const validateToken = (token, fileId, userId = 'test-user') => {
  try {
    const expected = generateToken(fileId, userId);
    return token === expected;
  } catch {
    return false;
  }
};

const ok = (res, headers = {}) => {
  Object.entries(headers).forEach(([k, v]) => { if (v !== undefined) res.setHeader(k, v); });
  return res.status(200).send();
};

const conflict = (res, currentLock, reason) => {
  res.setHeader('X-WOPI-Lock', currentLock || '');
  if (reason) res.setHeader('X-WOPI-LockFailureReason', reason);
  return res.status(409).send();
};

const getLockHeader = (req) => (req.headers['x-wopi-lock'] || '').toString();
const getOldLockHeader = (req) => (req.headers['x-wopi-oldlock'] || '').toString();

// FIXED: Better lock parsing and comparison
const parseLock = (lockStr) => {
  if (!lockStr || lockStr === '') return null;
  try {
    return JSON.parse(lockStr);
  } catch {
    return lockStr; // return as string if not JSON
  }
};

// FIXED: Much simpler and more reliable lock comparison
const locksMatch = (lock1, lock2) => {
  if (lock1 === lock2) return true; // Quick string match
  if (!lock1 || !lock2) return false;
  
  const parsed1 = parseLock(lock1);
  const parsed2 = parseLock(lock2);
  
  // If both parsed successfully as objects, compare S and F properties
  if (typeof parsed1 === 'object' && typeof parsed2 === 'object') {
    return parsed1.S === parsed2.S && parsed1.F === parsed2.F;
  }
  
  // Fallback to string comparison
  return String(lock1) === String(lock2);
};

// Log all WOPI requests
app.use('/wopi/files/:fileId', (req, res, next) => {
  if (DEBUG) {
    console.log('\n=== WOPI REQUEST ===');
    console.log(`📨 ${req.method} ${req.path}`);
    console.log(`🔑 File: ${req.params.fileId}`);
    console.log(`🔧 Override: ${req.headers['x-wopi-override'] || 'NONE'}`);
    console.log(`🔒 Lock: ${req.headers['x-wopi-lock'] || 'NONE'}`);
    console.log(`🔄 OldLock: ${req.headers['x-wopi-oldlock'] || 'NONE'}`);
    console.log(`📦 Content-Type: ${req.headers['content-type']}`);
    console.log(`📏 Content-Length: ${req.headers['content-length']}`);
  }
  next();
});

// --- Bootstrap sample.docx if missing ---
(() => {
  const id = 'sample.docx';
  const p = filePathOf(id);
  if (!fs.existsSync(p)) {
    const minimal = Buffer.from('UEsDBBQAAAAIAJySrVHHW+cgBAAAAAQAAAAIAAAAbWltZXR5cGVhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtb2ZmaWNlZG9jdW1lbnQud29yZHByb2Nlc3NpbmdtbC5kb2N1bWVudFBLBQYAAAAAAQABAD4AAAA0AAAAAAA=', 'base64');
    writeFile(id, minimal);
  } else {
    versions.set(id, fs.statSync(p).mtimeMs.toString());
  }
})();

// ===================== WOPI ROUTES =====================

// CheckFileInfo
app.get('/wopi/files/:fileId', (req, res) => {
  const { fileId } = req.params;
  const token = req.query.access_token;
  
  if (DEBUG) console.log(`🔍 [CheckFileInfo] ${fileId}`);
  
  if (!validateToken(token, fileId)) return res.status(401).json({ error: 'Invalid token' });

  const buf = readFileOrNull(fileId);
  if (!buf) return res.status(404).json({ error: 'File not found' });

  const stat = fs.statSync(filePathOf(fileId));
  const currentLock = locks.get(fileId);
  
  const info = {
    BaseFileName: fileId,
    OwnerId: 'test-owner',
    Size: buf.length,
    UserId: 'test-user',
    Version: versions.get(fileId) || stat.mtimeMs.toString(),
    UserCanWrite: true,
    ReadOnly: false,
    SupportsUpdate: true,
    SupportsLocks: true,
    SupportsGetLock: true,
    SupportsRename: false,
    SupportsDeleteFile: false,
    SupportsCobalt: false,
    SupportsPutRelativeFile: false,
    // CRITICAL FOR AUTO-SAVE:
    SupportsExtendedLockLength: true,
    SupportsEcosystem: false,
    SupportsGetLock: true,
    SupportsFolders: false,
    LastModifiedTime: new Date(stat.mtime).toISOString(),
    SHA256: computeSha256(buf),
    UserFriendlyName: 'Test User',
    HostEditUrl: `${PUBLIC_BASE}/wopi/files/${encodeURIComponent(fileId)}`,
    HostViewUrl: `${PUBLIC_BASE}/wopi/files/${encodeURIComponent(fileId)}`,
    UserCanNotWriteRelative: true,
    BreadcrumbBrandName: 'WOPI Test',
    BreadcrumbBrandUrl: PUBLIC_BASE,
    BreadcrumbFolderName: 'files',
    BreadcrumbDocName: fileId
  };

  if (DEBUG) console.log(`✅ [CheckFileInfo Response]`, { 
    size: info.Size, 
    version: info.Version,
    locked: !!currentLock 
  });

  res.json(info);
});

// GetFile (download content)
app.get('/wopi/files/:fileId/contents', (req, res) => {
  const { fileId } = req.params;
  const token = req.query.access_token;
  
  if (DEBUG) console.log(`📥 [GetFile] ${fileId}`);
  
  if (!validateToken(token, fileId)) return res.status(401).json({ error: 'Invalid token' });

  const buf = readFileOrNull(fileId);
  if (!buf) return res.status(404).json({ error: 'File not found' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Length', buf.length);
  res.send(buf);
  
  if (DEBUG) console.log(`✅ [GetFile Sent] ${buf.length} bytes`);
});

// PutFile (save) - FIXED LOCK HANDLING
app.post('/wopi/files/:fileId/contents', (req, res) => {
  const { fileId } = req.params;
  const token = req.query.access_token;
  
  if (DEBUG) console.log(`💾 [PutFile] ${fileId} - Starting save operation`);
  
  if (!validateToken(token, fileId)) return res.status(401).json({ error: 'Invalid token' });

  const incomingLock = getLockHeader(req);
  const currentLock = locks.get(fileId) || null;
  const oldLock = getOldLockHeader(req);

  if (DEBUG) console.log(`🔒 [PutFile Locks] current: ${currentLock}, incoming: ${incomingLock}, old: ${oldLock}`);

  // FIXED: Use intelligent lock comparison
  if (currentLock) {
    const incomingMatchesCurrent = locksMatch(incomingLock, currentLock);
    const oldLockMatchesCurrent = locksMatch(oldLock, currentLock);
    
    if (DEBUG) console.log(`🔍 [PutFile Lock Check] incomingMatch: ${incomingMatchesCurrent}, oldLockMatch: ${oldLockMatchesCurrent}`);
    
    if (!incomingMatchesCurrent && !oldLockMatchesCurrent) {
      if (DEBUG) console.log(`❌ [PutFile Conflict] Lock mismatch - returning 409`);
      return conflict(res, currentLock, 'File is locked by another session');
    }
  }

  const body = req.rawBody;
  if (!body || !Buffer.isBuffer(body) || body.length === 0) {
    if (DEBUG) console.log(`❌ [PutFile Error] No content provided`);
    return res.status(400).json({ error: 'No file content' });
  }

  if (DEBUG) console.log(`📦 [PutFile Content] ${body.length} bytes received`);

  const newVersion = writeFile(fileId, body);
  
  // Update lock if this is a new/extended lock
  if (incomingLock && !locksMatch(incomingLock, currentLock)) {
    locks.set(fileId, incomingLock);
    if (DEBUG) console.log(`🔒 [PutFile Lock Updated] ${incomingLock}`);
  }

  if (DEBUG) console.log(`✅ [PutFile Success] Saved ${body.length} bytes, version: ${newVersion}`);

  return ok(res, {
    'X-WOPI-ItemVersion': newVersion,
    'X-WOPI-Lock': locks.get(fileId) || ''
  });
});

// LOCK / UNLOCK / REFRESH_LOCK / GET_LOCK - FIXED LOCK HANDLING
app.post('/wopi/files/:fileId', (req, res) => {
  const { fileId } = req.params;
  const token = req.query.access_token;
  
  if (!validateToken(token, fileId)) return res.status(401).json({ error: 'Invalid token' });

  const override = (req.headers['x-wopi-override'] || '').toString().toUpperCase();
  const incomingLock = getLockHeader(req);
  const currentLock = locks.get(fileId) || null;

  if (DEBUG) console.log(`🛠️ [Override: ${override}] ${fileId} - lock: ${incomingLock}`);

  switch (override) {
    case 'LOCK': {
      if (!incomingLock) {
        if (DEBUG) console.log(`❌ [LOCK Error] No lock provided`);
        return res.status(400).send();
      }
      
      // FIXED: Use intelligent lock comparison
      const locksMatchResult = locksMatch(incomingLock, currentLock);
      if (DEBUG) console.log(`🔍 [LOCK Comparison] current: ${currentLock}, incoming: ${incomingLock}, match: ${locksMatchResult}`);
      
      if (!currentLock || locksMatchResult) {
        // Always update the lock to the latest version (Office extends locks with more metadata)
        locks.set(fileId, incomingLock);
        if (DEBUG) console.log(`🔒 [LOCK Acquired/Extended] ${incomingLock}`);
        return ok(res);
      }
      
      if (DEBUG) console.log(`❌ [LOCK Conflict] Already locked by: ${currentLock}`);
      return conflict(res, currentLock, 'Already locked by a different lock');
    }

    case 'UNLOCK': {
      if (!incomingLock) {
        if (DEBUG) console.log(`❌ [UNLOCK Error] No lock provided`);
        return res.status(400).send();
      }
      
      // FIXED: Use intelligent lock comparison
      if (!currentLock || locksMatch(incomingLock, currentLock)) {
        locks.delete(fileId);
        if (DEBUG) console.log(`🔓 [UNLOCK Success] Lock removed`);
        return ok(res);
      }
      
      if (DEBUG) console.log(`❌ [UNLOCK Conflict] Wrong lock: ${incomingLock}, expected: ${currentLock}`);
      return conflict(res, currentLock, 'Unlock with wrong lock');
    }

    case 'REFRESH_LOCK': {
      if (!incomingLock) {
        if (DEBUG) console.log(`❌ [REFRESH_LOCK Error] No lock provided`);
        return res.status(400).send();
      }
      
      // FIXED: Use intelligent lock comparison
      if (locksMatch(incomingLock, currentLock)) {
        // Update to the latest lock version
        locks.set(fileId, incomingLock);
        if (DEBUG) console.log(`🔄 [REFRESH_LOCK Success] Lock refreshed and updated`);
        return ok(res);
      }
      
      if (DEBUG) console.log(`❌ [REFRESH_LOCK Conflict] Wrong lock: ${incomingLock}, expected: ${currentLock}`);
      return conflict(res, currentLock, 'Refresh with wrong lock');
    }

    case 'GET_LOCK': {
      if (DEBUG) console.log(`🔍 [GET_LOCK] Current lock: ${currentLock || 'NONE'}`);
      if (currentLock) res.setHeader('X-WOPI-Lock', currentLock);
      return res.status(200).send();
    }

    case 'PUT_RELATIVE': {
      if (DEBUG) console.log(`❌ [PUT_RELATIVE] Disabled for now`);
      return res.status(501).json({ error: 'PutRelative not supported' });
    }

    default:
      if (DEBUG) console.log(`❌ [UNKNOWN Override] ${override}`);
      return res.status(400).json({ error: 'Unsupported X-WOPI-Override' });
  }
});

// ---------- Convenience APIs ----------
app.get('/api/generate-wopi-url', (req, res) => {
  const fileId = 'sample.docx';
  if (!readFileOrNull(fileId)) return res.status(404).json({ error: 'File not found' });

  const token = generateToken(fileId);
  const wopiSrc = `${PUBLIC_BASE}/wopi/files/${encodeURIComponent(fileId)}`;
  const editUrl = `https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=${encodeURIComponent(wopiSrc)}&access_token=${encodeURIComponent(token)}&ui=en-US&rs=en-US`;

  if (DEBUG) console.log(`🔗 [Generate URL] ${fileId}`);
  
  res.json({ fileId, wopiSrc, accessToken: token, editUrl });
});

app.get('/api/files', (req, res) => {
  const all = fs.readdirSync(FILES_DIR).map(f => {
    const st = fs.statSync(path.join(FILES_DIR, f));
    return {
      fileId: f,
      size: st.size,
      version: versions.get(f) || st.mtimeMs.toString(),
      lastModified: new Date(st.mtime).toISOString()
    };
  });
  res.json({ files: all, total: all.length });
});

app.get('/api/debug-file/:fileId', (req, res) => {
  const { fileId } = req.params;
  const p = filePathOf(fileId);
  if (!fs.existsSync(p)) return res.status(404).json({ error: 'File not found' });
  const buf = fs.readFileSync(p);
  const st = fs.statSync(p);
  res.json({
    fileId,
    size: buf.length,
    version: versions.get(fileId) || st.mtimeMs.toString(),
    lastModified: new Date(st.mtime).toISOString(),
    sha256: computeSha256(buf).slice(0, 16) + '...',
    locked: !!locks.get(fileId),
    lockValue: locks.get(fileId) || null
  });
});

// Clear locks endpoint
app.post('/api/clear-locks', (req, res) => {
  const count = locks.size;
  locks.clear();
  if (DEBUG) console.log(`🧹 [Clear Locks] Cleared ${count} locks`);
  res.json({ message: `Cleared ${count} locks` });
});

// Reset sample file
app.post('/api/reset-sample', (req, res) => {
  const fileId = 'sample.docx';
  const minimal = Buffer.from('UEsDBBQAAAAIAJySrVHHW+cgBAAAAAQAAAAIAAAAbWltZXR5cGVhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvrm1hdHMtb2ZmaWNlZG9jdW1lbnQud29yZHByb2Nlc3NpbmdtbC5kb2N1bWVudFBLBQYAAAAAAQABAD4AAAA0AAAAAAA=', 'base64');
  writeFile(fileId, minimal);
  locks.delete(fileId);
  if (DEBUG) console.log(`🔄 [Reset Sample] Reset sample.docx`);
  res.json({ message: 'Sample file reset' });
});

app.get('/', (_req, res) => {
  res.json({
    message: 'WOPI Server Running (Fixed Lock Comparison)',
    base: PUBLIC_BASE,
    endpoints: {
      health: '/api/files',
      wopiUrl: '/api/generate-wopi-url',
      clearLocks: '/api/clear-locks (POST)',
      resetSample: '/api/reset-sample (POST)',
      debugFile: '/api/debug-file/sample.docx'
    }
  });
});

app.listen(PORT, () => {
  console.log(`🚀 WOPI Server (Fixed Lock Comparison) listening on http://localhost:${PORT}`);
  console.log(`📊 Debug logging: ${DEBUG}`);
  console.log(`📁 Files directory: ${FILES_DIR}`);
});