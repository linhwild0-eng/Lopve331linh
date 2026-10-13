require('dotenv').config();

const express = require('express');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const sqlite3 = require('sqlite3').verbose();
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;
const DB_PATH = path.resolve(process.env.NTTCLASS_DB_PATH || path.join(__dirname, 'nttclass.db'));
const PUBLIC_DIR = path.join(__dirname, 'public');
const INDEX_PATH = path.join(PUBLIC_DIR, 'index.html');
const ACCOUNT_CREATION_CODE = process.env.NTTCLASS_ACCOUNT_CREATION_CODE || '';
const SESSION_TTL_MS = 8 * 60 * 60 * 1000;
const sessions = new Map();
let appDataWriteQueue = Promise.resolve();

if (!fs.existsSync(INDEX_PATH)) {
  throw new Error(`Required website file not found: ${INDEX_PATH}. Deploy the public/index.html file with the backend.`);
}
if (ACCOUNT_CREATION_CODE && Buffer.byteLength(ACCOUNT_CREATION_CODE, 'utf8') < 12) {
  throw new Error('NTTCLASS_ACCOUNT_CREATION_CODE must be at least 12 bytes long.');
}

fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });

const isProduction = process.env.NODE_ENV === 'production';
const DEFAULT_ADMIN = {
  username: process.env.NTTCLASS_ADMIN_USERNAME || (isProduction ? '' : 'admin'),
  password: process.env.NTTCLASS_ADMIN_PASSWORD || (isProduction ? '' : 'ChangeMe123!'),
  displayName: process.env.NTTCLASS_ADMIN_DISPLAY_NAME || (isProduction ? '' : 'Local Admin'),
  role: 'admin'
};

if (!DEFAULT_ADMIN.username || !DEFAULT_ADMIN.password || !DEFAULT_ADMIN.displayName.trim()) {
  throw new Error('Missing administrator seed settings. Set NTTCLASS_ADMIN_USERNAME, NTTCLASS_ADMIN_PASSWORD, and NTTCLASS_ADMIN_DISPLAY_NAME before starting the app.');
}

if (!/^[a-zA-Z0-9_.-]{3,40}$/.test(DEFAULT_ADMIN.username) ||
    DEFAULT_ADMIN.password.length < 8 ||
    DEFAULT_ADMIN.displayName.trim().length > 80) {
  throw new Error('Invalid administrator seed settings. Check NTTCLASS_ADMIN_* environment variables.');
}

const DEFAULT_PERMISSIONS = {
  teacher: {
    classes: { view: true, create: false, update: false, delete: false },
    students: { view: true, create: true, update: true, delete: false },
    attendance: { view: true, manage: true },
    grades: { view: true, manage: true },
    staffAttendance: { view: true, manage: true },
    payroll: { view: false, manage: false },
    finance: { view: false, manage: false },
    reports: { view: true }
  },
  assistant: {
    classes: { view: true, create: false, update: false, delete: false },
    students: { view: true, create: true, update: true, delete: false },
    attendance: { view: true, manage: true },
    grades: { view: false, manage: false },
    staffAttendance: { view: true, manage: true },
    payroll: { view: false, manage: false },
    finance: { view: false, manage: false },
    reports: { view: true }
  }
};

const COLLECTIONS = {
  classes: ['classes', (item) => item.id],
  students: ['students', (item) => item.id],
  attendance: ['attendance', (item) => `${item.studentId}:${item.date}`],
  staffAttendance: ['staffAttendance', (item) => `${item.userId}:${item.date}`],
  assessments: ['grades', (item) => item.id],
  payments: ['finance', (item) => `${item.studentId}:${item.month}`]
};

const db = new sqlite3.Database(DB_PATH, (error) => {
  if (error) {
    console.error('Failed to open database:', error.message);
    process.exit(1);
    return;
  }
  initializeDatabase().then(() => {
    app.listen(PORT, () => {
      console.log(`NTTCLASS server listening on port ${PORT}`);
    });
  }).catch((initializationError) => {
    console.error('Failed to initialize database:', initializationError.message);
    db.close(() => process.exit(1));
  });
});

function serializeDataWrite(operation) {
  const result = appDataWriteQueue.then(operation, operation);
  appDataWriteQueue = result.then(() => undefined, () => undefined);
  return result;
}

function makeSaltHex() {
  return crypto.randomBytes(16).toString('hex');
}

function hashPassword(password, saltHex) {
  return crypto.pbkdf2Sync(password, Buffer.from(saltHex, 'hex'), 150000, 32, 'sha256').toString('hex');
}

function matchesAccountCreationCode(candidate) {
  if (typeof candidate !== 'string' || !ACCOUNT_CREATION_CODE) return false;
  const candidateBytes = Buffer.from(candidate, 'utf8');
  const configuredBytes = Buffer.from(ACCOUNT_CREATION_CODE, 'utf8');
  return candidateBytes.length === configuredBytes.length &&
    crypto.timingSafeEqual(candidateBytes, configuredBytes);
}

function buildDefaultData() {
  const salt = makeSaltHex();
  return {
    classes: [],
    staff: [],
    students: [],
    assessments: [],
    attendance: [],
    payments: [],
    accounts: [{
      id: crypto.randomUUID(),
      displayName: DEFAULT_ADMIN.displayName,
      username: DEFAULT_ADMIN.username,
      role: DEFAULT_ADMIN.role,
      salt,
      passwordHash: hashPassword(DEFAULT_ADMIN.password, salt)
    }],
    permissions: structuredClone(DEFAULT_PERMISSIONS)
  };
}

function ensureConfiguredAdminAccount(data) {
  if (!data || !Array.isArray(data.accounts)) return data;
  const envUsername = String(process.env.NTTCLASS_ADMIN_USERNAME || '').trim();
  const envPassword = String(process.env.NTTCLASS_ADMIN_PASSWORD || '');
  const envDisplayName = String(process.env.NTTCLASS_ADMIN_DISPLAY_NAME || '').trim();
  if (!envUsername || !envPassword || !envDisplayName) return data;

  const adminAccount = data.accounts.find(account => account.role === 'admin') || null;
  if (!adminAccount) {
    const salt = makeSaltHex();
    data.accounts.push({
      id: crypto.randomUUID(),
      displayName: envDisplayName,
      username: envUsername,
      role: 'admin',
      salt,
      passwordHash: hashPassword(envPassword, salt)
    });
    return data;
  }

  const needsUpdate = adminAccount.username !== envUsername || adminAccount.displayName !== envDisplayName ||
    adminAccount.passwordHash !== hashPassword(envPassword, adminAccount.salt || makeSaltHex());

  if (!needsUpdate) return data;

  adminAccount.displayName = envDisplayName;
  adminAccount.username = envUsername;
  adminAccount.salt = makeSaltHex();
  adminAccount.passwordHash = hashPassword(envPassword, adminAccount.salt);
  return data;
}

function normalizePermissions(value) {
  const normalized = structuredClone(DEFAULT_PERMISSIONS);
  for (const role of ['teacher', 'assistant']) {
    const supplied = value && value[role];
    if (!supplied || typeof supplied !== 'object') continue;
    for (const module of Object.keys(normalized[role])) {
      const oldValue = supplied[module];
      if (oldValue && typeof oldValue === 'object') {
        for (const action of Object.keys(normalized[role][module])) {
          if (typeof oldValue[action] === 'boolean') normalized[role][module][action] = oldValue[action];
        }
      } else if (typeof oldValue === 'boolean') {
        if (module === 'classes' || module === 'students') {
          normalized[role][module].view = oldValue;
          normalized[role][module].create = oldValue;
          normalized[role][module].update = oldValue;
        } else if (module === 'attendance' || module === 'grades' || module === 'staffAttendance' || module === 'finance') {
          normalized[role][module].view = oldValue;
          normalized[role][module].manage = oldValue;
        }
      }
    }
    for (const module of ['attendance', 'grades', 'staffAttendance', 'finance']) {
      if (normalized[role][module].manage) normalized[role][module].view = true;
    }
    for (const module of ['classes', 'students']) {
      if (['create', 'update', 'delete'].some(action => normalized[role][module][action])) {
        normalized[role][module].view = true;
      }
    }
  }
  return normalized;
}

function normalizeData(value) {
  const data = value && typeof value === 'object' ? value : {};
  return {
    ...data,
    classes: Array.isArray(data.classes) ? data.classes : [],
    staff: Array.isArray(data.staff) ? data.staff : [],
    students: Array.isArray(data.students)
      ? data.students.map(({ sessionTuition, ...student }) => student)
      : [],
    assessments: Array.isArray(data.assessments) ? data.assessments : [],
    attendance: Array.isArray(data.attendance) ? data.attendance : [],
    staffAttendance: Array.isArray(data.staffAttendance) ? data.staffAttendance : [],
    payments: Array.isArray(data.payments) ? data.payments : [],
    accounts: Array.isArray(data.accounts) ? data.accounts : [],
    permissions: normalizePermissions(data.permissions)
  };
}

function initializeDatabase() {
  return new Promise((resolve, reject) => {
    db.run('CREATE TABLE IF NOT EXISTS app_state (key TEXT PRIMARY KEY, value TEXT NOT NULL)', (error) => {
      if (error) return reject(error);
      db.get("SELECT value FROM app_state WHERE key = 'app_data'", (readError, row) => {
        if (readError) return reject(readError);
        if (!row) {
          return writeAppData(buildDefaultData()).then(() => {
            console.log('Default administrator account initialized.');
            resolve();
          }, reject);
        }
        try {
          const normalized = normalizeData(JSON.parse(row.value));
          writeAppData(normalized).then(() => {
            console.log('Database ready.');
            resolve();
          }, reject);
        } catch (parseError) {
          reject(parseError);
        }
      });
    });
  });
}

function readAppData() {
  return new Promise((resolve, reject) => {
    db.get("SELECT value FROM app_state WHERE key = 'app_data'", (error, row) => {
      if (error) return reject(error);
      if (!row) return resolve(ensureConfiguredAdminAccount(buildDefaultData()));
      try {
        resolve(ensureConfiguredAdminAccount(normalizeData(JSON.parse(row.value))));
      } catch (parseError) {
        reject(parseError);
      }
    });
  });
}

function writeAppData(data) {
  return new Promise((resolve, reject) => {
    const normalized = normalizeData(ensureConfiguredAdminAccount(data));
    db.run(
      `INSERT INTO app_state (key, value) VALUES ('app_data', ?)
       ON CONFLICT(key) DO UPDATE SET value = excluded.value`,
      [JSON.stringify(normalized)],
      (error) => error ? reject(error) : resolve()
    );
  });
}

function sanitizeAccount(account) {
  const { salt, passwordHash, hourlyRate, ...safe } = account;
  return safe;
}

function dataForUser(data, user) {
  const safe = {
    ...data,
    accounts: data.accounts.map(sanitizeAccount)
  };
  if (user.role !== 'admin') {
    const permissions = data.permissions[user.role] || {};
    safe.accounts = [];
    safe.staff = [];
    if (!permissions.students?.view) {
      safe.students = data.students.map(({ id, name, classId, tuition }) => ({
        id,
        name,
        ...(classId ? { classId } : {}),
        ...(permissions.finance?.view ? { tuition } : {})
      }));
    }
    if (!data.permissions[user.role]?.finance?.view) {
      safe.classes = safe.classes.map(({ tuition, ...item }) => item);
      safe.students = safe.students.map(({ tuition, ...item }) => item);
      safe.payments = [];
    }
    if (!data.permissions[user.role]?.attendance?.view) safe.attendance = [];
    if (!data.permissions[user.role]?.grades?.view) safe.assessments = [];
    if (!data.permissions[user.role]?.staffAttendance?.view) safe.staffAttendance = [];
    else safe.staffAttendance = safe.staffAttendance.filter(item => item.userId === user.id);
  }
  return safe;
}

function can(user, data, module, action) {
  return user.role === 'admin' || data.permissions[user.role]?.[module]?.[action] === true;
}

function makeError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}

function isValidDate(value) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const [year, month, day] = value.split('-').map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

function isValidMonth(value) {
  return typeof value === 'string' && /^\d{4}-(0[1-9]|1[0-2])$/.test(value);
}

async function authenticate(req, res, next) {
  const authorization = req.get('authorization') || '';
  const match = authorization.match(/^Bearer ([a-f0-9]{64})$/i);
  if (!match) return res.status(401).json({ error: 'Vui lòng đăng nhập lại.' });
  const session = sessions.get(match[1]);
  if (!session || session.expiresAt <= Date.now()) {
    sessions.delete(match[1]);
    return res.status(401).json({ error: 'Phiên đăng nhập đã hết hạn.' });
  }
  try {
    const data = await readAppData();
    const account = data.accounts.find(item => item.id === session.accountId);
    if (!account) {
      sessions.delete(match[1]);
      return res.status(401).json({ error: 'Tài khoản không còn hoạt động.' });
    }
    req.user = sanitizeAccount(account);
    req.sessionToken = match[1];
    req.appData = data;
    session.expiresAt = Date.now() + SESSION_TTL_MS;
    return next();
  } catch (error) {
    console.error('Authentication failed:', error.message);
    return res.status(500).json({ error: 'Không thể xác thực phiên đăng nhập.' });
  }
}

function validateCollectionItem(name, item) {
  if (!item || typeof item !== 'object' || Array.isArray(item)) {
    throw makeError(400, `Bản ghi ${name} không hợp lệ.`);
  }
  const requiredFields = {
    classes: ['id', 'name'],
    students: ['id', 'name'],
    attendance: ['studentId', 'date', 'present'],
    staffAttendance: ['userId', 'date', 'status'],
    assessments: ['id', 'studentId', 'title', 'date', 'score', 'maxScore', 'progress'],
    payments: ['studentId', 'month', 'paid']
  }[name];
  if (!requiredFields.every(field => item[field] !== undefined && item[field] !== '')) {
    throw makeError(400, `Thiếu dữ liệu bắt buộc cho ${name}.`);
  }
  if (name === 'classes' && (!Number.isFinite(Number(item.tuition)) || Number(item.tuition) < 0)) {
    throw makeError(400, 'Học phí lớp không hợp lệ.');
  }
  if (name === 'classes' && (typeof item.name !== 'string' || !item.name.trim() || item.name.length > 80)) {
    throw makeError(400, 'Tên lớp không hợp lệ.');
  }
  if (name === 'students' && (typeof item.name !== 'string' || !item.name.trim() || item.name.length > 80 ||
      (item.phone !== undefined && (typeof item.phone !== 'string' || item.phone.length > 30)) ||
      (item.classId !== undefined && typeof item.classId !== 'string'))) {
    throw makeError(400, 'Thông tin học sinh không hợp lệ.');
  }
  if (name === 'students' && item.tuition !== undefined && (!Number.isFinite(Number(item.tuition)) || Number(item.tuition) < 0)) {
    throw makeError(400, 'Học phí học sinh không hợp lệ.');
  }
  if (name === 'attendance' && typeof item.present !== 'boolean') {
    throw makeError(400, 'Trạng thái điểm danh không hợp lệ.');
  }
  if (name === 'attendance' && !isValidDate(item.date)) {
    throw makeError(400, 'Ngày điểm danh không hợp lệ.');
  }
  if (name === 'staffAttendance' &&
      (!isValidDate(item.date) ||
       !['present', 'late', 'absent', 'leave'].includes(item.status) ||
       (item.hours !== undefined && (!Number.isFinite(Number(item.hours)) || Number(item.hours) < 0 || Number(item.hours) > 24)) ||
       (item.hours !== undefined && Math.abs(Number(item.hours) * 4 - Math.round(Number(item.hours) * 4)) > 1e-8) ||
       (['absent', 'leave'].includes(item.status) && Number(item.hours || 0) > 0) ||
       (item.note !== undefined && (typeof item.note !== 'string' || item.note.length > 300)))) {
    throw makeError(400, 'Thông tin điểm danh nhân sự không hợp lệ.');
  }
  if (name === 'assessments') {
    const validProgress = ['not_started', 'in_progress', 'completed', 'needs_support'];
    if (!isValidDate(item.date) ||
        typeof item.title !== 'string' || !item.title.trim() || item.title.trim().length > 120 ||
        !Number.isFinite(Number(item.score)) ||
        !Number.isFinite(Number(item.maxScore)) ||
        Number(item.score) < 0 || Number(item.maxScore) <= 0 ||
        Number(item.score) > Number(item.maxScore) ||
        !validProgress.includes(item.progress) ||
        (item.note !== undefined && (typeof item.note !== 'string' || item.note.length > 500))) {
      throw makeError(400, 'Thông tin tiến độ hoặc điểm số không hợp lệ.');
    }
  }
  if (name === 'payments' && !isValidMonth(item.month)) {
    throw makeError(400, 'Tháng học phí không hợp lệ.');
  }
  if (name === 'payments' && typeof item.paid !== 'boolean') {
    throw makeError(400, 'Trạng thái học phí không hợp lệ.');
  }
}

async function saveDataChanges(user, currentData, changes) {
  if (!changes || typeof changes !== 'object' || Array.isArray(changes)) {
    throw makeError(400, 'Dữ liệu thay đổi không hợp lệ.');
  }
  const allowedNames = [...Object.keys(COLLECTIONS), 'permissions'];
  if (Object.keys(changes).some(name => !allowedNames.includes(name))) {
    throw makeError(400, 'Có trường dữ liệu không được hỗ trợ.');
  }

  const nextData = structuredClone(currentData);
  for (const [name, [module, keyOf]] of Object.entries(COLLECTIONS)) {
    const operations = changes[name] || [];
    if (!Array.isArray(operations) || operations.length > 5000) {
      throw makeError(400, `Danh sách thay đổi ${name} không hợp lệ.`);
    }
    const collection = [...nextData[name]];
    const seen = new Set();
    for (const operation of operations) {
      if (!operation || !['create', 'update', 'delete'].includes(operation.op)) {
        throw makeError(400, `Thao tác ${name} không hợp lệ.`);
      }
      const targetId = String(operation.id || '');
      if (!targetId || seen.has(targetId)) throw makeError(400, `Mã bản ghi ${name} không hợp lệ.`);
      seen.add(targetId);
      const managedModule = ['attendance', 'grades', 'finance', 'staffAttendance'].includes(module);
      const requiredAction = managedModule ? 'manage' : operation.op;
      if (!can(user, currentData, module, requiredAction)) {
        throw makeError(403, `Bạn không có quyền ${operation.op} dữ liệu ${module}.`);
      }
      const index = collection.findIndex(item => String(keyOf(item)) === targetId);
      if (operation.op === 'create') {
        validateCollectionItem(name, operation.item);
        if (String(keyOf(operation.item)) !== targetId || index !== -1) {
          throw makeError(409, `Bản ghi ${name} đã tồn tại hoặc mã không khớp.`);
        }
        if (name === 'students' && operation.item.classId &&
            !nextData.classes.some(item => item.id === operation.item.classId)) {
          throw makeError(400, 'Lớp học được chọn không tồn tại.');
        }
        if (['attendance', 'payments'].includes(name) &&
            !nextData.students.some(item => item.id === operation.item.studentId)) {
          throw makeError(400, 'Học sinh được chọn không tồn tại.');
        }
        if (name === 'assessments' && !nextData.students.some(item => item.id === operation.item.studentId)) {
          throw makeError(400, 'Học sinh được chọn không tồn tại.');
        }
        if (name === 'staffAttendance') {
          const staffAccount = nextData.accounts.find(item => item.id === operation.item.userId);
          if (!staffAccount || !['teacher', 'assistant'].includes(staffAccount.role)) {
            throw makeError(400, 'Tài khoản nhân sự được chọn không tồn tại.');
          }
          if (user.role !== 'admin' && operation.item.userId !== user.id) {
            throw makeError(403, 'Bạn chỉ được điểm danh cho chính mình.');
          }
          if (user.role !== 'admin' && operation.item.status === 'absent') {
            throw makeError(403, 'Chỉ quản trị viên được ghi nhận nhân sự vắng mặt.');
          }
          if (user.role !== 'admin' || operation.item.status === 'absent' || operation.item.status === 'leave') {
            operation.item.hours = 0;
          } else if (operation.item.hours === undefined) {
            operation.item.hours = 0;
          }
        }
        if (user.role !== 'admin' && ['students', 'classes'].includes(name) && !currentData.permissions[user.role]?.finance?.manage) {
          operation.item.tuition = 0;
        }
        collection.push(operation.item);
      } else {
        if (index === -1) throw makeError(409, `Bản ghi ${name} đã được thay đổi ở thiết bị khác.`);
        if (operation.op === 'delete') {
          if (name === 'staffAttendance' && user.role !== 'admin') {
            const record = collection[index];
            if (record.userId !== user.id) {
              throw makeError(403, 'Bạn chỉ được xóa điểm danh của chính mình.');
            }
            if (Number(record.hours || 0) > 0) {
              throw makeError(403, 'Chỉ quản trị viên được xóa điểm danh đã ghi nhận giờ công.');
            }
          }
          if (name === 'classes' && nextData.students.some(student => student.classId === targetId)) {
            throw makeError(409, 'Không thể xóa lớp đang có học sinh.');
          }
          if (name === 'students' && (nextData.attendance.some(item => item.studentId === targetId) ||
              nextData.payments.some(item => item.studentId === targetId) ||
              nextData.assessments.some(item => item.studentId === targetId))) {
            throw makeError(409, 'Không thể xóa học sinh đang có dữ liệu điểm danh, học phí hoặc đánh giá.');
          }
          collection.splice(index, 1);
        } else {
          validateCollectionItem(name, operation.item);
          if (String(keyOf(operation.item)) !== targetId) throw makeError(400, 'Mã bản ghi không khớp.');
          if (name === 'students' && operation.item.classId &&
              !nextData.classes.some(item => item.id === operation.item.classId)) {
            throw makeError(400, 'Lớp học được chọn không tồn tại.');
          }
          if (name === 'assessments' && !nextData.students.some(item => item.id === operation.item.studentId)) {
            throw makeError(400, 'Học sinh được chọn không tồn tại.');
          }
          if (['attendance', 'payments'].includes(name) &&
              !nextData.students.some(item => item.id === operation.item.studentId)) {
            throw makeError(400, 'Học sinh được chọn không tồn tại.');
          }
          if (name === 'staffAttendance') {
            const staffAccount = nextData.accounts.find(item => item.id === operation.item.userId);
            if (!staffAccount || !['teacher', 'assistant'].includes(staffAccount.role)) {
              throw makeError(400, 'Tài khoản nhân sự được chọn không tồn tại.');
            }
            if (user.role !== 'admin' && operation.item.userId !== user.id) {
              throw makeError(403, 'Bạn chỉ được điểm danh cho chính mình.');
            }
            if (user.role !== 'admin' && operation.item.status === 'absent') {
              throw makeError(403, 'Chỉ quản trị viên được ghi nhận nhân sự vắng mặt.');
            }
            if (user.role !== 'admin') {
              operation.item.hours = collection[index].hours || 0;
            }
            if (operation.item.status === 'absent' || operation.item.status === 'leave') {
              operation.item.hours = 0;
            } else if (operation.item.hours === undefined) {
              operation.item.hours = collection[index].hours || 0;
            }
          }
          if (module === 'students' && user.role !== 'admin' && !currentData.permissions[user.role]?.finance?.manage) {
            operation.item.tuition = collection[index].tuition;
          }
          if (module === 'classes' && user.role !== 'admin' && !currentData.permissions[user.role]?.finance?.manage) {
            operation.item.tuition = collection[index].tuition;
          }
          collection[index] = { ...collection[index], ...operation.item };
        }
      }
    }
    nextData[name] = collection;
  }

  if (Object.hasOwn(changes, 'permissions')) {
    if (user.role !== 'admin') throw makeError(403, 'Chỉ quản trị viên được thay đổi phân quyền.');
    nextData.permissions = normalizePermissions(changes.permissions);
  }
  await writeAppData(nextData);
  return nextData;
}

app.use(express.json({ limit: '2mb' }));
app.use(express.static(PUBLIC_DIR));

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, message: 'NTTCLASS backend is running.' });
});

app.post('/api/login', async (req, res) => {
  try {
    const { username, password } = req.body || {};
    if (typeof username !== 'string' || typeof password !== 'string' || !username.trim() || !password) {
      return res.status(400).json({ error: 'Tên đăng nhập và mật khẩu là bắt buộc.' });
    }
    const data = await readAppData();
    const account = data.accounts.find(item => item.username.toLowerCase() === username.trim().toLowerCase());
    if (!account) return res.status(401).json({ error: 'Tên đăng nhập hoặc mật khẩu không chính xác.' });
    const expectedHash = Buffer.from(account.passwordHash, 'hex');
    const actualHash = Buffer.from(hashPassword(password, account.salt), 'hex');
    if (expectedHash.length !== actualHash.length || !crypto.timingSafeEqual(expectedHash, actualHash)) {
      return res.status(401).json({ error: 'Tên đăng nhập hoặc mật khẩu không chính xác.' });
    }
    const token = crypto.randomBytes(32).toString('hex');
    sessions.set(token, { accountId: account.id, expiresAt: Date.now() + SESSION_TTL_MS });
    return res.json({ ok: true, token, account: sanitizeAccount(account) });
  } catch (error) {
    console.error('Login failed:', error.message);
    return res.status(500).json({ error: 'Không thể xử lý đăng nhập.' });
  }
});

app.post('/api/logout', authenticate, (req, res) => {
  sessions.delete(req.sessionToken);
  res.json({ ok: true });
});

app.get('/api/me', authenticate, (req, res) => {
  res.json({ account: req.user, permissions: req.appData.permissions[req.user.role] || {} });
});

app.get('/api/data', authenticate, (req, res) => {
  res.json(dataForUser(req.appData, req.user));
});

app.get('/api/attendance/export', authenticate, async (req, res) => {
  if (!can(req.user, req.appData, 'attendance', 'view')) {
    return res.status(403).json({ error: 'Bạn không có quyền xuất lịch sử điểm danh.' });
  }
  const classId = req.query.classId || 'all';
  if (typeof classId !== 'string') {
    return res.status(400).json({ error: 'Lớp cần xuất không hợp lệ.' });
  }
  if (classId !== 'all' && !req.appData.classes.some(item => item.id === classId)) {
    return res.status(400).json({ error: 'Lớp cần xuất không tồn tại.' });
  }
  try {
    const students = req.appData.students.filter(student => classId === 'all' || student.classId === classId);
    const studentsById = new Map(students.map(student => [student.id, student]));
    const attendance = req.appData.attendance
      .filter(record => studentsById.has(record.studentId))
      .sort((a, b) => {
        const dateOrder = b.date.localeCompare(a.date);
        if (dateOrder) return dateOrder;
        return (studentsById.get(a.studentId)?.name || '').localeCompare(
          studentsById.get(b.studentId)?.name || '',
          'vi'
        );
      });
    const classesById = new Map(req.appData.classes.map(item => [item.id, item.name]));
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'NTTCLASS';
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet('Lich su diem danh');
    worksheet.columns = [
      { header: 'STT', key: 'index', width: 10 },
      { header: 'Họ tên', key: 'student', width: 30 },
      { header: 'Lớp', key: 'className', width: 24 },
      { header: 'Ngày tháng', key: 'date', width: 18 },
      { header: 'Trạng thái', key: 'status', width: 16 }
    ];
    worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3159D9' } };
    worksheet.views = [{ state: 'frozen', ySplit: 1 }];
    worksheet.autoFilter = { from: 'A1', to: 'E1' };
    attendance.forEach((record, index) => {
      const student = studentsById.get(record.studentId);
      worksheet.addRow({
        index: index + 1,
        student: student.name,
        className: classesById.get(student.classId) || 'Chưa xếp lớp',
        date: new Date(`${record.date}T00:00:00.000Z`),
        status: record.present ? 'Có mặt' : 'Vắng'
      });
    });
    worksheet.getColumn('D').numFmt = 'dd/mm/yyyy';

    const file = await workbook.xlsx.writeBuffer();
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="nttclass-lich-su-diem-danh.xlsx"',
      'Cache-Control': 'no-store'
    });
    return res.send(Buffer.from(file));
  } catch (error) {
    console.error('Export attendance workbook failed:', error.message);
    return res.status(500).json({ error: 'Không thể xuất file Excel lịch sử điểm danh.' });
  }
});

function tuitionSummaryForMonth(data, month) {
  const attendanceCounts = new Map();
  for (const record of data.attendance) {
    if (record.present && typeof record.date === 'string' && record.date.startsWith(`${month}-`)) {
      attendanceCounts.set(record.studentId, (attendanceCounts.get(record.studentId) || 0) + 1);
    }
  }
  const paidStudents = new Set(data.payments
    .filter(item => item.month === month && item.paid)
    .map(item => item.studentId));
  const students = data.students.map(student => {
    const attendedSessions = attendanceCounts.get(student.id) || 0;
    const sessionTuition = Number(student.tuition || 0) / 8;
    const amountDue = attendedSessions * sessionTuition;
    return { studentId: student.id, attendedSessions, sessionTuition, amountDue, paid: paidStudents.has(student.id) };
  });
  const due = students.reduce((total, student) => total + student.amountDue, 0);
  const received = students.reduce((total, student) => total + (student.paid ? student.amountDue : 0), 0);
  return { students, summary: { due, received, outstanding: due - received } };
}

app.get('/api/finance', authenticate, (req, res) => {
  if (!can(req.user, req.appData, 'finance', 'view')) {
    return res.status(403).json({ error: 'Bạn không có quyền xem dữ liệu học phí.' });
  }
  const { month } = req.query;
  if (!isValidMonth(month)) {
    return res.status(400).json({ error: 'Tháng học phí không hợp lệ.' });
  }
  return res.json({ month, ...tuitionSummaryForMonth(req.appData, month) });
});

app.get('/api/finance/export', authenticate, async (req, res) => {
  if (!can(req.user, req.appData, 'finance', 'view')) {
    return res.status(403).json({ error: 'Bạn không có quyền xuất dữ liệu học phí.' });
  }
  const { month } = req.query;
  if (!isValidMonth(month)) {
    return res.status(400).json({ error: 'Tháng học phí không hợp lệ.' });
  }
  try {
    const tuition = tuitionSummaryForMonth(req.appData, month);
    const studentsById = new Map(req.appData.students.map(student => [student.id, student]));
    const classesById = new Map(req.appData.classes.map(item => [item.id, item.name]));
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'NTTCLASS';
    workbook.created = new Date();

    const details = workbook.addWorksheet('Hoc phi');
    details.columns = [
      { header: 'Học sinh', key: 'student', width: 30 },
      { header: 'Lớp', key: 'className', width: 24 },
      { header: 'Học phí / tháng (đ)', key: 'monthlyTuition', width: 22 },
      { header: 'Buổi có mặt', key: 'attendedSessions', width: 16 },
      { header: 'Học phí / buổi (đ)', key: 'sessionTuition', width: 20 },
      { header: 'Phải đóng (đ)', key: 'amountDue', width: 20 },
      { header: 'Trạng thái', key: 'status', width: 16 }
    ];
    details.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    details.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3159D9' } };
    details.views = [{ state: 'frozen', ySplit: 1 }];
    details.autoFilter = { from: 'A1', to: 'G1' };
    for (const item of tuition.students) {
      const student = studentsById.get(item.studentId);
      if (!student) throw new Error(`Student ${item.studentId} was not found while exporting tuition.`);
      details.addRow({
        student: student.name,
        className: classesById.get(student.classId) || 'Chưa xếp lớp',
        monthlyTuition: Number(student.tuition || 0),
        attendedSessions: item.attendedSessions,
        sessionTuition: item.sessionTuition,
        amountDue: item.amountDue,
        status: item.paid ? 'Đã đóng' : 'Chưa đóng'
      });
    }
    for (const column of ['C', 'E', 'F']) {
      details.getColumn(column).numFmt = '#,##0.###';
    }
    const summary = workbook.addWorksheet('Tổng hợp');
    summary.columns = [
      { header: 'Nội dung', key: 'label', width: 34 },
      { header: 'Giá trị', key: 'value', width: 22 }
    ];
    summary.addRows([
      { label: 'Tháng', value: month },
      { label: 'Tổng phải đóng theo buổi (đ)', value: tuition.summary.due },
      { label: 'Đã thu (đ)', value: tuition.summary.received },
      { label: 'Còn phải thu (đ)', value: tuition.summary.outstanding },
      { label: 'Công thức', value: 'Học phí tháng / 8 × số buổi có mặt' }
    ]);
    summary.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    summary.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3159D9' } };
    for (const row of [3, 4, 5]) summary.getCell(`B${row}`).numFmt = '#,##0.###';

    const file = await workbook.xlsx.writeBuffer();
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="nttclass-hoc-phi-${month}.xlsx"`,
      'Cache-Control': 'no-store'
    });
    return res.send(Buffer.from(file));
  } catch (error) {
    console.error('Export tuition workbook failed:', error.message);
    return res.status(500).json({ error: 'Không thể xuất file Excel học phí.' });
  }
});

app.put('/api/finance/monthly-tuition', authenticate, async (req, res) => {
  if (!can(req.user, req.appData, 'finance', 'manage')) {
    return res.status(403).json({ error: 'Bạn không có quyền thay đổi học phí.' });
  }
  const { studentId, tuition, allStudents } = req.body || {};
  if (!Number.isSafeInteger(tuition) || tuition < 0 || tuition > 1000000000 ||
      (allStudents === true ? studentId !== undefined : typeof studentId !== 'string' || !studentId)) {
    return res.status(400).json({ error: 'Học phí tháng phải là số nguyên từ 0 đến 1.000.000.000 đồng.' });
  }
  try {
    const updatedCount = await serializeDataWrite(async () => {
      const data = await readAppData();
      if (allStudents === true) {
        if (!data.students.length) throw makeError(404, 'Chưa có học sinh để cập nhật học phí.');
        for (const student of data.students) student.tuition = tuition;
        await writeAppData(data);
        return data.students.length;
      }
      const student = data.students.find(item => item.id === studentId);
      if (!student) throw makeError(404, 'Không tìm thấy học sinh.');
      student.tuition = tuition;
      await writeAppData(data);
      return 1;
    });
    return res.json({ ok: true, ...(allStudents === true ? { updatedCount } : { studentId }), tuition });
  } catch (error) {
    if (error.status) return res.status(error.status).json({ error: error.message });
    console.error('Update monthly tuition failed:', error.message);
    return res.status(500).json({ error: 'Không thể cập nhật học phí tháng.' });
  }
});

app.get('/api/reports', authenticate, (req, res) => {
  if (!can(req.user, req.appData, 'reports', 'view')) {
    return res.status(403).json({ error: 'Bạn không có quyền xem báo cáo.' });
  }
  const { month } = req.query;
  if (!isValidMonth(month)) {
    return res.status(400).json({ error: 'Tháng báo cáo không hợp lệ.' });
  }
  const financeVisible = can(req.user, req.appData, 'finance', 'view');
  const students = req.appData.students;
  const attendance = req.appData.attendance.filter(item => typeof item.date === 'string' && item.date.startsWith(`${month}-`));
  const tuitionSummary = tuitionSummaryForMonth(req.appData, month);
  const financeByStudent = new Map(tuitionSummary.students.map(item => [item.studentId, item]));
  const paid = req.appData.payments.filter(item => item.month === month && item.paid);
  const present = attendance.filter(item => item.present).length;
  const classes = req.appData.classes.map(cls => {
    const classStudents = students.filter(student => student.classId === cls.id);
    const classStudentIds = new Set(classStudents.map(student => student.id));
    const classAttendance = attendance.filter(item => classStudentIds.has(item.studentId));
    const classPresent = classAttendance.filter(item => item.present).length;
    const row = {
      name: cls.name,
      students: classStudents.length,
      sessions: classAttendance.length,
      present: classPresent,
      rate: classAttendance.length ? Math.round(classPresent / classAttendance.length * 100) : null
    };
    if (financeVisible) {
      row.amountDue = classStudents.reduce((sum, student) =>
        sum + (financeByStudent.get(student.id)?.amountDue || 0), 0);
      row.received = paid.reduce((sum, item) => sum + (classStudentIds.has(item.studentId)
        ? financeByStudent.get(item.studentId)?.amountDue || 0
        : 0), 0);
    }
    return row;
  });
  const summary = {
    present,
    attendanceCount: attendance.length,
    attendanceRate: attendance.length ? Math.round(present / attendance.length * 100) : null
  };
  if (financeVisible) {
    summary.received = tuitionSummary.summary.received;
    summary.due = tuitionSummary.summary.outstanding;
  }
  return res.json({ month, summary, classes, financeVisible });
});

app.get('/api/payroll', authenticate, async (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Chỉ quản trị viên được xem bảng lương.' });
  const { month } = req.query;
  if (!isValidMonth(month)) {
    return res.status(400).json({ error: 'Tháng tính lương không hợp lệ.' });
  }
  try {
    const data = await readAppData();
    const staff = data.accounts.filter(account => ['teacher', 'assistant'].includes(account.role));
    const entries = data.staffAttendance.filter(item => typeof item.date === 'string' && item.date.startsWith(`${month}-`));
    const employees = staff.map(account => {
      const records = entries.filter(item => item.userId === account.id);
      const hours = records.reduce((sum, item) => sum + Number(item.hours || 0), 0);
      const hourlyRate = Number(account.hourlyRate || 0);
      return {
        userId: account.id,
        displayName: account.displayName,
        username: account.username,
        role: account.role,
        hourlyRate,
        hours,
        grossPay: Math.round(hours * hourlyRate),
        daysAttended: records.filter(item => item.status === 'present' || item.status === 'late').length
      };
    });
    return res.json({
      month,
      employees,
      totalHours: employees.reduce((sum, item) => sum + item.hours, 0),
      totalPay: employees.reduce((sum, item) => sum + item.grossPay, 0)
    });
  } catch (error) {
    console.error('Read payroll failed:', error.message);
    return res.status(500).json({ error: 'Không thể tính bảng lương.' });
  }
});

app.put('/api/payroll/rates', authenticate, async (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Chỉ quản trị viên được thay đổi mức lương.' });
  const rates = req.body?.rates;
  if (!Array.isArray(rates) || rates.length > 500) {
    return res.status(400).json({ error: 'Danh sách mức lương không hợp lệ.' });
  }
  try {
    await serializeDataWrite(async () => {
      const data = await readAppData();
      const seen = new Set();
      for (const item of rates) {
        if (!item || typeof item.userId !== 'string' || seen.has(item.userId) ||
            !Number.isSafeInteger(item.hourlyRate) || item.hourlyRate < 0 || item.hourlyRate > 100000000) {
          throw makeError(400, 'Mức lương phải là số nguyên từ 0 đến 100.000.000 đồng/giờ.');
        }
        seen.add(item.userId);
        const account = data.accounts.find(entry => entry.id === item.userId && ['teacher', 'assistant'].includes(entry.role));
        if (!account) throw makeError(400, 'Tài khoản nhân sự được chọn không tồn tại.');
        account.hourlyRate = item.hourlyRate;
      }
      await writeAppData(data);
    });
    return res.json({ ok: true });
  } catch (error) {
    if (error.status) return res.status(error.status).json({ error: error.message });
    console.error('Save payroll rates failed:', error.message);
    return res.status(500).json({ error: 'Không thể lưu mức lương.' });
  }
});

app.put('/api/data', authenticate, async (req, res) => {
  try {
    const data = await serializeDataWrite(async () =>
      saveDataChanges(req.user, await readAppData(), req.body?.changes));
    res.json({ ok: true, data: dataForUser(data, req.user) });
  } catch (error) {
    if (error.status) return res.status(error.status).json({ error: error.message });
    console.error('Write data failed:', error.message);
    res.status(500).json({ error: 'Không thể lưu dữ liệu.' });
  }
});

app.post('/api/accounts', authenticate, async (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Chỉ quản trị viên được tạo tài khoản.' });
  if (!ACCOUNT_CREATION_CODE) {
    return res.status(503).json({ error: 'Chức năng tạo tài khoản chưa được cấu hình mã xác nhận.' });
  }
  try {
    const { displayName, username, password, role, confirmationCode } = req.body || {};
    if (!matchesAccountCreationCode(confirmationCode)) {
      return res.status(403).json({ error: 'Mã xác nhận tạo tài khoản không đúng.' });
    }
    if (typeof displayName !== 'string' || !displayName.trim() || displayName.length > 80 ||
        typeof username !== 'string' || !/^[a-zA-Z0-9_.-]{3,40}$/.test(username) ||
        typeof password !== 'string' || password.length < 8 ||
        !['admin', 'teacher', 'assistant'].includes(role)) {
      return res.status(400).json({ error: 'Thông tin tài khoản không hợp lệ.' });
    }
    const account = await serializeDataWrite(async () => {
      const data = await readAppData();
      if (data.accounts.some(item => item.username.toLowerCase() === username.toLowerCase())) {
        throw makeError(409, 'Tên đăng nhập đã tồn tại.');
      }
      const salt = makeSaltHex();
      const newAccount = {
        id: crypto.randomUUID(),
        displayName: displayName.trim(),
        username,
        role,
        hourlyRate: 0,
        salt,
        passwordHash: hashPassword(password, salt)
      };
      data.accounts.push(newAccount);
      await writeAppData(data);
      return newAccount;
    });
    res.status(201).json({ account: sanitizeAccount(account) });
  } catch (error) {
    if (error.status) return res.status(error.status).json({ error: error.message });
    console.error('Create account failed:', error.message);
    res.status(500).json({ error: 'Không thể tạo tài khoản.' });
  }
});

app.delete('/api/accounts/:id', authenticate, async (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Chỉ quản trị viên được xóa tài khoản.' });
  try {
    const target = await serializeDataWrite(async () => {
      const data = await readAppData();
      const account = data.accounts.find(entry => entry.id === req.params.id);
      if (!account) throw makeError(404, 'Không tìm thấy tài khoản.');
      if (account.id === req.user.id) throw makeError(400, 'Không thể xóa tài khoản đang đăng nhập.');
      if (account.role === 'admin' && data.accounts.filter(entry => entry.role === 'admin').length <= 1) {
        throw makeError(400, 'Không thể xóa quản trị viên cuối cùng.');
      }
      data.accounts = data.accounts.filter(entry => entry.id !== account.id);
      data.staffAttendance = data.staffAttendance.filter(entry => entry.userId !== account.id);
      await writeAppData(data);
      return account;
    });
    for (const [token, session] of sessions) {
      if (session.accountId === target.id) sessions.delete(token);
    }
    res.json({ ok: true });
  } catch (error) {
    if (error.status) return res.status(error.status).json({ error: error.message });
    console.error('Delete account failed:', error.message);
    res.status(500).json({ error: 'Không thể xóa tài khoản.' });
  }
});

app.use('/api', (_req, res) => {
  res.status(404).json({ error: 'Không tìm thấy API.' });
});

app.use((_req, res) => {
  res.sendFile(INDEX_PATH);
});
