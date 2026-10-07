const mongoose = require('mongoose');

// Lay chuoi ket noi tu bien moi truong Cloud
const MONGODB_READ_URI = process.env.MONGODB_READ_URI;
const MONGODB_READWRITE_URI = process.env.MONGODB_READWRITE_URI;

if (!MONGODB_READ_URI || !MONGODB_READWRITE_URI) {
  console.error('🔴 LOI: Thieu bien moi truong MONGODB_READ_URI hoac MONGODB_READWRITE_URI trong file .env');
  process.exit(1);
}

// 1. Luong Ket Noi Read-Only (Tai khoan: read_23IT151)
const readConnection = mongoose.createConnection(MONGODB_READ_URI, {
  maxPoolSize: 10
});

readConnection.on('connected', () => {
  console.log('🟢 [MongoDB Cloud] Connection 1 (Read-Only) initialized successfully with user: read_23IT151');
});

readConnection.on('error', (err) => {
  console.error('🔴 [MongoDB Cloud] Connection 1 (Read-Only) Error:', err.message);
});

// 2. Luong Ket Noi Read-Write (Tai khoan: readwrite_23IT151)
const writeConnection = mongoose.createConnection(MONGODB_READWRITE_URI, {
  maxPoolSize: 10
});

writeConnection.on('connected', () => {
  console.log('🟢 [MongoDB Cloud] Connection 2 (Read-Write) initialized successfully with user: readwrite_23IT151');
});

writeConnection.on('error', (err) => {
  console.error('🔴 [MongoDB Cloud] Connection 2 (Read-Write) Error:', err.message);
});

module.exports = {
  readConnection,
  writeConnection,
  MONGODB_READWRITE_URI
};
