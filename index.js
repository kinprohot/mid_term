require('dotenv').config();
const express = require('express');
const session = require('express-session');
const MongoStore = require('connect-mongo');
const path = require('path');

const { MONGODB_READWRITE_URI } = require('./config/database');
const bookRoutes = require('./routes/book.routes');

const app = express();
const PORT = process.env.PORT || 5000;

// Config Handlebars View Engine
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'hbs');

// Middleware for parsing form data and JSON
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// -----------------------------------------------------------------------------
// STATELESS SESSION STORE TẬP TRUNG TẠI CLOUD MONGODB ATLAS
// (Sử dụng luồng kết nối Write Connection: readwrite_23IT151)
// -----------------------------------------------------------------------------
app.use(session({
  secret: process.env.SESSION_SECRET || 'secret_cloud_23IT151',
  resave: false,
  saveUninitialized: false,
  store: MongoStore.create({
    mongoUrl: MONGODB_READWRITE_URI,
    collectionName: 'sessions',
    ttl: 24 * 60 * 60 // 1 ngày
  }),
  cookie: { maxAge: 24 * 60 * 60 * 1000 }
}));

// Đăng ký ứng dụng Routes
app.use('/', bookRoutes);

// Khởi chạy Máy chủ HTTP
app.listen(PORT, () => {
  console.log(`=================================================================`);
  console.log(`🚀 Cloud Web Server running at http://localhost:${PORT}`);
  console.log(`👤 Student: Nguyễn Hoàng Lực | MSSV: 23IT151`);
  console.log(`🔑 MSSV Prefix: 151 | Dynamic VAT: 5%`);
  console.log(`=================================================================`);
});
