const express = require('express');
const router = express.Router();
const bookController = require('../controllers/book.controller');

// Route GET: Trang chủ & Hiển thị danh sách sách (Luồng Read)
router.get('/', bookController.getHomePage);

// Route POST: Thêm mới sách (Luồng Write)
router.post('/add-book', bookController.createBook);

module.exports = router;
