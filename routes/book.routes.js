const express = require('express');
const router = express.Router();
const bookController = require('../controllers/book.controller');

// Route GET: Trang chủ & Hiển thị danh sách sách (Luồng Read)
router.get('/', bookController.getHomePage);

// Route POST: Thêm mới sách (Luồng Write: readwrite_23IT151)
router.post('/add-book', bookController.createBook);

// Route POST: Cập nhật sách (Luồng Write: readwrite_23IT151)
router.post('/update-book/:id', bookController.updateBook);

// Route GET: Xóa sách (Luồng Write: readwrite_23IT151)
router.get('/delete-book/:id', bookController.deleteBook);

module.exports = router;
