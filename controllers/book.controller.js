const { BookReadModel, BookWriteModel } = require('../models/book.model');

// HANG SO CA NHAN HOA THEO MSSV: 23IT151
const STUDENT_NAME = 'Nguyễn Hoàng Lực';
const STUDENT_ID = '23IT151';
const MSSV_PREFIX = '151'; // 3 so cuoi MSSV
const LAST_DIGIT = 1;     // Chu so cuoi MSSV 23IT151
const VAT_PERCENT = LAST_DIGIT + 4; // (1 + 4) = 5% theo yeu cau de thi
const VAT_RATE = VAT_PERCENT / 100; // 0.05

// [GET] / - Hien thi trang chu va danh sach sach (Routing den Read Connection)
exports.getHomePage = async (req, res) => {
  try {
    // Tang va luu luot xem trong Stateless Session (MongoDB Atlas Store)
    req.session.viewsCount = (req.session.viewsCount || 0) + 1;

    // DIEU HUONG READ: Su dung BookReadModel (Ket noi boi user read_23IT151)
    const books = await BookReadModel.find().sort({ createdAt: -1 }).lean();

    res.render('index', {
      studentName: STUDENT_NAME,
      studentId: STUDENT_ID,
      mssvPrefix: MSSV_PREFIX,
      vatPercent: VAT_PERCENT,
      books: books,
      viewsCount: req.session.viewsCount,
      sessionId: req.sessionID,
      error: req.query.error,
      success: req.query.success
    });
  } catch (err) {
    console.error('Loi khi truy van danh sach sach tu Read Connection:', err);
    res.status(500).send('Loi truy van du lieu tu Cloud MongoDB Atlas (Read Connection).');
  }
};

// [POST] /add-book - Xu ly them sach moi (Routing den Write Connection)
exports.createBook = async (req, res) => {
  try {
    const { bookId, title, author, price, imageUrl } = req.body;
    const cleanBookId = (bookId || '').trim();

    // 1. BO LOC DU LIEU: Bat buoc ma san pham co tien to la 3 so cuoi MSSV (151)
    if (!cleanBookId.startsWith(MSSV_PREFIX)) {
      const errorMsg = `Ma san pham khong hop le! Ma phai bat dau bang 3 so cuoi MSSV cua ban (${MSSV_PREFIX}). Vi du: 151-BOOK01.`;
      return res.redirect(`/?error=${encodeURIComponent(errorMsg)}`);
    }

    // 2. THUAT TOAN THUE VAT DONG: VAT = (Chu so cuoi MSSV + 4)% = (1 + 4)% = 5%
    const numPrice = parseFloat(price);
    if (isNaN(numPrice) || numPrice < 0) {
      return res.redirect(`/?error=${encodeURIComponent('Gia goc san pham phai la mot so duong hop le!')}`);
    }

    const finalPrice = Math.round(numPrice * (1 + VAT_RATE) * 100) / 100;

    // DIEU HUONG WRITE: Su dung BookWriteModel (Ket noi boi user readwrite_23IT151)
    await BookWriteModel.create({
      bookId: cleanBookId,
      title: title.trim(),
      author: author.trim(),
      price: numPrice,
      imageUrl: (imageUrl || '').trim(),
      vatRate: VAT_PERCENT,
      finalPrice: finalPrice
    });

    const successMsg = `Them sach "${title}" thanh cong! Gia sau thue VAT ${VAT_PERCENT}%: $${finalPrice}.`;
    res.redirect(`/?success=${encodeURIComponent(successMsg)}`);
  } catch (err) {
    console.error('Loi khi ghi du lieu len Write Connection:', err);
    let errorMsg = 'Loi he thong khi ghi du lieu len Cloud MongoDB Atlas.';
    if (err.code === 11000) {
      errorMsg = `Ma sach "${req.body.bookId}" da ton tai trong CSDL Cloud. Vui long nhap ma sach khac!`;
    }
    res.redirect(`/?error=${encodeURIComponent(errorMsg)}`);
  }
};
