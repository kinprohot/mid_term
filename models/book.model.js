const mongoose = require('mongoose');
const { readConnection, writeConnection } = require('../config/database');

// Dinh nghia Schema San Pham / Sach theo yeu cau bai lab
const bookSchema = new mongoose.Schema({
  bookId: {
    type: String,
    required: [true, 'Ma sach la bat buoc'],
    unique: true,
    trim: true
  },
  title: {
    type: String,
    required: [true, 'Ten sach la bat buoc'],
    trim: true
  },
  author: {
    type: String,
    required: [true, 'Tac gia la bat buoc'],
    trim: true
  },
  price: {
    type: Number,
    required: [true, 'Gia goc la bat buoc'],
    min: 0
  },
  imageUrl: {
    type: String,
    trim: true,
    default: ''
  },
  vatRate: {
    type: Number,
    default: 5 // (1 + 4)% VAT = 5% theo MSSV 23IT151
  },
  finalPrice: {
    type: Number,
    required: true
  },
  createdAt: {
    type: Date,
    default: Date.now
  }
});

// Rang buoc Models theo 2 luong ket noi phan quyen
// BookReadModel -> Gan voi Read Connection (read_23IT151)
const BookReadModel = readConnection.model('Book', bookSchema);

// BookWriteModel -> Gan voi Write Connection (readwrite_23IT151)
const BookWriteModel = writeConnection.model('Book', bookSchema);

module.exports = {
  BookReadModel,
  BookWriteModel
};
