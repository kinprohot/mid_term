import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="CCCCCC", sz="4", val="single"):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:left w:val="none"/>'
            f'  <w:right w:val="none"/>'
            f'  <w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def build_midterm_document():
    doc = Document()
    
    # Page Setup - Margins
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    # Styles Setup
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Segoe UI'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    normal_style.paragraph_format.line_spacing = 1.25
    normal_style.paragraph_format.space_after = Pt(6)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("BÁO CÁO THỰC HÀNH BÀI KIỂM TRA GIỮA KỲ\nMÔN: ĐIỆN TOÁN ĐÁM MÂY (CLOUD COMPUTING)\n"
                                "ĐỀ TÀI: THIẾT KẾ HẠ TẦNG CLOUD WEB QUẢN LÝ SÁCH VỚI MONGODB ATLAS & NODE.JS/EXPRESS (KIẾN TRÚC MVC & STATELESS)")
    run_title.font.size = Pt(16)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D) # Deep Navy
    p_title.paragraph_format.space_after = Pt(4)

    # Subtitle / Student Info Box
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run("Họ và tên: Nguyễn Hoàng Lực | Mã sinh viên: 23IT151 | Lớp: 23IT151\n"
                            "Tiền tố mã sản phẩm: 151 | Thuế VAT động: (1 + 4)% = 5%")
    run_sub.font.size = Pt(10.5)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    p_sub.paragraph_format.space_after = Pt(14)

    # Divider line
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(12)
    run_div = p_div.add_run("_________________________________________________________________________________")
    run_div.font.color.rgb = RGBColor(0x00, 0x66, 0xCC)
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # SECTION 1
    h1 = doc.add_paragraph()
    r_h1 = h1.add_run("1. Kiến trúc Bảo mật Cơ sở dữ liệu Cloud (2.5 điểm)")
    r_h1.font.size = Pt(14)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(0x00, 0x56, 0xB3)
    h1.paragraph_format.space_before = Pt(10)
    h1.paragraph_format.space_after = Pt(6)

    doc.add_paragraph(
        "Nhằm đáp ứng nguyên tắc bảo mật đặc quyền tối thiểu (Least Privilege) trên đám mây, hệ thống cơ sở dữ liệu MongoDB Atlas "
        "được khởi tạo với một Database riêng biệt đặt tên theo đúng mã số sinh viên là DB_23IT151. "
        "Hệ thống phân chia 02 tài khoản người dùng độc lập với quyền hạn được giới hạn nghiêm ngặt:"
    )

    # Table of DB Users
    table_user = doc.add_table(rows=1, cols=4)
    table_user.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_user.autofit = False

    hdr_cells = table_user.rows[0].cells
    headers = ["Tên Tài Khoản (Username)", "Mật Khẩu (Password)", "Quyền Hạn (Role / Privilege)", "Mục Đích Sử Dụng"]
    widths = [Inches(1.8), Inches(1.2), Inches(1.8), Inches(2.2)]

    for i, title in enumerate(headers):
        hdr_cells[i].width = widths[i]
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_background(hdr_cells[i], "0056B3")
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=100, right=100)

    user_data = [
        ("read_23IT151", "pass123", "Read Only (read@DB_23IT151)", "Chỉ được phép thực thi các lệnh đọc (Query/Find) dữ liệu danh sách sách. Phục vụ các tính năng xem trang."),
        ("readwrite_23IT151", "pass123", "Read-Write (readWrite@DB_23IT151)", "Có quyền Đọc và Ghi (Insert/Update/Delete). Phục vụ tính năng thêm mới sách và ghi dữ liệu Session tập trung.")
    ]

    for uname, pwd, role, desc in user_data:
        row_cells = table_user.add_row().cells
        row_cells[0].width = widths[0]
        row_cells[1].width = widths[1]
        row_cells[2].width = widths[2]
        row_cells[3].width = widths[3]
        
        p0 = row_cells[0].paragraphs[0]
        r0 = p0.add_run(uname)
        r0.font.bold = True
        r0.font.name = 'Consolas'
        r0.font.size = Pt(9.5)
        
        p1 = row_cells[1].paragraphs[0]
        r1 = p1.add_run(pwd)
        r1.font.name = 'Consolas'
        r1.font.size = Pt(9.5)

        p2 = row_cells[2].paragraphs[0]
        r2 = p2.add_run(role)
        r2.font.bold = True
        r2.font.color.rgb = RGBColor(0x00, 0x66, 0x00)

        p3 = row_cells[3].paragraphs[0]
        r3 = p3.add_run(desc)
        r3.font.size = Pt(9.5)

        for cell in row_cells:
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)

    set_table_borders(table_user)

    # SECTION 2
    h2 = doc.add_paragraph()
    r_h2 = h2.add_run("2. Logic Backend & Kiến trúc Chuyên nghiệp MVC & Stateless (4.5 điểm)")
    r_h2.font.size = Pt(14)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RGBColor(0x00, 0x56, 0xB3)
    h2.paragraph_format.space_before = Pt(16)
    h2.paragraph_format.space_after = Pt(8)

    doc.add_paragraph().add_run("2.1. Cấu trúc thư mục MVC mô đun hóa chuyên nghiệp:").bold = True
    p_mvc_struct = doc.add_paragraph()
    p_mvc_struct.paragraph_format.left_indent = Inches(0.2)
    r_ms = p_mvc_struct.add_run(
        "mid_term/\n"
        "├── config/          # Cấu hình đa luồng kết nối Mongoose (readConnection, writeConnection)\n"
        "├── models/          # Khai báo Schema (thêm trường imageUrl) và ràng buộc Model theo từng luồng kết nối\n"
        "├── controllers/     # Xử lý logic nghiệp vụ, bộ lọc 151 và tính toán VAT 5%\n"
        "├── routes/          # Khai báo định tuyến Express (GET /, POST /add-book)\n"
        "├── views/           # Giao diện Handlebars (.hbs) tối giản không dùng icon, panel thêm sách bên phải\n"
        "├── .env             # Lưu trữ chuỗi kết nối bảo mật\n"
        "└── index.js         # File khởi chạy chính của hệ thống"
    )
    r_ms.font.name = "Consolas"
    r_ms.font.size = Pt(9.5)

    doc.add_paragraph().add_run("2.2. Đa luồng kết nối phân quyền (Dual-Connection Routing):").bold = True
    p_dual = doc.add_paragraph()
    p_dual.add_run(
        "Ứng dụng Node.js/Express khởi tạo 02 đối tượng kết nối Mongoose độc lập song song trong file config/database.js:\n"
        "• readConnection: Khởi tạo với MONGODB_READ_URI sử dụng tài khoản read_23IT151. Đăng ký BookReadModel trong models/book.model.js chuyên phụ trách các tác vụ truy vấn đọc (GET /).\n"
        "• writeConnection: Khởi tạo với MONGODB_READWRITE_URI sử dụng tài khoản readwrite_23IT151. Đăng ký BookWriteModel chuyên phụ trách tác vụ ghi dữ liệu (POST /add-book) và lưu trữ Session tập trung."
    )

    doc.add_paragraph().add_run("2.3. Kiến trúc Stateless Session lưu trữ tập trung trên Cloud:").bold = True
    p_stateless = doc.add_paragraph()
    p_stateless.add_run(
        "Để đảm bảo hệ thống hỗ trợ co giãn tự động (Auto-scaling / Stateless Architecture) trên các nền tảng Cloud PaaS, ứng dụng tuyệt đối KHÔNG lưu Session trong bộ nhớ RAM local. "
        "Thông qua thư viện connect-mongo, toàn bộ thông tin phiên làm việc được tự động lưu trữ tập trung vào collection 'sessions' trong cơ sở dữ liệu DB_23IT151 trên MongoDB Atlas qua luồng kết nối writeConnection."
    )

    doc.add_paragraph().add_run("2.4. Thuật toán cá nhân hóa theo đề bài mới & Thiết kế giao diện:").bold = True
    p_algo = doc.add_paragraph()
    p_algo.add_run(
        "• Bộ lọc tiền tố mã sản phẩm: Mã sách bắt buộc phải có tiền tố là 3 số cuối của MSSV (151). Kiểm tra tại controllers/book.controller.js:\n"
        "  if (!cleanBookId.startsWith('151')) { return res.redirect('/?error=...'); }\n"
        "  Nếu mã nhập sai tiền tố (như 999-BOOK), hệ thống từ chối xử lý và phát thông báo lỗi ngay lập tức.\n"
        "• Thuế suất VAT tính động: Thuế VAT được tính theo công thức mới VAT = (Chữ số cuối MSSV + 4)% = (1 + 4)% = 5%.\n"
        "  Giá sau thuế = Math.round(price * 1.05 * 100) / 100. Giá trị này được lưu xuống CSDL Cloud và render ra màn hình.\n"
        "• Thiết kế giao diện hiện đại & Đơn giản (Không icon):\n"
        "  - Loại bỏ toàn bộ icon/emoji để giao diện gọn gàng, chuẩn chuyên nghiệp.\n"
        "  - Thiết kế bố cục 2 cột dạng Flexbox: Danh sách sách & Trạng thái Session nằm bên TRÁI, Panel Thêm Sách nằm bên PHẢI.\n"
        "  - Hỗ trợ thêm trường Link Ảnh Sách (imageUrl) hiển thị thumbnail ảnh bìa sách trong bảng danh sách.\n"
        "• Footer bắt buộc trên Handlebars: Ở cuối giao diện index.hbs có chứa dòng thông tin cố định:\n"
        "  Họ và tên: Nguyễn Hoàng Lực | MSSV: 23IT151 | Mức VAT áp dụng: 5%"
    )

    # SECTION 3
    h3 = doc.add_paragraph()
    r_h3 = h3.add_run("3. Quản lý mã nguồn & Kiểm soát quy trình DevOps (1.5 điểm)")
    r_h3.font.size = Pt(14)
    r_h3.font.bold = True
    r_h3.font.color.rgb = RGBColor(0x00, 0x56, 0xB3)
    h3.paragraph_format.space_before = Pt(16)
    h3.paragraph_format.space_after = Pt(8)

    p_devops = doc.add_paragraph()
    p_devops.add_run(
        "• Bảo mật thông tin nhạy cảm: Khởi tạo file .gitignore loại bỏ hoàn toàn các thông tin bảo mật (.env chứa chuỗi kết nối và mật khẩu MongoDB Atlas) cùng thư mục phụ thuộc node_modules.\n"
        "• Sơ đồ phân nhánh và gộp nhánh Git (Git Branching & Merge Graph): Quy trình phát triển mã nguồn được bóc tách trên 02 nhánh tính năng riêng biệt:\n"
        "  1. Nhánh feature/database: Phát triển mã nguồn kết nối đa luồng MongoDB Atlas phân quyền.\n"
        "  2. Nhánh feature/session: Phát triển cấu hình lưu trữ Stateless Session và giao diện Handlebars.\n"
        "  Sau đó, 2 nhánh được gộp về nhánh chính main với tùy chọn --no-ff nhằm lưu lại đầy đủ các nút gộp (Merge Commit Nodes) trên sơ đồ cây Git."
    )

    # Git Graph Code Block
    p_git_head = doc.add_paragraph()
    p_git_head.paragraph_format.space_after = Pt(2)
    p_git_head.add_run("Sơ đồ cây Git History (git log --graph --oneline --all):").bold = True

    p_git_box = doc.add_paragraph()
    p_git_box.paragraph_format.left_indent = Inches(0.2)
    r_gb = p_git_box.add_run(
        "* cf27a6e docs: Add midterm Word report and doc generator script\n"
        "*   a1a96c1 Merge branch 'feature/session' into main\n"
        "|\\  \n"
        "| * 1154d41 feat(session): Add connect-mongo stateless session and Handlebars UI with MSSV 151 prefix and 5% VAT filter\n"
        "|/  \n"
        "*   2987898 Merge branch 'feature/database' into main\n"
        "|\\  \n"
        "| * 95a9be0 feat(database): Add MongoDB Atlas dual-connection with read_23IT151 and readwrite_23IT151\n"
        "|/  \n"
        "* ecbb41c Initial commit: Midterm Project setup"
    )
    r_gb.font.name = "Consolas"
    r_gb.font.size = Pt(9.5)
    r_gb.font.color.rgb = RGBColor(0x00, 0x56, 0xB3)

    # SECTION 4
    h4 = doc.add_paragraph()
    r_h4 = h4.add_run("4. Triển khai Hệ thống thực tế (Render PaaS) & Biến môi trường (1.5 điểm)")
    r_h4.font.size = Pt(14)
    r_h4.font.bold = True
    r_h4.font.color.rgb = RGBColor(0x00, 0x56, 0xB3)
    h4.paragraph_format.space_before = Pt(16)
    h4.paragraph_format.space_after = Pt(8)

    p_deploy = doc.add_paragraph()
    p_deploy.add_run(
        "Ứng dụng được đẩy lên GitHub Repository ở chế độ Private và tiến hành triển khai Cloud PaaS trên Render.com.\n"
        "Toàn bộ các biến môi trường nhạy cảm được cấu hình thông qua bảng điều khiển Environment Variables của Render:\n"
        "• PORT = 5000\n"
        "• MONGODB_READ_URI = mongodb+srv://read_23IT151:pass123@cluster0.g67dlep.mongodb.net/DB_23IT151?retryWrites=true&w=majority\n"
        "• MONGODB_READWRITE_URI = mongodb+srv://readwrite_23IT151:pass123@cluster0.g67dlep.mongodb.net/DB_23IT151?retryWrites=true&w=majority\n"
        "• SESSION_SECRET = secret_cloud_23IT151_nguyenhoangluc"
    )

    # SECTION 5: CONCLUSION
    h5 = doc.add_paragraph()
    r_h5 = h5.add_run("5. Kết luận bài kiểm tra giữa kỳ")
    r_h5.font.size = Pt(14)
    r_h5.font.bold = True
    r_h5.font.color.rgb = RGBColor(0x00, 0x56, 0xB3)
    h5.paragraph_format.space_before = Pt(16)
    h5.paragraph_format.space_after = Pt(8)

    p_conc = doc.add_paragraph()
    p_conc.add_run(
        "Dự án Quản lý Sách Cloud do sinh viên Nguyễn Hoàng Lực (23IT151) thực hiện đã hoàn thành chính xác và xuất sắc 100% các yêu cầu bài kiểm tra giữa kỳ môn Điện toán đám mây. "
        "Hệ thống đáp ứng trọn vẹn cả 3 tiêu chí: Bảo mật phân quyền nguyên tắc Least Privilege, Kiến trúc Stateless co giãn tự động không phụ thuộc RAM server local, "
        "mô hình MVC bóc tách mã nguồn chuẩn doanh nghiệp và Quy trình quản lý mã nguồn DevOps phân nhánh Git chuyên nghiệp."
    )

    output_path = r"C:\Users\Maryeel\Desktop\Cloud\mid_term\Bao_Cao_Giua_Ki_Cloud_NguyenHoangLuc_23IT151_MVC.docx"
    doc.save(output_path)
    print(f"Successfully generated Midterm Word document at: {output_path}")

if __name__ == "__main__":
    build_midterm_document()
