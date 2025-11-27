const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const fs = require('fs');
const multer = require('multer');

const PDFDocument = require('pdfkit');
const ExcelJS = require('exceljs');
const app = express();
const db = new sqlite3.Database('./db2.sqlite');



// =====================
// MIDDLEWARE
// =====================
app.use(express.urlencoded({ extended: true }));

app.use(express.static('public'));
//app.use('/uploads', express.static('uploads'));

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// =====================
// MULTER CONFIG
// =====================
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, 'uploads/'),
  filename: (req, file, cb) => {
    const unique = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, unique + path.extname(file.originalname));
  }
});

app.use(express.json());

// =====================
// DATABASE INIT
// =====================
db.serialize(() => {
  db.run(`CREATE TABLE IF NOT EXISTS tickets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
     
    area TEXT NOT NULL,
    equipment TEXT NOT NULL,
    production_manager TEXT NOT NULL,
    equipment_staff TEXT NOT NULL,
    maintenance_staff TEXT NOT NULL,
    work_type TEXT NOT NULL,
    equipment_status TEXT NOT NULL,
    failure_reason TEXT NOT NULL,
    stop_time TEXT NOT NULL,
    lsx_number TEXT,
    status TEXT DEFAULT 'open',
    
    completion_time TEXT,
    total_processing_time TEXT,
    cause_recognition TEXT,
    solution TEXT,
    materials_used TEXT,
    equipment_status_after TEXT,
    recommendations TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    photo TEXT
  
  )`);
});


require('./middleware/formParser')(app);

const upload = require('./middleware/upload');
// Static
app.use(express.static('public'));
app.use('/uploads', express.static('uploads'));
// =====================
// ROUTES
// =====================

// Home
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Send ticket form
app.get('/send', (req, res) => {
  res.render('send');
});

// SUBMIT ticket — FIXED ✔✔✔
// NOW SUPPORTS TEXT + PHOTO UPLOAD
app.post('/send', upload.single('anh'), (req, res) => {
  console.log(req.body)
  console.log(req.file)

  
    const {
        area, equipment, production_manager, equipment_staff,
        maintenance_staff, work_type, equipment_status,
        failure_reason, stop_time, lsx_number
    } = req.body;
  const photo =  req.file.filename

 const sql = `
  INSERT INTO tickets (
    area, equipment, production_manager, equipment_staff,
    maintenance_staff, work_type, equipment_status, failure_reason,
    stop_time, lsx_number, photo
  ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;

db.run(sql, [
  area, equipment, production_manager, equipment_staff,
  maintenance_staff, work_type, equipment_status, failure_reason,
  stop_time, lsx_number, photo
], function(err) {
  if (err) {
    console.log(err);
    return res.status(500).send("Lỗi khi cập nhật");
  }

  res.send(`<script>alert('Ticket đã tạo! ID: ${this.lastID}'); window.location.href='/';</script>`);
});

});

// Dashboard
app.get('/dashboard', (req, res) => {
  db.all("SELECT * FROM tickets ORDER BY created_at DESC", (err, tickets) => {
    if (err) return res.status(500).send('Lỗi dashboard');
    res.render('dashboard', { tickets });
  });
});

// Fix ticket page
app.get('/fix', (req, res) => {
  db.all("SELECT * FROM tickets WHERE status != 'finished' ORDER BY created_at DESC", (err, tickets) => {
    if (err) return res.status(500).send('Lỗi lấy ticket');
    res.render('fix', { tickets });
  });
});
app.post('/update-status/:id', (req, res) => {
  const { id } = req.params;
  const { status } = req.body;

  const sql = `UPDATE tickets SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?`;

  db.run(sql, [status, id], function(err) {
    if (err) {
      console.error("SQLite Error:", err.message);
      return res.status(500).json({ success: false, message: "Lỗi khi cập nhật trạng thái" });
    }

    res.json({ success: true, message: "Trạng thái đã được cập nhật" });
  });
});

// Update ticket info
app.post('/fix/:id', (req, res) => {
  const { id } = req.params;
  const {
    completion_time, total_processing_time, cause_recognition,
    solution, materials_used, equipment_status_after, recommendations
  } = req.body;

  const sql = `UPDATE tickets SET
    completion_time=?, total_processing_time=?,
    cause_recognition=?, solution=?, materials_used=?,
    equipment_status_after=?, recommendations=?,
    status='finished', updated_at=CURRENT_TIMESTAMP
  WHERE id=?`;

  db.run(sql, [
    completion_time, total_processing_time, cause_recognition,
    solution, materials_used, equipment_status_after, recommendations, id
  ], (err) => {
    if (err) return res.status(500).send('Lỗi update ticket');
    res.send(`<script>alert("Cập nhật thành công!"); window.location.href="/fix";</script>`);
  });
});

app.get('/print/:id', (req, res) => {
  const id = req.params.id;

  db.get("SELECT * FROM tickets WHERE id = ?", [id], (err, ticket) => {
    if (err) return res.status(500).send('Database error');
    if (!ticket) return res.status(404).send('Ticket not found');

    // Create PDF document
    const doc = new PDFDocument({ margin: 50 });

    // Set response headers
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename=ticket_${id}.pdf`);


doc.registerFont('Roboto', 'fonts/Roboto-Regular.ttf');
doc.registerFont('Roboto-Bold', 'fonts/Roboto-Bold.ttf');
    // Pipe PDF to response
    doc.pipe(res);

    // Add title
    doc.fontSize(14).font('Roboto-Bold').text('CÔNG TY TNHH MTV CADIVI ĐỒNG NAI    CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { align: 'left' });
    doc.moveDown();
   doc.font('Roboto').fontSize(10).text(`Ticket ID: #${ticket.id}`, { align: 'left' });
    doc.moveDown();

    // Add ticket details
    doc.font('Roboto-Bold').fontSize(20).text('BIÊN BẢN DỪNG THIẾT BỊ',{ align: 'center' });
    doc.moveDown(0.5);
    
    doc.fontSize(10).font('Helvetica');
    doc.font('Roboto').text(`Chúng tôi gồm:`);
    doc.font('Roboto').text(`Ông(bà): ${ticket.production_manager || 'N/A'}     Quản lý sản xuất` );
    doc.font('Roboto').text(`Ông(bà): ${ticket.equipment_staff || 'N/A'}        Nhân viên thiết bị`);
    doc.font('Roboto').text(`Ông(bà): ${ticket.maintenance_staff || 'N/A'}      Nhân viên bảo trì`);
    doc.font('Roboto').text(`Cùng nhau tiến hành dừng thiết bị (tên máy): : ${ticket.equipment || 'N/A'}để thực hiện việc: ${ticket.work_type || 'N/A'}`);
    doc.font('Roboto').text(`Mô tả tình trạng thiết bị / nguyên nhân:`);
    doc.font('Roboto').text(`1.	Tình trạng thiết bị (dành cho nhân viên vận hành):`);




    


    doc.moveDown();

    doc.fontSize(14).font('Roboto-Bold').text('CHI TIẾT SỰ CỐ');
    doc.moveDown(0.5);
    
    doc.fontSize(10).font('Helvetica');
    doc.font('Roboto').text(`Loại công việc: ${ticket.work_type || 'N/A'}`);
    doc.font('Roboto').text(`Trạng thái thiết bị: ${ticket.equipment_status || 'N/A'}`);
    doc.font('Roboto').text(`Nguyên nhân hỏng: ${ticket.failure_reason || 'N/A'}`);
    doc.font('Roboto').text(`Thời gian dừng: ${ticket.stop_time || 'N/A'}`);
    doc.font('Roboto').text(`Số LSX: ${ticket.lsx_number || 'N/A'}`);
    doc.font('Roboto').text(`Trạng thái: ${ticket.status || 'N/A'}`);
    doc.moveDown();

    if (ticket.status === 'finished') {
      doc.fontSize(14).font('Roboto-Bold').text('THÔNG TIN XỬ LÝ');
      doc.moveDown(0.5);
      
      doc.fontSize(10).font('Helvetica');
      doc.font('Roboto').text(`Thời gian hoàn thành: ${ticket.completion_time || 'N/A'}`);
      doc.font('Roboto').text(`Tổng thời gian xử lý: ${ticket.total_processing_time || 'N/A'}`);
      doc.font('Roboto').text(`Nhận định nguyên nhân: ${ticket.cause_recognition || 'N/A'}`);
      doc.font('Roboto').text(`Giải pháp: ${ticket.solution || 'N/A'}`);
      doc.font('Roboto').text(`Vật tư sử dụng: ${ticket.materials_used || 'N/A'}`);
      doc.font('Roboto').text(`Trạng thái thiết bị sau: ${ticket.equipment_status_after || 'N/A'}`);
      doc.font('Roboto').text(`Kiến nghị: ${ticket.recommendations || 'N/A'}`);
      doc.moveDown();
    }

    doc.fontSize(14).font('Roboto-Bold').text('THÔNG TIN HỆ THỐNG');
    doc.moveDown(0.5);
    
    doc.fontSize(10).font('Helvetica');
    doc.font('Roboto').text(`Ngày tạo: ${ticket.created_at || 'N/A'}`);
    doc.font('Roboto').text(`Cập nhật lần cuối: ${ticket.updated_at || 'N/A'}`);

    // Add photo if exists
    if (ticket.photo) {
      const photoPath = path.join(__dirname, 'uploads', ticket.photo);
      if (fs.existsSync(photoPath)) {
        doc.moveDown();
        doc.fontSize(14).font('Roboto-Bold').text('HÌNH ẢNH');
        doc.moveDown(0.5);
        try {
          doc.image(photoPath, { fit: [400, 300], align: 'center' });
        } catch (e) {
          doc.fontSize(10).font('Roboto').text('(Không thể load ảnh)');
        }
      }
    }

    // Finalize PDF
    doc.end();
  });
});



// =====================
// START SERVER
// =====================
const PORT = process.env.PORT || 80;
app.listen(PORT, () => {
  console.log(`Server chạy tại http://localhost:${PORT}`);
});
