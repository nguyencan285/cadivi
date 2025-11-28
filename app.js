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
const db = new sqlite3.Database('./database.sqlite');



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
    photo_after TEXT
  
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

   const ticketId = this.lastID;
  const currentDate = new Date();
  const date = currentDate.toLocaleDateString('vi-VN');
  const time = currentDate.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' });


  // Create a nicely formatted HTML popup
  const popupHtml = `
    <style>
      .ticket-popup {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        max-width: 400px;
        margin: 20px auto;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 15px;
        padding: 25px;
        box-shadow: 0 10px 40px rgba(0,0,0,0.3);
        color: white;
      }
      .ticket-header {
        text-align: center;
        border-bottom: 2px solid rgba(255,255,255,0.3);
        padding-bottom: 15px;
        margin-bottom: 20px;
      }
      .ticket-title {
        font-size: 24px;
        font-weight: bold;
        margin: 0 0 5px 0;
      }
      .ticket-id {
        font-size: 18px;
        opacity: 0.9;
      }
      .ticket-row {
        display: flex;
        justify-content: space-between;
        margin: 12px 0;
        padding: 8px;
        background: rgba(255,255,255,0.1);
        border-radius: 8px;
      }
      .ticket-label {
        font-weight: 600;
        opacity: 0.9;
      }
      .ticket-value {
        font-weight: 500;
        text-align: right;
      }
      .status-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        background: rgba(255,255,255,0.2);
        font-size: 14px;
      }
      .btn-container {
        text-align: center;
        margin-top: 20px;
      }
      .btn-ok {
        background: white;
        color: #667eea;
        border: none;
        padding: 12px 40px;
        border-radius: 25px;
        font-size: 16px;
        font-weight: bold;
        cursor: pointer;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        transition: transform 0.2s;
      }
      .btn-ok:hover {
        transform: scale(1.05);
      }
    </style>
    <div class="ticket-popup">
      <div class="ticket-header">
        <div class="ticket-title">PHIẾU BÁO HỎNG</div>
        <div class="ticket-id">ID: ${ticketId}</div>
      </div>
      <div class="ticket-row">
        <span class="ticket-label">NGÀY, GIỜ::</span>
        <span class="ticket-value">${date} | ${time}</span>
      </div>
      <div class="ticket-row">
        <span class="ticket-label">KHU VỰC:</span>
        <span class="ticket-value">${area} - ${equipment}</span>
          <span class="ticket-label">Thiết bị:</span>
        <span class="ticket-value"> ${equipment}</span>
      </div>
      <div class="ticket-row">
        <span class="ticket-label">NGƯỜI GHI NHẬN:</span>
        <span class="ticket-value">${equipment_staff}</span>
      </div>
      <div class="ticket-row">
        <span class="ticket-label">MÔ TẢ KHIẾM KHUYẾT:</span>
        <span class="ticket-value"><span class="status-badge">${equipment_status}</span></span>
      </div>
      <div class="btn-container">
        <button class="btn-ok" onclick="window.location.href='/'">OK</button>
      </div>
    </div>
  `;

  res.send(popupHtml);
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

// fix
/*
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
}); */
// Update ticket info with photo upload
app.post('/fix/:id', upload.single('photo_after'), (req, res) => {
  const { id } = req.params;
  const {
    completion_time, total_processing_time, cause_recognition,
    solution, materials_used, equipment_status_after, recommendations
  } = req.body;

  // Get photo filename if uploaded
  const photo_after = req.file ? req.file.filename : null;

  // Update SQL to include photo_after
  const sql = `UPDATE tickets SET
    completion_time=?, total_processing_time=?,
    cause_recognition=?, solution=?, materials_used=?,
    equipment_status_after=?, recommendations=?, photo_after=?,
    status='finished', updated_at=CURRENT_TIMESTAMP
  WHERE id=?`;

  db.run(sql, [
    completion_time, total_processing_time, cause_recognition,
    solution, materials_used, equipment_status_after, recommendations,
    photo_after, id
  ], (err) => {
    if (err) {
      console.log(err);
      return res.status(500).send('Lỗi update ticket');
    }
    
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
    doc.fontSize(12).font('Roboto-Bold').text('CÔNG TY TNHH MTV CADIVI ĐỒNG NAI                  CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { align: 'right' });
    
        doc.fontSize(10).font('Roboto-Bold').text(' Độc lập-Tự do-Hạnh phúc ', { align: 'right' });
         doc.fontSize(10).font('Roboto-Bold').text('ĐƠN VỊ: CN CADIVI TÂN Á  ' , { align: 'left' });
          doc.fontSize(10).font('Roboto-Bold').text('Số .............-......../BB-KTCĐ  ', { align: 'left' } );


    doc.moveDown(1);
  // doc.font('Roboto').fontSize(10).text(`Ticket ID: #${ticket.id}`, { align: 'left' });
 

    // Add ticket details
    doc.font('Roboto-Bold').fontSize(20).text('BIÊN BẢN DỪNG THIẾT BỊ',{ align: 'center' });
    doc.moveDown(0.5);
    
    doc.fontSize(10).font('Helvetica');
    doc.font('Roboto').text(`Hôm nay lúc: ${ticket.stop_time}`);
    doc.font('Roboto').text(`Chúng tôi gồm:`);
    doc.font('Roboto').text(`Ông(bà): ${ticket.production_manager || 'N/A'}     Quản lý sản xuất` );
    doc.font('Roboto').text(`Ông(bà): ${ticket.equipment_staff || 'N/A'}        Nhân viên thiết bị`);
    doc.font('Roboto').text(`Ông(bà): ${ticket.maintenance_staff || 'N/A'}      Nhân viên bảo trì`);
    doc.font('Roboto').text(`Cùng nhau tiến hành dừng thiết bị (tên máy): : ${ticket.equipment || 'N/A'}  để thực hiện việc: ${ticket.work_type || 'N/A'}`);
    doc.font('Roboto-Bold').text(`Mô tả tình trạng thiết bị / nguyên nhân:`);
    doc.font('Roboto').text(`1.	Tình trạng thiết bị (dành cho nhân viên vận hành):${ticket.equipment_status || 'N/A'}`);
    doc.moveDown(1)
    doc.font('Roboto').text(`2.	Nguyên nhân hư hỏng (dành cho nhân viên bảo trì):${ticket.failure_reason || 'N/A'}`);
    doc.font('Roboto-Bold').text(`Thời điểm dừng, thiết bị đang được triển khai LSX số (theo SAP):${ticket.lsx_number || 'N/A'}`);
    doc.font('Roboto').text(`Mô tả tình trạng thiết bị / nguyên nhân:`);
    doc.moveDown(2)
    doc.font('Roboto-Bold').text(`CNVH thiết bị              NV P.KTCĐ             Quản lý sản xuất`,{ align: 'adjust' });
    doc.font('Roboto').text(`(Ký tên)                     (Ký tên)                (Ký tên)`,{ align: 'adjust' });

    // Add photo if exists
    if (ticket.photo) {
      const photoPath = path.join(__dirname, 'uploads', ticket.photo);
      if (fs.existsSync(photoPath)) {
        doc.moveDown();
        doc.fontSize(14).font('Roboto-Bold').text('Hình ảnh');
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




// BIEN BAN BAO GIAO//

app.get('/print/fix/:id', (req, res) => {
  const id = req.params.id;

  db.get("SELECT * FROM tickets WHERE id = ?", [id], (err, ticket) => {
    if (err) return res.status(500).send('Database error');
    if (!ticket) return res.status(404).send('Ticket not found');

    const doc = new PDFDocument({ margin: 50 });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename=fix_report_${id}.pdf`);

    doc.registerFont('Roboto', 'fonts/Roboto-Regular.ttf');
    doc.registerFont('Roboto-Bold', 'fonts/Roboto-Bold.ttf');
    
    doc.pipe(res);

    // Header
    doc.fontSize(12).font('Roboto-Bold').text('CÔNG TY TNHH MTV CADIVI ĐỒNG NAI                  CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { align: 'right' });
    doc.fontSize(10).font('Roboto-Bold').text(' Độc lập-Tự do-Hạnh phúc ', { align: 'right' });
    doc.fontSize(10).font('Roboto-Bold').text('ĐƠN VỊ: CN CADIVI TÂN Á  ', { align: 'left' });
    doc.fontSize(10).font('Roboto-Bold').text(`Số ${id}-${new Date().getFullYear()}/BB-SC  `, { align: 'left' });

    doc.moveDown(1);

    // Title
    doc.font('Roboto-Bold').fontSize(20).text('BIÊN BẢN SỬA CHỮA THIẾT BỊ', { align: 'center' });
    doc.moveDown(0.5);
    //Contetn
     doc.fontSize(10).font('Helvetica');
    doc.font('Roboto').text(`Hôm nay lúc: ${ticket.completion_time }`);
    doc.font('Roboto').text(`Chúng tôi gồm:`);
    doc.font('Roboto').text(`Ông(bà): ${ticket.production_manager || 'N/A'}     Quản lý sản xuất` );
    doc.font('Roboto').text(`Ông(bà): ${ticket.equipment_staff || 'N/A'}        Nhân viên thiết bị`);
    doc.font('Roboto').text(`Ông(bà): ${ticket.maintenance_staff || 'N/A'}      Nhân viên bảo trì`);
    doc.font('Roboto').text(`Cùng nhau tiến hành bàn giao thiết bị (tên máy): : ${ticket.equipment || 'N/A'}  theo nội dung công việc: ${ticket.work_type || 'N/A'}`);
    doc.font('Roboto').text(`1.Thời gian bắt đầu bảo trì: (Bắt đầu SC) ${ticket.stop_time}`);
    doc.font('Roboto').text(`2.Thời gian kết thúc bảo trì: (Bắt đầu SC) ${ticket.completion_time}`);
// Calculate duration if both times exist
    let durationText = '....';
    if (ticket.stop_time && ticket.completion_time) {
      const startTime = new Date(ticket.stop_time);
      const endTime = new Date(ticket.completion_time);
      const diffMs = endTime - startTime;
      
      if (diffMs >= 0) {
        const hours = Math.floor(diffMs / (1000 * 60 * 60));
        const minutes = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
        durationText = `${hours} giờ ${minutes} phút`;
      }
    }
    doc.font('Roboto').text(`3.Tổng thời gian bảo trì: ${durationText}`);
    doc.font('Roboto').text(`4.Thời gian chờ vật tư (khác):${ticket.total_processing_time}`);

    doc.font('Roboto-Bold').text(`Khối lượng công việc đã giải quyết:`);
    doc.font('Roboto').text(`1.Biện pháp khắc phục:${ticket.solution}`);
        doc.font('Roboto').text(`2.Vật tư sử dụng (nếu có):${ticket.materials_used}`);
        doc.font('Roboto').text(`3.Tình trạng thiết bị sau khi bảo trì:${ticket.equipment_status_after}`);
        doc.font('Roboto-Bold').text(`Đề nghị/ khuyến cáo:${ticket.recommendations}`);

                doc.font('Roboto').text(`Kết luận:  □ Đồng ý đưa vào sản xuất  □ Không đồng ý  □ Cần theo dõi `, { align: 'adjust' });
 doc.font('Roboto-Bold').text(`CNVH thiết bị              NV P.KTCĐ             Quản lý sản xuất`,{ align: 'adjust' });
    doc.font('Roboto').text(`(Ký tên)                     (Ký tên)                (Ký tên)`,{ align: 'adjust' });


    // Photo
    if (ticket.photo) {
      const photoPath = path.join(__dirname, 'uploads', ticket.photo_after);
      if (fs.existsSync(photoPath)) {
        doc.moveDown();
        doc.fontSize(12).font('Roboto-Bold').text('Hình ảnh sau sửa chữa:');
        doc.moveDown(0.5);
        try {
          doc.image(photoPath, { fit: [400, 300], align: 'center' });
        } catch (e) {
          //doc.fontSize(10).font('Roboto').text('(Không thể load ảnh)');
        }
      }
    }

    
    doc.end();
  });
});
// =====================
// EXPORT ENTIRE DATABASE TO EXCEL
// =====================



// Export entire database as Excel
app.get('/export/database', async (req, res) => {
  try {
    // Get all tickets from database
    db.all("SELECT * FROM tickets ORDER BY created_at DESC", async (err, tickets) => {
      if (err) {
        console.error("Database error:", err);
        return res.status(500).send('Database error');
      }

      // Create workbook
      const workbook = new ExcelJS.Workbook();
      
      // Set workbook properties
      workbook.creator = 'Ticket Management System';
      workbook.created = new Date();
      workbook.modified = new Date();

      // =====================
      // Sheet 1: All Tickets (Full Data)
      // =====================
      const worksheet = workbook.addWorksheet('Tất cả Tickets', {
        views: [{ state: 'frozen', ySplit: 1 }] // Freeze header row
      });

      // Define columns with Vietnamese headers
      worksheet.columns = [
        { header: 'ID', key: 'id', width: 8 },
        { header: 'Khu vực', key: 'area', width: 20 },
        { header: 'Thiết bị', key: 'equipment', width: 25 },
        { header: 'Quản lý SX', key: 'production_manager', width: 20 },
        { header: 'NV Thiết bị', key: 'equipment_staff', width: 20 },
        { header: 'NV Bảo trì', key: 'maintenance_staff', width: 20 },
        { header: 'Loại công việc', key: 'work_type', width: 18 },
        { header: 'TT Thiết bị', key: 'equipment_status', width: 18 },
        { header: 'Nguyên nhân hỏng', key: 'failure_reason', width: 35 },
        { header: 'Thời gian dừng', key: 'stop_time', width: 18 },
        { header: 'Số LSX', key: 'lsx_number', width: 15 },
        { header: 'Trạng thái', key: 'status', width: 15 },
        { header: 'Ảnh', key: 'photo', width: 20 },
        { header: 'TG Hoàn thành', key: 'completion_time', width: 20 },
        { header: 'Tổng TG XL', key: 'total_processing_time', width: 18 },
        { header: 'Nhận định NN', key: 'cause_recognition', width: 35 },
        { header: 'Giải pháp', key: 'solution', width: 35 },
        { header: 'Vật tư sử dụng', key: 'materials_used', width: 25 },
        { header: 'TT TB sau XL', key: 'equipment_status_after', width: 20 },
        { header: 'Kiến nghị', key: 'recommendations', width: 35 },
        { header: 'Ngày tạo', key: 'created_at', width: 22 },
        { header: 'Cập nhật', key: 'updated_at', width: 22 }
      ];

      // Style header row
      worksheet.getRow(1).font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
      worksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF2C3E50' }
      };
      worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
      worksheet.getRow(1).height = 25;

      // Add data
      tickets.forEach((ticket, index) => {
        const row = worksheet.addRow(ticket);
        
        // Alternate row colors
        if (index % 2 === 0) {
          row.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF8F9FA' }
          };
        }

        // Color code status
        const statusCell = row.getCell('status');
        if (ticket.status === 'finished') {
          statusCell.font = { color: { argb: 'FF28A745' }, bold: true };
        } else if (ticket.status === 'open') {
          statusCell.font = { color: { argb: 'FFDC3545' }, bold: true };
        }

        // Wrap text for long fields
        row.getCell('failure_reason').alignment = { wrapText: true };
        row.getCell('cause_recognition').alignment = { wrapText: true };
        row.getCell('solution').alignment = { wrapText: true };
        row.getCell('recommendations').alignment = { wrapText: true };
      });

      // Add borders
      worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell) => {
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
            left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
            bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
            right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
          };
        });
      });

      // Auto-filter
      worksheet.autoFilter = {
        from: 'A1',
        to: 'V1'
      };

      // =====================
      // Sheet 2: Statistics Summary
      // =====================
      const statsSheet = workbook.addWorksheet('Thống kê');
      
      // Calculate statistics
      const totalTickets = tickets.length;
      const openTickets = tickets.filter(t => t.status === 'open').length;
      const finishedTickets = tickets.filter(t => t.status === 'finished').length;
      
      // Count by area
      const areaStats = {};
      tickets.forEach(t => {
        const area = t.area || 'N/A';
        areaStats[area] = (areaStats[area] || 0) + 1;
      });

      // Count by equipment
      const equipmentStats = {};
      tickets.forEach(t => {
        const equipment = t.equipment || 'N/A';
        equipmentStats[equipment] = (equipmentStats[equipment] || 0) + 1;
      });

      // Add title
      statsSheet.mergeCells('A1:D1');
      statsSheet.getCell('A1').value = 'THỐNG KÊ TỔNG QUAN';
      statsSheet.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
      statsSheet.getCell('A1').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF2C3E50' }
      };
      statsSheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
      statsSheet.getRow(1).height = 30;

      // Overall stats
      statsSheet.addRow([]);
      statsSheet.addRow(['Tổng số tickets:', totalTickets]);
      statsSheet.addRow(['Tickets đang mở:', openTickets]);
      statsSheet.addRow(['Tickets đã hoàn thành:', finishedTickets]);
      statsSheet.addRow(['Tỷ lệ hoàn thành:', `${((finishedTickets/totalTickets)*100).toFixed(1)}%`]);

      // Style stats
      ['A3', 'A4', 'A5', 'A6'].forEach(cell => {
        statsSheet.getCell(cell).font = { bold: true };
      });
      statsSheet.getColumn('A').width = 25;
      statsSheet.getColumn('B').width = 20;

      // Stats by area
      statsSheet.addRow([]);
      statsSheet.addRow(['THỐNG KÊ THEO KHU VỰC']);
      statsSheet.getCell('A8').font = { bold: true, size: 12 };
      statsSheet.addRow(['Khu vực', 'Số lượng']);
      statsSheet.getRow(9).font = { bold: true };
      
      Object.entries(areaStats).forEach(([area, count]) => {
        statsSheet.addRow([area, count]);
      });

      // Stats by equipment
      const equipmentStartRow = statsSheet.rowCount + 2;
      statsSheet.addRow([]);
      statsSheet.addRow(['THỐNG KÊ THEO THIẾT BỊ']);
      statsSheet.getCell(`A${equipmentStartRow}`).font = { bold: true, size: 12 };
      statsSheet.addRow(['Thiết bị', 'Số lượng']);
      statsSheet.getRow(equipmentStartRow + 1).font = { bold: true };
      
      Object.entries(equipmentStats).forEach(([equipment, count]) => {
        statsSheet.addRow([equipment, count]);
      });

      // =====================
      // Sheet 3: Open Tickets Only
      // =====================
      const openSheet = workbook.addWorksheet('Tickets đang mở');
      openSheet.columns = worksheet.columns; // Same columns
      
      // Header styling
      openSheet.getRow(1).font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
      openSheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFDC3545' }
      };
      openSheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
      openSheet.getRow(1).height = 25;

      // Add only open tickets
      const openTicketsData = tickets.filter(t => t.status === 'open');
      openTicketsData.forEach((ticket, index) => {
        const row = openSheet.addRow(ticket);
        if (index % 2 === 0) {
          row.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFF8F8' }
          };
        }
      });

      openSheet.autoFilter = { from: 'A1', to: 'V1' };

      // =====================
      // Sheet 4: Finished Tickets Only
      // =====================
      const finishedSheet = workbook.addWorksheet('Tickets đã hoàn thành');
      finishedSheet.columns = worksheet.columns; // Same columns
      
      // Header styling
      finishedSheet.getRow(1).font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
      finishedSheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF28A745' }
      };
      finishedSheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
      finishedSheet.getRow(1).height = 25;

      // Add only finished tickets
      const finishedTicketsData = tickets.filter(t => t.status === 'finished');
      finishedTicketsData.forEach((ticket, index) => {
        const row = finishedSheet.addRow(ticket);
        if (index % 2 === 0) {
          row.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF8FFF8' }
          };
        }
      });

      finishedSheet.autoFilter = { from: 'A1', to: 'V1' };

      // =====================
      // Generate and send file
      // =====================
      const filename = `database_backup_${new Date().toISOString().split('T')[0]}.xlsx`;
      
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

      await workbook.xlsx.write(res);
      res.end();

      console.log(`✅ Database exported: ${filename} (${totalTickets} tickets)`);
    });

  } catch (error) {
    console.error('Export error:', error);
    res.status(500).send('Error exporting database: ' + error.message);
  }
});

// =====================
// Add button to dashboard to export database
// =====================
// You can add this HTML to your dashboard.ejs:


// =====================
// START SERVER
// =====================
const PORT = process.env.PORT || 80;
app.listen(PORT, () => {
  console.log(`Server chạy tại http://localhost:${PORT}`);
});
