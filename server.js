process.env.TZ = 'Asia/Ho_Chi_Minh'; // Luôn dùng giờ Việt Nam, không phụ thuộc múi giờ VPS

const express = require('express');
const fs = require('fs');
const cors = require('cors');
const path = require('path');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer'); // Chỉ dùng cho API gửi email (tùy chọn)

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || '0.0.0.0';

// Thư mục lưu dữ liệu: trên Render đặt DATA_DIR=/data (Disk), chạy trên máy thì dùng thư mục dự án
const DATA_DIR = process.env.DATA_DIR || __dirname;
const filePath = path.join(DATA_DIR, 'database.json');
const uploadDir = path.join(DATA_DIR, 'uploads');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });

// Lần đầu chạy trên Disk: copy database.json mẫu trong dự án sang Disk
const seedFile = path.join(__dirname, 'database.json');
if (!fs.existsSync(filePath) && fs.existsSync(seedFile)) {
    fs.copyFileSync(seedFile, filePath);
}

const upload = multer({ dest: uploadDir });
// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ==========================================
// HÀM HỖ TRỢ ĐỌC & GHI FILE DATABASE.JSON
// ==========================================
const DEFAULT_DATA = {
    danhSachNhanVien: [],
    lichSuDatCom: [],
    keHoachNhap: [],
    trackingList: [],
    vanDonDiGiao: [],
    vanDonThuHoi: [],
    nhaVuyenChuyen: ["247", "Nhất Tín", "Viettel", "GNNB"],
    danhSachTraNVC: [],
    nhaVuyenChuyenTra: ["247", "Nhất Tín", "Viettel"],
    chamCongList: [],
    nhaCungUngList: [
        { ten: "Hoa Anh Đào", donGia: 42500 },
        { ten: "Flex", donGia: 43000 },
        { ten: "Topaz", donGia: 60000 }
    ],
    vattu: []
};

function writeData(data) {
    const tmp = filePath + '.tmp';
    fs.writeFileSync(tmp, JSON.stringify(data, null, 2), 'utf8');
    fs.renameSync(tmp, filePath); // đổi tên là thao tác nguyên tử, không bao giờ ghi dở file
}

function readData() {
    if (!fs.existsSync(filePath)) {
        const fresh = JSON.parse(JSON.stringify(DEFAULT_DATA));
        writeData(fresh);
        return fresh;
    }
    let data;
    try {
        data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    } catch (error) {
        // File hỏng: giữ lại bản sao để cứu dữ liệu, KHÔNG trả về dữ liệu rỗng để tránh ghi đè mất sạch
        fs.copyFileSync(filePath, filePath + '.hong-' + Date.now());
        console.error('database.json bị lỗi định dạng:', error);
        throw new Error('database.json bị lỗi, đã sao lưu file hỏng. Hãy kiểm tra trên server.');
    }
    Object.keys(DEFAULT_DATA).forEach(function (k) {
        if (data[k] === undefined) data[k] = JSON.parse(JSON.stringify(DEFAULT_DATA[k]));
    });
    return data;
}

function ngayHomNayVN() {
    const now = new Date();
    return String(now.getDate()).padStart(2, '0') + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + now.getFullYear();
}

// ==========================================
// 1. API NHÂN VIÊN & ĐẶT CƠM
// ==========================================
app.get('/api/nhan-vien', (req, res) => {
    let db = readData();
    let danhSach = db.danhSachNhanVien || [];
    let danhSachHopLe = danhSach.filter(nv => nv.nhom && nv.nhom.toUpperCase() !== 'GNNB');
    res.json({ success: true, data: danhSachHopLe });
});

app.post('/api/them-nhan-vien', (req, res) => {
    try {
        const { hoVaTen, nhom } = req.body;
        if (!hoVaTen || !nhom) {
            return res.status(400).json({ success: false, message: 'Thiếu thông tin nhân viên!' });
        }
        let db = readData();
        let existing = db.danhSachNhanVien.find(nv => nv.hoVaTen.toLowerCase() === hoVaTen.toLowerCase());
        if (existing) {
            return res.json({ success: false, message: 'Nhân viên này đã có trong hệ thống!' });
        }
        db.danhSachNhanVien.push({ hoVaTen, nhom });
        writeData(db);
        res.json({ success: true, message: `Đã lưu thành công nhân viên ${hoVaTen}!` });
    } catch (err) {
        res.status(500).json({ success: false, message: err.message });
    }
});

app.delete('/api/nhan-vien/:ten', (req, res) => {
    try {
        const tenNV = decodeURIComponent(req.params.ten);
        let db = readData();
        let initialLength = db.danhSachNhanVien.length;
        db.danhSachNhanVien = db.danhSachNhanVien.filter(nv => nv.hoVaTen.toLowerCase() !== tenNV.toLowerCase());
        if (db.danhSachNhanVien.length === initialLength) {
            return res.json({ success: false, message: 'Không tìm thấy nhân viên cần xóa!' });
        }
        writeData(db);
        res.json({ success: true, message: `Đã xóa thành công nhân viên ${tenNV}!` });
    } catch (err) {
        res.status(500).json({ success: false, message: err.message });
    }
});

app.post('/api/dat-com', (req, res) => {
    try {
        const { ngayDat, chiTietSuatAn, tongSoLuong } = req.body;
        let db = readData();
        db.lichSuDatCom.push({
            ngayDat: ngayDat || "Hôm nay",
            chiTiet: chiTietSuatAn || [],
            tongSoLuong: tongSoLuong || 0,
            thoiGianLuu: new Date().toLocaleString('vi-VN')
        });
        writeData(db);
        res.json({ success: true, message: 'Đã lưu lịch sử đặt cơm thành công!' });
    } catch (err) {
        res.status(500).json({ success: false, message: err.message });
    }
});

app.get('/api/lich-su-dat-com', (req, res) => {
    try {
        let db = readData();
        res.json({ success: true, data: db.lichSuDatCom || [] });
    } catch (err) {
        res.status(500).json({ success: false, message: err.message });
    }
});

// ==========================================
// 2. API KẾ HOẠCH NHẬP HÀNG & TRACKING
// ==========================================
app.post('/api/upload-ke-hoach', upload.single('keHoachFile'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ success: false, message: 'Chưa chọn tệp Excel kế hoạch!' });
    }
    try {
        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        const ngayHomNay = ngayHomNayVN();

        let newRecords = sheetData.map((row) => ({
            ngay: ngayHomNay,
            kho: row['Kho tổng'] || row['kho'] || 'HO-HCM',
            nganhHang: row['Ngành hàng'] || row['nganhHang'] || '',
            loaiYc: row['Loại yêu cầu'] || row['loaiYc'] || 'Nhập',
            ncc: row['NCC'] || row['ncc'] || '',
            soDinhDanh: row['Số định danh'] || row['soDinhDanh'] || '',
            slKeHoach: parseInt(row['Số lượng'] || row['slKeHoach']) || 0,
            slThucNhan: parseInt(row['Thực nhận'] || row['slThucNhan']) || 0,
            ghiChu: row['Note'] || row['ghiChu'] || '',
            nguon: row['Source'] || row['nguon'] || 'Excel'
        }));

        let db = readData();
        db.keHoachNhap = db.keHoachNhap.concat(newRecords);
        db.keHoachNhap.sort((a, b) => (a.ncc || "").localeCompare(b.ncc || "", 'vi', { sensitivity: 'base' }));
        db.keHoachNhap.forEach((item, idx) => { item.stt = idx + 1; });
        writeData(db);

        res.json({ success: true, message: 'Đã tải lên và gộp dữ liệu thành công!', data: db.keHoachNhap });
    } catch (error) {
        console.error(error);
        res.status(500).json({ success: false, message: 'Lỗi đọc file Excel kế hoạch!' });
    } finally {
        // Luôn dọn file tạm, kể cả khi lỗi
        try { if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path); } catch (e) { }
    }
});

app.get('/api/lay-ke-hoach', (req, res) => {
    try {
        let db = readData();
        res.json({ success: true, data: db.keHoachNhap || [] });
    } catch (error) {
        res.status(500).json({ success: false, message: "Lỗi đọc kế hoạch!" });
    }
});

app.get('/api/lay-tracking', (req, res) => {
    try {
        let db = readData();
        res.json({ success: true, data: db.trackingList || [] });
    } catch (error) {
        res.status(500).json({ success: false, message: "Lỗi đọc tracking!" });
    }
});

app.post('/api/them-ke-hoach', (req, res) => {
    try {
        let db = readData();
        const donMoi = {
            stt: 1,
            ngay: ngayHomNayVN(),
            kho: 'HO-HCM',
            nganhHang: 'Apple',
            loaiYc: 'Nhập',
            ncc: 'Nhà Cung Cấp Mới',
            soDinhDanh: '',
            slKeHoach: 0,
            slThucNhan: 0,
            ghiChu: '',
            nguon: 'Admin'
        };
        db.keHoachNhap.unshift(donMoi);
        db.keHoachNhap.forEach((item, idx) => { item.stt = idx + 1; });
        writeData(db);
        res.json({ success: true, message: 'Đã thêm đơn mới thành công!', data: db.keHoachNhap });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

app.post('/api/xoa-ke-hoach', (req, res) => {
    try {
        const { index } = req.body;
        let db = readData();
        if (Number.isInteger(index) && index >= 0 && db.keHoachNhap.length > index) {
            db.keHoachNhap.splice(index, 1);
            db.keHoachNhap.forEach((item, idx) => { item.stt = idx + 1; });
            writeData(db);
            res.json({ success: true, message: 'Đã xóa đơn thành công!', data: db.keHoachNhap });
        } else {
            res.status(400).json({ success: false, message: 'Không tìm thấy vị trí đơn cần xóa!' });
        }
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

app.post('/api/cap-nhat-ke-hoach', (req, res) => {
    try {
        const danhSachMoi = req.body.danhSach;
        if (!Array.isArray(danhSachMoi)) {
            return res.status(400).json({ success: false, message: "Dữ liệu kế hoạch không hợp lệ!" });
        }
        let db = readData();
        db.keHoachNhap = danhSachMoi;
        writeData(db);
        res.json({ success: true, message: "Lưu kế hoạch thành công!" });
    } catch (error) {
        console.error(error);
        res.status(500).json({ success: false, message: "Lỗi server khi lưu kế hoạch!" });
    }
});

app.post('/api/cap-nhat-tracking', (req, res) => {
    try {
        const danhSachTracking = req.body.trackingList || [];
        let db = readData();
        db.trackingList = danhSachTracking;
        writeData(db);
        res.json({ success: true, message: "Đã ghi nhận tracking vào database thành công!" });
    } catch (error) {
        console.error(error);
        res.status(500).json({ success: false, message: "Lỗi ghi tracking vào database!" });
    }
});

// ==========================================
// 3. API THU HỒI VẬN ĐƠN
// ==========================================
app.get('/api/lay-thu-hoi-van-don', (req, res) => {
    try {
        let db = readData();
        res.json({
            success: true,
            vanDonDiGiao: db.vanDonDiGiao || [],
            vanDonThuHoi: db.vanDonThuHoi || [],
            nhaVuyenChuyen: db.nhaVuyenChuyen || ["247", "Nhất Tín", "Viettel", "GNNB"]
        });
    } catch (e) {
        res.status(500).json({ success: false, message: "Lỗi đọc dữ liệu thu hồi!" });
    }
});

app.post('/api/cap-nhat-thu-hoi', (req, res) => {
    try {
        let db = readData();
        db.vanDonDiGiao = req.body.vanDonDiGiao || [];
        db.vanDonThuHoi = req.body.vanDonThuHoi || [];
        if (req.body.nhaVuyenChuyen) {
            db.nhaVuyenChuyen = req.body.nhaVuyenChuyen;
        }
        writeData(db);
        res.json({ success: true, message: "Đã lưu dữ liệu thu hồi thành công!" });
    } catch (e) {
        res.status(500).json({ success: false, message: "Lỗi ghi dữ liệu thu hồi!" });
    }
});

// ==========================================
// 4. API TRẢ VẬN ĐƠN CHO NHÀ VẬN CHUYỂN
// ==========================================
app.get('/api/lay-tra-van-don', (req, res) => {
    try {
        let db = readData();
        res.json({
            success: true,
            danhSachTraNVC: db.danhSachTraNVC || [],
            nhaVuyenChuyenTra: db.nhaVuyenChuyenTra || ["247", "Nhất Tín", "Viettel"]
        });
    } catch (e) {
        res.status(500).json({ success: false, message: "Lỗi đọc dữ liệu trả vận đơn!" });
    }
});

app.post('/api/cap-nhat-tra-van-don', (req, res) => {
    try {
        let db = readData();
        db.danhSachTraNVC = req.body.danhSachTraNVC || [];
        if (req.body.nhaVuyenChuyenTra) {
            db.nhaVuyenChuyenTra = req.body.nhaVuyenChuyenTra;
        }
        writeData(db);
        res.json({ success: true, message: "Đã lưu dữ liệu trả vận đơn thành công!" });
    } catch (e) {
        res.status(500).json({ success: false, message: "Lỗi lưu dữ liệu trả vận đơn!" });
    }
});

// ==========================================
// 5. API CHẤM CÔNG & THỜI VỤ
// ==========================================
app.get('/api/lay-cham-cong-thoi-vu', (req, res) => {
    try {
        let db = readData();
        res.json({
            success: true,
            chamCongList: db.chamCongList || [],
            nhaCungUngList: db.nhaCungUngList || DEFAULT_DATA.nhaCungUngList
        });
    } catch (e) {
        res.status(500).json({ success: false, message: "Lỗi đọc dữ liệu chấm công thời vụ!" });
    }
});

app.post('/api/cap-nhat-cham-cong-thoi-vu', (req, res) => {
    try {
        let db = readData();
        db.chamCongList = req.body.chamCongList || [];
        if (req.body.nhaCungUngList) {
            db.nhaCungUngList = req.body.nhaCungUngList;
        }
        writeData(db);
        res.json({ success: true, message: "Đã lưu dữ liệu chấm công thời vụ thành công!" });
    } catch (e) {
        res.status(500).json({ success: false, message: "Lỗi lưu dữ liệu chấm công thời vụ!" });
    }
});

// ==========================================
// 6. API GỬI EMAIL BIÊN BẢN (TÙY CHỌN - cần biến môi trường MAIL_USER, MAIL_PASS)
// ==========================================
const transporter = (process.env.MAIL_USER && process.env.MAIL_PASS)
    ? nodemailer.createTransport({
        service: 'gmail',
        auth: { user: process.env.MAIL_USER, pass: process.env.MAIL_PASS }
    })
    : null;

app.post('/api/gui-email-bien-ban', async (req, res) => {
    if (!transporter) {
        return res.status(503).json({ success: false, message: "Server chưa cấu hình MAIL_USER / MAIL_PASS!" });
    }
    try {
        const { soPhieu, daiDienNCC, nguoiGiao, danhSachHang, ccList } = req.body;
        let noiDungHang = (danhSachHang || []).map((item, idx) =>
            `<tr>
                <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">${idx + 1}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">${item.maHang || ''}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">${item.tenHang || ''}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">${item.imei || ''}</td>
                <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">${item.soLuong || 1}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">${item.tinhTrang || ''}</td>
            </tr>`
        ).join('');
        let mailOptions = {
            from: `"Kho Tổng FPT Shop Miền Nam" <${process.env.MAIL_USER}>`,
            to: 'namnvh@fpt.com',
            cc: (ccList || []).join(','),
            subject: `[BIÊN BẢN KIỂM HÀNG] ${soPhieu} - NCC: ${daiDienNCC}`,
            html: `
                <h3>Kính gửi các anh chị,</h3>
                <p>Kho Tổng Miền Nam xin gửi biên bản kiểm hàng không đạt điều kiện nhập kho:</p>
                <ul>
                    <li><b>Số phiếu:</b> ${soPhieu}</li>
                    <li><b>Bên bàn giao:</b> ${nguoiGiao} (${daiDienNCC})</li>
                </ul>
                <table style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 13px;">
                    <thead>
                        <tr style="background: #1e293b; color: white;">
                            <th style="border: 1px solid #ddd; padding: 8px;">STT</th>
                            <th style="border: 1px solid #ddd; padding: 8px;">Mã Hàng</th>
                            <th style="border: 1px solid #ddd; padding: 8px;">Tên Hàng</th>
                            <th style="border: 1px solid #ddd; padding: 8px;">IMEI / Serial</th>
                            <th style="border: 1px solid #ddd; padding: 8px;">SL</th>
                            <th style="border: 1px solid #ddd; padding: 8px;">Tình Trạng Lỗi</th>
                        </tr>
                    </thead>
                    <tbody>${noiDungHang}</tbody>
                </table>
                <p>Trân trọng,<br><b>Kho Tổng FPT Shop Miền Nam</b></p>
            `
        };
        await transporter.sendMail(mailOptions);
        res.json({ success: true, message: "Đã gửi email thành công qua server!" });
    } catch (error) {
        console.error("Lỗi gửi email:", error);
        res.status(500).json({ success: false, message: "Lỗi gửi email từ server!" });
    }
});

// ==========================================
// 7. API QUẢN LÝ VẬT TƯ KHO
// ==========================================
app.get('/api/vattu', (req, res) => {
    try {
        const ngay = req.query.ngay;
        let db = readData();

        // 1. Ngày được chọn đã có dữ liệu
        const exactRecord = db.vattu.find(item => item.ngay === ngay);
        if (exactRecord) {
            return res.json({ success: true, data: exactRecord.danhSach });
        }

        // 2. Chưa có: kế thừa số tồn của ngày gần nhất trước đó
        const sortedVattu = [...db.vattu].sort((a, b) => new Date(b.ngay) - new Date(a.ngay));
        const previousRecord = sortedVattu.find(item => item.ngay < ngay);
        if (previousRecord) {
            const carriedData = previousRecord.danhSach.map(item => ({
                tenVatTu: item.tenVatTu,
                nhap: 0,
                xuat: 0,
                ton: item.ton
            }));
            return res.json({ success: true, data: carriedData });
        }

        // 3. Hoàn toàn mới
        res.json({ success: true, data: [] });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

app.post('/api/vattu', (req, res) => {
    try {
        const { ngay, danhSachVatTu } = req.body;
        if (!ngay || !Array.isArray(danhSachVatTu)) {
            return res.status(400).json({ success: false, message: "Thiếu thông tin ngày hoặc danh sách vật tư!" });
        }

        let db = readData();
        const existingIndex = db.vattu.findIndex(item => item.ngay === ngay);

        if (existingIndex >= 0) {
            const record = db.vattu[existingIndex];
            danhSachVatTu.forEach(newItem => {
                const oldItem = record.danhSach.find(i => i.tenVatTu.toLowerCase() === newItem.tenVatTu.toLowerCase());
                if (oldItem) {
                    oldItem.nhap = (parseFloat(oldItem.nhap) || 0) + (parseFloat(newItem.nhap) || 0);
                    oldItem.xuat = (parseFloat(oldItem.xuat) || 0) + (parseFloat(newItem.xuat) || 0);
                    oldItem.ton = (parseFloat(oldItem.ton) || 0) + (parseFloat(newItem.nhap) || 0) - (parseFloat(newItem.xuat) || 0);
                } else {
                    record.danhSach.push(newItem);
                }
            });
            // Vật tư đã bị xóa trên giao diện thì xóa luôn khỏi ngày này
            const tenConLai = danhSachVatTu.map(i => i.tenVatTu.toLowerCase());
            record.danhSach = record.danhSach.filter(i => tenConLai.includes(i.tenVatTu.toLowerCase()));
            record.updatedAt = new Date().toISOString();
        } else {
            db.vattu.push({
                ngay: ngay,
                danhSach: danhSachVatTu,
                updatedAt: new Date().toISOString()
            });
        }

        writeData(db);
        res.json({ success: true, message: "Đã lưu dữ liệu vật tư thành công!" });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

// ==========================================
// XỬ LÝ LỖI & KHỞI ĐỘNG SERVER
// ==========================================
app.use('/api', (req, res) => res.status(404).json({ success: false, message: 'API không tồn tại' }));

app.use((err, req, res, next) => {
    console.error(err);
    res.status(500).json({ success: false, message: err.message || 'Lỗi server' });
});

app.listen(PORT, '0.0.0.0', () => {
  console.log("Server Kho đang chạy ở cổng " + PORT);
});