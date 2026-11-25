// ==========================================
// KHAI BÁO BIẾN TOÀN CỤC CHUNG
// ==========================================
var globalKeHoachData = [];
var globalLichSuData = [];
var scannedTrackingList = [];
var bienBanTimer = null;
var danhSachHangBienBan = [
    { id: 1, maHang: "", tenHang: "", imei: "", soLuong: 1, tinhTrang: "Móp méo / Không đạt" }
];

// ==========================================
// 1. HÀM THÊM NHÂN VIÊN MỚI TỪ GIAO DIỆN
// ==========================================
async function addNewStaff(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var nameInput = document.getElementById('newStaffName');
    var groupSelect = document.getElementById('newStaffGroup');
    
    var hoVaTen = nameInput.value.trim();
    var nhom = groupSelect.value;
    if (!hoVaTen) {
        alert('Vui lòng nhập mã hoặc họ tên nhân viên!');
        return;
    }
    try {
        let response = await fetch('http://localhost:3000/api/them-nhan-vien', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ hoVaTen: hoVaTen, nhom: nhom })
        });
        let result = await response.json();
        if (result.success) {
            alert('✅ ' + result.message);
            nameInput.value = '';
        } else {
            alert('⚠️ ' + result.message);
        }
    } catch (error) {
        console.error("Lỗi kết nối:", error);
        alert('Không thể kết nối đến Server Back End để lưu nhân viên mới!');
    }
    hienThiLoading(false);
}

// ==========================================
// 2. XỬ LÝ ĐỌC FILE EXCEL ĐIỂM DANH & KHỞI TẠO TRANG
// ==========================================
document.addEventListener('DOMContentLoaded', function() {
    var lastTab = localStorage.getItem('activeSubKeHoachTab') || 'hienTai';
    switchSubKeHoachTab(lastTab);
    
    var btnNhanDien = document.getElementById('btnNhanDienExcel') || document.querySelector('.btn-upload');
    var ngayDatInput = document.getElementById('ngayDat');
    if (ngayDatInput && !ngayDatInput.value) {
        ngayDatInput.valueAsDate = new Date();
    }
    
    if (btnNhanDien) {
        btnNhanDien.addEventListener('click', async function(e) {
            e.preventDefault();
            var fileInput = document.getElementById('excelFile');
            if (!fileInput || fileInput.files.length === 0) {
                alert('Vui lòng chọn tệp Excel điểm danh trước khi nhấn nhận diện!');
                return;
            }
            try {
                let res = await fetch('http://localhost:3000/api/nhan-vien');
                let result = await res.json();
                var danhSachChuan = result.data || [];
                var file = fileInput.files[0];
                var reader = new FileReader();
                reader.onload = function(e) {
                    var data = new Uint8Array(e.target.result);
                    var workbook = XLSX.read(data, { type: 'array' });
                    var worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    var danhSachDiemDanh = [];
                    var nhanVienDiMuon = [];
                    jsonData.forEach(function(row, idx) {
                        if (idx > 0 && row[1]) {
                            var tenNV = String(row[1]).trim();
                            danhSachDiemDanh.push(tenNV);
                            var gioVao = row[2] ? String(row[2]).trim() : ""; 
                            if (gioVao && isLateTime(gioVao)) {
                                nhanVienDiMuon.push(tenNV + " (Vào lúc: " + gioVao + ")");
                            }
                        }
                    });
                    
                    var nhanVienDiLam = danhSachChuan.filter(function(nv) {
                        return danhSachDiemDanh.some(function(ten) {
                            return ten.toLowerCase().includes(nv.hoVaTen.toLowerCase());
                        });
                    });
                    
                    var nhanVienVangMat = danhSachChuan.filter(function(nv) {
                        return !danhSachDiemDanh.some(function(ten) {
                            return ten.toLowerCase().includes(nv.hoVaTen.toLowerCase());
                        });
                    });
                    
                    displayAttendanceWarnings(nhanVienDiMuon, nhanVienVangMat);
                    
                    var statusEl = document.getElementById('status');
                    if(statusEl) {
                        statusEl.innerText = "✅ Nhận diện thành công! Đi làm: " + nhanVienDiLam.length + " | Vắng mặt: " + nhanVienVangMat.length;
                        statusEl.style.color = "#16a34a";
                    }
                    var resultArea = document.getElementById('resultArea');
                    if(resultArea) {
                        resultArea.style.display = 'block';
                    }
                    
                    renderMenuTableDynamic(nhanVienDiLam);
                };
                reader.readAsArrayBuffer(file);
            } catch (error) {
                console.error("Lỗi kết nối Server:", error);
                alert("Không thể kết nối đến Server Back End!");
            }
        });
    }
    var savedTracking = localStorage.getItem('scannedTrackingList');
    if (savedTracking) {
        try {
            scannedTrackingList = JSON.parse(savedTracking);
            renderWebTrackingTable();
        } catch(e) {
            scannedTrackingList = [];
        }
    }
    
    loadKeHoachData();
    loadThuHoiData();
    loadTraVNDData();
    loadThoiVuData();
});

function isLateTime(timeStr,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var parts = timeStr.split(':');
    if (parts.length >= 2) {
        var hour = parseInt(parts[0], 10);
        var minute = parseInt(parts[1], 10);
        if (hour > 8 || (hour === 8 && minute > 30)) {
            return true;
        }
    }
    hienThiLoading(false);
    return false;
    
}

function displayAttendanceWarnings(diMuonList, vangMatList,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var warningArea = document.getElementById('attendanceWarningArea');
    var lateListEl = document.getElementById('lateStaffList');
    var absentListEl = document.getElementById('absentStaffList');
    
    if (!warningArea) return;
    warningArea.style.display = 'flex';
    
    if (diMuonList.length > 0) {
        lateListEl.innerHTML = diMuonList.map(function(item) { return "<li>" + item + "</li>"; }).join('');
    } else {
        lateListEl.innerHTML = "<li><em>Không có nhân viên nào đi muộn sau 08:30.</em></li>";
    }
    
    if (vangMatList.length > 0) {
        absentListEl.innerHTML = vangMatList.map(function(item) { return "<li>" + item.hoVaTen + " (" + item.nhom + ")</li>"; }).join('');
    } else {
        absentListEl.innerHTML = "<li><em>Tuyệt vời! Không có nhân viên nào vắng mặt.</em></li>";
    }
    hienThiLoading(false);
}

// ==========================================
// 3. RENDER BẢNG THỰC ĐƠN ĐẶT CƠM
// ==========================================
function renderMenuTableDynamic(nhanVienDiLamList,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var tbody = document.getElementById('tableBody');
    if (!tbody) return;
    var groupCounts = {};
    nhanVienDiLamList.forEach(function(nv) {
        var nhomKey = nv.nhom || "Khác";
        if (!groupCounts[nhomKey]) groupCounts[nhomKey] = 0;
        groupCounts[nhomKey]++;
    });
    var html = "";
    var danhSachBoPhan = ["Kho Bán Mới", "Kho Ecom", "Kho Chờ Xử Lý", "Quản lý", "Khách"];
    danhSachBoPhan.forEach(function(nhom) {
        var sl = groupCounts[nhom] || 0;
        var sang = 0, chaySang = 0, trua = sl, chayTrua = 0, toi = 0, chayToi = 0;
        var tongSo = sang + chaySang + trua + chayTrua + toi + chayToi;
        
        html += "<tr>";
        if (nhom === "Khách") {
            html += "<td style='font-weight:bold; text-align:left; padding: 10px; color: #ea580c;'>👤 Khách / Đối tác (Tự điền)</td>";
        } else {
            html += "<td style='font-weight:bold; text-align:left; padding: 10px;'>" + nhom + " (Đi làm: <span style='color:#16a34a;'>" + sl + "</span>)</td>";
        }
        html += "<td><input type='number' class='input-sang' value='" + sang + "' oninput='tinhTongRow(this)'></td>";
        html += "<td><input type='number' class='input-chay-sang' value='" + chaySang + "' oninput='tinhTongRow(this)'></td>";
        
        if (nhom === "Khách") {
            html += "<td><input type='number' class='input-trua' value='" + trua + "' oninput='tinhTongRow(this)'></td>";
            html += "<td><input type='number' class='input-chay-trua' value='" + chayTrua + "' oninput='tinhTongRow(this)'></td>";
        } else {
            html += "<td><input type='number' class='input-trua' value='" + trua + "' oninput='tinhLaiTongSo(this, " + sl + ")'></td>";
            html += "<td><input type='number' class='input-chay-trua' value='" + chayTrua + "' oninput='tinhLaiTongSo(this, " + sl + ")'></td>";
        }
        html += "<td><input type='number' class='input-toi' value='" + toi + "' oninput='tinhTongRow(this)'></td>";
        html += "<td><input type='number' class='input-chay-toi' value='" + chayToi + "' oninput='tinhTongRow(this)'></td>";
        html += "<td style='font-weight:bold; color: #2563eb;' class='col-tong'>" + tongSo + "</td>";
        html += "</tr>";
    });
    tbody.innerHTML = html;
    hienThiLoading(false);
}

function tinhTongRow(inputEl,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var row = inputEl.closest('tr');
    var inputs = row.querySelectorAll('input[type="number"]');
    var sum = 0;
    inputs.forEach(function(inp) {
        sum += parseInt(inp.value) || 0;
    });
    var totalCell = row.querySelector('.col-tong');
    if (totalCell) {
        totalCell.innerText = sum;
    }
    hienThiLoading(false);
}

function tinhLaiTongSo(inputEl, maxSl,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var row = inputEl.closest('tr');
    var truaInput = row.querySelector('.input-trua');
    var chayTruaInput = row.querySelector('.input-chay-trua');
    var valChay = parseInt(chayTruaInput.value) || 0;
    
    if (valChay > maxSl) {
        valChay = maxSl;
        chayTruaInput.value = valChay;
    }
    var valMan = maxSl - valChay;
    truaInput.value = valMan;
    tinhTongRow(inputEl);
    hienThiLoading(false);
}

// ==========================================
// 4. MỞ & QUẢN LÝ POP-UP XÁC NHẬN & BÁO CÁO KỲ
// ==========================================
function openPopupXacNhan(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var ngayDatVal = document.getElementById('ngayDat') ? document.getElementById('ngayDat').value : '';
    var formattedDate = "03-11-2026";
    if (ngayDatVal) {
        var parts = ngayDatVal.split('-');
        if (parts.length === 3) {
            formattedDate = parts[2] + '-' + parts[1] + '-' + parts[0];
        }
    }
    var dateTextEl = document.getElementById('popupDateText');
    if (dateTextEl) {
        dateTextEl.innerText = "Ngày đặt: " + formattedDate;
    }
    var totalSang = 0, totalSangChay = 0;
    var totalTrua = 0, totalTruaChay = 0;
    var totalToi = 0, totalToiChay = 0;
    var rows = document.querySelectorAll('#tableBody tr');
    rows.forEach(function(row) {
        var inputs = row.querySelectorAll('input[type="number"]');
        if (inputs.length >= 6) {
            totalSang += parseInt(inputs[0].value) || 0;
            totalSangChay += parseInt(inputs[1].value) || 0;
            totalTrua += parseInt(inputs[2].value) || 0;
            totalTruaChay += parseInt(inputs[3].value) || 0;
            totalToi += parseInt(inputs[4].value) || 0;
            totalToiChay += parseInt(inputs[5].value) || 0;
        }
    });
    var grandTotal = totalSang + totalSangChay + totalTrua + totalTruaChay + totalToi + totalToiChay;
    var mealListEl = document.getElementById('popupMealList');
    if (mealListEl) {
        mealListEl.innerHTML = `
            <div class="popup-row"><span>Ăn sáng:</span> <span class="val">${totalSang}</span></div>
            <div class="popup-row"><span>Ăn sáng chay:</span> <span class="val">${totalSangChay}</span></div>
            <div class="popup-row"><span>Ăn trưa:</span> <span class="val">${totalTrua}</span></div>
            <div class="popup-row"><span>Ăn trưa chay:</span> <span class="val">${totalTruaChay}</span></div>
            <div class="popup-row"><span>Ăn tối:</span> <span class="val">${totalToi}</span></div>
            <div class="popup-row"><span>Ăn tối chay:</span> <span class="val">${totalToiChay}</span></div>
            <div class="popup-total"><span>Tổng số lượng:</span> <span style="color: #ea580c;">${grandTotal} phần</span></div>
        `;
    }
    var popupModal = document.getElementById('popupModal');
    if (popupModal) {
        popupModal.style.display = 'flex';
    }
    hienThiLoading(false);
}

function closePopup(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var popupModal = document.getElementById('popupModal');
    if (popupModal) {
        popupModal.style.display = 'none';
    }
    hienThiLoading(false);
}

function copyPopupAsImage(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var popupCard = document.querySelector('.popup-card');
    if (!popupCard) return;
    var footerButtons = popupCard.querySelector('.popup-footer');
    if (footerButtons) {
        footerButtons.style.display = 'none';
    }
    html2canvas(popupCard, { scale: 2, backgroundColor: '#ffffff' }).then(function(canvas) {
        if (footerButtons) {
            footerButtons.style.display = 'flex';
        }
        canvas.toBlob(function(blob) {
            if (!blob) {
                alert("Không thể tạo hình ảnh từ Pop-up!");
                return;
            }
            try {
                navigator.clipboard.write([
                    new ClipboardItem({ 'image/png': blob })
                ]).then(function() {
                    alert("📸 Đã sao chép bảng Pop-up dưới dạng ảnh thành công!\nBạn hãy sang Zalo và nhấn Ctrl + V để dán gửi nhé.");
                    submitDatCom();
                }).catch(function(err) {
                    console.error(err);
                    alert("Trình duyệt không hỗ trợ quyền sao chép hình ảnh trực tiếp.");
                });
            } catch (e) {
                console.error(e);
            }
        });
    }).catch(function(err) {
        if (footerButtons) {
            footerButtons.style.display = 'flex';
        }
        console.error(err);
    });
    hienThiLoading(false);
}

async function submitDatCom(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var ngayDatVal = document.getElementById('ngayDat') ? document.getElementById('ngayDat').value : "Hôm nay";
    var chiTietSuatAn = [];
    var totalGrand = 0;
    var rows = document.querySelectorAll('#tableBody tr');
    rows.forEach(function(row) {
        var boPhanText = row.cells[0].innerText;
        var inputs = row.querySelectorAll('input[type="number"]');
        if (inputs.length >= 6) {
            var item = {
                boPhan: boPhanText,
                sang: parseInt(inputs[0].value) || 0,
                sangChay: parseInt(inputs[1].value) || 0,
                trua: parseInt(inputs[2].value) || 0,
                truaChay: parseInt(inputs[3].value) || 0,
                toi: parseInt(inputs[4].value) || 0,
                toiChay: parseInt(inputs[5].value) || 0
            };
            var sumRow = item.sang + item.sangChay + item.trua + item.truaChay + item.toi + item.toiChay;
            item.tongRow = sumRow;
            totalGrand += sumRow;
            chiTietSuatAn.push(item);
        }
    });
    try {
        let response = await fetch('http://localhost:3000/api/dat-com', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                ngayDat: ngayDatVal,
                chiTietSuatAn: chiTietSuatAn,
                tongSoLuong: totalGrand
            })
        });
        let result = await response.json();
        if (result.success) {
            closePopup();
            console.log("✅ Đã lưu lịch sử đặt cơm thành công!");
        }
    } catch (error) {
        console.error("Lỗi kết nối:", error);
    }
    hienThiLoading(false);
}

async function openDeleteStaffModal(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var modal = document.getElementById('deleteStaffModal');
    var tbody = document.getElementById('staffListTableBody');
    if (!modal || !tbody) return;
    try {
        let res = await fetch('http://localhost:3000/api/nhan-vien');
        let result = await res.json();
        var danhSach = result.data || [];
        var html = "";
        danhSach.forEach(function(nv) {
            html += "<tr style='border-bottom: 1px solid #f1f5f9;'>";
            html += "<td style='padding: 10px; font-weight: 600;'>" + nv.hoVaTen + "</td>";
            html += "<td style='padding: 10px; color: #64748b;'>" + nv.nhom + "</td>";
            html += "<td style='padding: 10px; text-align: center;'><button type='button' onclick=\"deleteStaff('" + nv.hoVaTen + "')\" style='background: #fee2e2; color: #dc2626; border: none; padding: 6px 12px; border-radius: 4px; font-weight: bold; cursor: pointer;'>[-] Xóa</button></td>";
            html += "</tr>";
        });
        tbody.innerHTML = html;
        modal.style.display = 'flex';
    } catch (e) {
        console.error(e);
        alert("Không thể tải danh sách nhân viên từ Server!");
    }
    hienThiLoading(false);
}

function closeDeleteStaffModal(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var modal = document.getElementById('deleteStaffModal');
    if (modal) modal.style.display = 'none';
    hienThiLoading(false);
}

async function deleteStaff(hoVaTen,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    if (!confirm("Bạn có chắc chắn muốn xóa nhân viên [" + hoVaTen + "] ra khỏi hệ thống không?")) return;
    try {
        let res = await fetch('http://localhost:3000/api/nhan-vien/' + encodeURIComponent(hoVaTen), {
            method: 'DELETE'
        });
        let result = await res.json();
        if (result.success) {
            alert("✅ " + result.message);
            openDeleteStaffModal();
        } else {
            alert("⚠️ " + result.message);
        }
    } catch (e) {
        console.error(e);
        alert("Lỗi kết nối khi xóa nhân viên!");
    }
    hienThiLoading(false);
}

async function loadBaoCaoKy(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var monthVal = parseInt(document.getElementById('reportMonthSelect').value, 10);
    var yearVal = parseInt(document.getElementById('reportYearSelect').value, 10);
    var startMonth = monthVal - 1;
    var startYear = yearVal;
    if (startMonth === 0) {
        startMonth = 12;
        startYear = yearVal - 1;
    }
    var startDate = new Date(startYear, startMonth - 1, 26);
    var endDate = new Date(yearVal, monthVal - 1, 25);
    var titleEl = document.getElementById('reportTitlePeriod');
    if (titleEl) {
        titleEl.innerText = `📊 BÁO CÁO KỲ TỪ ${pad(startDate.getDate())}-${pad(startDate.getMonth() + 1)}-${startDate.getFullYear()} ĐẾN ${pad(endDate.getDate())}-${pad(endDate.getMonth() + 1)}-${endDate.getFullYear()}:`;
    }
    try {
        let res = await fetch('http://localhost:3000/api/lich-su-dat-com');
        let result = await res.json();
        var allHistory = result.data || [];
        var tbody = document.getElementById('reportTableBody');
        var html = "";
        
        var filteredData = allHistory.filter(function(item) {
            var parts = item.ngayDat.split('-');
            var itemDate;
            if (parts[0].length === 4) {
                itemDate = new Date(parts[0], parts[1] - 1, parts[2]);
            } else {
                itemDate = new Date(parts[2], parts[1] - 1, parts[0]);
            }
            return itemDate >= startDate && itemDate <= endDate;
        });
        
        if (filteredData.length === 0) {
            tbody.innerHTML = `<tr><td colspan="6" style="padding: 20px; color: #64748b;">Không có dữ liệu đặt cơm trong kỳ này.</td></tr>`;
            // Xóa bảng tổng nếu có sẵn từ lần tra cứu trước
            var containerTongCu = document.getElementById('khuVucBangTongThang');
            if (containerTongCu) containerTongCu.innerHTML = "";
            return;
        }

        // Biến cộng dồn cho bảng tổng tháng
        var tongSuatSang = 0, tongSuatTrua = 0, tongSuatToi = 0;
        var donGiaMacDinh = 30000; // Đơn giá theo logic tính hiện tại của bạn

        filteredData.forEach(function(record) {
            var chiTiet = record.chiTiet || [];
            var sumSang = 0, sumTrua = 0, sumToi = 0;
            chiTiet.forEach(function(row) {
                var s = (row.sang || 0) + (row.sangChay || 0);
                var t = (row.trua || 0) + (row.truaChay || 0);
                var o = (row.toi || 0) + (row.toiChay || 0);
                
                sumSang += s;
                sumTrua += t;
                sumToi += o;

                // Cộng dồn vào tổng tháng
                tongSuatSang += s;
                tongSuatTrua += t;
                tongSuatToi += o;
            });
            var totalRow = sumSang + sumTrua + sumToi;
            var ThanhTien = totalRow * donGiaMacDinh;
            
            html += `<tr>`;
            html += `<td style="font-weight: bold;">${record.ngayDat}</td>`;
            html += `<td>${sumSang}</td>`;
            html += `<td>${sumTrua}</td>`;
            html += `<td>${sumToi}</td>`;
            html += `<td style="font-weight: bold; color: #ea580c;">${totalRow}</td>`;
            html += `<td style="font-weight: bold; color: #16a34a;">${ThanhTien.toLocaleString('vi-VN')} VNĐ</td>`;
            html += `</tr>`;
        });
        tbody.innerHTML = html;

        // --- RENDER BẢNG TỔNG HỢP SUẤT ĂN & THÀNH TIỀN PHÍA DƯỚI ---
        renderBangTongThang(tongSuatSang, tongSuatTrua, tongSuatToi, donGiaMacDinh);

    } catch (e) {
        console.error(e);
        alert("Lỗi khi tải dữ liệu báo cáo từ Server!");
    }
    hienThiLoading(false);
}

// HÀM TẠO VÀ HIỂN THỊ BẢNG TỔNG THÁNG BÊN DƯỚI
function renderBangTongThang(tongSang, tongTrua, tongToi, donGia,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var tienSang = tongSang * donGia;
    var tienTrua = tongTrua * donGia;
    var tienToi = tongToi * donGia;
    var tongCongSuat = tongSang + tongTrua + tongToi;
    var tongCongTien = tienSang + tienTrua + tienToi;

    var containerTong = document.getElementById('khuVucBangTongThang');
    if (!containerTong) {
        var parentDiv = document.getElementById('reportTableBody').closest('div').parentNode;
        containerTong = document.createElement('div');
        containerTong.id = 'khuVucBangTongThang';
        containerTong.style.marginTop = '30px';
        parentDiv.appendChild(containerTong);
    }

    containerTong.innerHTML = `
        <h3 style="color: #1e293b; margin-bottom: 12px; font-size: 16px;">📑 TỔNG HỢP SỐ LƯỢNG SUẤT ĂN & CHI PHÍ TRONG KỲ</h3>
        <div style="overflow-x: auto;">
            <table border="1" style="width: 100%; border-collapse: collapse; text-align: center; background: white; font-size: 13px;">
                <thead>
                    <tr style="background: #0f172a; color: white;">
                        <th style="padding: 12px;">BUỔI ĂN</th>
                        <th style="padding: 12px;">TỔNG SỐ SUẤT</th>
                        <th style="padding: 12px;">ĐƠN GIÁ (VNĐ)</th>
                        <th style="padding: 12px; text-align: right; color: #4ade80;">THÀNH TIỀN (VNĐ)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="padding: 10px; font-weight: bold; text-align: left;">Sáng</td>
                        <td style="padding: 10px;">${tongSang.toLocaleString('vi-VN')}</td>
                        <td style="padding: 10px;">${donGia.toLocaleString('vi-VN')} đ</td>
                        <td style="padding: 10px; text-align: right; font-weight: bold;">${tienSang.toLocaleString('vi-VN')} đ</td>
                    </tr>
                    <tr>
                        <td style="padding: 10px; font-weight: bold; text-align: left;">Trưa</td>
                        <td style="padding: 10px;">${tongTrua.toLocaleString('vi-VN')}</td>
                        <td style="padding: 10px;">${donGia.toLocaleString('vi-VN')} đ</td>
                        <td style="padding: 10px; text-align: right; font-weight: bold;">${tienTrua.toLocaleString('vi-VN')} đ</td>
                    </tr>
                    <tr>
                        <td style="padding: 10px; font-weight: bold; text-align: left;">Tối</td>
                        <td style="padding: 10px;">${tongToi.toLocaleString('vi-VN')}</td>
                        <td style="padding: 10px;">${donGia.toLocaleString('vi-VN')} đ</td>
                        <td style="padding: 10px; text-align: right; font-weight: bold;">${tienToi.toLocaleString('vi-VN')} đ</td>
                    </tr>
                    <tr style="background: #f8fafc; font-weight: bold; font-size: 14px;">
                        <td style="padding: 12px; text-align: left; color: #1e293b;">TỔNG CỘNG</td>
                        <td style="padding: 12px; color: #2563eb;">${tongCongSuat.toLocaleString('vi-VN')} suất</td>
                        <td style="padding: 12px;">--</td>
                        <td style="padding: 12px; text-align: right; color: #16a34a; font-size: 15px;">${tongCongTien.toLocaleString('vi-VN')} VNĐ</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `;
    hienThiLoading(false);
}

function pad(n) {
    return n < 10 ? '0' + n : n;
}

// ==========================================
// 5. CHỨC NĂNG 1: KẾ HOẠCH NHẬP HÀNG
// ==========================================
async function uploadKeHoachFile(inputElement,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    if (!inputElement.files || inputElement.files.length === 0) return;
    var file = inputElement.files[0];
    var formData = new FormData();
    formData.append('keHoachFile', file);
    try {
        let res = await fetch('http://localhost:3000/api/upload-ke-hoach', {
            method: 'POST',
            body: formData
        });
        let result = await res.json();
        if (result.success) {
            alert('✅ ' + result.message);
            globalKeHoachData = result.data;
            renderKeHoachTable(globalKeHoachData);
        } else {
            alert('⚠️ ' + result.message);
        }
    } catch (e) {
        console.error(e);
        alert('Lỗi kết nối khi tải lên file kế hoạch!');
    }
}

async function loadKeHoachData(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    try {
        let res = await fetch('http://localhost:3000/api/lay-ke-hoach');
        let result = await res.json();
        if (result.success) {
            globalKeHoachData = result.data || [];
            renderKeHoachTable(globalKeHoachData);
        }
    } catch (e) {
        console.error(e);
    }
    hienThiLoading(false);
}

function renderKeHoachTable(dataList,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var tbody = document.getElementById('keHoachTableBody');
    if (!tbody) return;
    if (!dataList || dataList.length === 0) {
        tbody.innerHTML = `<tr><td colspan="13" style="padding: 30px; color: #64748b; text-align: center;">Chưa có dữ liệu kế hoạch.</td></tr>`;
        return;
    }
    var now = new Date();
    var dd = String(now.getDate()).padStart(2, '0');
    var mm = String(now.getMonth() + 1).padStart(2, '0');
    var yyyy = now.getFullYear();
    var ngayHienTaiStr = dd + '-' + mm + '-' + yyyy;
    
    var txtNgay = document.getElementById('txtNgayHienTai');
    if (txtNgay) txtNgay.innerText = ngayHienTaiStr;
    var dataHomNay = [];
    dataList.forEach(function(item, originalIdx) {
        if (!item.ngay || item.ngay === '1/1/1970') {
            dataHomNay.push({ item: item, index: originalIdx });
            return;
        }
        var itemDateStr = item.ngay.toString().trim();
        if (itemDateStr === ngayHienTaiStr || itemDateStr.includes(yyyy + '-' + mm + '-' + dd) || itemDateStr.includes(dd + '-' + mm + '-' + yyyy)) {
            dataHomNay.push({ item: item, index: originalIdx });
        }
    });
    if (dataHomNay.length === 0) {
        tbody.innerHTML = `<tr><td colspan="13" style="padding: 40px; color: #64748b; text-align: center;">Chưa có dữ liệu kế hoạch cho ngày hôm nay (${ngayHienTaiStr}).</td></tr>`;
        return;
    }
    var html = "";
    dataHomNay.forEach(function(obj, displayIdx) {
        var item = obj.item;
        var originalIdx = obj.index;
        var slKh = parseInt(item.slKeHoach) || 0;
        var slTn = parseInt(item.slThucNhan) || 0;
        var slConLai = slKh - slTn;
        var hienThiNgay = (item.ngay && item.ngay !== '1/1/1970') ? item.ngay : ngayHienTaiStr;
        var isNewRow = (item.nguon === 'Admin' || item.ncc === 'Nhà Cung Cấp Mới');
        
        html += `<tr data-index="${originalIdx}" style="border-bottom: 1px solid #e2e8f0; text-align: center;">`;
        html += `<td style="padding: 10px; border: 1px solid #e2e8f0; font-weight: bold;">${displayIdx + 1}</td>`;
        html += `<td style="padding: 10px; border: 1px solid #e2e8f0; white-space: nowrap; background: #f8fafc; font-weight: 500;">${hienThiNgay}</td>`;
        html += `<td style="padding: 10px; border: 1px solid #e2e8f0; white-space: nowrap; font-weight: 600; background: #f8fafc;">${item.kho || 'HO-HCM'}</td>`;
        
        if (isNewRow) {
            html += `<td style="padding: 8px; border: 1px solid #e2e8f0;"><input type="text" class="input-nganhhang" value="${item.nganhHang || ''}" style="width: 100px; padding: 5px; text-align: center; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 12px;" placeholder="Ngành hàng..."></td>`;
            html += `<td style="padding: 8px; border: 1px solid #e2e8f0;"><input type="text" class="input-loaiyc" value="${item.loaiYc || 'Nhập'}" style="width: 85px; padding: 5px; text-align: center; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 12px;" placeholder="Loại YC..."></td>`;
            html += `<td style="padding: 8px; border: 1px solid #e2e8f0;"><input type="text" class="input-ncc" value="${item.ncc || ''}" style="width: 180px; padding: 5px; text-align: left; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 12px; font-weight: bold;" placeholder="Nhập tên NCC..."></td>`;
            html += `<td style="padding: 8px; border: 1px solid #e2e8f0;"><input type="text" class="input-sodinhdanh" value="${item.soDinhDanh || ''}" style="width: 100px; padding: 5px; text-align: center; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 12px;" placeholder="Mã định danh..."></td>`;
            html += `<td style="padding: 8px; border: 1px solid #e2e8f0;"><input type="number" class="input-slkh" value="${slKh}" style="width: 75px; padding: 5px; text-align: center; border: 1.5px solid #2563eb; border-radius: 4px; font-weight: bold; color: #2563eb; font-size: 12px;"></td>`;
            html += `<td style="padding: 8px; border: 1px solid #e2e8f0;"><input type="number" class="input-thucnhan" value="${slTn}" style="width: 75px; padding: 5px; text-align: center; border: 1.5px solid #10b981; border-radius: 4px; font-weight: bold; color: #059669; font-size: 12px;"></td>`;
            html += `<td style="padding: 10px; border: 1px solid #e2e8f0; font-weight: bold; color: #dc2626; font-size: 13px;" class="col-conlai">${slConLai}</td>`;
            html += `<td style="padding: 8px; border: 1px solid #e2e8f0;"><input type="text" class="input-ghichu" value="${item.ghiChu || ''}" style="width: 140px; padding: 5px; text-align: left; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 12px;" placeholder="Ghi chú..."></td>`;
        } else {
            html += `<td style="padding: 10px; border: 1px solid #e2e8f0; white-space: nowrap;">${item.nganhHang || ''}</td>`;
            html += `<td style="padding: 10px; border: 1px solid #e2e8f0; white-space: nowrap;">${item.loaiYc || 'Nhập'}</td>`;
            html += `<td style="padding: 10px 12px; border: 1px solid #e2e8f0; text-align: left; font-weight: bold; color: #1e293b; max-width: 220px; white-space: normal; word-break: break-word;">${item.ncc || ''}</td>`;
            html += `<td style="padding: 8px; border: 1px solid #e2e8f0;"><input type="text" class="input-sodinhdanh" value="${item.soDinhDanh || ''}" style="width: 100px; padding: 5px; text-align: center; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 12px;" placeholder="Mã định danh..."></td>`;
            html += `<td style="padding: 10px; border: 1px solid #e2e8f0; font-weight: bold; font-size: 13px;" class="col-sl-kh">${slKh}</td>`;
            html += `<td style="padding: 8px; border: 1px solid #e2e8f0;"><input type="number" class="input-thucnhan" value="${slTn}" style="width: 75px; padding: 5px; text-align: center; border: 1.5px solid #10b981; border-radius: 4px; font-weight: bold; color: #059669; font-size: 12px;"></td>`;
            html += `<td style="padding: 10px; border: 1px solid #e2e8f0; font-weight: bold; color: #dc2626; font-size: 13px;" class="col-conlai">${slConLai}</td>`;
            html += `<td style="padding: 10px 12px; border: 1px solid #e2e8f0; text-align: left; color: #475569; max-width: 180px; white-space: normal; word-break: break-word;">${item.ghiChu || ''}</td>`;
        }
        
        html += `<td style="padding: 10px; border: 1px solid #e2e8f0; white-space: nowrap;"><span style="background: #f1f5f9; padding: 3px 6px; border-radius: 4px; font-weight: 600; font-size: 11px; border: 1px solid #e2e8f0;">${item.nguon || ''}</span></td>`;
        html += `<td style="padding: 10px; border: 1px solid #e2e8f0;"><button type="button" onclick="xoaDongKeHoach(this)" style="background: #fee2e2; color: #dc2626; border: none; padding: 5px 10px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 11px;">Xóa</button></td>`;
        html += `</tr>`;
    });
    tbody.innerHTML = html;
    hienThiLoading(false);
}

function themDonKeHoachMoi(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var now = new Date();
    var defaultDate = String(now.getDate()).padStart(2, '0') + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + now.getFullYear();
    
    var rowMoi = {
        ngay: defaultDate,
        kho: "HO-HCM",
        nganhHang: "",
        loaiYc: "Nhập",
        ncc: "",
        soDinhDanh: "",
        slKeHoach: 0,
        slThucNhan: 0,
        ghiChu: "",
        nguon: "Admin"
    };
    if (!globalKeHoachData) {
        globalKeHoachData = [];
    }
    globalKeHoachData.push(rowMoi);
    renderKeHoachTable(globalKeHoachData);
    window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
    hienThiLoading(false);
}

async function xoaDongKeHoach(btn,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var row = btn.closest('tr');
    var originalIdx = row.getAttribute('data-index');
    if (originalIdx !== null && originalIdx !== undefined) {
        originalIdx = parseInt(originalIdx, 10);
    } else {
        return;
    }
    
    if (confirm('Bạn có chắc chắn muốn xóa đơn này không?')) {
        try {
            let res = await fetch('http://localhost:3000/api/xoa-ke-hoach', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ index: originalIdx })
            });
            let result = await res.json();
            if (result.success) {
                globalKeHoachData = result.data;
                renderKeHoachTable(globalKeHoachData);
            } else {
                alert('⚠️ ' + result.message);
            }
        } catch (e) {
            console.error(e);
            alert('Lỗi kết nối khi xóa đơn!');
        }
    }
    hienThiLoading(false);
}

function capNhatMangDuLieuTam(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var rows = document.querySelectorAll('#keHoachTableBody tr');
    rows.forEach(function(r) {
        var originalIdx = r.getAttribute('data-index');
        if (originalIdx === null || originalIdx === undefined) return;
        originalIdx = parseInt(originalIdx, 10);
        
        var item = globalKeHoachData[originalIdx] ? Object.assign({}, globalKeHoachData[originalIdx]) : {};
        var inputs = r.querySelectorAll('input');
        
        if (inputs.length >= 7) {
            item.nganhHang = inputs[0] ? inputs[0].value : (item.nganhHang || '');
            item.loaiYc = inputs[1] ? inputs[1].value : (item.loaiYc || '');
            item.ncc = inputs[2] ? inputs[2].value : (item.ncc || '');
            item.soDinhDanh = inputs[3] ? inputs[3].value : (item.soDinhDanh || '');
            item.slKeHoach = inputs[4] ? (parseInt(inputs[4].value) || 0) : (item.slKeHoach || 0);
            item.slThucNhan = inputs[5] ? (parseInt(inputs[5].value) || 0) : (item.slThucNhan || 0);
            item.ghiChu = inputs[6] ? inputs[6].value : (item.ghiChu || '');
        } else if (inputs.length >= 2) {
            item.soDinhDanh = inputs[0] ? inputs[0].value : (item.soDinhDanh || '');
            item.slThucNhan = inputs[1] ? (parseInt(inputs[1].value) || 0) : (item.slThucNhan || 0);
        }
        globalKeHoachData[originalIdx] = item;
    });
    hienThiLoading(false);
}

document.addEventListener('input', function(e) {
    if (e.target.closest('#keHoachTableBody')) {
        var row = e.target.closest('tr');
        if (!row) return;
        if (e.target.classList.contains('input-thucnhan') || e.target.type === 'number') {
            var allInputsInRow = row.querySelectorAll('input');
            var slKh = 0, slTn = 0;
            
            if (allInputsInRow.length >= 7) {
                slKh = parseInt(allInputsInRow[4].value) || 0;
                slTn = parseInt(allInputsInRow[5].value) || 0;
            } else {
                let cellKhText = row.querySelector('.col-sl-kh') ? row.querySelector('.col-sl-kh').innerText : "0";
                slKh = parseInt(cellKhText) || 0;
                let tnInput = row.querySelector('.input-thucnhan');
                slTn = tnInput ? (parseInt(tnInput.value) || 0) : 0;
            }
            var slConLai = slKh - slTn;
            var slConLaiCell = row.querySelector('.col-conlai');
            if (slConLaiCell) {
                slConLaiCell.innerText = slConLai;
            }
        }
        capNhatMangDuLieuTam();
    }
});

async function luuKeHoachThuCong(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    capNhatMangDuLieuTam();
    var btnLuu = document.getElementById('btnLuuThayDoi');
    var originalText = "";
    if (btnLuu) {
        originalText = btnLuu.innerHTML;
        btnLuu.innerHTML = "⏳ Đang lưu & đồng bộ...";
        btnLuu.disabled = true;
        btnLuu.style.opacity = "0.7";
        btnLuu.style.cursor = "not-allowed";
    }
    try {
        let response = await fetch('http://localhost:3000/api/cap-nhat-ke-hoach', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ danhSach: globalKeHoachData })
        });
        let result = await response.json();
        
        if (result.success) {
            var saved = localStorage.getItem('scannedTrackingList');
            var currentTrackingList = [];
            if (saved) {
                try { currentTrackingList = JSON.parse(saved); } catch(e) { currentTrackingList = []; }
            }
            var now = new Date();
            var timeStr = now.toLocaleTimeString('vi-VN') + ' ' + String(now.getDate()).padStart(2, '0') + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + now.getFullYear();
            
            currentTrackingList.forEach(function(t) {
                if (t.status && t.status.includes('Khớp')) {
                    t.status = "Đã nhận (Chờ khớp KH)";
                    t.ncc = "-- Chưa chọn NCC --";
                    t.timeNhap = "-- Chưa nhập --";
                }
            });
            
            globalKeHoachData.forEach(function(item) {
                var ma = item.soDinhDanh ? item.soDinhDanh.trim() : "";
                var sl = parseInt(item.slThucNhan) || 0;
                
                if (ma !== "" && sl > 0) {
                    var codeLower = ma.toLowerCase();
                    var found = currentTrackingList.find(function(t) {
                        return t.code && t.code.toString().trim().toLowerCase() === codeLower;
                    });
                    if (found) {
                        found.ncc = item.ncc || "Nhà cung cấp không tên";
                        found.status = "✔ Khớp kế hoạch (Đã nhập)";
                        found.timeNhap = timeStr;
                    } else {
                        currentTrackingList.unshift({
                            id: Date.now() + Math.random(),
                            code: ma,
                            status: "✔ Khớp kế hoạch (Đã nhập)",
                            ncc: item.ncc || "Nhà cung cấp không tên",
                            timeNhan: timeStr,
                            timeNhap: timeStr
                        });
                    }
                }
            });
            localStorage.setItem('scannedTrackingList', JSON.stringify(currentTrackingList));
            scannedTrackingList = currentTrackingList;
            await fetch('http://localhost:3000/api/cap-nhat-tracking', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ trackingList: currentTrackingList })
            });
            if (typeof renderWebTrackingTable === 'function') {
                renderWebTrackingTable();
            }
            alert("✅ Đã lưu thay đổi và đồng bộ trạng thái thành công!");
        } else {
            alert("⚠️ Lỗi khi lưu dữ liệu từ server!");
        }
    } catch (error) {
        console.error("Lỗi kết nối server:", error);
        alert("❌ Không thể kết nối tới server để lưu!");
    } finally {
        if (btnLuu) {
            btnLuu.innerHTML = originalText;
            btnLuu.disabled = false;
            btnLuu.style.opacity = "1";
            btnLuu.style.cursor = "pointer";
        }
    }
    hienThiLoading(false);
}

// ==========================================
// 6. POPUP BÁO CÁO NHẬP & CHỤP ẢNH CLIPBOARD
// ==========================================
function moPopupBaoCaoNhap(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var popup = document.getElementById('baoCaoPopup');
    var tbody = document.getElementById('popupTableBody');
    var tongKetTopDiv = document.getElementById('popupTongKetTop');
    var titleEl = document.getElementById('popupTitle');
    
    if (!popup || !tbody) return;
    var now = new Date();
    var ngayStr = String(now.getDate()).padStart(2, '0') + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + now.getFullYear();
    if (titleEl) {
        titleEl.innerText = "KHO MIỀN NAM GỬI BÁO CÁO NHẬP - NGÀY " + ngayStr;
    }
    var html = "";
    var totalKh = 0;
    var totalThucNhan = 0;
    var totalConLai = 0;
    
    if (!globalKeHoachData || globalKeHoachData.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" style="padding: 20px; text-align: center; color: #64748b;">Chưa có dữ liệu báo cáo.</td></tr>`;
        popup.style.display = 'flex';
        return;
    }
    
    globalKeHoachData.forEach(function(item, index) {
        var slKh = parseInt(item.slKeHoach) || 0;
        var slTn = parseInt(item.slThucNhan) || 0;
        var slCl = slKh - slTn;
        totalKh += slKh;
        totalThucNhan += slTn;
        totalConLai += slCl;
        
        html += `<tr style="border-bottom: 1px solid #e2e8f0; text-align: center;">`;
        html += `<td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold;">${index + 1}</td>`;
        html += `<td style="padding: 8px; border: 1px solid #e2e8f0; white-space: nowrap;">${item.ngay}</td>`;
        html += `<td style="padding: 8px; border: 1px solid #e2e8f0; white-space: nowrap; font-weight: 600;">${item.kho || 'HO-HCM'}</td>`;
        html += `<td style="padding: 8px 10px; border: 1px solid #e2e8f0; text-align: left; font-weight: bold; color: #1e293b; max-width: 220px; white-space: normal; word-break: break-word;">${item.ncc || ''}</td>`;
        html += `<td style="padding: 8px; border: 1px solid #e2e8f0; white-space: nowrap;">${item.loaiYc || 'Nhập'}</td>`;
        html += `<td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold;">${slKh}</td>`;
        html += `<td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold; color: #059669;">${slTn}</td>`;
        html += `<td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold; color: #dc2626;">${slCl}</td>`;
        
        var ghiChuDisplay = item.ghiChu || '';
        if (slCl < 0) {
            ghiChuDisplay = `<span style="background: #fff5f5; color: #dc2626; padding: 3px 8px; border-radius: 4px; font-weight: bold; border: 1px solid #feb2b2; display: inline-block;">⚠️ NGOÀI KẾ HOẠCH</span> ${item.ghiChu ? '- ' + item.ghiChu : ''}`;
        }
        html += `<td style="padding: 8px 10px; border: 1px solid #e2e8f0; text-align: left; max-width: 200px; white-space: normal; word-break: break-word;">${ghiChuDisplay}</td>`;
        html += `</tr>`;
    });
    tbody.innerHTML = html;
    
    if (tongKetTopDiv) {
        tongKetTopDiv.innerHTML = `📊 TỔNG KẾ HOẠCH: ${totalKh} &nbsp;&nbsp;|&nbsp;&nbsp; TỔNG THỰC NHẬN: ${totalThucNhan} &nbsp;&nbsp;|&nbsp;&nbsp; TỔNG CÒN LẠI: ${totalConLai}`;
    }
    popup.style.display = 'flex';
    hienThiLoading(false);
}

function dongPopupBaoCaoNhap(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var popup = document.getElementById('baoCaoPopup');
    if (popup) popup.style.display = 'none';
    hienThiLoading(false);
}

function copyBaoCaoClipboard(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var textToCopy = "📊 KHO MIỀN NAM GỬI BÁO CÁO NHẬP\n----------------------------------\n";
    globalKeHoachData.forEach(function(item, index) {
        var slKh = parseInt(item.slKeHoach) || 0;
        var slTn = parseInt(item.slThucNhan) || 0;
        var slCl = slKh - slTn;
        textToCopy += `${index + 1}. ${item.ncc} | KH: ${slKh} | Thực nhận: ${slTn} | Còn lại: ${slCl} | Note: ${item.ghiChu || 'Không'}\n`;
    });
    navigator.clipboard.writeText(textToCopy).then(function() {
        alert("📋 Đã sao chép nội dung báo cáo vào bộ nhớ tạm thành công!");
    });
    hienThiLoading(false);
}

// ==========================================
// 7. ĐIỀU HƯỚNG TAB CON KẾ HOẠCH
// ==========================================

function switchSubKeHoachTab(type,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    localStorage.setItem('activeSubKeHoachTab', type);
    var viewHienTai = document.getElementById('subViewHienTai');
    var viewLichSu = document.getElementById('subViewLichSu');
    var viewTracking = document.getElementById('subViewTracking');
    var viewBienBan = document.getElementById('subViewBienBan');
    
    var btnHienTai = document.getElementById('subTabKeHoachBtn');
    var btnLichSu = document.getElementById('subTabLichSuBtn');
    var btnTracking = document.getElementById('subTabTrackingBtn');
    var btnBienBan = document.getElementById('subTabBienBanBtn');
    
    if (viewHienTai) viewHienTai.style.display = 'none';
    if (viewLichSu) viewLichSu.style.display = 'none';
    if (viewTracking) viewTracking.style.display = 'none';
    if (viewBienBan) viewBienBan.style.display = 'none';
    
    
    var inactiveBg = '#f1f5f9';
    var inactiveColor = '#475569';
    var inactiveBorder = '1px solid #cbd5e1';
    
    if (btnHienTai) { btnHienTai.style.background = inactiveBg; btnHienTai.style.color = inactiveColor; btnHienTai.style.border = inactiveBorder; }
    if (btnLichSu) { btnLichSu.style.background = inactiveBg; btnLichSu.style.color = inactiveColor; btnLichSu.style.border = inactiveBorder; }
    if (btnTracking) { btnTracking.style.background = inactiveBg; btnTracking.style.color = inactiveColor; btnTracking.style.border = inactiveBorder; }
    if (btnBienBan) { btnBienBan.style.background = inactiveBg; btnBienBan.style.color = inactiveColor; btnBienBan.style.border = inactiveBorder; }
    
    var activeBg = '#2563eb';
    var activeColor = '#ffffff';
    var activeBorder = 'none';
    
    if (type === 'hienTai') {
        if (viewHienTai) viewHienTai.style.display = 'block';
        if (btnHienTai) { btnHienTai.style.background = activeBg; btnHienTai.style.color = activeColor; btnHienTai.style.border = activeBorder; }
        loadKeHoachData();
    } else if (type === 'lichSu') {
        if (viewLichSu) viewLichSu.style.display = 'block';
        if (btnLichSu) { btnLichSu.style.background = activeBg; btnLichSu.style.color = activeColor; btnLichSu.style.border = activeBorder; }
        loadLichSuNhapHangData();
    } else if (type === 'tracking') {
        if (viewTracking) viewTracking.style.display = 'block';
        if (btnTracking) { btnTracking.style.background = activeBg; btnTracking.style.color = activeColor; btnTracking.style.border = activeBorder; }
        renderWebTrackingTable();
        setTimeout(function() {
            var input = document.getElementById('webTrackingInput');
            if (input) input.focus();
        }, 100);
    } else if (type === 'bienBan') {
        if (viewBienBan) viewBienBan.style.display = 'block';
        if (btnBienBan) { btnBienBan.style.background = activeBg; btnBienBan.style.color = activeColor; btnBienBan.style.border = activeBorder; }
        khoiTaoBienBan();
    }
    hienThiLoading(false);
}

// ==========================================
// 8. XỬ LÝ LỊCH SỬ NHẬP HÀNG & LỌC (FIX TẬN GỐC - ĐẢM BẢO HIỂN THỊ 100%)
// ==========================================
// ==========================================
// 2. KHỞI TẠO TRANG KHI VỪA TẢI XONG
// ==========================================
document.addEventListener('DOMContentLoaded', function() {
    var lastTab = localStorage.getItem('activeSubKeHoachTab') || 'hienTai';
    switchSubKeHoachTab(lastTab);
    
    // Gọi đồng thời để nạp sẵn dữ liệu cho cả Kế hoạch và Lịch sử
    loadKeHoachData();
    loadLichSuNhapHangData(); // <-- BỔ SUNG DÒNG NÀY ĐỂ NẠP SẴN DỮ LIỆU LỊCH SỬ

    var ngayDatInput = document.getElementById('ngayDat');
    if (ngayDatInput && !ngayDatInput.value) {
        ngayDatInput.valueAsDate = new Date();
    }
});
async function loadLichSuNhapHangData(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var tbody = document.getElementById('lichSuTableBody');
    if (tbody) {
        tbody.innerHTML = `<tr><td colspan="10" style="padding: 40px; color: #64748b; text-align: center;">⏳ Đang đồng bộ dữ liệu lịch sử...</td></tr>`;
    }

    // 1. Ưu tiên lấy từ biến toàn cục globalKeHoachData nếu đã có sẵn trên RAM
    if (typeof globalKeHoachData !== 'undefined' && globalKeHoachData.length > 0) {
        globalLichSuData = globalKeHoachData;
        renderLichSuTable(globalLichSuData);
        return;
    }

    // 2. Nếu chưa có, gọi API lấy từ Server
    try {
        let res = await fetch('http://localhost:3000/api/lay-ke-hoach');
        let result = await res.json();
        
        if (result.success && result.data && result.data.length > 0) {
            globalLichSuData = result.data;
            globalKeHoachData = result.data; // Đồng bộ ngược lại
            renderLichSuTable(globalLichSuData);
            return;
        }
    } catch (e) {
        console.error("Lỗi gọi API lịch sử:", e);
    }

    // 3. Dự phòng cuối cùng: Lấy từ localStorage nếu server/RAM trống
    try {
        var localData = localStorage.getItem('scannedTrackingList');
        if (localData) {
            var parsed = JSON.parse(localData);
            if (parsed.length > 0) {
                globalLichSuData = parsed.map(function(item, idx) {
                    return {
                        stt: idx + 1,
                        ngay: item.timeNhan ? item.timeNhan.split(' ')[1] : 'Hôm nay',
                        kho: 'HO-HCM',
                        nganhHang: 'Tổng hợp',
                        ncc: item.ncc || 'Nhà cung cấp',
                        soDinhDanh: item.code || '',
                        slKeHoach: 1,
                        slThucNhan: 1,
                        nguon: 'Tracking'
                    };
                });
                renderLichSuTable(globalLichSuData);
                return;
            }
        }
    } catch (err) {
        console.error(err);
    }

    // Nếu mọi cách đều không có dữ liệu
    if (tbody) {
        tbody.innerHTML = `<tr><td colspan="10" style="padding: 40px; color: #dc2626; text-align: center;">⚠️ Chưa có dữ liệu lịch sử nhập hàng. Bạn hãy nhập hoặc tải file Excel ở tab Kế Hoạch trước nhé!</td></tr>`;
    }
    hienThiLoading(false);
}

function locLichSuNhapHang(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var ngayVal = document.getElementById('filterLichSuNgay').value; // YYYY-MM-DD
    var thangVal = document.getElementById('filterLichSuThang').value; // YYYY-MM
    
    if (!globalLichSuData || globalLichSuData.length === 0) return;
    var filtered = globalLichSuData.filter(function(item) {
        if (!item.ngay) return false;
        var itemNgayStr = item.ngay.toString().trim();
        
        if (ngayVal) {
            var p = ngayVal.split('-'); 
            var chuanDDMMYYYY = p[2] + '-' + p[1] + '-' + p[0];
            var chuanYYYYMMDD = ngayVal;
            if (itemNgayStr !== chuanDDMMYYYY && itemNgayStr !== chuanYYYYMMDD) {
                return false;
            }
        }
        
        if (thangVal) {
            var pThang = thangVal.split('-'); 
            var yearStr = pThang[0];
            var monthStr = pThang[1];
            var chuanMMYYYY = monthStr + '-' + yearStr;
            var chuanYYYYMM = yearStr + '-' + monthStr;
            
            if (!itemNgayStr.includes(monthStr + '-' + yearStr) && 
                !itemNgayStr.includes(chuanMMYYYY) && 
                !itemNgayStr.includes(chuanYYYYMM)) {
                return false;
            }
        }
        return true;
    });
    
    renderLichSuTable(filtered);
    hienThiLoading(false);
}

function resetBoLocLichSu(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var inputNgay = document.getElementById('filterLichSuNgay');
    var inputThang = document.getElementById('filterLichSuThang');
    if (inputNgay) inputNgay.value = '';
    if (inputThang) inputThang.value = '';
    
    renderLichSuTable(globalLichSuData);
    hienThiLoading(false);
}

function renderLichSuTable(dataList,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var tbody = document.getElementById('lichSuTableBody');
    if (!tbody) return;
    
    if (!dataList || dataList.length === 0) {
        tbody.innerHTML = `<tr><td colspan="10" style="padding: 40px; color: #64748b; text-align: center;">Không có dữ liệu lịch sử phù hợp với bộ lọc.</td></tr>`;
        return;
    }
    
    var html = "";
    dataList.forEach(function(item, index) {
        var slKh = parseInt(item.slKeHoach) || 0;
        var slTn = parseInt(item.slThucNhan) || 0;
        var slConLai = slKh - slTn;
        
        html += `<tr style="border-bottom: 1px solid #e2e8f0; text-align: center;">`;
        html += `<td style="padding: 12px 10px; border-right: 1px solid #f1f5f9; font-weight: bold;">${index + 1}</td>`;
        html += `<td style="padding: 12px 10px; border-right: 1px solid #f1f5f9; background: #f8fafc; font-weight: 500;">${item.ngay || '--'}</td>`;
        html += `<td style="padding: 12px 10px; border-right: 1px solid #f1f5f9; background: #f8fafc;">${item.kho || 'HO-HCM'}</td>`;
        html += `<td style="padding: 12px 10px; border-right: 1px solid #f1f5f9;">${item.nganhHang || ''}</td>`;
        html += `<td style="padding: 12px 15px; border-right: 1px solid #f1f5f9; text-align: left; font-weight: bold; color: #1e293b;">${item.ncc || ''}</td>`;
        html += `<td style="padding: 12px 10px; border-right: 1px solid #f1f5f9;">${item.soDinhDanh || ''}</td>`;
        html += `<td style="padding: 12px 10px; border-right: 1px solid #f1f5f9; font-weight: bold;">${slKh}</td>`;
        html += `<td style="padding: 12px 10px; border-right: 1px solid #f1f5f9; font-weight: bold; color: #059669;">${slTn}</td>`;
        html += `<td style="padding: 12px 10px; border-right: 1px solid #f1f5f9; font-weight: bold; color: #dc2626;">${slConLai}</td>`;
        html += `<td style="padding: 12px 10px;"><span style="background: #f1f5f9; padding: 4px 8px; border-radius: 6px; font-weight: 600; font-size: 11px; border: 1px solid #e2e8f0;">${item.nguon || 'Hệ thống'}</span></td>`;
        html += `</tr>`;
    });
    tbody.innerHTML = html;
    hienThiLoading(false);
}

function taiFileExcelLichSu(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    if (!globalLichSuData || globalLichSuData.length === 0) {
        alert("⚠️ Không có dữ liệu lịch sử để xuất file Excel!");
        return;
    }
    var dataExport = globalLichSuData.map(function(item, idx) {
        return {
            "STT": idx + 1,
            "Ngày": item.ngay || "",
            "Kho": item.kho || "HO-HCM",
            "Ngành Hàng": item.nganhHang || "",
            "Nhà Cung Cấp": item.ncc || "",
            "Số Định Danh": item.soDinhDanh || "",
            "SL Kế Hoạch": parseInt(item.slKeHoach) || 0,
            "SL Thực Nhận": parseInt(item.slThucNhan) || 0,
            "SL Còn Lại": (parseInt(item.slKeHoach) || 0) - (parseInt(item.slThucNhan) || 0),
            "Nguồn": item.nguon || "Hệ thống"
        };
    });
    var worksheet = XLSX.utils.json_to_sheet(dataExport);
    var workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "LichSuNhapHang");
    XLSX.writeFile(workbook, "BaoCao_LichSu_NhapHang.xlsx");
    hienThiLoading(false);
}

// ==========================================
// 9. CHỨC NĂNG THEO DÕI NHẬP HÀNG & QUÉT MÃ (TRACKING)
// ==========================================
async function syncTrackingToServer(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    try {
        await fetch('http://localhost:3000/api/cap-nhat-tracking', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ trackingList: scannedTrackingList })
        });
    } catch (e) {
        console.error("Lỗi đồng bộ tracking lên server:", e);
    }
    hienThiLoading(false);
}

async function taiLaiDuLieuTracking(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var btn = document.getElementById('btnTaiLaiTracking');
    if (btn) {
        btn.innerHTML = "⏳ Đang tải...";
        btn.disabled = true;
    }
    try {
        let [resKeHoach, resTracking] = await Promise.all([
            fetch('http://localhost:3000/api/lay-ke-hoach'),
            fetch('http://localhost:3000/api/lay-tracking')
        ]);
        let dataKeHoach = await resKeHoach.json();
        let dataTracking = await resTracking.json();
        if (dataKeHoach.success) {
            globalKeHoachData = dataKeHoach.data || [];
            if (typeof renderKeHoachTable === 'function') {
                renderKeHoachTable(globalKeHoachData);
            }
        }
        if (dataTracking.success) {
            scannedTrackingList = dataTracking.data || [];
            localStorage.setItem('scannedTrackingList', JSON.stringify(scannedTrackingList));
            renderWebTrackingTable();
        }
    } catch (e) {
        console.error("Lỗi tải lại dữ liệu:", e);
        alert("⚠️ Không thể kết nối đến server để làm mới dữ liệu!");
    } finally {
        if (btn) {
            btn.innerHTML = "🔄 Tải Lại Dữ Liệu";
            btn.disabled = false;
        }
    }
    hienThiLoading(false);
}

async function xuLyThemSoDinhDanh(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var inputEl = document.getElementById('webTrackingInput');
    if (!inputEl) return;
    var val = inputEl.value.trim();
    if (!val) return;
    
    var now = new Date();
    var timeNhanStr = now.toLocaleTimeString('vi-VN') + ' ' + String(now.getDate()).padStart(2, '0') + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + now.getFullYear();
    
    var currentList = [];
    var saved = localStorage.getItem('scannedTrackingList');
    if (saved) {
        try { currentList = JSON.parse(saved); } catch(e) { currentList = []; }
    }
    
    var existingIndex = currentList.findIndex(function(item) {
        return item.code.toString().trim().toLowerCase() === val.toLowerCase();
    });
    if (existingIndex !== -1) {
        alert("⚠️ Số định danh [" + val + "] đã được quét trước đó lúc " + currentList[existingIndex].timeNhan + "!");
        inputEl.value = '';
        inputEl.focus();
        return;
    }
    
    var newItem = {
        id: Date.now(),
        code: val,
        status: "Đã nhận (Chờ khớp KH)",
        ncc: "-- Chưa chọn NCC --",
        timeNhan: timeNhanStr,
        timeNhap: "-- Chưa nhập --"
    };
    currentList.unshift(newItem);
    localStorage.setItem('scannedTrackingList', JSON.stringify(currentList));
    scannedTrackingList = currentList;
    
    renderWebTrackingTable();
    
    fetch('http://localhost:3000/api/cap-nhat-tracking', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ trackingList: currentList })
    }).catch(e => console.error(e));
    
    var msg = document.getElementById('trackingStatusMsg');
    if (msg) {
        msg.style.display = 'block';
        msg.style.background = '#dcfce7';
        msg.style.color = '#16a34a';
        msg.innerText = "✔ Đã ghi nhận thời gian nhận cho số định danh: " + val;
    }
    inputEl.value = '';
    inputEl.focus();
    hienThiLoading(false);
}

function renderWebTrackingTable(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var tbody = document.getElementById('webTrackingTableBody');
    if (!tbody) return;
    
    var saved = localStorage.getItem('scannedTrackingList');
    if (saved) {
        try { scannedTrackingList = JSON.parse(saved); } catch(e) {}
    }
    if (!scannedTrackingList || scannedTrackingList.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" style="padding: 40px; color: #64748b; text-align: center;">Chưa có số định danh nào được quét.</td></tr>`;
        return;
    }
    var html = "";
    scannedTrackingList.forEach(function(item, idx) {
        var isKhop = item.status && item.status.includes('Khớp');
        var statusBg = isKhop ? '#dcfce7' : '#fee2e2';
        var statusColor = isKhop ? '#16a34a' : '#dc2626';
        var statusText = isKhop ? '✔ Đã nhập' : '⚠️ Đã nhận';
        html += `<tr style="border-bottom: 1px solid #e2e8f0; text-align: center;">`;
        html += `<td style="padding: 12px 10px; border-right: 1px solid #f1f5f9; font-weight: bold;">${idx + 1}</td>`;
        html += `<td style="padding: 12px 15px; border-right: 1px solid #f1f5f9; text-align: left; font-weight: bold; color: #2563eb;">${item.code || ''}</td>`;
        html += `<td style="padding: 10px; border-right: 1px solid #f1f5f9;"><span style="background:${statusBg}; color:${statusColor}; padding:4px 10px; border-radius:6px; font-weight:bold; font-size:12px;">${statusText}</span></td>`;
        html += `<td style="padding: 12px 15px; border-right: 1px solid #f1f5f9; text-align: left; font-weight: 600; color: #1e293b;">${item.ncc || '-- Chưa có NCC --'}</td>`;
        html += `<td style="padding: 12px 15px; border-right: 1px solid #f1f5f9;">${item.timeNhan || ''}</td>`;
        html += `<td style="padding: 12px 15px; border-right: 1px solid #f1f5f9; color: #475569; font-weight: 500;">${item.timeNhap || ''}</td>`;
        html += `<td style="padding: 12px 10px;"><button type="button" onclick="xoaDongTracking(${item.id})" style="background: #fee2e2; color: #dc2626; border: none; padding: 6px 12px; border-radius: 6px; font-weight: bold; cursor: pointer; font-size: 12px;">Xóa</button></td>`;
        html += `</tr>`;
    });
    tbody.innerHTML = html;
    hienThiLoading(false);
}

async function xoaDongTracking(id,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    scannedTrackingList = scannedTrackingList.filter(function(item) {
        return String(item.id) !== String(id);
    });
    localStorage.setItem('scannedTrackingList', JSON.stringify(scannedTrackingList));
    await syncTrackingToServer();
    renderWebTrackingTable();
    hienThiLoading(false);
}

async function xoaTatCaTracking(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    if (confirm("⚠️ Bạn có chắc chắn muốn xóa toàn bộ lịch sử quét định danh không?")) {
        scannedTrackingList = [];
        localStorage.removeItem('scannedTrackingList');
        await syncTrackingToServer();
        renderWebTrackingTable();
        alert("🧹 Đã làm sạch toàn bộ dữ liệu theo dõi nhập hàng!");
    }
    hienThiLoading(false);
}

// ==========================================
// 10. QUẢN LÝ THU HỒI VẬN ĐƠN (TAB 5)
// ==========================================
var danhSachNVC = ["247", "Nhất Tín", "Viettel", "GNNB"];
var vanDonDiGiaoList = [];
var vanDonThuHoiList = [];

function xacDinhNhaVuyenChuyen(maVd,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    if (!maVd) return null;
    var upper = maVd.trim().toUpperCase();
    if (upper.startsWith("FV") || upper.startsWith("FO")) return "Viettel";
    if (upper.startsWith("FE") || upper.startsWith("FF")) return "247";
    if (upper.startsWith("FS") || upper.startsWith("FA") || upper.startsWith("FB")) return "Nhất Tín";
    if (upper.startsWith("FRT") || upper.startsWith("2")) return "GNNB";
    return null;
    hienThiLoading(false);
}

function switchSubThuHoiTab(tabName,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    ['scan', 'lichSu', 'caiDatNVC'].forEach(function(t) {
        var view = document.getElementById('subViewThuHoi' + (t==='scan'?'Scan':t==='lichSu'?'LichSu':'CaiDatNVC'));
        var btn = document.getElementById('subTabThuHoi' + (t==='scan'?'Scan':t==='lichSu'?'LichSu':'NVC') + 'Btn');
        if (view) view.style.display = (t === tabName) ? 'block' : 'none';
        if (btn) {
            btn.style.background = (t === tabName) ? '#2563eb' : '#f1f5f9';
            btn.style.color = (t === tabName) ? '#ffffff' : '#475569';
            btn.style.border = (t === tabName) ? 'none' : '1px solid #cbd5e1';
        }
    });
    if (tabName === 'lichSu') renderLichSuThuHoiTable();
    if (tabName === 'caiDatNVC') renderCaiDatNVCList();
    hienThiLoading(false);
}

async function xuLyScanDiGiao(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var input = document.getElementById('inputVdDiGiao');
    if (!input) return;
    var ma = input.value.trim();
    if (!ma) return;
    var nvc = xacDinhNhaVuyenChuyen(ma);
    if (!nvc) {
        alert("❌ Mã vận đơn [" + ma + "] không tồn tại hoặc không đúng định dạng quy ước!");
        input.value = '';
        input.focus();
        return;
    }
    var itemTrung = vanDonDiGiaoList.find(function(item) { return item.ma.toLowerCase() === ma.toLowerCase(); });
    if (itemTrung) {
        alert("⚠️ CẢNH BÁO: Vận đơn [" + ma + "] đã được scan đi giao trước đó!\n🕒 Thời gian: " + itemTrung.thoiGian);
        input.value = '';
        input.focus();
        return;
    }
    var now = new Date();
    var timeStr = now.toLocaleTimeString('vi-VN') + ' ' + String(now.getDate()).padStart(2, '0') + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + now.getFullYear();
    vanDonDiGiaoList.unshift({
        id: Date.now(),
        ma: ma,
        nvc: nvc,
        thoiGian: timeStr,
        trangThai: "Đang giao"
    });
    input.value = '';
    input.focus();
    renderBangThuHoi();
    await syncThuHoiToServer();
    hienThiLoading(false);
}

async function xuLyScanThuHoi(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var input = document.getElementById('inputVdThuHoi');
    if (!input) return;
    var ma = input.value.trim();
    if (!ma) return;
    var nvc = xacDinhNhaVuyenChuyen(ma);
    if (!nvc) {
        alert("❌ Mã vận đơn [" + ma + "] không tồn tại hoặc không đúng định dạng quy ước!");
        input.value = '';
        input.focus();
        return;
    }
    var itemTrungThuHoi = vanDonThuHoiList.find(function(item) { return item.ma.toLowerCase() === ma.toLowerCase(); });
    if (itemTrungThuHoi) {
        alert("⚠️ CẢNH BÁO: Vận đơn [" + ma + "] đã được scan thu hồi trước đó rồi!\n🕒 Thời gian thu hồi: " + itemTrungThuHoi.thoiGian);
        input.value = '';
        input.focus();
        return;
    }
    var now = new Date();
    var timeStr = now.toLocaleTimeString('vi-VN') + ' ' + String(now.getDate()).padStart(2, '0') + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + now.getFullYear();
    vanDonThuHoiList.unshift({
        id: Date.now(),
        ma: ma,
        nvc: nvc,
        thoiGian: timeStr
    });
    var diGiaoItem = vanDonDiGiaoList.find(function(item) { return item.ma.toLowerCase() === ma.toLowerCase(); });
    if (diGiaoItem) {
        diGiaoItem.trangThai = "Đã thu hồi";
    }
    input.value = '';
    input.focus();
    renderBangThuHoi();
    await syncThuHoiToServer();
    hienThiLoading(false);
}

async function thucHienThuHoiNhanh(maVd,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var nvc = xacDinhNhaVuyenChuyen(maVd);
    var now = new Date();
    var timeStr = now.toLocaleTimeString('vi-VN') + ' ' + String(now.getDate()).padStart(2, '0') + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + now.getFullYear();
    vanDonThuHoiList.unshift({
        id: Date.now(),
        ma: maVd,
        nvc: nvc,
        thoiGian: timeStr
    });
    renderBangThuHoi();
    await syncThuHoiToServer();
    hienThiLoading(false);
}

function renderBangThuHoi(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var tbodyDiGiao = document.getElementById('tableDiGiaoBody');
    var tbodyThuHoi = document.getElementById('tableThuHoiBody');
    var baoCaoBox = document.getElementById('baoCaoNVCBox');
    
    if (tbodyDiGiao) {
        var danhSachDiGiaoChuaThuHoi = vanDonDiGiaoList.filter(function(diGiao) {
            return !vanDonThuHoiList.some(function(thuHoi) {
                return thuHoi.ma.toLowerCase() === diGiao.ma.toLowerCase();
            });
        });
        if (danhSachDiGiaoChuaThuHoi.length === 0) {
            tbodyDiGiao.innerHTML = `<tr><td colspan="5" style="padding: 20px; text-align: center; color: #64748b;">Chưa có vận đơn đi giao cần xử lý.</td></tr>`;
        } else {
            tbodyDiGiao.innerHTML = danhSachDiGiaoChuaThuHoi.map(function(item, idx) {
                return `<tr style="border-bottom: 1px solid #e2e8f0; text-align: center;">
                    <td style="padding: 10px; font-weight: bold;">${idx + 1}</td>
                    <td style="padding: 10px; text-align: left; font-weight: bold; color: #2563eb;">${item.ma}</td>
                    <td style="padding: 10px;"><span style="background: #e0f2fe; color: #0369a1; padding: 3px 8px; border-radius: 4px; font-weight: bold;">${item.nvc}</span></td>
                    <td style="padding: 10px;">${item.thoiGian}</td>
                    <td style="padding: 10px;"><button type="button" onclick="xoaVDDiGiao(${item.id})" style="background: #fee2e2; color: #dc2626; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-weight: bold;">Xóa</button></td>
                </tr>`;
            }).join('');
        }
    }
    
    if (tbodyThuHoi) {
        var danhSachChuaThuHoi = vanDonDiGiaoList.filter(function(diGiao) {
            return !vanDonThuHoiList.some(function(thuHoi) {
                return thuHoi.ma.toLowerCase() === diGiao.ma.toLowerCase();
            });
        });
        if (danhSachChuaThuHoi.length === 0) {
            tbodyThuHoi.innerHTML = `<tr><td colspan="5" style="padding: 20px; text-align: center; color: #64748b;">Tuyệt vời! Không còn vận đơn nào cần thu hồi.</td></tr>`;
        } else {
            tbodyThuHoi.innerHTML = danhSachChuaThuHoi.map(function(item, idx) {
                return `<tr style="border-bottom: 1px solid #e2e8f0; text-align: center; background: #fffbeb;">
                    <td style="padding: 10px; font-weight: bold;">${idx + 1}</td>
                    <td style="padding: 10px; text-align: left; font-weight: bold; color: #d97706;">${item.ma}</td>
                    <td style="padding: 10px;"><span style="background: #fef3c7; color: #b45309; padding: 3px 8px; border-radius: 4px; font-weight: bold;">${item.nvc}</span></td>
                    <td style="padding: 10px;">${item.thoiGian}</td>
                    <td style="padding: 10px;">
                        <button type="button" onclick="thucHienThuHoiNhanh('${item.ma}')" style="background: #16a34a; color: white; border: none; padding: 5px 12px; border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 12px;">✔ Thu Hồi</button>
                    </td>
                </tr>`;
            }).join('');
        }
    }
    
    if (baoCaoBox) {
        var htmlBaoCao = "";
        danhSachNVC.forEach(function(nvcName) {
            var tongDiGiao = vanDonDiGiaoList.filter(function(i) { return i.nvc === nvcName; }).length;
            var tongThuHoi = vanDonThuHoiList.filter(function(i) { return i.nvc === nvcName; }).length;
            var chuaThuHoi = Math.max(0, tongDiGiao - tongThuHoi);
            var bgColor = "#ffffff"; 
            var borderColor = "#cbd5e1";
            var accentColor = "#2563eb";
            if (nvcName.toLowerCase().includes("247")) {
                bgColor = "#dcfce7"; borderColor = "#86efac"; accentColor = "#16a34a";
            } else if (nvcName.toLowerCase().includes("nhất tín") || nvcName.toLowerCase().includes("nhat tin")) {
                bgColor = "#fef3c7"; borderColor = "#fde047"; accentColor = "#d97706";
            } else if (nvcName.toLowerCase().includes("viettel")) {
                bgColor = "#ffffff"; borderColor = "#cbd5e1"; accentColor = "#2563eb";
            } else if (nvcName.toLowerCase().includes("gnnb")) {
                bgColor = "#e0f2fe"; borderColor = "#7dd3fc"; accentColor = "#0284c7";
            }
            htmlBaoCao += `
                <div style="background: ${bgColor}; padding: 18px; border-radius: 12px; border: 1px solid ${borderColor}; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                    <div style="font-size: 15px; font-weight: bold; color: #1e293b; margin-bottom: 8px;">🚚 ${nvcName}</div>
                    <div style="font-size: 24px; font-weight: bold; color: ${accentColor};">${chuaThuHoi} <span style="font-size: 13px; color: #64748b; font-weight: normal;">chưa thu hồi</span></div>
                </div>
            `;
        });
        baoCaoBox.innerHTML = htmlBaoCao;
    }
    hienThiLoading(false);
}

async function xoaVDDiGiao(id,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    vanDonDiGiaoList = vanDonDiGiaoList.filter(function(i) { return i.id !== id; });
    renderBangThuHoi();
    await syncThuHoiToServer();
    hienThiLoading(false);
}

async function xoaVDThuHoi(id,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    vanDonThuHoiList = vanDonThuHoiList.filter(function(i) { return i.id !== id; });
    renderBangThuHoi();
    await syncThuHoiToServer();
    hienThiLoading(false);
}

function renderLichSuThuHoiTable(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    var tbody = document.getElementById('tableLichSuThuHoiBody');
    if (!tbody) return;
    if (vanDonDiGiaoList.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" style="padding: 30px; text-align: center; color: #64748b;">Chưa có dữ liệu vận đơn đi giao để đối soát.</td></tr>`;
        return;
    }
    var html = "";
    vanDonDiGiaoList.forEach(function(item, idx) {
        var thuHoiMatch = vanDonThuHoiList.find(function(t) { return t.ma.toLowerCase() === item.ma.toLowerCase(); });
        var thoiGianThuHoi = thuHoiMatch ? thuHoiMatch.thoiGian : "--- Chưa thu hồi ---";
        var badgeColor = thuHoiMatch ? '#dcfce7' : '#fee2e2';
        var badgeText = thuHoiMatch ? '#16a34a' : '#dc2626';
        var statusLabel = thuHoiMatch ? '✔ Đã thu hồi' : '⏳ Chưa thu hồi';
        html += `<tr style="border-bottom: 1px solid #e2e8f0; text-align: center;">
            <td style="padding: 12px; font-weight: bold;">${idx + 1}</td>
            <td style="padding: 12px; text-align: left; font-weight: bold; color: #2563eb;">${item.ma}</td>
            <td style="padding: 12px; font-weight: 600;">${item.nvc}</td>
            <td style="padding: 12px;">${item.thoiGian}</td>
            <td style="padding: 12px; font-weight: 500; color: ${thuHoiMatch ? '#16a34a':'#dc2626'}">${thoiGianThuHoi}</td>
            <td style="padding: 12px;"><span style="background: ${badgeColor}; color: ${badgeText}; padding: 4px 10px; border-radius: 6px; font-weight: bold; font-size: 12px;">${statusLabel}</span></td>
        </tr>`;
    });
    tbody.innerHTML = html;
    hienThiLoading(false);
}

function renderCaiDatNVCList(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var listEl = document.getElementById('listNhaVuyenChuyen');
    if (!listEl) return;
    listEl.innerHTML = danhSachNVC.map(function(nvc) {
        return `<li style="display: flex; justify-content: space-between; align-items: center; padding: 10px 15px; border-bottom: 1px solid #f1f5f9; font-weight: bold;">
            <span>🚚 ${nvc}</span>
            <button type="button" onclick="xoaNhaVuyenChuyen('${nvc}')" style="background: #fee2e2; color: #dc2626; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer; font-size: 12px;">Xóa</button>
        </li>`;
    }).join('');
    hienThiLoading(false);
}

async function themNhaVuyenChuyen(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var input = document.getElementById('inputTenNVC');
    if (!input) return;
    var ten = input.value.trim();
    if (!ten) return;
    if (danhSachNVC.includes(ten)) {
        alert("⚠️ Nhà vận chuyển này đã tồn tại!");
        return;
    }
    danhSachNVC.push(ten);
    input.value = '';
    renderCaiDatNVCList();
    renderBangThuHoi();
    await syncThuHoiToServer();
    hienThiLoading(false);
}

async function xoaNhaVuyenChuyen(ten,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    if (confirm("Bạn có chắc muốn xóa nhà vận chuyển [" + ten + "] không?")) {
        danhSachNVC = danhSachNVC.filter(function(n) { return n !== ten; });
        renderCaiDatNVCList();
        renderBangThuHoi();
        await syncThuHoiToServer();
    }
    hienThiLoading(false);
}

async function syncThuHoiToServer(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    try {
        await fetch('http://localhost:3000/api/cap-nhat-thu-hoi', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                vanDonDiGiao: vanDonDiGiaoList,
                vanDonThuHoi: vanDonThuHoiList,
                nhaVuyenChuyen: danhSachNVC
            })
        });
    } catch (e) {
        console.error("Lỗi đồng bộ thu hồi:", e);
    }
    hienThiLoading(false);
}

async function loadThuHoiData(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    try {
        let res = await fetch('http://localhost:3000/api/lay-thu-hoi-van-don');
        let result = await res.json();
        if (result.success) {
            vanDonDiGiaoList = result.vanDonDiGiao || [];
            vanDonThuHoiList = result.vanDonThuHoi || [];
            if (result.nhaVuyenChuyen && result.nhaVuyenChuyen.length > 0) {
                danhSachNVC = result.nhaVuyenChuyen;
            }
            renderBangThuHoi();
        }
    } catch (e) {
        console.error("Lỗi tải dữ liệu thu hồi:", e);
    }
    hienThiLoading(false);
}

// ==========================================
// 11. QUẢN LÝ TRẢ VẬN ĐƠN CHO NVC
// ==========================================
var danhSachNVCTra = ["247", "Nhất Tín", "Viettel"];
var tamDanhSachTra = [];
var lichSuTraNVCList = [];

function switchSubTraNVCTab(tabName,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    ['scan', 'lichSu', 'caiDatNVC'].forEach(function(t) {
        var view = document.getElementById('subViewTra' + (t==='scan'?'Scan':t==='lichSu'?'LichSu':'CaiDatNVC'));
        var btn = document.getElementById('subTabTra' + (t==='scan'?'Scan':t==='lichSu'?'LichSu':'NVC') + 'Btn');
        if (view) view.style.display = (t === tabName) ? 'block' : 'none';
        if (btn) {
            btn.style.background = (t === tabName) ? '#2563eb' : '#f1f5f9';
            btn.style.color = (t === tabName) ? '#ffffff' : '#475569';
            btn.style.border = (t === tabName) ? 'none' : '1px solid #cbd5e1';
        }
    });
    if (tabName === 'lichSu') renderLichSuTraNVC();
    if (tabName === 'caiDatNVC') renderCaiDatNVCTraList();
    hienThiLoading(false);
}

function xuLyThemVDTra(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var selectNVC = document.getElementById('selectNVCTra');
    var inputVd = document.getElementById('inputVdTra');
    if (!selectNVC || !inputVd) return;
    var nvc = selectNVC.value;
    var ma = inputVd.value.trim();
    if (!nvc) {
        alert("⚠️ Vui lòng chọn Nhà vận chuyển trước khi scan!");
        return;
    }
    if (!ma) return;
    
    var itemTrungTam = tamDanhSachTra.find(function(item) { return item.ma.toLowerCase() === ma.toLowerCase(); });
    if (itemTrungTam) {
        alert("⚠️ CẢNH BÁO: Số vận đơn [" + ma + "] đã có trong danh sách chờ trả!\n🕒 Thời gian: " + itemTrungTam.thoiGian);
        inputVd.value = '';
        inputVd.focus();
        return;
    }
    
    var itemTrungLichSu = lichSuTraNVCList.find(function(item) { return item.ma.toLowerCase() === ma.toLowerCase(); });
    if (itemTrungLichSu) {
        alert("⚠️ CẢNH BÁO: Số vận đơn [" + ma + "] đã được trả cho nhà vận chuyển trước đó!\n🕒 Thời gian trả: " + itemTrungLichSu.thoiGian);
        inputVd.value = '';
        inputVd.focus();
        return;
    }
    
    var now = new Date();
    var timeStr = now.toLocaleTimeString('vi-VN') + ' ' + String(now.getDate()).padStart(2, '0') + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + now.getFullYear();
    tamDanhSachTra.unshift({
        id: Date.now() + Math.random(),
        ma: ma,
        nvc: nvc,
        thoiGian: timeStr,
        trangThai: "Đã trả NVC"
    });
    inputVd.value = '';
    inputVd.focus();
    renderBangTamTraNVC();
    hienThiLoading(false);
}

function renderBangTamTraNVC(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var tbody = document.getElementById('tableTraNVCBody');
    if (!tbody) return;
    if (tamDanhSachTra.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="padding: 30px; text-align: center; color: #64748b;">Chưa có vận đơn nào được scan.</td></tr>`;
        return;
    }
    tbody.innerHTML = tamDanhSachTra.map(function(item, idx) {
        return `<tr style="border-bottom: 1px solid #e2e8f0; text-align: center;">
            <td style="padding: 10px; font-weight: bold;">${idx + 1}</td>
            <td style="padding: 10px; text-align: left; font-weight: bold; color: #2563eb;">${item.ma}</td>
            <td style="padding: 10px;"><span style="background: #e0f2fe; color: #0369a1; padding: 3px 8px; border-radius: 4px; font-weight: bold;">${item.nvc}</span></td>
            <td style="padding: 10px;">${item.thoiGian}</td>
            <td style="padding: 10px;"><button type="button" onclick="xoaVDTamTra(${item.id})" style="background: #fee2e2; color: #dc2626; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-weight: bold;">Xóa</button></td>
        </tr>`;
    }).join('');
    hienThiLoading(false);
}

function xoaVDTamTra(id,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    tamDanhSachTra = tamDanhSachTra.filter(function(i) { return i.id !== id; });
    renderBangTamTraNVC();
    hienThiLoading(false);
}

async function luuDanhSachTraNVC(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    if (tamDanhSachTra.length === 0) {
        alert("⚠️ Không có vận đơn nào để lưu!");
        return;
    }
    tamDanhSachTra.forEach(function(item) {
        lichSuTraNVCList.unshift(item);
    });
    tamDanhSachTra = [];
    renderBangTamTraNVC();
    await syncTraNVCToServer();
    alert("✅ Đã lưu dữ liệu trả vận đơn vào database thành công!");
    hienThiLoading(false);
}

function renderSelectNVCTra(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var select = document.getElementById('selectNVCTra');
    if (!select) return;
    if (!danhSachNVCTra || danhSachNVCTra.length === 0) {
        danhSachNVCTra = ["247", "Nhất Tín", "Viettel"];
    }
    select.innerHTML = '<option value="">-- Chọn nhà vận chuyển --</option>' + danhSachNVCTra.map(function(nvc) {
        return `<option value="${nvc}">${nvc}</option>`;
    }).join('');
    hienThiLoading(false);
}

function renderLichSuTraNVC(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var tbody = document.getElementById('tableLichSuTraNVCBody');
    if (!tbody) return;
    if (lichSuTraNVCList.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="padding: 30px; text-align: center; color: #64748b;">Chưa có dữ liệu lịch sử trả vận đơn.</td></tr>`;
        return;
    }
    tbody.innerHTML = lichSuTraNVCList.map(function(item, idx) {
        return `<tr style="border-bottom: 1px solid #e2e8f0; text-align: center;">
            <td style="padding: 12px; font-weight: bold;">${idx + 1}</td>
            <td style="padding: 12px; text-align: left; font-weight: bold; color: #16a34a;">${item.ma}</td>
            <td style="padding: 12px; font-weight: 600;">${item.nvc}</td>
            <td style="padding: 12px;">${item.thoiGian}</td>
            <td style="padding: 12px;"><span style="background: #dcfce7; color: #16a34a; padding: 4px 10px; border-radius: 6px; font-weight: bold; font-size: 12px;">✔ ${item.trangThai}</span></td>
        </tr>`;
    }).join('');
    hienThiLoading(false);
}

function renderCaiDatNVCTraList(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var listEl = document.getElementById('listNVCTra');
    if (!listEl) return;
    listEl.innerHTML = danhSachNVCTra.map(function(nvc) {
        return `<li style="display: flex; justify-content: space-between; align-items: center; padding: 10px 15px; border-bottom: 1px solid #f1f5f9; font-weight: bold;">
            <span>🚚 ${nvc}</span>
            <button type="button" onclick="xoaNVCTra('${nvc}')" style="background: #fee2e2; color: #dc2626; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer; font-size: 12px;">Xóa</button>
        </li>`;
    }).join('');
    hienThiLoading(false);
}

async function themNVCTra(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var input = document.getElementById('inputTenNVCTra');
    if (!input) return;
    var ten = input.value.trim();
    if (!ten) return;
    if (danhSachNVCTra.includes(ten)) {
        alert("⚠️ Nhà vận chuyển này đã tồn tại!");
        return;
    }
    danhSachNVCTra.push(ten);
    input.value = '';
    renderCaiDatNVCTraList();
    renderSelectNVCTra();
    await syncTraNVCToServer();
    hienThiLoading(false);
}

async function xoaNVCTra(ten,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    if (confirm("Bạn có chắc muốn xóa nhà vận chuyển [" + ten + "] không?")) {
        danhSachNVCTra = danhSachNVCTra.filter(function(n) { return n !== ten; });
        renderCaiDatNVCTraList();
        renderSelectNVCTra();
        await syncTraNVCToServer();
    }
    hienThiLoading(false);
}

async function syncTraNVCToServer(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    try {
        await fetch('http://localhost:3000/api/cap-nhat-tra-van-don', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                danhSachTraNVC: lichSuTraNVCList,
                nhaVuyenChuyenTra: danhSachNVCTra
            })
        });
    } catch (e) {
        console.error("Lỗi đồng bộ trả vận đơn:", e);
    }
    hienThiLoading(false);
}

async function loadTraVNDData(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);             // Bật icon xoay tròn
    try {
        let res = await fetch('http://localhost:3000/api/lay-tra-van-don');
        let result = await res.json();
        if (result.success) {
            lichSuTraNVCList = result.danhSachTraNVC || [];
            if (result.nhaVuyenChuyenTra && result.nhaVuyenChuyenTra.length > 0) {
                danhSachNVCTra = result.nhaVuyenChuyenTra;
            }
            renderSelectNVCTra();
        }
    } catch (e) {
        console.error("Lỗi tải dữ liệu trả vận đơn:", e);
        renderSelectNVCTra();
    }
    hienThiLoading(false);
}

// ==========================================
// 12. CHẤM CÔNG & QUẢN LÝ NHÂN SỰ THỜI VỤ
// ==========================================
var nhaCungUngThoiVuList = [
    { ten: "Hoa Anh Đào", donGia: 42500 },
    { ten: "Flex", donGia: 43000 },
    { ten: "Topaz", donGia: 60000 }
];
var chamCongThoiVuList = [];

function switchSubThoiVuTab(tabName,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    ['nhapLieu', 'baoCao', 'caiDat'].forEach(function(t) {
        var view = document.getElementById('subViewThoiVu' + (t==='nhapLieu'?'NhapLieu':t==='baoCao'?'BaoCao':'CaiDat'));
        var btn = document.getElementById('subTabThoiVu' + (t==='nhapLieu'?'NhapLieu':t==='baoCao'?'BaoCao':'CaiDat') + 'Btn');
        if (view) view.style.display = (t === tabName) ? 'block' : 'none';
        if (btn) {
            btn.style.background = (t === tabName) ? '#2563eb' : '#f1f5f9';
            btn.style.color = (t === tabName) ? '#ffffff' : '#475569';
            btn.style.border = (t === tabName) ? 'none' : '1px solid #cbd5e1';
        }
    });
    if (tabName === 'nhapLieu') {
        renderBangNhapLieuThoiVu();
        var ngayInput = document.getElementById('inputThoiVuNgay');
        if (ngayInput) {
            ngayInput.valueAsDate = new Date();
        }
    }
    if (tabName === 'baoCao') renderBaoCaoThoiVu();
    if (tabName === 'caiDat') renderCaiDatNCCThoiVu();
    hienThiLoading(false);
}

function renderBangNhapLieuThoiVu(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var tbody = document.getElementById('tableNhapLieuThoiVuBody');
    if (!tbody) return;
    var ngayInput = document.getElementById('inputThoiVuNgay');
    if (ngayInput && !ngayInput.value) {
        ngayInput.valueAsDate = new Date();
    }
    if (nhaCungUngThoiVuList.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="padding: 30px; text-align: center; color: #64748b;">Chưa có nhà cung ứng nào.</td></tr>`;
        return;
    }
    tbody.innerHTML = nhaCungUngThoiVuList.map(function(ncc, idx) {
        return `<tr style="border-bottom: 1px solid #e2e8f0; text-align: center;">
            <td style="padding: 12px; text-align: left; font-weight: bold; color: #1e293b;">🚚 ${ncc.ten}</td>
            <td style="padding: 12px; font-weight: 600; color: #0284c7;">${ncc.donGia.toLocaleString('vi-VN')} VNĐ</td>
            <td style="padding: 12px;">
                <input type="number" id="sl_ns_${idx}" value="0" min="0" oninput="tinhThanhTienThoiVu(${idx}, ${ncc.donGia})" style="width: 100px; padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px; text-align: center; font-weight: bold;">
            </td>
            <td style="padding: 12px;">
                <input type="number" id="so_gio_${idx}" value="8" min="0" step="0.5" oninput="tinhThanhTienThoiVu(${idx}, ${ncc.donGia})" style="width: 100px; padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px; text-align: center; font-weight: bold;">
            </td>
            <td style="padding: 12px; text-align: right; font-weight: bold; color: #16a34a; font-size: 15px;" id="thanh_tien_${idx}">0 VNĐ</td>
        </tr>`;
    }).join('');
    hienThiLoading(false);
}

function tinhThanhTienThoiVu(idx, donGia,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var sl = parseFloat(document.getElementById('sl_ns_' + idx).value) || 0;
    var gio = parseFloat(document.getElementById('so_gio_' + idx).value) || 0;
    var thanhTien = sl * gio * donGia;
    var el = document.getElementById('thanh_tien_' + idx);
    if (el) {
        el.innerText = thanhTien.toLocaleString('vi-VN') + " VNĐ";
    }
    hienThiLoading(false);
}

async function luuChamCongThoiVu(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var ngayInput = document.getElementById('inputThoiVuNgay');
    if (!ngayInput || !ngayInput.value) {
        alert("⚠️ Vui lòng chọn ngày chấm công!");
        return;
    }
    var ngayChon = ngayInput.value;
    nhaCungUngThoiVuList.forEach(function(ncc, idx) {
        var sl = parseFloat(document.getElementById('sl_ns_' + idx).value) || 0;
        var gio = parseFloat(document.getElementById('so_gio_' + idx).value) || 0;
        var thanhTien = sl * gio * ncc.donGia;
        if (sl > 0) {
            var existingIndex = chamCongThoiVuList.findIndex(function(item) {
                return item.ngay === ngayChon && item.nhaCungUng === ncc.ten;
            });
            var record = {
                id: existingIndex >= 0 ? chamCongThoiVuList[existingIndex].id : Date.now() + idx,
                ngay: ngayChon,
                nhaCungUng: ncc.ten,
                donGia: ncc.donGia,
                soLuongNS: sl,
                soGio: gio,
                thanhTien: thanhTien
            };
            if (existingIndex !== -1) {
                chamCongThoiVuList[existingIndex] = record;
            } else {
                chamCongThoiVuList.unshift(record);
            }
        }
    });
    await syncThoiVuToServer();
    alert("✅ Đã lưu và cập nhật dữ liệu chấm công thời vụ thành công!");
    switchSubThoiVuTab('baoCao');
    hienThiLoading(false);
}

function renderBaoCaoThoiVu(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var container = document.getElementById('subViewThoiVuBaoCao');
    if (!container) return;
    var now = new Date();
    var currentMonthYear = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
    
    if (!document.getElementById('filterThoiVuThang')) {
        container.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; flex-wrap: wrap; gap: 15px;">
                <h3 style="margin: 0; color: #1e293b;">📊 Báo Cáo Chi Phí Thời Vụ & Tổng Hợp</h3>
                <div style="display: flex; gap: 12px; align-items: center;">
                    <div>
                        <label style="font-weight: 600; font-size: 13px; margin-right: 8px;">📅 Lọc Theo Tháng:</label>
                        <input type="month" id="filterThoiVuThang" value="${currentMonthYear}" onchange="renderBaoCaoThoiVu()" style="padding: 8px 12px; border: 1px solid #cbd5e1; border-radius: 6px; font-weight: bold; background: white;">
                    </div>
                    <button type="button" onclick="taiExcelBaoCaoThoiVu()" style="padding: 9px 18px; background: #16a34a; color: white; border: none; border-radius: 6px; font-weight: bold; cursor: pointer;">
                        📥 Tải Excel Báo Cáo
                    </button>
                </div>
            </div>
            <div style="margin-bottom: 25px;">
                <h4 style="color: #334155; margin-bottom: 10px;">📌 Tổng Hợp Chi Phí Theo Nhà Cung Ứng Trong Tháng</h4>
                <div style="overflow-x: auto; border: 1px solid #cbd5e1; border-radius: 8px; background: #f8fafc;">
                    <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                        <thead>
                            <tr style="background: #0f172a; color: white; text-align: center;">
                                <th style="padding: 12px; text-align: left;">NHÀ CUNG ỨNG</th>
                                <th style="padding: 12px;">TỔNG SỐ LƯỢNG NS</th>
                                <th style="padding: 12px;">TỔNG SỐ GIỜ LÀM</th>
                                <th style="padding: 12px; text-align: right; color: #4ade80;">TỔNG TIỀN TRONG THÁNG (VNĐ)</th>
                            </tr>
                        </thead>
                        <tbody id="tableTongHopNCCBody"></tbody>
                    </table>
                </div>
            </div>
            <div>
                <h4 style="color: #334155; margin-bottom: 10px;">📋 Chi Tiết Chấm Công Hàng Ngày</h4>
                <div style="overflow-x: auto; border: 1px solid #cbd5e1; border-radius: 8px; background: white;">
                    <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                        <thead>
                            <tr style="background: #1e293b; color: white; text-align: center;">
                                <th style="padding: 12px;">STT</th>
                                <th style="padding: 12px;">NGÀY</th>
                                <th style="padding: 12px; text-align: left;">NHÀ CUNG ỨNG</th>
                                <th style="padding: 12px;">ĐƠN GIÁ</th>
                                <th style="padding: 12px;">SỐ LƯỢNG NS</th>
                                <th style="padding: 12px;">SỐ GIỜ/NGƯỜI</th>
                                <th style="padding: 12px; text-align: right; color: #4ade80;">THÀNH TIỀN (VNĐ)</th>
                                <th style="padding: 12px;">THAO TÁC</th>
                            </tr>
                        </thead>
                        <tbody id="tableBaoCaoThoiVuBody"></tbody>
                    </table>
                </div>
            </div>
        `;
    }
    
    var filterInput = document.getElementById('filterThoiVuThang');
    var selectedMonth = filterInput ? filterInput.value : currentMonthYear;
    var filteredList = chamCongThoiVuList.filter(function(item) {
        if (!item.ngay) return false;
        return item.ngay.startsWith(selectedMonth);
    });
    
    var tbodyChiTiet = document.getElementById('tableBaoCaoThoiVuBody');
    if (tbodyChiTiet) {
        if (filteredList.length === 0) {
            tbodyChiTiet.innerHTML = `<tr><td colspan="8" style="padding: 30px; text-align: center; color: #64748b;">Không có dữ liệu chấm công trong tháng ${selectedMonth}.</td></tr>`;
        } else {
            tbodyChiTiet.innerHTML = filteredList.map(function(item, idx) {
                return `<tr style="border-bottom: 1px solid #e2e8f0; text-align: center;">
                    <td style="padding: 12px; font-weight: bold;">${idx + 1}</td>
                    <td style="padding: 12px; font-weight: 600;">${item.ngay}</td>
                    <td style="padding: 12px; text-align: left; font-weight: bold; color: #2563eb;">${item.nhaCungUng}</td>
                    <td style="padding: 12px;">${item.donGia.toLocaleString('vi-VN')} đ</td>
                    <td style="padding: 12px; font-weight: bold;">${item.soLuongNS}</td>
                    <td style="padding: 12px;">${item.soGio} h</td>
                    <td style="padding: 12px; text-align: right; font-weight: bold; color: #16a34a; font-size: 14px;">${item.thanhTien.toLocaleString('vi-VN')} VNĐ</td>
                    <td style="padding: 12px;"><button type="button" onclick="xoaBanGhiChamCong(${item.id})" style="background: #fee2e2; color: #dc2626; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-weight: bold;">Xóa</button></td>
                </tr>`;
            }).join('');
        }
    }
    
    var tbodyTongHop = document.getElementById('tableTongHopNCCBody');
    if (tbodyTongHop) {
        var tongHopMap = {};
        nhaCungUngThoiVuList.forEach(function(ncc) {
            tongHopMap[ncc.ten] = { tongNS: 0, tongGio: 0, tongTien: 0 };
        });
        filteredList.forEach(function(item) {
            if (!tongHopMap[item.nhaCungUng]) {
                tongHopMap[item.nhaCungUng] = { tongNS: 0, tongGio: 0, tongTien: 0 };
            }
            tongHopMap[item.nhaCungUng].tongNS += Number(item.soLuongNS) || 0;
            tongHopMap[item.nhaCungUng].tongGio += (Number(item.soLuongNS) || 0) * (Number(item.soGio) || 0);
            tongHopMap[item.nhaCungUng].tongTien += Number(item.thanhTien) || 0;
        });
        var tongHopHTML = "";
        var totalAllMoney = 0;
        Object.keys(tongHopMap).forEach(function(tenNCC) {
            var data = tongHopMap[tenNCC];
            totalAllMoney += data.tongTien;
            tongHopHTML += `
                <tr style="border-bottom: 1px solid #cbd5e1; text-align: center; background: white;">
                    <td style="padding: 12px; text-align: left; font-weight: bold; color: #1e293b;">🚚 ${tenNCC}</td>
                    <td style="padding: 12px; font-weight: bold; color: #0284c7;">${data.tongNS.toLocaleString('vi-VN')} lượt NS</td>
                    <td style="padding: 12px; font-weight: bold;">${data.tongGio.toLocaleString('vi-VN')} giờ</td>
                    <td style="padding: 12px; text-align: right; font-weight: bold; color: #16a34a; font-size: 15px;">${data.tongTien.toLocaleString('vi-VN')} VNĐ</td>
                </tr>
            `;
        });
        tongHopHTML += `
            <tr style="background: #f1f5f9; font-weight: bold; text-align: center;">
                <td style="padding: 12px; text-align: left; color: #dc2626;">💥 TỔNG CỘNG TOÀN BỘ CHI PHÍ THỜI VỤ</td>
                <td colspan="2" style="padding: 12px;">---</td>
                <td style="padding: 12px; text-align: right; color: #dc2626; font-size: 16px;">${totalAllMoney.toLocaleString('vi-VN')} VNĐ</td>
            </tr>
        `;
        tbodyTongHop.innerHTML = tongHopHTML;
    }
hienThiLoading(false);
}

function taiExcelBaoCaoThoiVu(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var filterInput = document.getElementById('filterThoiVuThang');
    var selectedMonth = filterInput ? filterInput.value : "ALL";
    var filteredList = chamCongThoiVuList.filter(function(item) {
        if (!selectedMonth) return true;
        return item.ngay.startsWith(selectedMonth);
    });
    if (filteredList.length === 0) {
        alert("⚠️ Không có dữ liệu trong tháng này để xuất Excel!");
        return;
    }
    var dataToExport = filteredList.map(function(item, idx) {
        return {
            "STT": idx + 1,
            "Ngày Chấm Công": item.ngay,
            "Nhà Cung Ứng": item.nhaCungUng,
            "Đơn Giá (VNĐ/Giờ)": item.donGia,
            "Số Lượng Nhân Sự": item.soLuongNS,
            "Số Giờ Làm/Người": item.soGio,
            "Thành Tiền (VNĐ)": item.thanhTien
        };
    });
    var worksheet = XLSX.utils.json_to_sheet(dataToExport);
    var workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "BaoCaoThoiVu");
    XLSX.writeFile(workbook, "BaoCao_ChiPhi_ThoiVu_" + selectedMonth + ".xlsx");
    hienThiLoading(false);
}

async function xoaBanGhiChamCong(id,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    if (confirm("Bạn có chắc muốn xóa bản ghi chấm công này không?")) {
        chamCongThoiVuList = chamCongThoiVuList.filter(function(i) { return i.id !== id; });
        renderBaoCaoThoiVu();
        await syncThoiVuToServer();
    }
    hienThiLoading(false);
}

function renderCaiDatNCCThoiVu(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var listEl = document.getElementById('listNCCThoiVu');
    if (!listEl) return;
    listEl.innerHTML = nhaCungUngThoiVuList.map(function(ncc) {
        return `<li style="display: flex; justify-content: space-between; align-items: center; padding: 12px 15px; border-bottom: 1px solid #f1f5f9; font-weight: bold;">
            <span>🚚 ${ncc.ten} <span style="color: #0284c7; font-size: 12px; margin-left: 10px;">(${ncc.donGia.toLocaleString('vi-VN')} VNĐ/giờ)</span></span>
            <button type="button" onclick="xoaNCCThoiVu('${ncc.ten}')" style="background: #fee2e2; color: #dc2626; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer; font-size: 12px;">Xóa</button>
        </li>`;
    }).join('');
    hienThiLoading(false);
}

async function themNhaCungUngThoiVu(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var inputTen = document.getElementById('inputTenNCC');
    var inputGia = document.getElementById('inputDonGiaNCC');
    if (!inputTen || !inputGia) return;
    var ten = inputTen.value.trim();
    var gia = parseFloat(inputGia.value);
    if (!ten || isNaN(gia)) {
        alert("⚠️ Vui lòng nhập đầy đủ tên nhà cung ứng và đơn giá hợp lệ!");
        return;
    }
    var isExist = nhaCungUngThoiVuList.some(function(n) { return n.ten.toLowerCase() === ten.toLowerCase(); });
    if (isExist) {
        alert("⚠️ Nhà cung ứng này đã tồn tại!");
        return;
    }
    nhaCungUngThoiVuList.push({ ten: ten, donGia: gia });
    inputTen.value = '';
    inputGia.value = '';
    renderCaiDatNCCThoiVu();
    renderBangNhapLieuThoiVu();
    await syncThoiVuToServer();
    hienThiLoading(false);
}

async function xoaNCCThoiVu(ten,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    if (confirm("Bạn có chắc muốn xóa nhà cung ứng [" + ten + "] không?")) {
        nhaCungUngThoiVuList = nhaCungUngThoiVuList.filter(function(n) { return n.ten !== ten; });
        renderCaiDatNCCThoiVu();
        renderBangNhapLieuThoiVu();
        await syncThoiVuToServer();
    }
    hienThiLoading(false);
}

async function syncThoiVuToServer(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    try {
        await fetch('http://localhost:3000/api/cap-nhat-cham-cong-thoi-vu', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                chamCongList: chamCongThoiVuList,
                nhaCungUngList: nhaCungUngThoiVuList
            })
        });
    } catch (e) {
        console.error("Lỗi đồng bộ chấm công thời vụ:", e);
    }
    hienThiLoading(false);
}

async function loadThoiVuData(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    try {
        let res = await fetch('http://localhost:3000/api/lay-cham-cong-thoi-vu');
        let result = await res.json();
        if (result.success) {
            chamCongThoiVuList = result.chamCongList || [];
            if (result.nhaCungUngList && result.nhaCungUngList.length > 0) {
                nhaCungUngThoiVuList = result.nhaCungUngList;
            }
            renderBangNhapLieuThoiVu();
        }
    } catch (e) {
        console.error("Lỗi tải dữ liệu chấm công thời vụ:", e);
    }
    hienThiLoading(false);
}

// ==========================================
// QUẢN LÝ BIÊN BẢN KIỂM HÀNG & IN ẤN
// ==========================================
function khoiTaoBienBan(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    if (bienBanTimer) clearInterval(bienBanTimer);
    capNhatThoiGianBienBan();
    bienBanTimer = setInterval(capNhatThoiGianBienBan, 1000);
    renderBangBienBan();
    hienThiLoading(false);
}

function capNhatThoiGianBienBan(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var now = new Date();
    var gio = String(now.getHours()).padStart(2, '0');
    var phut = String(now.getMinutes()).padStart(2, '0');
    var giay = String(now.getSeconds()).padStart(2, '0');
    
    var ngay = String(now.getDate()).padStart(2, '0');
    var thang = String(now.getMonth() + 1).padStart(2, '0');
    var nam = now.getFullYear();
    
    var thoiGianStr = gio + ":" + phut + ":" + giay + " ngày " + ngay + "-" + thang + "-" + nam;
    var el = document.getElementById('bbThoiGianLap');
    if (el) {
        el.innerText = "Thời gian lập: " + thoiGianStr;
    }
    hienThiLoading(false);
}

function renderBangBienBan(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var tbody = document.getElementById('tableBienBanBody');
    if (!tbody) return;
    
    if (danhSachHangBienBan.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" style="padding: 20px; text-align: center; color: #64748b;">Chưa có sản phẩm lỗi. Nhấn nút [+] Thêm Sản Phẩm Lỗi ở bên dưới.</td></tr>`;
        return;
    }
    
    var html = "";
    var tongSoLuong = 0;
    
    danhSachHangBienBan.forEach(function(item, idx) {
        var sl = parseInt(item.soLuong) || 1;
        tongSoLuong += sl;
        
        html += `
            <tr style="border-bottom: 1px solid #cbd5e1; text-align: center;">
                <td style="padding: 8px; border: 1px solid #cbd5e1; font-weight: bold;">${idx + 1}</td>
                <td style="padding: 6px; border: 1px solid #cbd5e1;">
                    <input type="text" value="${item.maHang || ''}" oninput="capNhatItemBienBan(${idx}, 'maHang', this.value)" style="width: 100%; padding: 6px; border: 1px solid #cbd5e1; border-radius: 4px; text-align: left; font-size: 13px;" placeholder="Mã hàng...">
                </td>
                <td style="padding: 6px; border: 1px solid #cbd5e1;">
                    <input type="text" value="${item.tenHang || ''}" oninput="capNhatItemBienBan(${idx}, 'tenHang', this.value)" style="width: 100%; padding: 6px; border: 1px solid #cbd5e1; border-radius: 4px; text-align: left; font-size: 13px;" placeholder="Tên hàng...">
                </td>
                <td style="padding: 6px; border: 1px solid #cbd5e1;">
                    <input type="text" value="${item.imei || ''}" oninput="capNhatItemBienBan(${idx}, 'imei', this.value)" style="width: 100%; padding: 6px; border: 1px solid #cbd5e1; border-radius: 4px; text-align: left; font-size: 13px;" placeholder="IMEI / Serial...">
                </td>
                <td style="padding: 6px; border: 1px solid #cbd5e1;">
                    <input type="number" value="${sl}" min="1" oninput="capNhatItemBienBan(${idx}, 'soLuong', this.value)" style="width: 60px; padding: 6px; border: 1px solid #cbd5e1; border-radius: 4px; text-align: center; font-weight: bold; font-size: 13px;">
                </td>
                <td style="padding: 6px; border: 1px solid #cbd5e1;">
                    <input type="text" value="${item.tinhTrang || 'Móp méo / Không đạt'}" oninput="capNhatItemBienBan(${idx}, 'tinhTrang', this.value)" style="width: 100%; padding: 6px; border: 1px solid #cbd5e1; border-radius: 4px; text-align: left; font-size: 13px;">
                </td>
                <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center;">
                    <button type="button" onclick="xoaHangBienBan(${idx})" style="background: #fee2e2; color: #dc2626; border: none; padding: 6px 10px; border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 12px;">Xóa</button>
                </td>
            </tr>
        `;
    });
    
    tbody.innerHTML = html;
    hienThiLoading(false);
}

function capNhatItemBienBan(idx, field, val,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    danhSachHangBienBan[idx][field] = val;
    hienThiLoading(false);
}

function themDongBienBan(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    danhSachHangBienBan.push({
        id: Date.now(),
        maHang: "",
        tenHang: "",
        imei: "",
        soLuong: 1,
        tinhTrang: "Móp méo / Không đạt"
    });
    renderBangBienBan();
    hienThiLoading(false);
}

function xoaHangBienBan(index,event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    danhSachHangBienBan.splice(index, 1);
    renderBangBienBan();
    hienThiLoading(false);
}

function inBienBan(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var containerBienBan = document.getElementById('subViewBienBan');
    if (!containerBienBan) {
        window.print();
        return;
    }
    var originalHTML = containerBienBan.innerHTML;
    
    var nguoiGiaoVal = document.getElementById('bbNguoiGiao') ? document.getElementById('bbNguoiGiao').value : '';
    var daiDienNCCVal = document.getElementById('bbDaiDienNCC') ? document.getElementById('bbDaiDienNCC').value : '';
    var hoaDonVal = document.getElementById('bbHoaDon') ? document.getElementById('bbHoaDon').value : '';
    var vanDonVal = document.getElementById('bbVanDon') ? document.getElementById('bbVanDon').value : '';
    var nhanVienNhanVal = document.getElementById('bbNhanVienNhan') ? document.getElementById('bbNhanVienNhan').value : 'NGUYỄN MINH THẢO';
    var daiDienDvVal = document.getElementById('bbDaiDienDV') ? document.getElementById('bbDaiDienDV').value : 'Kho Tổng FPT SHOP Miền Nam';
    var diaChiKhoVal = document.getElementById('bbDiaChiKho') ? document.getElementById('bbDiaChiKho').value : 'Lô L1, đường N2, KCN Hựu Thạnh, Xã Đức Hòa, tỉnh Tây Ninh';
    var sdtVal = document.getElementById('bbSdt') ? document.getElementById('bbSdt').value : '0909427775';
    
    var gridContainer = containerBienBan.querySelector('.bb-grid-container');
    if (gridContainer) {
        gridContainer.innerHTML = `
            <div style="display: flex; flex-direction: column; gap: 8px; font-size: 13px;">
                <div><b>A/ Bên bàn giao:</b> <span style="font-weight: bold; border-bottom: 1px dotted #000; padding-bottom: 2px;">${nguoiGiaoVal || '---'}</span></div>
                <div><b>Đại diện NCC:</b> <span style="font-weight: bold; border-bottom: 1px dotted #000; padding-bottom: 2px;">${daiDienNCCVal || '---'}</span></div>
                <div><b>Hóa đơn số:</b> <span style="font-weight: bold; border-bottom: 1px dotted #000; padding-bottom: 2px;">${hoaDonVal || '---'}</span></div>
                <div><b>Vận đơn số:</b> <span style="font-weight: bold; border-bottom: 1px dotted #000; padding-bottom: 2px;">${vanDonVal || '---'}</span></div>
            </div>
            <div style="display: flex; flex-direction: column; gap: 8px; font-size: 13px;">
                <div><b>B/ Bên nhận bàn giao:</b> <span style="font-weight: bold;">${nhanVienNhanVal}</span></div>
                <div><b>Đại diện đơn vị:</b> <span style="font-weight: bold;">${daiDienDvVal}</span></div>
                <div><b>Địa chỉ kho:</b> <span style="font-weight: bold;">${diaChiKhoVal}</span></div>
                <div><b>SĐT liên hệ:</b> <span style="font-weight: bold;">${sdtVal}</span></div>
            </div>
        `;
    }
    
    var tbody = document.getElementById('tableBienBanBody');
    if (tbody && typeof danhSachHangBienBan !== 'undefined') {
        var tongSoLuong = 0;
        var rowsHTML = danhSachHangBienBan.map(function(item, idx) {
            var sl = parseInt(item.soLuong) || 1;
            tongSoLuong += sl;
            return `<tr style="border-bottom: 1px solid #000; text-align: center;">
                <td style="padding: 6px; border: 1.5px solid #000; font-weight: bold;">${idx + 1}</td>
                <td style="padding: 6px; border: 1.5px solid #000; text-align: left; font-weight: bold;">${item.maHang || ''}</td>
                <td style="padding: 6px; border: 1.5px solid #000; text-align: left; font-weight: bold;">${item.tenHang || ''}</td>
                <td style="padding: 6px; border: 1.5px solid #000; text-align: left; font-weight: bold;">${item.imei || ''}</td>
                <td style="padding: 6px; border: 1.5px solid #000; font-weight: bold;">${sl}</td>
                <td style="padding: 6px; border: 1.5px solid #000; text-align: left; font-weight: bold;">${item.tinhTrang || ''}</td>
            </tr>`;
        }).join('');
        rowsHTML += `
            <tr style="border: 1.5px solid #000; font-weight: bold; background: #f8fafc;">
                <td colspan="4" style="padding: 6px; border: 1.5px solid #000; text-align: right;">TỔNG CỘNG</td>
                <td style="padding: 6px; border: 1.5px solid #000; text-align: center;">${tongSoLuong}</td>
                <td style="padding: 6px; border: 1.5px solid #000;"></td>
            </tr>
        `;
        tbody.innerHTML = rowsHTML;
    }
    
    window.print();
    
    setTimeout(function() {
        containerBienBan.innerHTML = originalHTML;
        if (typeof renderBangBienBan === 'function') {
            renderBangBienBan();
        }
    }, 500);
    hienThiLoading(false);
}

function taiLaiBienBan(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    var soPhieuInput = document.getElementById('bbSoPhieu');
    if (soPhieuInput) {
        var currentSoPhieu = soPhieuInput.value;
        var match = currentSoPhieu.match(/^(.*?)(\d+)$/);
        if (match) {
            var prefix = match[1];
            var numStr = match[2];
            var nextNum = parseInt(numStr, 10) + 1;
            var paddedNum = String(nextNum).padStart(numStr.length, '0');
            soPhieuInput.value = prefix + paddedNum;
        } else {
            soPhieuInput.value = currentSoPhieu + '-1';
        }
    }
    
    if(document.getElementById('bbNguoiGiao')) document.getElementById('bbNguoiGiao').value = '';
    if(document.getElementById('bbDaiDienNCC')) document.getElementById('bbDaiDienNCC').value = '';
    if(document.getElementById('bbHoaDon')) document.getElementById('bbHoaDon').value = '';
    if(document.getElementById('bbVanDon')) document.getElementById('bbVanDon').value = '';
    
    danhSachHangBienBan = [
        { id: 1, maHang: "", tenHang: "", imei: "", soLuong: 1, tinhTrang: "Móp méo / Không đạt" }
    ];
    renderBangBienBan();
    capNhatThoiGianBienBan();
    hienThiLoading(false);
}

function lamMoiBienBan() {
    taiLaiBienBan();
}
// ==========================================
// CHỨC NĂNG GỬI EMAIL BIÊN BẢN QUA OUTLOOK (CÓ LOẠI TRỪ CC VÀ HTML CHUẨN)
// ==========================================
async function guiEmailBienBan(event) {
    if (event) event.preventDefault(); // Chặn hành vi submit/reload ngầm của form
    hienThiLoading(true);
    // 1. Lấy thông tin từ giao diện biên bản kiểm hàng
    var soPhieuInput = document.getElementById('bienBanSoPhieu') || document.getElementById('soPhieu');
    var hoaDonInput = document.getElementById('bienBanHoaDon') || document.getElementById('hoaDon');
    var vanDonInput = document.getElementById('bienBanVanDon') || document.getElementById('vanDon');
    var taiXeInput = document.getElementById('bienBanTaiXe') || document.getElementById('nguoiGiao');
    var nccInput = document.getElementById('bienBanNCC') || document.getElementById('nccBienBan');
    var sdtKhoInput = document.getElementById('bienBanSDTKho') || document.getElementById('sdtKho') || { value: '0900000000' };

    var soPhieu = soPhieuInput ? soPhieuInput.value : ('BB-' + Date.now());
    var hoaDon = hoaDonInput ? hoaDonInput.value : 'N/A';
    var vanDon = vanDonInput ? vanDonInput.value : 'N/A';
    var taiXe = taiXeInput ? taiXeInput.value : 'Bên giao hàng';
    var ncc = nccInput ? nccInput.value : 'Nhà Cung Cấp';
    var sdtKho = sdtKhoInput ? sdtKhoInput.value : '0900000000';
    var thoiGianHienTai = new Date().toLocaleString('vi-VN');

    // Lấy danh sách hàng hóa lỗi (danhSachHangBienBan)
    var danhSachHang = typeof danhSachHangBienBan !== 'undefined' ? danhSachHangBienBan : [];
    if (!danhSachHang || danhSachHang.length === 0) {
        alert("⚠️ Biên bản chưa có hàng hóa lỗi để gửi email!");
        return;
    }

    // Format danh sách sản phẩm thành text chi tiết cho thẻ <pre>
    var danhSachSpText = danhSachHang.map(function(item, idx) {
        return `[${idx + 1}] Mã: ${item.maHang || ''} | Tên: ${item.tenHang || ''} | IMEI: ${item.imei || 'Không có'} | SL: ${item.soLuong || 1} | Lỗi: ${item.tinhTrang || ''}`;
    }).join('\n');

    // Gom dữ liệu vào object data theo đúng yêu cầu của bạn
    var data = {
        thoiGian: thoiGianHienTai,
        hoaDon: hoaDon,
        vanDon: vanDon,
        taiXe: taiXe,
        ncc: ncc,
        sdtKho: sdtKho,
        danhSachSpText: danhSachSpText
    };

    // 2. Danh sách 4 tài khoản CC mặc định
    var danhSachCC = ['namnvh@fpt.com', 'thongnm6@fpt.com', 'phucvd6@fpt.com', 'thaonm2@fpt.com'];

    // Yêu cầu người gửi nhập email Outlook đang dùng để tiến hành loại trừ khỏi CC
    var nguoiGuiHienTai = prompt("Vui lòng nhập email Outlook bạn đang sử dụng để gửi biên bản (để hệ thống tự động loại trừ khỏi danh sách CC):", "namnvh@fpt.com");
    if (!nguoiGuiHienTai) return; // Hủy nếu không nhập
    nguoiGuiHienTai = nguoiGuiHienTai.trim().toLowerCase();

    // Loại trừ email của người gửi khỏi danh sách CC nếu trùng
    var ccFiltered = danhSachCC.filter(function(email) {
        return email.toLowerCase() !== nguoiGuiHienTai;
    });

    // 3. Xây dựng nội dung HTML chuẩn theo yêu cầu
    var emailBodyHTML = 
        "Chào anh/chị,<br><br>" +
        "Kho Tổng Miền Nam xin gửi thông tin kiểm tra hàng hóa không đạt điều kiện nhập kho:<br>" +
        "--------------------------------------------------<br>" +
        "• <b>Thời gian lập:</b> " + data.thoiGian + "<br>" +
        "• <b>Hóa đơn số:</b> " + data.hoaDon + "<br>" +
        "• <b>Vận đơn số:</b> " + data.vanDon + "<br>" +
        "• <b>Bên bàn giao:</b> " + data.taiXe + " (" + data.ncc + ")<br>" +
        "• <b>SĐT liên hệ kho:</b> " + data.sdtKho + "<br><br>" +
        "<b>CHI TIẾT HÀNG HƯ HỎNG / LỖI:</b><br>" +
        "<pre style='font-family: Arial, sans-serif; font-size: 13px;'>" + data.danhSachSpText + "</pre><br>" +
        "<span style='color: #e30000; font-weight: bold;'>NHỜ ANH CHỊ CHO KHO HƯỚNG XỬ LÍ TIẾP THEO</span><br><br>" +
        "Trân trọng,<br>" +
        "<b>Kho Tổng FPT Shop Miền Nam</b>";

    var subject = `[BIÊN BẢN KIỂM HÀNG] Phiếu: ${soPhieu} - NCC: ${data.ncc}`;
    var recipientTo = "namnvh@fpt.com"; // Gửi chính đến quản lý hoặc hòm thư tiếp nhận chính

    // 4. Mở ứng dụng / Web Outlook qua giao thức mailto
    // Lưu ý: mailto hỗ trợ định dạng text/plain hoặc HTML cơ bản tùy trình duyệt
    var mailtoLink = "mailto:" + encodeURIComponent(recipientTo) +
                     "?cc=" + encodeURIComponent(ccFiltered.join(',')) +
                     "&subject=" + encodeURIComponent(subject) +
                     "&body=" + encodeURIComponent(emailBodyHTML.replace(/<br>/g, '\n').replace(/<\/?b>/g, ''));

    // Kích hoạt mở Outlook
    window.location.href = mailtoLink;
    alert("✅ Đã mở Outlook/Webmail với đầy đủ nội dung và tự động loại trừ email của bạn khỏi danh sách CC!");
    hienThiLoading(false);
}
// Hàm hỗ trợ đổi định dạng YYYY-MM-DD từ input date sang DD/MM/YYYY cho đẹp khi lên popup
function formatDateVN(dateStr) {
    if (!dateStr) return '';
    var parts = dateStr.split('-');
    if (parts.length === 3) {
        return parts[2] + '/' + parts[1] + '/' + parts[0];
    }
    return dateStr;
}

function themDongThongBao() {
    var tbody = document.getElementById('bangNhapThongBaoBody');
    var rowCount = tbody.rows.length + 1;
    
    var tr = document.createElement('tr');
    tr.innerHTML = `
        <td style="padding: 8px 4px; border: 1px solid #cbd5e1; font-weight: bold;">${rowCount}</td>
        <td style="padding: 8px 6px; border: 1px solid #cbd5e1;">
            <select class="tb-ten-hang" style="width: 100%; padding: 6px 4px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 12px; box-sizing: border-box;">
                <option value="ĐIỆN THOẠI + MTB...">Điện thoại di động, máy tính bảng, máy tính xách tay,...</option>
                <option value="PHỤ KIỆN - GIA DỤNG">Phụ kiện - Gia Dụng</option>
                <option value="ĐIỆN MÁY">Điện máy</option>
            </select>
        </td>
        <td style="padding: 8px 4px; border: 1px solid #cbd5e1;"><input type="number" class="tb-sl-vandon" value="10" style="width: 100%; padding: 6px 4px; text-align: center; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 12px; box-sizing: border-box;"></td>
        <td style="padding: 8px 4px; border: 1px solid #cbd5e1;"><input type="number" class="tb-sl-pcs" value="20" style="width: 100%; padding: 6px 4px; text-align: center; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 12px; box-sizing: border-box;"></td>
        <td style="padding: 8px 4px; border: 1px solid #cbd5e1;">
            <select class="tb-hinh-thuc" style="width: 100%; padding: 6px 4px; border: 1px solid #cbd5e1; border-radius: 4px; font-weight: bold; font-size: 12px; box-sizing: border-box;">
                <option value="CHUYỂN PHÁT NHANH">Chuyển phát nhanh</option>
                <option value="ĐƯỜNG BỘ">Đường bộ</option>
                <option value="HỎA TỐC">Hỏa tốc</option>
                <option value="HỎA TỐC ĐẶC BIỆT">Hỏa tốc đặc biệt</option>
            </select>
        </td>
        <td style="padding: 8px 4px; border: 1px solid #cbd5e1;"><input type="time" class="tb-tu-gio" value="10:00" style="width: 100%; padding: 6px 2px; text-align: center; border: 1px solid #cbd5e1; border-radius: 4px; font-weight: bold; font-size: 12px; box-sizing: border-box;"></td>
        <td style="padding: 8px 4px; border: 1px solid #cbd5e1;"><input type="time" class="tb-den-gio" value="17:30" style="width: 100%; padding: 6px 2px; text-align: center; border: 1px solid #cbd5e1; border-radius: 4px; font-weight: bold; font-size: 12px; box-sizing: border-box;"></td>
        <td style="padding: 8px 4px; border: 1px solid #cbd5e1;"><input type="date" class="tb-tu-ngay" value="2026-10-13" style="width: 100%; padding: 6px 2px; text-align: center; border: 1px solid #cbd5e1; border-radius: 4px; font-weight: bold; font-size: 12px; box-sizing: border-box;"></td>
        <td style="padding: 8px 4px; border: 1px solid #cbd5e1;"><input type="date" class="tb-den-ngay" value="2026-10-14" style="width: 100%; padding: 6px 2px; text-align: center; border: 1px solid #cbd5e1; border-radius: 4px; font-weight: bold; font-size: 12px; box-sizing: border-box;"></td>
        <td style="padding: 8px 4px; border: 1px solid #cbd5e1;"><button type="button" onclick="xoaDongThongBao(this)" style="background: #ef4444; color: white; border: none; padding: 6px 10px; border-radius: 4px; cursor: pointer; font-weight: bold;">X</button></td>
    `;
    tbody.appendChild(tr);
}

function xoaDongThongBao(btn) {
    var row = btn.closest('tr');
    row.remove();
    var rows = document.querySelectorAll('#bangNhapThongBaoBody tr');
    rows.forEach((r, idx) => {
        r.cells[0].innerText = idx + 1;
    });
}

function taoPopupThongBaoNVC() {
    var nvc = document.getElementById('thongBaoNVC').value;
    var lanGiao = document.getElementById('thongBaoLanGiao').value;
    
    var rows = document.querySelectorAll('#bangNhapThongBaoBody tr');
    if (rows.length === 0) {
        alert("⚠️ Vui lòng thêm ít nhất một dòng hàng!");
        return;
    }

    var htmlRows = "";
    rows.forEach((r, idx) => {
        var tenHang = r.querySelector('.tb-ten-hang').value;
        var slVD = r.querySelector('.tb-sl-vandon').value;
        var slPCS = r.querySelector('.tb-sl-pcs').value;
        var hinhThuc = r.querySelector('.tb-hinh-thuc').value;
        var tuGio = r.querySelector('.tb-tu-gio').value;
        var denGio = r.querySelector('.tb-den-gio').value;
        var tuNgay = formatDateVN(r.querySelector('.tb-tu-ngay').value);
        var denNgay = formatDateVN(r.querySelector('.tb-den-ngay').value);

        var khungGioText = (tuGio && denGio) ? (tuGio + " Tới " + denGio) : (tuGio || denGio);

        htmlRows += `
            <tr style="text-align: center; font-size: 14px; color: #000;">
                <td style="border: 2px solid #000; padding: 10px; font-weight: bold;">${idx + 1}</td>
                <td style="border: 2px solid #000; padding: 10px; text-align: left; font-weight: bold; color: #dc2626;">${tenHang}</td>
                <td style="border: 2px solid #000; padding: 10px; font-weight: bold;">${slVD}</td>
                <td style="border: 2px solid #000; padding: 10px; font-weight: bold;">${slPCS}</td>
                <td style="border: 2px solid #000; padding: 10px; font-weight: bold;">${hinhThuc}</td>
                <td style="border: 2px solid #000; padding: 10px; font-weight: bold; color: #1e3a8a;">${khungGioText}</td>
                <td style="border: 2px solid #000; padding: 10px; font-weight: bold; color: #b91c1c;">${tuNgay}</td>
                <td style="border: 2px solid #000; padding: 10px; font-weight: bold; color: #b91c1c;">${denNgay}</td>
            </tr>
        `;
    });

    var existingModal = document.getElementById('modalThongBaoNVC');
    if (existingModal) existingModal.remove();

    var modalHTML = `
        <div id="modalThongBaoNVC" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); display: flex; justify-content: center; align-items: center; z-index: 9999; backdrop-filter: blur(4px);">
            <div style="background: white; padding: 30px; border-radius: 12px; width: 980px; max-width: 95%; box-shadow: 0 20px 25px -5px rgb(0 0 0 / 0.2);">
                
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
                    <h3 style="margin: 0; color: #1e293b;">🖼️ Xem Trước Bản Thông Báo (Sẵn sàng sao chép)</h3>
                    <button type="button" onclick="document.getElementById('modalThongBaoNVC').remove()" style="background: #cbd5e1; border: none; width: 32px; height: 32px; border-radius: 50%; font-weight: bold; cursor: pointer;">✕</button>
                </div>

                <!-- KHUNG CHỤP ẢNH -->
                <div id="captureAreaNVC" style="background: #ffffff; padding: 25px; border: 3px solid #000; font-family: Arial, sans-serif;">
                    
                    <div style="border: 2px solid #000; text-align: center; padding: 12px; font-weight: bold; font-size: 20px; color: #c2410c; background: #fff; margin-bottom: -2px;">
                        THÔNG BÁO KẾ HOẠCH GIAO HÀNG CỦA KHO TỔNG FPT
                    </div>

                    <div style="display: grid; grid-template-columns: 2fr 1fr; border: 2px solid #000; padding: 12px 15px; font-weight: bold; font-size: 16px; background: #f8fafc; margin-bottom: 15px;">
                        <div>TỚI NHÀ VẬN CHUYỂN : <span style="color: #15803d; text-transform: uppercase;">${nvc}</span></div>
                        <div style="text-align: right; color: #b91c1c;">${lanGiao}</div>
                    </div>

                    <table style="width: 100%; border-collapse: collapse; border: 2px solid #000;">
                        <thead>
                            <tr style="background: #f1f5f9; color: #000; font-size: 13px; text-align: center;">
                                <th style="border: 2px solid #000; padding: 10px; width: 50px;" rowspan="2">STT</th>
                                <th style="border: 2px solid #000; padding: 10px;" rowspan="2">TÊN HÀNG GỬI</th>
                                <th style="border: 2px solid #000; padding: 10px; width: 90px;" rowspan="2">SỐ LƯỢNG VẬN ĐƠN</th>
                                <th style="border: 2px solid #000; padding: 10px; width: 90px;" rowspan="2">SỐ LƯỢNG HÀNG - PCS</th>
                                <th style="border: 2px solid #000; padding: 10px; width: 140px;" rowspan="2">HÌNH THỨC KÝ GỬI HÀNG HÓA</th>
                                <th style="border: 2px solid #000; padding: 10px;" colspan="3">THỜI GIAN GIAO HÀNG DỰ KIẾN</th>
                            </tr>
                            <tr style="background: #f1f5f9; color: #000; font-size: 12px; text-align: center;">
                                <th style="border: 2px solid #000; padding: 6px; width: 140px;">KHUNG GIỜ</th>
                                <th style="border: 2px solid #000; padding: 6px; width: 110px;">TỪ NGÀY</th>
                                <th style="border: 2px solid #000; padding: 6px; width: 110px;">ĐẾN NGÀY</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${htmlRows}
                        </tbody>
                    </table>

                    <div style="margin-top: 15px; font-weight: bold; font-size: 13px; color: #000; text-align: left;">
                        LƯU Ý : <span style="font-weight: normal;">Khi có thay đổi kế hoạch so với dự kiến kho tổng sẽ thông báo sau.</span>
                    </div>

                </div>

                <div style="display: flex; justify-content: flex-end; gap: 15px; margin-top: 25px;">
                    <button type="button" onclick="document.getElementById('modalThongBaoNVC').remove()" style="padding: 10px 20px; background: #e2e8f0; border: none; border-radius: 6px; font-weight: bold; cursor: pointer;">Đóng</button>
                    <button type="button" onclick="saoChepAnhThongBao()" style="padding: 10px 24px; background: #2563eb; color: white; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; display: flex; align-items: center; gap: 8px;">📋 Sao Chép Hình Ảnh</button>
                </div>

            </div>
        </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHTML);
}

async function saoChepAnhThongBao() {
    var captureElement = document.getElementById('captureAreaNVC');
    try {
        if (typeof html2canvas === 'undefined') {
            alert("⚠️ Thuvien html2canvas chưa được tải!");
            return;
        }
        var canvas = await html2canvas(captureElement, { scale: 2, backgroundColor: '#ffffff' });
        canvas.toBlob(async function(blob) {
            try {
                await navigator.clipboard.write([
                    new ClipboardItem({ 'image/png': blob })
                ]);
                alert("✅ Đã sao chép hình ảnh thông báo vào bộ nhớ tạm! Bạn có thể dán (Ctrl+V) trực tiếp vào Zalo hoặc Email.");
            } catch (err) {
                var link = document.createElement('a');
                link.download = 'ThongBao_VanChuyen.png';
                link.href = canvas.toDataURL();
                link.click();
                alert("✅ Đã tải xuống bức ảnh thông báo!");
            }
        });
    } catch (error) {
        console.error("Lỗi khi chụp ảnh:", error);
        alert("❌ Không thể tạo hình ảnh thông báo!");
    }
}
// ==========================================
// QUẢN LÝ DANH MỤC ĐỘNG (NHÀ VẬN CHUYỂN & HÌNH THỨC KÝ GỬI)
// ==========================================

// Khởi tạo danh sách mặc định nếu chưa có
function layDanhSachNVC() {
    var stored = localStorage.getItem('danhSachNVC_Custom');
    return stored ? JSON.parse(stored) : ['247 EXPRESS', 'NHẤT TÍN EXPRESS', 'VIETTEL POST', 'GNNB'];
}

function layDanhSachHinhThuc() {
    var stored = localStorage.getItem('danhSachHinhThuc_Custom');
    return stored ? JSON.parse(stored) : ['CHUYỂN PHÁT NHANH', 'ĐƯỜNG BỘ', 'HỎA TỐC', 'HỎA TỐC ĐẶC BIỆT'];
}

// Cập nhật lại toàn bộ các thẻ <select> trên giao diện
function capNhatTatCaDropdownSelect() {
    var nvcList = layDanhSachNVC();
    var hinhThucList = layDanhSachHinhThuc();

    // 1. Cập nhật dropdown chọn Nhà Vận Chuyển chính
    var selectNVC = document.getElementById('thongBaoNVC');
    if (selectNVC) {
        var currentVal = selectNVC.value;
        selectNVC.innerHTML = nvcList.map(item => `<option value="${item}">${item}</option>`).join('');
        if (nvcList.includes(currentVal)) selectNVC.value = currentVal;
    }

    // 2. Cập nhật tất cả các dropdown Hình Thức Ký Gửi trong bảng
    var selectHinhThucs = document.querySelectorAll('.tb-hinh-thuc');
    selectHinhThucs.forEach(sel => {
        var cur = sel.value;
        sel.innerHTML = hinhThucList.map(item => `<option value="${item}">${item}</option>`).join('');
        if (hinhThucList.includes(cur)) sel.value = cur;
    });

    // 3. Render lại danh sách tại bảng cài đặt
    renderBangCaiDatDanhMuc();
}

function renderBangCaiDatDanhMuc() {
    var nvcList = layDanhSachNVC();
    var hinhThucList = layDanhSachHinhThuc();

    // Render NVC
    var ulNVC = document.getElementById('listCaiDatNVC');
    if (ulNVC) {
        ulNVC.innerHTML = nvcList.map((item, idx) => `
            <li style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: white; margin-bottom: 4px; border-radius: 4px; border: 1px solid #e2e8f0; font-size: 13px;">
                <span><b>${idx + 1}.</b> ${item}</span>
                <button type="button" onclick="xoaNVC('${item}')" style="background: #ef4444; color: white; border: none; padding: 2px 6px; border-radius: 3px; cursor: pointer; font-size: 11px;">Xóa</button>
            </li>
        `).join('');
    }

    // Render Hình Thức Ký Gửi
    var ulHT = document.getElementById('listCaiDatHinhThuc');
    if (ulHT) {
        ulHT.innerHTML = hinhThucList.map((item, idx) => `
            <li style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: white; margin-bottom: 4px; border-radius: 4px; border: 1px solid #e2e8f0; font-size: 13px;">
                <span><b>${idx + 1}.</b> ${item}</span>
                <button type="button" onclick="xoaHinhThuc('${item}')" style="background: #ef4444; color: white; border: none; padding: 2px 6px; border-radius: 3px; cursor: pointer; font-size: 11px;">Xóa</button>
            </li>
        `).join('');
    }
}

function themNhaVanChuyenMoi() {
    var input = document.getElementById('inputThemNVC');
    var val = input.value.trim().toUpperCase();
    if (!val) {
        alert("⚠️ Vui lòng nhập tên nhà vận chuyển!");
        return;
    }
    var list = layDanhSachNVC();
    if (list.includes(val)) {
        alert("⚠️ Nhà vận chuyển này đã có trong danh sách!");
        return;
    }
    list.push(val);
    localStorage.setItem('danhSachNVC_Custom', JSON.stringify(list));
    input.value = '';
    capNhatTatCaDropdownSelect();
}

function xoaNVC(item) {
    var list = layDanhSachNVC();
    list = list.filter(i => i !== item);
    localStorage.setItem('danhSachNVC_Custom', JSON.stringify(list));
    capNhatTatCaDropdownSelect();
}

function themHinhThucKyGuiMoi() {
    var input = document.getElementById('inputThemHinhThuc');
    var val = input.value.trim().toUpperCase();
    if (!val) {
        alert("⚠️ Vui lòng nhập hình thức ký gửi!");
        return;
    }
    var list = layDanhSachHinhThuc();
    if (list.includes(val)) {
        alert("⚠️ Hình thức ký gửi này đã có trong danh sách!");
        return;
    }
    list.push(val);
    localStorage.setItem('danhSachHinhThuc_Custom', JSON.stringify(list));
    input.value = '';
    capNhatTatCaDropdownSelect();
}

function xoaHinhThuc(item) {
    var list = layDanhSachHinhThuc();
    list = list.filter(i => i !== item);
    localStorage.setItem('danhSachHinhThuc_Custom', JSON.stringify(list));
    capNhatTatCaDropdownSelect();
}

// Gọi tự động khi vừa tải trang để đồng bộ các dropdown
document.addEventListener('DOMContentLoaded', function() {
    setTimeout(capNhatTatCaDropdownSelect, 200);
});

// ==========================================
// 8. MODULE 8: THEO DÕI VẬT TƯ KHO (ĐÃ CLEAN & CHUẨN XÁC)
// ==========================================
window.addEventListener('DOMContentLoaded', (event) => {
    var today = new Date();
    var vtNgay = document.getElementById('inputVatTuNgay');
    if (vtNgay && !vtNgay.value) {
        vtNgay.valueAsDate = today;
        taiDuLieuVatTuTheoNgay();
    }
});

document.addEventListener('DOMContentLoaded', () => {
    var inputNgay = document.getElementById('inputVatTuNgay');
    if (inputNgay) {
        inputNgay.addEventListener('change', taiDuLieuVatTuTheoNgay);
    }
});

// Hàm tính toán tự động cột số 4 (Tổng số lượng còn lại) dựa trên tồn gốc cộng dồn phát sinh
function tinhToanTonKho(element) {
    var row = element.closest('tr');
    var inputNhap = row.querySelector('.vt-nhap');
    var inputXuat = row.querySelector('.vt-xuat');
    var cellTon = row.querySelector('.vt-ton');

    if (!inputNhap || !inputXuat || !cellTon) return;

    var tonGoc = parseFloat(cellTon.getAttribute('data-tong-goc')) || 0;
    var nhapMoi = parseFloat(inputNhap.value) || 0;
    var xuatMoi = parseFloat(inputXuat.value) || 0;
    
    // Tổng số lượng còn lại = Tồn gốc + Nhập mới - Xuất mới
    var tongConLai = tonGoc + nhapMoi - xuatMoi;
    
    cellTon.innerText = tongConLai;
    cellTon.style.color = tongConLai < 0 ? '#dc2626' : '#16a34a';
}

function capNhatLaiTatCaTonKho() {
    var rows = document.querySelectorAll('#tableVatTuBody tr');
    rows.forEach(function(row) {
        var inputNhap = row.querySelector('.vt-nhap');
        if (inputNhap) {
            tinhToanTonKho(inputNhap);
        }
    });
}

// Hàm tải dữ liệu vật tư từ Server theo ngày (Đảm bảo lôi data và hiển thị cột số 4)
async function taiDuLieuVatTuTheoNgay() {
    var ngayChon = document.getElementById('inputVatTuNgay').value;
    if (!ngayChon) return;

    try {
        var response = await fetch(`http://localhost:3000/api/vattu?ngay=${ngayChon}`);
        var result = await response.json();
        
        var tbody = document.getElementById('tableVatTuBody');
        if (!tbody) return;
        tbody.innerHTML = '';

        var dsVatTuHienThi = [];
        if (result.success && result.data && result.data.length > 0) {
            dsVatTuHienThi = result.data;
        } else {
            dsVatTuHienThi = [
                { tenVatTu: "Băng keo đỏ", nhap: 0, xuat: 0, ton: 0 },
                { tenVatTu: "Thùng carton", nhap: 0, xuat: 0, ton: 0 },
                { tenVatTu: "Tem bạc FPT", nhap: 0, xuat: 0, ton: 0 }
            ];
        }

        dsVatTuHienThi.forEach((item, index) => {
            var nhapVal = parseFloat(item.nhap) || 0;
            var xuatVal = parseFloat(item.xuat) || 0;
            var tonGoc = item.ton !== undefined ? parseFloat(item.ton) : (nhapVal - xuatVal);

            var tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="padding: 12px; text-align: center; border-bottom: 1px solid #e2e8f0; font-weight: bold;">${index + 1}</td>
                <td style="padding: 12px 20px; border-bottom: 1px solid #e2e8f0; font-weight: bold; color: #0f172a;">${item.tenVatTu}</td>
                <td style="padding: 10px 15px; border-bottom: 1px solid #e2e8f0; text-align: center;">
                    <input type="number" class="vt-nhap" value="0" min="0" oninput="tinhToanTonKho(this)" style="width: 100px; padding: 6px; text-align: center; border: 1px solid #cbd5e1; border-radius: 6px; font-weight: bold;">
                </td>
                <td style="padding: 10px 15px; border-bottom: 1px solid #e2e8f0; text-align: center;">
                    <input type="number" class="vt-xuat" value="0" min="0" oninput="tinhToanTonKho(this)" style="width: 100px; padding: 6px; text-align: center; border: 1px solid #cbd5e1; border-radius: 6px; font-weight: bold;">
                </td>
                <td style="padding: 10px 15px; border-bottom: 1px solid #e2e8f0; text-align: center; font-weight: bold; font-size: 15px; color: #16a34a;" class="vt-ton" data-tong-goc="${tonGoc}">${tonGoc}</td>
                <td style="padding: 10px 12px; border-bottom: 1px solid #e2e8f0; text-align: center;">
                    <button type="button" onclick="xoaDongVatTu(this)" style="background: #fee2e2; color: #dc2626; border: none; padding: 6px 10px; border-radius: 6px; font-weight: bold; cursor: pointer;">🗑️</button>
                </td>
            `;
            tbody.appendChild(tr);
        });

        // Kích hoạt quét tính toán để hiện số liệu đầy đủ ở cột 4
        capNhatLaiTatCaTonKho();
    } catch (error) {
        console.error("Lỗi tải dữ liệu vật tư từ server:", error);
    }
}

function openThemVatTuModal() {
    var inputMoi = document.getElementById('inputTenVatTuMoi');
    if (inputMoi) inputMoi.value = '';
    var modal = document.getElementById('themVatTuModal');
    if (modal) modal.style.display = 'flex';
    if (inputMoi) inputMoi.focus();
}

function closeThemVatTuModal() {
    var modal = document.getElementById('themVatTuModal');
    if (modal) modal.style.display = 'none';
}

function xacNhanThemVatTu() {
    var inputMoi = document.getElementById('inputTenVatTuMoi');
    if (!inputMoi) return;
    
    var tenVatTu = inputMoi.value.trim();
    if (!tenVatTu) {
        alert('Vui lòng nhập tên vật tư!');
        return;
    }
    
    var tbody = document.getElementById('tableVatTuBody');
    var existingRows = tbody.querySelectorAll('tr');
    var found = false;
    
    existingRows.forEach(row => {
        var nameCell = row.cells[1].innerText.trim();
        if (nameCell.toLowerCase() === tenVatTu.toLowerCase()) {
            found = true;
            var themSoLuong = prompt(`Vật tư "${tenVatTu}" đã có trong danh sách. Nhập số lượng cần cộng dồn thêm:`, "0");
            if (themSoLuong !== null && !isNaN(themSoLuong)) {
                var inputNhap = row.querySelector('.vt-nhap');
                inputNhap.value = parseFloat(inputNhap.value || 0) + parseFloat(themSoLuong);
                tinhToanTonKho(inputNhap);
                alert(`Đã cộng dồn thành công ${themSoLuong} vào ${tenVatTu}!`);
            }
        }
    });

    if (!found) {
        var rowCount = tbody.rows.length + 1;
        var newRow = document.createElement('tr');
        newRow.innerHTML = `
            <td style="padding: 12px; text-align: center; border-bottom: 1px solid #e2e8f0; font-weight: bold;">${rowCount}</td>
            <td style="padding: 12px 20px; border-bottom: 1px solid #e2e8f0; font-weight: bold; color: #0f172a;">${tenVatTu}</td>
            <td style="padding: 10px 15px; border-bottom: 1px solid #e2e8f0; text-align: center;"><input type="number" class="vt-nhap" value="0" min="0" oninput="tinhToanTonKho(this)" style="width: 100px; padding: 6px; text-align: center; border: 1px solid #cbd5e1; border-radius: 6px; font-weight: bold;"></td>
            <td style="padding: 10px 15px; border-bottom: 1px solid #e2e8f0; text-align: center;"><input type="number" class="vt-xuat" value="0" min="0" oninput="tinhToanTonKho(this)" style="width: 100px; padding: 6px; text-align: center; border: 1px solid #cbd5e1; border-radius: 6px; font-weight: bold;"></td>
            <td style="padding: 10px 15px; border-bottom: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: #16a34a; font-size: 15px;" class="vt-ton" data-tong-goc="0">0</td>
            <td style="padding: 10px 12px; border-bottom: 1px solid #e2e8f0; text-align: center;"><button type="button" onclick="xoaDongVatTu(this)" style="background: #fee2e2; color: #dc2626; border: none; padding: 6px 10px; border-radius: 6px; font-weight: bold; cursor: pointer;">🗑️</button></td>
        `;
        tbody.appendChild(newRow);
    }

    closeThemVatTuModal();
    capNhatLaiTatCaTonKho();
}

function xoaDongVatTu(btn) {
    if (confirm('Bạn có chắc chắn muốn xóa vật tư này khỏi danh sách?')) {
        btn.closest('tr').remove();
        capNhatSTTVatTu();
    }
}

function capNhatSTTVatTu() {
    var rows = document.querySelectorAll('#tableVatTuBody tr');
    rows.forEach(function(row, index) {
        row.cells[0].innerText = index + 1;
    });
}

// Hàm lưu dữ liệu vật tư lên Server (Cộng dồn vào database.json)
async function luuDuLieuVatTu() {
    var ngayChon = document.getElementById('inputVatTuNgay').value;
    if (!ngayChon) {
        alert("⚠️ Vui lòng chọn ngày kiểm tra vật tư!");
        return;
    }

    var rows = document.querySelectorAll('#tableVatTuBody tr');
    var danhSachVatTu = [];

    rows.forEach(function(row) {
        var tenVatTu = row.cells[1].innerText.trim();
        var soLuongNhap = parseFloat(row.querySelector('.vt-nhap').value) || 0;
        var soLuongXuat = parseFloat(row.querySelector('.vt-xuat').value) || 0;
        var tonKho = parseFloat(row.querySelector('.vt-ton').innerText) || 0;

        danhSachVatTu.push({
            tenVatTu: tenVatTu,
            nhap: soLuongNhap,
            xuat: soLuongXuat,
            ton: tonKho
        });
    });

    var payload = {
        ngay: ngayChon,
        danhSachVatTu: danhSachVatTu
    };

    try {
        var response = await fetch('http://localhost:3000/api/vattu', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        var result = await response.json();
        if (result.success) {
            alert("✅ " + result.message);
            taiDuLieuVatTuTheoNgay(); // Tải lại để cập nhật số tồn gốc mới nhất
        } else {
            alert("❌ Lỗi khi lưu: " + result.message);
        }
    } catch (error) {
        console.error("Lỗi kết nối server:", error);
        alert("❌ Không thể kết nối đến server để lưu dữ liệu vật tư!");
    }
}
// MẪU CHUẨN XỬ LÝ KHÔNG RELOAD TRANG DÙNG LOADING ICON
async function luuDuLieuKhongReload(event, apiEndpoint, payloadData, onSuccessCallback) {
    if (event) event.preventDefault(); // Chặn reload trang mặc định

    // 1. Hiển thị Icon Loading (Ví dụ bật một phần tử loading chung trên màn hình)
    hienThiLoading(true);

    try {
        // 2. Gửi dữ liệu ngầm bằng Fetch API
        let response = await fetch(apiEndpoint, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payloadData)
        });

        let result = await response.json();

        if (result.success || response.ok) {
            // 3. Thông báo thành công mượt mà (Toast message / Alert nhẹ)
            hienThongBaoNhanh("💾 Lưu dữ liệu thành công!", "success");
            
            // 4. Gọi callback cập nhật lại bảng dữ liệu trên giao diện mà không tải lại web
            if (typeof onSuccessCallback === 'function') {
                onSuccessCallback(result);
            }
        } else {
            alert("Lỗi: " + (result.message || "Không thể lưu dữ liệu!"));
        }
    } catch (error) {
        console.error("Lỗi kết nối:", error);
        alert("Lỗi kết nối đến server. Vui lòng kiểm tra lại!");
    } finally {
        // 5. Tắt Icon Loading dù thành công hay thất bại
        hienThiLoading(false);
    }
}