const EVENT_DB_ID = '1xlGdKp9WZY2J6WcUvkfYrx1GTJHvrdrLvzeFmQPY86o';
const ACCOUNT_DB_ID = '1RBXdJs6HZB0rPCRU0YyOcEUt_usBGhYQq0ePGpLLUCg';

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    
    if (action === 'login') {
      return responseJSON(login(params.username, params.password));
    } else if (action === 'getEvents') {
      return responseJSON(getEvents());
    } else if (action === 'getRegistrations') {
      return responseJSON(getRegistrations(params.maSuKien));
    }
    
    return responseJSON({ success: false, message: 'Hành động không hợp lệ.' });
  } catch (error) {
    return responseJSON({ success: false, message: error.toString() });
  }
}

// Hàm chuẩn hóa dữ liệu từ mảng 2D thành mảng Object
function sheetDataToObjectArray(values) {
  const headers = values[0];
  const result = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    result.push(obj);
  }
  return result;
}

// 1. Chức năng đăng nhập
function login(username, password) {
  const sheet = SpreadsheetApp.openById(ACCOUNT_DB_ID).getSheetByName('TaiKhoanQuanTri');
  const data = sheetDataToObjectArray(sheet.getDataRange().getValues());
  
  const user = data.find(u => 
    u.TenDangNhap == username && 
    u.MatKhau == password && 
    u.TrangThaiHoatDong == 'Hoạt động'
  );
  
  if (user) {
    return { success: true, data: { ten: user.Ten, username: user.TenDangNhap } };
  } else {
    return { success: false, message: 'Sai tên đăng nhập, mật khẩu hoặc tài khoản đã bị khóa.' };
  }
}

// 2. Lấy danh sách sự kiện (Sắp xếp Ngày tổ chức lớn nhất lên đầu)
function getEvents() {
  const sheet = SpreadsheetApp.openById(EVENT_DB_ID).getSheetByName('ThongTinSuKien');
  const values = sheet.getDataRange().getValues();
  let events = sheetDataToObjectArray(values);
  
  // Xử lý ngày tháng và sắp xếp
  events = events.map(e => {
    // Lưu lại chuỗi ngày hiển thị để gửi về FE
    e._NgayToChucStr = e.NgayToChuc instanceof Date ? Utilities.formatDate(e.NgayToChuc, Session.getScriptTimeZone(), 'dd/MM/yyyy') : e.NgayToChuc;
    // Giá trị timestamp để sắp xếp
    e._timestamp = e.NgayToChuc instanceof Date ? e.NgayToChuc.getTime() : 0;
    return e;
  });
  
  // Sort giảm dần (lớn nhất lên đầu)
  events.sort((a, b) => b._timestamp - a._timestamp);
  
  return { success: true, data: events };
}

// 3. Lấy danh sách đăng ký theo sự kiện
function getRegistrations(maSuKien) {
  const sheet = SpreadsheetApp.openById(EVENT_DB_ID).getSheetByName('DSDKSuKien');
  // Dùng getDisplayValues để lấy nguyên text ngày tháng, tránh lỗi timezone
  const values = sheet.getDataRange().getDisplayValues(); 
  const allRegs = sheetDataToObjectArray(values);
  
  const filteredRegs = allRegs.filter(r => r.MaSuKienWorkshop == maSuKien);
  
  return { success: true, data: filteredRegs };
}

// Trả về JSON
function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}