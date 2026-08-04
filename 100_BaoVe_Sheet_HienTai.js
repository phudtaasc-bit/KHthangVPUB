/**
 * KHÓA / MỞ BẢO VỆ TRÊN SHEET HIỆN TẠI
 * Chỉ chủ sở hữu file và phudt.entiz@gmail.com được thao tác.
 * Không thay đổi lịch tự động khóa đang có.
 */

const BV_EMAIL_DUOC_PHEP = 'phudt.entiz@gmail.com';
const BV_TEN_MENU = '🔐 Bảo vệ sheet';

/**
 * Chạy 1 lần sau khi push để cài menu bằng installable onOpen trigger.
 * Cách này không xung đột với hàm onOpen() hiện có trong dự án.
 */
function BV_caiDatMenuBaoVeSheet() {
  BV_kiemTraQuyenThaoTac_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const handler = 'BV_taoMenuBaoVeSheet';

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(handler)
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  BV_taoMenuBaoVeSheet();
  SpreadsheetApp.getUi().alert('Đã cài menu "' + BV_TEN_MENU + '".');
}

/** Tạo menu khi mở Google Sheet. */
function BV_taoMenuBaoVeSheet() {
  SpreadsheetApp.getUi()
    .createMenu(BV_TEN_MENU)
    .addItem('🔓 Mở bảo vệ sheet hiện tại', 'BV_moBaoVeSheetHienTai')
    .addItem('🔒 Khóa bảo vệ sheet hiện tại', 'BV_khoaBaoVeSheetHienTai')
    .addToUi();
}

/**
 * Mở tạm thời tất cả Protection của sheet hiện tại.
 * Chuyển sang chế độ cảnh báo, không xóa phạm vi bảo vệ.
 */
function BV_moBaoVeSheetHienTai() {
  BV_kiemTraQuyenThaoTac_();

  const sheet = SpreadsheetApp.getActiveSheet();
  const protections = BV_layProtectionSheetHienTai_(sheet);

  if (protections.length === 0) {
    SpreadsheetApp.getUi().alert(
      'Sheet "' + sheet.getName() + '" chưa có vùng hoặc sheet được bảo vệ.'
    );
    return;
  }

  protections.forEach(function(protection) {
    protection.setWarningOnly(true);
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Đã mở bảo vệ trên sheet "' + sheet.getName() + '".\n' +
    'Các phạm vi bảo vệ vẫn được giữ nguyên để khóa lại.'
  );
}

/**
 * Khóa lại tất cả Protection của sheet hiện tại.
 * Chỉ chủ sở hữu và email được cấu hình có quyền chỉnh sửa vùng bảo vệ.
 */
function BV_khoaBaoVeSheetHienTai() {
  BV_kiemTraQuyenThaoTac_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const protections = BV_layProtectionSheetHienTai_(sheet);

  if (protections.length === 0) {
    SpreadsheetApp.getUi().alert(
      'Sheet "' + sheet.getName() + '" chưa có vùng hoặc sheet được bảo vệ.'
    );
    return;
  }

  const ownerEmail = BV_layEmailChuSoHuu_(ss);
  const allowedEmails = [ownerEmail, BV_EMAIL_DUOC_PHEP]
    .filter(String)
    .map(function(email) { return email.toLowerCase(); });

  protections.forEach(function(protection) {
    protection.setWarningOnly(false);

    protection.getEditors().forEach(function(editor) {
      const email = String(editor.getEmail() || '').toLowerCase();
      if (email && allowedEmails.indexOf(email) === -1) {
        protection.removeEditor(editor);
      }
    });

    if (ownerEmail) protection.addEditor(ownerEmail);
    protection.addEditor(BV_EMAIL_DUOC_PHEP);

    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Đã khóa bảo vệ trên sheet "' + sheet.getName() + '".\n\n' +
    'Người được phép chỉnh sửa:\n' +
    '- Chủ sở hữu file\n' +
    '- ' + BV_EMAIL_DUOC_PHEP
  );
}

/** Lấy cả bảo vệ toàn sheet và bảo vệ theo vùng của sheet hiện tại. */
function BV_layProtectionSheetHienTai_(sheet) {
  return sheet
    .getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .concat(sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE));
}

/** Chỉ chủ sở hữu hoặc email được cấu hình được phép thao tác. */
function BV_kiemTraQuyenThaoTac_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentEmail = String(Session.getEffectiveUser().getEmail() || '')
    .trim()
    .toLowerCase();
  const ownerEmail = BV_layEmailChuSoHuu_(ss).toLowerCase();

  if ([ownerEmail, BV_EMAIL_DUOC_PHEP.toLowerCase()].indexOf(currentEmail) === -1) {
    throw new Error(
      'Bạn không có quyền mở/khóa bảo vệ. Chỉ chủ sở hữu file và ' +
      BV_EMAIL_DUOC_PHEP + ' được phép thao tác.'
    );
  }
}

/** Lấy email chủ sở hữu Google Sheet. */
function BV_layEmailChuSoHuu_(ss) {
  const owner = DriveApp.getFileById(ss.getId()).getOwner();
  if (!owner) {
    throw new Error('Không xác định được chủ sở hữu Google Sheet.');
  }
  return String(owner.getEmail() || '').trim();
}
