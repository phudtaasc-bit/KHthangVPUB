/**
 * KHÓA / MỞ BẢO VỆ TRÊN SHEET HIỆN TẠI
 * - Mở: lưu cấu hình bảo vệ hiện tại rồi gỡ Protection để sheet thực sự chỉnh sửa được.
 * - Khóa: khôi phục đúng các Protection đã lưu.
 * - Chỉ chủ sở hữu file và phudt.entiz@gmail.com được thao tác.
 * - Không thay đổi lịch tự động khóa hàng tháng đang có.
 */

function BV_taoMenuBaoVeSheet() {
  SpreadsheetApp.getUi()
    .createMenu('🔐 Bảo vệ sheet')
    .addItem('🔓 Mở bảo vệ sheet hiện tại', 'BV_moBaoVeSheetHienTai')
    .addItem('🔒 Khóa bảo vệ sheet hiện tại', 'BV_khoaBaoVeSheetHienTai')
    .addToUi();
}

/**
 * Chạy một lần nếu dự án chưa có trigger tạo menu riêng.
 * Không cần chạy lại nếu menu đã xuất hiện.
 */
function BV_caiDatMenuBaoVeSheet() {
  BV_kiemTraQuyenThaoTac_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var handler = 'BV_taoMenuBaoVeSheet';

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
  SpreadsheetApp.getUi().alert('Đã cài menu Bảo vệ sheet.');
}

/**
 * Mở bảo vệ sheet hiện tại.
 * Lưu toàn bộ cấu hình Protection vào Document Properties rồi gỡ Protection.
 */
function BV_moBaoVeSheetHienTai() {
  BV_kiemTraQuyenThaoTac_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var protections = BV_layProtectionSheetHienTai_(sheet);

  if (protections.length === 0) {
    SpreadsheetApp.getUi().alert(
      'Sheet "' + sheet.getName() + '" hiện không có Protection để mở.'
    );
    return;
  }

  var snapshot = protections.map(function(protection) {
    var isSheetProtection = protection.getProtectionType() === SpreadsheetApp.ProtectionType.SHEET;
    var editors = [];

    try {
      editors = protection.getEditors().map(function(user) {
        return user.getEmail();
      }).filter(String);
    } catch (err) {
      editors = [];
    }

    var item = {
      type: isSheetProtection ? 'SHEET' : 'RANGE',
      description: protection.getDescription() || '',
      editors: editors,
      domainEdit: protection.canDomainEdit(),
      warningOnly: protection.isWarningOnly()
    };

    if (isSheetProtection) {
      item.unprotectedRanges = protection.getUnprotectedRanges().map(function(range) {
        return range.getA1Notation();
      });
    } else {
      item.rangeA1 = protection.getRange().getA1Notation();
    }

    return item;
  });

  PropertiesService.getDocumentProperties().setProperty(
    BV_layKhoaLuuCauHinh_(sheet),
    JSON.stringify(snapshot)
  );

  protections.forEach(function(protection) {
    protection.remove();
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Đã mở hoàn toàn bảo vệ trên sheet "' + sheet.getName() + '".\n' +
    'Cấu hình bảo vệ đã được lưu để khóa lại.'
  );
}

/**
 * Khóa lại sheet hiện tại theo đúng cấu hình đã lưu khi mở.
 */
function BV_khoaBaoVeSheetHienTai() {
  BV_kiemTraQuyenThaoTac_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var properties = PropertiesService.getDocumentProperties();
  var key = BV_layKhoaLuuCauHinh_(sheet);
  var raw = properties.getProperty(key);

  if (!raw) {
    SpreadsheetApp.getUi().alert(
      'Không tìm thấy cấu hình bảo vệ đã lưu cho sheet "' + sheet.getName() + '".\n' +
      'Hãy dùng chức năng Mở bảo vệ trước, hoặc chờ cơ chế tự động khóa hiện có.'
    );
    return;
  }

  // Xóa Protection hiện có để tránh trùng lặp trước khi khôi phục.
  BV_layProtectionSheetHienTai_(sheet).forEach(function(protection) {
    protection.remove();
  });

  var snapshot = JSON.parse(raw);
  var ownerEmail = BV_layEmailChuSoHuu_(ss);
  var allowedEditors = [ownerEmail, 'phudt.entiz@gmail.com']
    .filter(String)
    .map(function(email) { return email.toLowerCase(); });

  snapshot.forEach(function(item) {
    var protection;

    if (item.type === 'SHEET') {
      protection = sheet.protect();
      if (item.unprotectedRanges && item.unprotectedRanges.length > 0) {
        protection.setUnprotectedRanges(
          item.unprotectedRanges.map(function(a1) {
            return sheet.getRange(a1);
          })
        );
      }
    } else {
      protection = sheet.getRange(item.rangeA1).protect();
    }

    if (item.description) {
      protection.setDescription(item.description);
    }

    protection.setWarningOnly(false);

    // Xóa editor không được phép.
    protection.getEditors().forEach(function(user) {
      var email = String(user.getEmail() || '').toLowerCase();
      if (email && allowedEditors.indexOf(email) === -1) {
        protection.removeEditor(user);
      }
    });

    if (ownerEmail) {
      protection.addEditor(ownerEmail);
    }
    protection.addEditor('phudt.entiz@gmail.com');

    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  });

  properties.deleteProperty(key);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Đã khóa lại sheet "' + sheet.getName() + '".\n\n' +
    'Người được phép chỉnh sửa vùng bảo vệ:\n' +
    '- Chủ sở hữu file\n' +
    '- phudt.entiz@gmail.com'
  );
}

function BV_layProtectionSheetHienTai_(sheet) {
  return sheet
    .getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .concat(sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE));
}

function BV_layKhoaLuuCauHinh_(sheet) {
  return 'BV_PROTECTION_SNAPSHOT_' + sheet.getSheetId();
}

function BV_kiemTraQuyenThaoTac_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentEmail = String(Session.getEffectiveUser().getEmail() || '')
    .trim()
    .toLowerCase();
  var ownerEmail = BV_layEmailChuSoHuu_(ss).toLowerCase();

  if ([ownerEmail, 'phudt.entiz@gmail.com'].indexOf(currentEmail) === -1) {
    throw new Error(
      'Bạn không có quyền mở/khóa bảo vệ. Chỉ chủ sở hữu file và ' +
      'phudt.entiz@gmail.com được phép thao tác.'
    );
  }
}

function BV_layEmailChuSoHuu_(ss) {
  var owner = DriveApp.getFileById(ss.getId()).getOwner();
  if (!owner) {
    throw new Error('Không xác định được chủ sở hữu Google Sheet.');
  }
  return String(owner.getEmail() || '').trim();
}
