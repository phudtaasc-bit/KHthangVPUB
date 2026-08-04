/**
 * KHÓA / MỞ BẢO VỆ TRÊN SHEET HIỆN TẠI
 *
 * Quyền thao tác:
 * - Chủ sở hữu Google Sheet;
 * - phudt.entiz@gmail.com.
 *
 * Lưu ý quan trọng:
 * - Hàm BV_capQuyenPhuToanBoFile() phải được CHỦ SỞ HỮU chạy một lần.
 * - Hàm này bổ sung phudt.entiz@gmail.com vào toàn bộ Protection hiện có.
 * - Sau đó email này mới có quyền gỡ/khôi phục Protection bằng menu.
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
 * CHỦ SỞ HỮU chạy hàm này một lần trong Apps Script Editor.
 * Bổ sung phudt.entiz@gmail.com vào toàn bộ Protection của tất cả sheet.
 */
function BV_capQuyenPhuToanBoFile() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ownerEmail = BV_layEmailChuSoHuu_(ss).toLowerCase();
  var currentEmail = String(Session.getEffectiveUser().getEmail() || '')
    .trim()
    .toLowerCase();

  if (currentEmail !== ownerEmail) {
    throw new Error(
      'Hàm BV_capQuyenPhuToanBoFile phải do chủ sở hữu Google Sheet chạy. ' +
      'Tài khoản hiện tại: ' + (currentEmail || '(không xác định)')
    );
  }

  var targetEmail = 'phudt.entiz@gmail.com';
  var total = 0;
  var errors = [];

  ss.getSheets().forEach(function(sheet) {
    var protections = BV_layProtectionSheetHienTai_(sheet);

    protections.forEach(function(protection) {
      try {
        if (protection.isWarningOnly()) {
          protection.setWarningOnly(false);
        }

        protection.addEditor(targetEmail);

        if (protection.canDomainEdit()) {
          protection.setDomainEdit(false);
        }

        total++;
      } catch (err) {
        errors.push(sheet.getName() + ': ' + err.message);
      }
    });
  });

  SpreadsheetApp.flush();

  var message =
    'Đã cấp quyền chỉnh sửa Protection cho ' + targetEmail + '.\n' +
    'Số Protection cập nhật: ' + total;

  if (errors.length > 0) {
    message += '\n\nCác lỗi còn lại:\n' + errors.join('\n');
  }

  SpreadsheetApp.getUi().alert(message);
}

/**
 * Mở hoàn toàn Protection của sheet hiện tại.
 * Lưu cấu hình rồi gỡ Protection.
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
    var isSheetProtection =
      protection.getProtectionType() === SpreadsheetApp.ProtectionType.SHEET;

    var item = {
      type: isSheetProtection ? 'SHEET' : 'RANGE',
      description: protection.getDescription() || ''
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
    try {
      protection.remove();
    } catch (err) {
      throw new Error(
        'Không thể gỡ Protection trên sheet "' + sheet.getName() + '". ' +
        'Chủ sở hữu cần chạy BV_capQuyenPhuToanBoFile() một lần trước. ' +
        'Chi tiết: ' + err.message
      );
    }
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Đã mở hoàn toàn bảo vệ trên sheet "' + sheet.getName() + '".\n' +
    'Cấu hình bảo vệ đã được lưu để khóa lại.'
  );
}

/**
 * Khóa lại sheet hiện tại theo cấu hình đã lưu.
 * Chỉ chủ sở hữu và phudt.entiz@gmail.com được chỉnh sửa Protection.
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
      'Hãy dùng chức năng Mở bảo vệ trước.'
    );
    return;
  }

  BV_layProtectionSheetHienTai_(sheet).forEach(function(protection) {
    protection.remove();
  });

  var snapshot = JSON.parse(raw);
  var ownerEmail = BV_layEmailChuSoHuu_(ss);
  var targetEmail = 'phudt.entiz@gmail.com';

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
    protection.addEditor(targetEmail);

    if (ownerEmail) {
      protection.addEditor(ownerEmail);
    }

    protection.getEditors().forEach(function(user) {
      var email = String(user.getEmail() || '').toLowerCase();
      if (
        email &&
        email !== ownerEmail.toLowerCase() &&
        email !== targetEmail.toLowerCase()
      ) {
        protection.removeEditor(user);
      }
    });

    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  });

  properties.deleteProperty(key);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Đã khóa lại sheet "' + sheet.getName() + '".\n\n' +
    'Người được phép chỉnh sửa:\n' +
    '- Chủ sở hữu file\n' +
    '- ' + targetEmail
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
  var targetEmail = 'phudt.entiz@gmail.com';

  if ([ownerEmail, targetEmail].indexOf(currentEmail) === -1) {
    throw new Error(
      'Bạn không có quyền mở/khóa bảo vệ. Chỉ chủ sở hữu file và ' +
      targetEmail + ' được phép thao tác.'
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
