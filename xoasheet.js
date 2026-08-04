function cleanSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetNames = [
    "UBNCSP – T01.2026",
    "UBNCSP – T02.2026",
    "UBNCSP – T03.2026",
    "UBNCSP – T04.2026",
    "UBNCSP – T05.2026",
    "UBNCSP – T06.2026",
    "UBNCSP – T07.2026",
    "UBNCSP – T08.2026",
    "UBNCSP – T09.2026",
    "UBNCSP – T10.2026",
    "UBNCSP – T11.2026",
    "UBNCSP – T12.2026"
  ];

  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);

    if (!sheet) {
      Logger.log("Không tìm thấy sheet: " + name);
      return;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return;

    var data = sheet.getRange(1, 1, lastRow, 2).getDisplayValues();

    for (var i = 0; i < lastRow; i++) {
      var rowNum = i + 1;

      var valA = String(data[i][0]).trim();
      var valB = String(data[i][1]).trim();

      var rowRange = sheet.getRange(rowNum, 1, 1, 26);
      var cellA = sheet.getRange(rowNum, 1);
      var cellB = sheet.getRange(rowNum, 2);

      // Nếu hàng có ô merge thì phải tách toàn bộ vùng merge trong hàng trước
      var needBreakMerge =
        (valA !== "∑" && /^\d+(\.\d+)*$/.test(valA)) ||
        (valB !== "" && valB.indexOf("Công tác phát sinh khác trong kỳ") === -1) ||
        (valA === "∑" || valB.indexOf("Công tác phát sinh khác trong kỳ") !== -1);

      if (needBreakMerge) {
        breakMergedRangesInRange_(rowRange);
      }

      // --- Cột A: xóa số thứ tự, giữ ∑ ---
      if (valA !== "∑" && /^\d+(\.\d+)*$/.test(valA)) {
        cellA.clearContent();
      }

      // --- Cột B: xóa tất cả trừ hàng "Công tác phát sinh khác trong kỳ" ---
      if (valB !== "" && valB.indexOf("Công tác phát sinh khác trong kỳ") === -1) {
        cellB.clearContent();
      }
    }

    Logger.log("✅ Xong: " + name);
  });

  SpreadsheetApp.getUi().alert("✅ Hoàn tất 12 sheet! Kiểm tra Nhật ký thực thi để xem chi tiết.");
}


/**
 * Tách toàn bộ các vùng ô đã hợp nhất nằm trong một range.
 * Tránh lỗi: "Bạn phải chọn tất cả các ô trong dải ô đã hợp nhất..."
 */
function breakMergedRangesInRange_(range) {
  var mergedRanges = range.getMergedRanges();

  mergedRanges.forEach(function(mergedRange) {
    try {
      mergedRange.breakApart();
    } catch (e) {
      Logger.log("Không tách được vùng merge: " + mergedRange.getA1Notation() + " | " + e.message);
    }
  });
}