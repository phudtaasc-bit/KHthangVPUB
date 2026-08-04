/**
 * KH tuần WBS - Cập nhật kết quả tuần về KH tháng
 *
 * Bản nhẹ:
 * - Chỉ xử lý sheet tháng đang chọn tại B2.
 * - Chỉ xử lý tuần đang chọn tại E2.
 * - Chỉ đọc các dòng đang hiển thị trên KH_TUAN_WBS.
 * - Không quét KH_TUAN_DATA.
 * - Không quét nhiều tháng / nhiều tuần.
 */

function khTuanWbs_pushWeekResultToMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const config = KH_TUAN_WBS_CONFIG;
  const sheetTuan = ss.getSheetByName(config.SHEET_TUAN);

  if (!sheetTuan) {
    ui.alert('Không tìm thấy sheet KH_TUAN_WBS.');
    return;
  }

  const sourceSheetName = String(sheetTuan.getRange(config.SOURCE_CELL).getDisplayValue() || '').trim();
  const weekName = String(sheetTuan.getRange(config.WEEK_CELL).getDisplayValue() || '').trim();

  if (!sourceSheetName || !weekName) {
    ui.alert('Thiếu B2 hoặc E2 trên KH_TUAN_WBS.');
    return;
  }

  const monthSheet = ss.getSheetByName(sourceSheetName);
  if (!monthSheet) {
    ui.alert('Không tìm thấy sheet tháng: ' + sourceSheetName);
    return;
  }

  const monthWarning = khTuanWbs_getMonthScopeWarning_(sourceSheetName);
  if (monthWarning) {
    const confirmMonth = ui.alert(
      'Cảnh báo tháng cập nhật',
      monthWarning,
      ui.ButtonSet.YES_NO
    );

    if (confirmMonth !== ui.Button.YES) {
      ss.toast('Đã hủy cập nhật kết quả tuần về KH tháng.', 'KH_TUAN_WBS', 5);
      return;
    }
  }

  const targetCols = khTuanWbs_findMonthResultColumns_(monthSheet);
  const missing = [];

  if (!targetCols.tyLeHt) missing.push('Tỷ lệ HT (%)');
  if (!targetCols.ngayHoanThanh) missing.push('Ngày hoàn thành thực tế');
  if (!targetCols.danhGia) missing.push('Đánh giá');
  if (!targetCols.yKienChiDao) missing.push('Ý kiến chỉ đạo');

  if (missing.length) {
    ui.alert(
      'Không tìm đủ cột đích trên sheet tháng:\n- ' + missing.join('\n- ') +
      '\n\nDừng cập nhật để tránh ghi nhầm dữ liệu.'
    );
    return;
  }

  const rows = khTuanWbs_readExistingRows_(sheetTuan);
  const candidates = khTuanWbs_collectPushCandidates_(rows, sourceSheetName, targetCols);

  const confirmMessage = [
    'Bạn đang cập nhật kết quả ' + weekName + ' về sheet ' + sourceSheetName + '.',
    '',
    'Số dòng đủ điều kiện cập nhật: ' + candidates.toUpdate.length,
    'Số dòng bỏ qua do K không phải Có: ' + candidates.skipNotMarked,
    'Số dòng bỏ qua do không có SourceRow: ' + candidates.skipNoSourceRow,
    'Số dòng bỏ qua do không có dữ liệu H/I/J/L: ' + candidates.skipNoData,
    '',
    'Nguyên tắc cập nhật:',
    '- Hàng nào từ KH tháng đổ sang KH tuần thì trả về đúng hàng đó.',
    '- Không cập nhật cột G của KH tuần.',
    '- Ô trống trên KH tuần không ghi đè dữ liệu tháng.',
    '- Dòng việc tuần con/manual/carry không có SourceRow sẽ bị bỏ qua.',
    '',
    'Bạn có chắc chắn thực hiện không?'
  ].join('\n');

  const confirm = ui.alert(
    'Cập nhật kết quả tuần về KH tháng',
    confirmMessage,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    ss.toast('Đã hủy cập nhật kết quả tuần về KH tháng.', 'KH_TUAN_WBS', 5);
    return;
  }

  const result = khTuanWbs_applyPushCandidates_(monthSheet, candidates.toUpdate);

  ss.toast(
    'Đã cập nhật ' + result.updatedRows + ' dòng / ' + result.updatedCells + ' ô về ' + sourceSheetName + '.',
    'KH_TUAN_WBS',
    8
  );

  return {
    sourceSheetName: sourceSheetName,
    weekName: weekName,
    updatedRows: result.updatedRows,
    updatedCells: result.updatedCells,
    skipNotMarked: candidates.skipNotMarked,
    skipNoSourceRow: candidates.skipNoSourceRow,
    skipNoData: candidates.skipNoData
  };
}

function khTuanWbs_getMonthScopeWarning_(sourceSheetName) {
  const info = khTuanWbs_getMonthInfoFromSourceSheet_(sourceSheetName);
  if (!info) return '';

  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth() + 1;

  const selectedYear = Number(info.nam);
  const selectedMonth = Number(info.thang);

  const currentMonthIndex = currentYear * 12 + currentMonth;
  const selectedMonthIndex = selectedYear * 12 + selectedMonth;

  const currentLabel = khTuanWbs_formatMonthLabel_(currentYear, currentMonth);
  const previousDate = new Date(currentYear, currentMonth - 2, 1);
  const previousLabel = khTuanWbs_formatMonthLabel_(previousDate.getFullYear(), previousDate.getMonth() + 1);
  const selectedLabel = khTuanWbs_formatMonthLabel_(selectedYear, selectedMonth);

  // Cho qua nếu là tháng hiện tại hoặc tháng liền trước.
  if (selectedMonthIndex === currentMonthIndex || selectedMonthIndex === currentMonthIndex - 1) {
    return '';
  }

  return [
    'Bạn đang cập nhật kết quả về ' + sourceSheetName + ' (' + selectedLabel + ').',
    '',
    'Tháng hiện tại là: ' + currentLabel + '.',
    'Theo quy trình, chỉ nên cập nhật kết quả cho:',
    '- Tháng hiện tại: ' + currentLabel,
    '- Tháng liền trước: ' + previousLabel,
    '',
    'Bạn có chắc chắn muốn tiếp tục không?'
  ].join('\n');
}

function khTuanWbs_findMonthResultColumns_(monthSheet) {
  const lastCol = monthSheet.getLastColumn();
  const headerRows = Math.min(10, monthSheet.getMaxRows());
  const values = monthSheet.getRange(1, 1, headerRows, lastCol).getDisplayValues();

  const result = {
    tyLeHt: 0,
    ngayHoanThanh: 0,
    danhGia: 0,
    yKienChiDao: 0
  };

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const text = khTuanWbs_normalizeText_(values[r][c]);

      if (!result.tyLeHt && text.indexOf('tỷ lệ ht') !== -1) {
        result.tyLeHt = c + 1;
      }

      if (!result.ngayHoanThanh && text.indexOf('ngày hoàn thành thực tế') !== -1) {
        result.ngayHoanThanh = c + 1;
      }

      if (!result.danhGia && text === 'đánh giá') {
        result.danhGia = c + 1;
      }

      if (!result.yKienChiDao && text.indexOf('ý kiến chỉ đạo') !== -1) {
        result.yKienChiDao = c + 1;
      }
    }
  }

  return result;
}

function khTuanWbs_collectPushCandidates_(rows, expectedSourceSheet, targetCols) {
  const config = KH_TUAN_WBS_CONFIG;
  const col = config.COL;

  const output = {
    toUpdate: [],
    skipNotMarked: 0,
    skipNoSourceRow: 0,
    skipNoData: 0
  };

  (rows || []).forEach(function(item) {
    const row = khTuanWbs_normalizeRowLength_(item.values || item);

    const updateFlag = khTuanWbs_normalizeText_(row[col.CAP_NHAT - 1]);
    const sourceSheet = String(row[col.SOURCE_SHEET - 1] || '').trim();
    const sourceRow = Number(row[col.SOURCE_ROW - 1] || 0);

    if (updateFlag !== 'có' && updateFlag !== 'co') {
      output.skipNotMarked++;
      return;
    }

    if (!sourceSheet || sourceSheet !== expectedSourceSheet || !sourceRow) {
      output.skipNoSourceRow++;
      return;
    }

    const payload = {};

    const tyLeHt = row[col.TY_LE_HT - 1];
    const ngayHoanThanh = row[col.NGAY_HT - 1];
    const danhGia = row[col.DANH_GIA - 1];
    const yKienChiDao = row[col.Y_KIEN_CHI_DAO - 1];

    if (String(tyLeHt || '').trim() !== '') {
      payload[targetCols.tyLeHt] = khTuanWbs_normalizePercentDropdownText_(tyLeHt);
    }

    if (String(ngayHoanThanh || '').trim() !== '') {
      payload[targetCols.ngayHoanThanh] = ngayHoanThanh;
    }

    if (String(danhGia || '').trim() !== '') {
      payload[targetCols.danhGia] = danhGia;
    }

    if (String(yKienChiDao || '').trim() !== '') {
      payload[targetCols.yKienChiDao] = yKienChiDao;
    }

    if (!Object.keys(payload).length) {
      output.skipNoData++;
      return;
    }

    output.toUpdate.push({
      sourceRow: sourceRow,
      payload: payload
    });
  });

  return output;
}

function khTuanWbs_applyPushCandidates_(monthSheet, candidates) {
  let updatedRows = 0;
  let updatedCells = 0;

  (candidates || []).forEach(function(item) {
    const sourceRow = Number(item.sourceRow || 0);
    const payload = item.payload || {};

    if (!sourceRow) return;

    let rowUpdated = false;

    Object.keys(payload).forEach(function(colText) {
      const colIndex = Number(colText);
      const value = payload[colIndex];

      if (!colIndex) return;

      monthSheet.getRange(sourceRow, colIndex).setValue(value);
      updatedCells++;
      rowUpdated = true;
    });

    if (rowUpdated) updatedRows++;
  });

  return {
    updatedRows: updatedRows,
    updatedCells: updatedCells
  };
}

function khTuanWbs_normalizePercentForMonth_(value) {
  if (value === null || typeof value === 'undefined') return '';

  const raw = String(value).trim();
  if (!raw) return '';

  // Nếu người dùng nhập dạng "100%" hoặc "80%"
  if (raw.indexOf('%') !== -1) {
    const n = Number(raw.replace('%', '').replace(',', '.').trim());
    if (isNaN(n)) return value;
    return n / 100;
  }

  const n = Number(raw.toString().replace(',', '.'));
  if (isNaN(n)) return value;

  /*
   * Quy ước:
   * - Nếu nhập 0.8 hoặc 1 thì hiểu là tỷ lệ thật.
   * - Nếu nhập 80 hoặc 100 thì hiểu là phần trăm người dùng nhập.
   */
  if (n > 1) return n / 100;

  return n;
}

function khTuanWbs_normalizePercentDropdownText_(value) {
  if (value === null || typeof value === 'undefined') return '';

  const raw = String(value).trim();
  if (!raw) return '';

  let n;

  if (raw.indexOf('%') !== -1) {
    n = Number(raw.replace('%', '').replace(',', '.').trim());
  } else {
    n = Number(raw.replace(',', '.'));
    if (!isNaN(n) && n <= 1) {
      n = n * 100;
    }
  }

  if (isNaN(n)) return raw;

  // Làm tròn để khớp các option dropdown kiểu 0%, 10%, 20%, ...
  const rounded = Math.round(n);

  return String(rounded) + '%';
}

