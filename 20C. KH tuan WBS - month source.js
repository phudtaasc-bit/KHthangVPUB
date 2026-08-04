/**
 * KH tuần WBS - Đọc kế hoạch tháng
 * Tách từ: 20. KH tuan WBS.js
 * Lưu ý: không tự ý đổi tên hàm public nếu chưa cập nhật menu/onEdit.
 */
function khTuanWbs_buildRowsFromMonthSheet_(monthSheet, sourceSheetName, weekName) {
  const config = KH_TUAN_WBS_CONFIG;
  const startRow = config.START_ROW;
  const doneSourceRows = khTuanWbs_buildDoneSourceRowsBeforeWeek_(sourceSheetName, weekName);
  const lastRow = khTuanWbs_getLastRowByColumns_(monthSheet, startRow, [
    config.MONTH_COL.WBS,
    config.MONTH_COL.NOI_DUNG
  ]);
  const soDong = Math.max(lastRow - startRow + 1, 0);

  if (soDong <= 0) {
    return [];
  }

  const rawValues = monthSheet.getRange(startRow, 1, soDong, config.MONTH_COL.PHE_DUYET).getValues();
  const displayWbsValues = monthSheet.getRange(startRow, config.MONTH_COL.WBS, soDong, 1).getDisplayValues();
  const displayTextValues = monthSheet.getRange(startRow, config.MONTH_COL.NOI_DUNG, soDong, 2).getDisplayValues();
  const result = [];
  let maDuAnHienTai = '';
  let tenDuAnHienTai = '';

  rawValues.forEach(function(row, index) {
    const sourceRow = startRow + index;
    const wbs = String(displayWbsValues[index][0] || '').trim();
    const noiDung = String(displayTextValues[index][0] || '').trim();
    const maDuAn = String(displayTextValues[index][1] || '').trim();

    if (!wbs && !noiDung) {
      return;
    }

    if (khTuanWbs_normalizeText_(noiDung) === khTuanWbs_normalizeText_(config.TEN_DONG_PHAT_SINH_KHAC)) {
      return;
    }

    const outputRow = khTuanWbs_createBlankRow_();
    const taskKey = sourceSheetName + '|' + weekName + '|R' + sourceRow;

    outputRow[config.COL.WBS - 1] = String(wbs || '');
    outputRow[config.COL.SOURCE_SHEET - 1] = sourceSheetName;
    outputRow[config.COL.SOURCE_ROW - 1] = sourceRow;
    outputRow[config.COL.TASK_KEY - 1] = taskKey;

    if (wbs === '∑' || wbs === 'Σ') {
      outputRow[config.COL.WBS - 1] = '∑';
      outputRow[config.COL.CV_THANG - 1] = noiDung;
      outputRow[config.COL.CAP_NHAT - 1] = 'Không';
      outputRow[config.COL.MA_DU_AN - 1] = maDuAnHienTai;
      outputRow[config.COL.TEN_DU_AN - 1] = tenDuAnHienTai;
      outputRow[config.COL.LOAI_DONG - 1] = 'Tổng';
      result.push(outputRow);
      return;
    }

    if (wbs === 'Dự án' || khTuanWbs_isRoman_(wbs)) {
      maDuAnHienTai = maDuAn;
      tenDuAnHienTai = noiDung;

      outputRow[config.COL.CV_THANG - 1] = maDuAn ? maDuAn + ' | ' + noiDung : noiDung;
      outputRow[config.COL.CAP_NHAT - 1] = 'Không';
      outputRow[config.COL.MA_DU_AN - 1] = maDuAnHienTai;
      outputRow[config.COL.TEN_DU_AN - 1] = tenDuAnHienTai;
      outputRow[config.COL.LOAI_DONG - 1] = 'Dự án';
      result.push(outputRow);
      return;
    }

    if (khTuanWbs_isNumericWbs_(wbs)) {
      if (doneSourceRows[String(sourceRow)]) {
        return;
      }

      outputRow[config.COL.CV_THANG - 1] = noiDung;
      outputRow[config.COL.CHU_TRI - 1] = row[config.MONTH_COL.CHU_TRI - 1] || '';
      outputRow[config.COL.PHE_DUYET - 1] = row[config.MONTH_COL.PHE_DUYET - 1] || '';
      outputRow[config.COL.DEADLINE - 1] = row[config.MONTH_COL.DEADLINE - 1] || '';
      outputRow[config.COL.CAP_NHAT - 1] = 'Có';
      outputRow[config.COL.MA_DU_AN - 1] = maDuAnHienTai;
      outputRow[config.COL.TEN_DU_AN - 1] = tenDuAnHienTai;
      outputRow[config.COL.LOAI_DONG - 1] = 'Gốc';
      result.push(outputRow);
    }
  });

  return result;
}

function khTuanWbs_buildDoneSourceRowsBeforeWeek_(sourceSheetName, weekName) {
  const weekIndex = khTuanWbs_getWeekIndex_(weekName);
  const doneSourceRows = {};

  if (weekIndex <= 1) {
    return doneSourceRows;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(KH_TUAN_WBS_CONFIG.SHEET_DATA);

  if (!dataSheet) {
    return doneSourceRows;
  }

  const config = KH_TUAN_WBS_CONFIG;
  const col = config.COL;

  khTuanWbs_readDataRows_(dataSheet).forEach(function(item) {
    if (item.sourceSheet !== sourceSheetName) return;

    const itemWeekIndex = khTuanWbs_getWeekIndex_(item.weekName);
    if (itemWeekIndex < 1 || itemWeekIndex >= weekIndex) return;

    const row = khTuanWbs_normalizeRowLength_(item.values);
    const sourceRow = String(row[col.SOURCE_ROW - 1] || '').trim();

    if (!sourceRow) return;
    if (!khTuanWbs_isDoneStatus_(row[col.DANH_GIA - 1])) return;

    doneSourceRows[sourceRow] = true;
  });

  return doneSourceRows;
}
