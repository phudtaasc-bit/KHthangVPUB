/**
 * KH tuần WBS - UI và entry points
 * Tách từ: 20. KH tuan WBS.js
 * Lưu ý: không tự ý đổi tên hàm public nếu chưa cập nhật menu/onEdit.
 */
function onEdit(e) {
  khTuanWbs_onEdit(e);
}

function khTuanWbs_onEdit(e) {
  try {
    if (!e || !e.range) {
      return;
    }

    const range = e.range;
    const sheet = range.getSheet();
    const config = KH_TUAN_WBS_CONFIG;
    const a1 = range.getA1Notation();

    if (sheet.getName() !== config.SHEET_TUAN) {
      return;
    }
/*
     * Chỉ khi sửa B2/E2 mới lưu và nạp lại dữ liệu tuần.
     */
    if (a1 !== config.SOURCE_CELL && a1 !== config.WEEK_CELL) {
      return;
    }

    khTuanWbs_saveCurrentWeek_(e);

    if (a1 === config.SOURCE_CELL) {
      khTuanWbs_updateWeekDropdown_(sheet, String(range.getDisplayValue() || '').trim());
    }

    napKhTuanWbsTuSheetThang({ skipSaveCurrent: true, ignoreCurrentRows: true });
  } catch (error) {
    Logger.log('Lỗi khTuanWbs_onEdit: %s', error.stack || error);
    throw error;
  }
}


function napKhTuanWbsTuSheetThang(options) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = KH_TUAN_WBS_CONFIG;
    const opts = options || {};
    const sheetTuan = ss.getSheetByName(config.SHEET_TUAN);

    if (!sheetTuan) {
      throw new Error('Không tìm thấy sheet tuần: ' + config.SHEET_TUAN);
    }

    const sourceSheetName = String(sheetTuan.getRange(config.SOURCE_CELL).getDisplayValue() || '').trim();
    const weekName = String(sheetTuan.getRange(config.WEEK_CELL).getDisplayValue() || '').trim();

    if (!sourceSheetName) {
      throw new Error('Chưa chọn sheet tháng nguồn tại ' + config.SHEET_TUAN + '!' + config.SOURCE_CELL);
    }

    if (!weekName) {
      throw new Error('Chưa chọn tuần tại ' + config.SHEET_TUAN + '!' + config.WEEK_CELL);
    }

    if (opts.skipSaveCurrent !== true) {
      khTuanWbs_saveCurrentWeek_();
    }

    const monthSheet = ss.getSheetByName(sourceSheetName);
    if (!monthSheet) {
      throw new Error('Không tìm thấy sheet tháng nguồn: ' + sourceSheetName);
    }

    const savedRows = khTuanWbs_loadSavedWeek_(sourceSheetName, weekName);
    const newRowsFromMonth = khTuanWbs_buildRowsFromMonthSheet_(monthSheet, sourceSheetName, weekName);
    const carryForwardRows = khTuanWbs_getCarryForwardRows_(sourceSheetName, weekName);
    const baseExistingRows = savedRows.length
      ? khTuanWbs_wrapRows_(savedRows)
      : (opts.ignoreCurrentRows === true ? [] : khTuanWbs_readExistingRows_(sheetTuan));
    const mergedRows = khTuanWbs_mergeRows_(newRowsFromMonth, baseExistingRows, carryForwardRows);

    khTuanWbs_writeRowsToWeekSheet_(sheetTuan, mergedRows);
    khTuanWbs_applyFormats_(sheetTuan);

    ss.toast(
      'Đã nạp KH tuần WBS từ ' + sourceSheetName + ' - ' + weekName + '. Dòng: ' + mergedRows.length,
      'KH_TUAN_WBS',
      5
    );

    return {
      sourceSheet: sourceSheetName,
      weekName: weekName,
      soDongTuThang: newRowsFromMonth.length,
      soDongDaLuu: savedRows.length,
      soDongChuyenKy: carryForwardRows.length,
      soDongSauMerge: mergedRows.length
    };
  } catch (error) {
    Logger.log('Lỗi napKhTuanWbsTuSheetThang: %s', error.stack || error);
    throw error;
  }
}

function khTuanWbs_updateWeekDropdown_(sheetTuan, sourceSheetName) {
  const config = KH_TUAN_WBS_CONFIG;
  const thongTinThang = khTuanWbs_getMonthInfoFromSourceSheet_(sourceSheetName);
  if (!thongTinThang) {
    sheetTuan.getRange(config.WEEK_CELL).clearDataValidations();
    return [];
  }

  const weekNames = khTuanWbs_buildWorkingWeeks_(thongTinThang.nam, thongTinThang.thang)
    .map(function(item) {
      return item.name;
    });
  const weekRange = sheetTuan.getRange(config.WEEK_CELL);

  if (!weekNames.length) {
    weekRange.clearDataValidations();
    weekRange.clearContent();
    return [];
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(weekNames, true)
    .setAllowInvalid(false)
    .build();

  weekRange.setDataValidation(rule);
  weekRange.setValue(weekNames[0]);

  return weekNames;
}

function khTuanWbs_applyFormats_(sheetTuan) {
  const config = KH_TUAN_WBS_CONFIG;
  const soDong = Math.max(sheetTuan.getMaxRows() - config.START_ROW + 1, 1);

  sheetTuan.getRange(config.START_ROW, config.COL.DEADLINE, soDong, 1).setNumberFormat('dd/mm/yyyy');
  sheetTuan.getRange(config.START_ROW, config.COL.NGAY_HT, soDong, 1).setNumberFormat('dd/mm/yyyy');
  sheetTuan.getRange(config.START_ROW, config.COL.TY_LE_HT, soDong, 1).setNumberFormat('0"%"');
}

function khTuanWbs_writeRowsToWeekSheet_(sheetTuan, rows) {
  const config = KH_TUAN_WBS_CONFIG;
  const oldLastRow = khTuanWbs_getLastRowByColumns_(sheetTuan, config.START_ROW, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]);
  const oldRowCount = Math.max(oldLastRow - config.START_ROW + 1, 0);

  if (rows.length > 0) {
    sheetTuan.getRange(config.START_ROW, 1, rows.length, 1).setNumberFormat('@');
    sheetTuan.getRange(config.START_ROW, 1, rows.length, config.TOTAL_COLS).setValues(rows);
  }

  if (oldRowCount > rows.length) {
    sheetTuan
      .getRange(config.START_ROW + rows.length, 1, oldRowCount - rows.length, config.TOTAL_COLS)
      .clearContent();
  }
}

function khTuanWbs_wrapRows_(rows) {
  const wrapped = (rows || []).map(function(row, index) {
    return {
      rowIndex: index,
      values: khTuanWbs_normalizeRowLength_(row)
    };
  });

  return khTuanWbs_annotateParentForItems_(wrapped);
}

function khTuanWbs_getPreviousControlValue_(sheetTuan, a1, e) {
  const config = KH_TUAN_WBS_CONFIG;

  if (e && e.range && e.range.getSheet().getName() === config.SHEET_TUAN && e.range.getA1Notation() === a1) {
    return String(e.oldValue || '').trim();
  }

  return String(sheetTuan.getRange(a1).getDisplayValue() || '').trim();
}

