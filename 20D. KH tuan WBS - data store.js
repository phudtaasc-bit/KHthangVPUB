/**
 * KH tuần WBS - Data store
 * Tách từ: 20. KH tuan WBS.js
 * Lưu ý: không tự ý đổi tên hàm public nếu chưa cập nhật menu/onEdit.
 */
function khTuanWbs_saveCurrentWeek_(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = KH_TUAN_WBS_CONFIG;
  const sheetTuan = ss.getSheetByName(config.SHEET_TUAN);

  if (!sheetTuan) {
    throw new Error('Không tìm thấy sheet tuần: ' + config.SHEET_TUAN);
  }

  const sourceSheet = khTuanWbs_getPreviousControlValue_(sheetTuan, config.SOURCE_CELL, e);
  const weekName = khTuanWbs_getPreviousControlValue_(sheetTuan, config.WEEK_CELL, e);

  if (!sourceSheet || !weekName) {
    return { daLuu: false, lyDo: 'Thiếu SourceSheet hoặc WeekName.' };
  }

  const existingRows = khTuanWbs_readExistingRows_(sheetTuan)
    .filter(function(item) {
      return khTuanWbs_rowHasData_(item.values);
    });

  const dataSheet = khTuanWbs_getOrCreateDataSheet_(ss);
  const dataRows = khTuanWbs_readDataRows_(dataSheet);
  const keptRows = dataRows.filter(function(item) {
    return !(item.sourceSheet === sourceSheet && item.weekName === weekName);
  });
  const rowsToWrite = keptRows.map(function(item) {
    return item.raw;
  });
  const carryInitByTaskKey = khTuanWbs_indexCarryForwardRowsByTaskKey_(
    khTuanWbs_getCarryForwardRows_(sourceSheet, weekName, { ignoreCurrentRows: true })
  );

  let soDongManualKhongCoParent = 0;
  existingRows.forEach(function(item) {
    const normalizedRow = khTuanWbs_normalizeRowLength_(item.values);
    const sourceRow = String(normalizedRow[config.COL.SOURCE_ROW - 1] || '').trim();
    let taskKey = String(normalizedRow[config.COL.TASK_KEY - 1] || '').trim();
    let parentSourceRow = '';
    let parentTaskKey = '';
    let manualOrder = '';
    let rowOrigin = 'BASE';
    let carryFromWeek = '';
    let carryFromTaskKey = '';
    let isUserEdited = '';
    let editedAt = '';

    if (sourceRow && !taskKey) {
      taskKey = sourceSheet + '|' + weekName + '|R' + sourceRow;
      normalizedRow[config.COL.TASK_KEY - 1] = taskKey;
    }

    if (item.isManual) {
      parentSourceRow = String(item.parentSourceRow || '').trim();
      parentTaskKey = String(item.parentTaskKey || '').trim();
      manualOrder = Number(item.manualOrder || 0);

      if (!parentSourceRow || !parentTaskKey || !manualOrder) {
        soDongManualKhongCoParent++;
        return;
      }

      if (khTuanWbs_isCarryTaskKey_(taskKey)) {
        rowOrigin = 'CARRY_INIT';
        carryFromWeek = khTuanWbs_getPreviousWeekName_(weekName);
        carryFromTaskKey = khTuanWbs_extractCarryFromTaskKey_(taskKey);
      } else {
        taskKey = sourceSheet + '|' + weekName + '|MANUAL|PARENT:' + parentTaskKey + '|ORDER:' + manualOrder;
        rowOrigin = 'MANUAL';
      }
      normalizedRow[config.COL.TASK_KEY - 1] = taskKey;
    } else if (khTuanWbs_isCarryTaskKey_(taskKey)) {
      rowOrigin = 'CARRY_INIT';
      carryFromWeek = khTuanWbs_getPreviousWeekName_(weekName);
      carryFromTaskKey = khTuanWbs_extractCarryFromTaskKey_(taskKey);
    }

    if (khTuanWbs_isCarryTaskKey_(taskKey)) {
      const initRow = carryInitByTaskKey[taskKey];
      if (!initRow || khTuanWbs_hasUserEditableDiff_(normalizedRow, initRow)) {
        rowOrigin = 'CARRY_EDITED';
        isUserEdited = 'TRUE';
        editedAt = new Date();
      } else {
        rowOrigin = 'CARRY_INIT';
        isUserEdited = 'FALSE';
      }
    }

    if (!normalizedRow[config.COL.SOURCE_SHEET - 1]) {
      normalizedRow[config.COL.SOURCE_SHEET - 1] = sourceSheet;
    }

    rowsToWrite.push([sourceSheet, weekName, taskKey].concat(normalizedRow, [
      parentSourceRow,
      parentTaskKey,
      manualOrder,
      rowOrigin,
      carryFromWeek,
      carryFromTaskKey,
      isUserEdited,
      editedAt
    ]));
  });

  khTuanWbs_rewriteDataSheet_(dataSheet, rowsToWrite);

  return {
    daLuu: true,
    sourceSheet: sourceSheet,
    weekName: weekName,
    soDong: existingRows.length - soDongManualKhongCoParent,
    soDongManualKhongCoParent: soDongManualKhongCoParent
  };
}

function khTuanWbs_loadSavedWeek_(sourceSheet, weekName) {
  return khTuanWbs_loadSavedWeekItems_(sourceSheet, weekName)
    .map(function(item) {
      return item.values;
    });
}

function khTuanWbs_loadSavedWeekItems_(sourceSheet, weekName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(KH_TUAN_WBS_CONFIG.SHEET_DATA);

  if (!dataSheet) {
    return [];
  }

  const items = khTuanWbs_readDataRows_(dataSheet)
    .filter(function(item) {
      return item.sourceSheet === sourceSheet && item.weekName === weekName;
    });
  return khTuanWbs_orderSavedWeekItems_(items);
}

function khTuanWbs_clearCurrentWeekDataForTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = KH_TUAN_WBS_CONFIG;
  const sheetTuan = ss.getSheetByName(config.SHEET_TUAN);

  if (!sheetTuan) {
    throw new Error('Không tìm thấy sheet tuần: ' + config.SHEET_TUAN);
  }

  const sourceSheet = String(sheetTuan.getRange(config.SOURCE_CELL).getDisplayValue() || '').trim();
  const weekName = String(sheetTuan.getRange(config.WEEK_CELL).getDisplayValue() || '').trim();

  if (!sourceSheet || !weekName) {
    throw new Error('Chưa chọn SourceSheet hoặc WeekName tại ' + config.SHEET_TUAN + '.');
  }

  const dataSheet = khTuanWbs_getOrCreateDataSheet_(ss);
  const dataRows = khTuanWbs_readDataRows_(dataSheet);
  const keptRows = dataRows
    .filter(function(item) {
      return !(item.sourceSheet === sourceSheet && item.weekName === weekName);
    })
    .map(function(item) {
      return item.raw;
    });
  const removedCount = dataRows.length - keptRows.length;

  khTuanWbs_rewriteDataSheet_(dataSheet, keptRows);
  ss.toast(
    'Đã xóa dữ liệu test của ' + sourceSheet + ' - ' + weekName + '. Dòng xóa: ' + removedCount,
    'KH_TUAN_WBS',
    5
  );

  return {
    sourceSheet: sourceSheet,
    weekName: weekName,
    soDongDaXoa: removedCount
  };
}

function khTuanWbs_resetAllDataForTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KH_TUAN_WBS_CONFIG.SHEET_DATA);

  if (!sheet) {
    SpreadsheetApp.getActive().toast('Không tìm thấy KH_TUAN_DATA.', 'KH_TUAN_WBS', 6);
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1) {
    SpreadsheetApp.getActive().toast('KH_TUAN_DATA không có dữ liệu để xóa.', 'KH_TUAN_WBS', 6);
    return;
  }

  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  SpreadsheetApp.getActive().toast(
    'Đã xóa toàn bộ dữ liệu trong KH_TUAN_DATA. Header được giữ nguyên.',
    'KH_TUAN_WBS',
    8
  );
}

function khTuanWbs_getOrCreateDataSheet_(ss) {
  const config = KH_TUAN_WBS_CONFIG;
  let sheet = ss.getSheetByName(config.SHEET_DATA);

  if (!sheet) {
    sheet = ss.insertSheet(config.SHEET_DATA);
  }

  khTuanWbs_ensureDataSheetHeaders_(sheet);
  sheet.hideSheet();

  return sheet;
}

function khTuanWbs_readDataRows_(dataSheet) {
  const config = KH_TUAN_WBS_CONFIG;
  const totalCols = khTuanWbs_getDataTotalCols_();
  const lastRow = dataSheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  return dataSheet.getRange(2, 1, lastRow - 1, totalCols).getValues()
    .map(function(row) {
      const normalizedRaw = khTuanWbs_normalizeDataRawRow_(row);
      const values = khTuanWbs_normalizeRowLength_(normalizedRaw.slice(config.DATA_META_COLS, config.DATA_META_COLS + config.TOTAL_COLS));
      return {
        raw: normalizedRaw,
        sourceSheet: String(normalizedRaw[0] || '').trim(),
        weekName: String(normalizedRaw[1] || '').trim(),
        taskKey: String(normalizedRaw[2] || '').trim(),
        values: values,
        isManual: !String(values[config.COL.SOURCE_ROW - 1] || '').trim() && khTuanWbs_hasManualUserData_(values),
        parentSourceRow: String(normalizedRaw[config.DATA_META_COLS + config.TOTAL_COLS] || '').trim(),
        parentTaskKey: String(normalizedRaw[config.DATA_META_COLS + config.TOTAL_COLS + 1] || '').trim(),
        manualOrder: Number(normalizedRaw[config.DATA_META_COLS + config.TOTAL_COLS + 2] || 0),
        rowOrigin: String(normalizedRaw[config.DATA_META_COLS + config.TOTAL_COLS + 3] || '').trim(),
        carryFromWeek: String(normalizedRaw[config.DATA_META_COLS + config.TOTAL_COLS + 4] || '').trim(),
        carryFromTaskKey: String(normalizedRaw[config.DATA_META_COLS + config.TOTAL_COLS + 5] || '').trim(),
        isUserEdited: String(normalizedRaw[config.DATA_META_COLS + config.TOTAL_COLS + 6] || '').trim(),
        editedAt: normalizedRaw[config.DATA_META_COLS + config.TOTAL_COLS + 7] || ''
      };
    })
    .filter(function(item) {
      return item.sourceSheet || item.weekName || item.taskKey || khTuanWbs_rowHasData_(item.values);
    });
}

function khTuanWbs_rewriteDataSheet_(dataSheet, rowsToWrite) {
  const config = KH_TUAN_WBS_CONFIG;
  const totalCols = khTuanWbs_getDataTotalCols_();
  const oldRows = Math.max(dataSheet.getLastRow() - 1, 0);

  khTuanWbs_ensureDataSheetHeaders_(dataSheet);

  if (oldRows > 0) {
    dataSheet.getRange(2, 1, oldRows, totalCols).clearContent();
  }

  if (rowsToWrite.length > 0) {
    dataSheet.getRange(2, 1, rowsToWrite.length, totalCols).setValues(
      rowsToWrite.map(khTuanWbs_normalizeDataRawRow_)
    );
  }

  dataSheet.getRange(2, config.DATA_META_COLS + config.COL.DEADLINE, Math.max(dataSheet.getMaxRows() - 1, 1), 1).setNumberFormat('dd/mm/yyyy');
  dataSheet.getRange(2, config.DATA_META_COLS + config.COL.NGAY_HT, Math.max(dataSheet.getMaxRows() - 1, 1), 1).setNumberFormat('dd/mm/yyyy');
  dataSheet.getRange(2, config.DATA_META_COLS + config.COL.TY_LE_HT, Math.max(dataSheet.getMaxRows() - 1, 1), 1).setNumberFormat('0"%"');
}

function khTuanWbs_ensureDataSheetHeaders_(sheet) {
  const headers = [
    'SourceSheet',
    'WeekName',
    'TaskKey',
    'STT/WBS',
    'Công việc tháng',
    'Việc tuần',
    'Chủ trì',
    'Phê duyệt',
    'Deadline',
    'Kết quả / cập nhật thực hiện',
    'Tỷ lệ HT (%)',
    'Ngày hoàn thành thực tế',
    'Đánh giá',
    'Cập nhật vào KH tháng?',
    'Ý kiến chỉ đạo của Phụ trách khối',
    '_SourceSheet',
    '_SourceRow',
    '_MaDuAn',
    '_TenDuAn',
    '_LoaiDong',
    '_TaskKey',
    '_ParentSourceRow',
    '_ParentTaskKey',
    '_ManualOrder',
    '_RowOrigin',
    '_CarryFromWeek',
    '_CarryFromTaskKey',
    '_IsUserEdited',
    '_EditedAt'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function khTuanWbs_orderSavedWeekItems_(items) {
  const sourceItems = [];
  const manualByParentTaskKey = {};
  const orphanItems = [];

  (items || []).forEach(function(item) {
    const row = khTuanWbs_normalizeRowLength_(item.values);
    const sourceRow = String(row[KH_TUAN_WBS_CONFIG.COL.SOURCE_ROW - 1] || '').trim();

    if (sourceRow) {
      sourceItems.push(item);
      return;
    }

    if (item.isManual && item.parentTaskKey) {
      if (!manualByParentTaskKey[item.parentTaskKey]) {
        manualByParentTaskKey[item.parentTaskKey] = [];
      }
      manualByParentTaskKey[item.parentTaskKey].push(item);
    } else {
      orphanItems.push(item);
    }
  });

  const ordered = [];
  sourceItems
    .sort(function(a, b) {
      const rowA = khTuanWbs_normalizeRowLength_(a.values);
      const rowB = khTuanWbs_normalizeRowLength_(b.values);
      return Number(rowA[KH_TUAN_WBS_CONFIG.COL.SOURCE_ROW - 1] || 0) - Number(rowB[KH_TUAN_WBS_CONFIG.COL.SOURCE_ROW - 1] || 0);
    })
    .forEach(function(item) {
    ordered.push(item);
    const taskKey = String(item.values[KH_TUAN_WBS_CONFIG.COL.TASK_KEY - 1] || '').trim();
    const manualItems = manualByParentTaskKey[taskKey] || [];
    manualItems
      .sort(function(a, b) {
        return Number(a.manualOrder || 0) - Number(b.manualOrder || 0);
      })
      .forEach(function(manualItem) {
        ordered.push(manualItem);
      });
  });

  return ordered.concat(orphanItems.filter(function(item) {
    return !item.isManual;
  }));
}

function khTuanWbs_annotateParentForItems_(items) {
  const config = KH_TUAN_WBS_CONFIG;
  let currentParentSourceSheet = '';
  let currentParentSourceRow = '';
  let currentParentTaskKey = '';
  let manualOrder = 0;

  return (items || []).map(function(item) {
    const row = khTuanWbs_normalizeRowLength_(item.values);
    const sourceSheet = String(row[config.COL.SOURCE_SHEET - 1] || '').trim();
    const sourceRow = String(row[config.COL.SOURCE_ROW - 1] || '').trim();
    const taskKey = String(row[config.COL.TASK_KEY - 1] || '').trim();
    const loaiDong = String(row[config.COL.LOAI_DONG - 1] || '').trim();
    const isSourceRow = !!sourceRow;
    const isParentRow = isSourceRow && loaiDong === 'Gốc';
    const isManual = !isSourceRow && khTuanWbs_hasManualUserData_(row);

    item.values = row;
    item.isManual = isManual;
    item.parentSourceSheet = '';
    item.parentSourceRow = '';
    item.parentTaskKey = '';
    item.manualOrder = '';

    if (isParentRow) {
      currentParentSourceSheet = sourceSheet;
      currentParentSourceRow = sourceRow;
      currentParentTaskKey = taskKey;
      manualOrder = 0;
    } else if (isSourceRow) {
      currentParentSourceSheet = '';
      currentParentSourceRow = '';
      currentParentTaskKey = '';
      manualOrder = 0;
    } else if (isManual && currentParentSourceRow && currentParentTaskKey) {
      manualOrder++;
      item.parentSourceSheet = currentParentSourceSheet;
      item.parentSourceRow = currentParentSourceRow;
      item.parentTaskKey = currentParentTaskKey;
      item.manualOrder = manualOrder;
    }

    return item;
  });
}

function khTuanWbs_hasManualUserData_(row) {
  const config = KH_TUAN_WBS_CONFIG;
  const col = config.COL;

  if (!row) return false;

  const sourceRow = String(row[col.SOURCE_ROW - 1] || '').trim();
  const viecTuan = String(row[col.VIEC_TUAN - 1] || '').trim();

  /*
   * Quy tắc:
   * - Dòng manual/carry không có SourceRow thì bắt buộc phải có C - Việc tuần.
   * - Không lưu dòng manual trống chỉ có D/E/F hoặc dropdown K.
   * - D/E/F chỉ là dữ liệu bổ sung của một dòng việc tuần đã có C.
   */
  if (!sourceRow && !viecTuan) {
    return false;
  }

  const colsCanSave = [
    col.VIEC_TUAN,
    col.CHU_TRI,
    col.PHE_DUYET,
    col.DEADLINE,
    col.KET_QUA,
    col.TY_LE_HT,
    col.NGAY_HT,
    col.DANH_GIA,
    col.CAP_NHAT_KH_THANG,
    col.Y_KIEN_CHI_DAO
  ];

  return colsCanSave.some(function(columnIndex) {
    return String(row[columnIndex - 1] || '').trim() !== '';
  });
}

function khTuanWbs_buildParentSourceKey_(sourceSheet, sourceRow) {
  const s = String(sourceSheet || '').trim();
  const r = String(sourceRow || '').trim();
  return s && r ? s + '|R' + r : '';
}

function khTuanWbs_hasUserEditableDiff_(rowA, rowB) {
  const config = KH_TUAN_WBS_CONFIG;
  const normalizedA = khTuanWbs_normalizeRowLength_(rowA);
  const normalizedB = khTuanWbs_normalizeRowLength_(rowB);
  const editableCols = [
    config.COL.VIEC_TUAN,
    config.COL.CHU_TRI,
    config.COL.PHE_DUYET,
    config.COL.DEADLINE,
    config.COL.KET_QUA,
    config.COL.TY_LE_HT,
    config.COL.NGAY_HT,
    config.COL.DANH_GIA,
    config.COL.CAP_NHAT,
    config.COL.Y_KIEN_CHI_DAO
  ];

  return editableCols.some(function(col) {
    return !khTuanWbs_valuesEqual_(normalizedA[col - 1], normalizedB[col - 1]);
  });
}

function khTuanWbs_readExistingRows_(sheetTuan) {
  const config = KH_TUAN_WBS_CONFIG;
  const lastRow = khTuanWbs_getLastRowByColumns_(sheetTuan, config.START_ROW, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]);
  const soDong = Math.max(lastRow - config.START_ROW + 1, 0);

  if (soDong <= 0) {
    return [];
  }

  let currentParentSourceSheet = '';
  let currentParentSourceRow = '';
  let currentParentTaskKey = '';
  let manualOrder = 0;

  return sheetTuan
    .getRange(config.START_ROW, 1, soDong, config.TOTAL_COLS)
    .getValues()
    .map(function(row, index) {
      const normalizedRow = khTuanWbs_normalizeRowLength_(row);
      const sourceSheet = String(normalizedRow[config.COL.SOURCE_SHEET - 1] || '').trim();
      const sourceRow = String(normalizedRow[config.COL.SOURCE_ROW - 1] || '').trim();
      const taskKey = String(normalizedRow[config.COL.TASK_KEY - 1] || '').trim();
      const loaiDong = String(normalizedRow[config.COL.LOAI_DONG - 1] || '').trim();
      const isSourceRow = !!sourceRow;
      const isParentRow = isSourceRow && loaiDong === 'Gốc';
      const isManual = !isSourceRow && khTuanWbs_hasManualUserData_(normalizedRow);
      let parentSourceSheet = '';
      let parentSourceRow = '';
      let parentTaskKey = '';
      let itemManualOrder = '';

      if (isParentRow) {
        currentParentSourceSheet = sourceSheet;
        currentParentSourceRow = sourceRow;
        currentParentTaskKey = taskKey;
        manualOrder = 0;
      } else if (isSourceRow) {
        currentParentSourceSheet = '';
        currentParentSourceRow = '';
        currentParentTaskKey = '';
        manualOrder = 0;
      } else if (isManual && currentParentSourceRow && currentParentTaskKey) {
        manualOrder++;
        parentSourceSheet = currentParentSourceSheet;
        parentSourceRow = currentParentSourceRow;
        parentTaskKey = currentParentTaskKey;
        itemManualOrder = manualOrder;
      }

      return {
        rowIndex: index,
        sheetRow: config.START_ROW + index,
        values: normalizedRow,
        isManual: isManual,
        parentSourceSheet: parentSourceSheet,
        parentSourceRow: parentSourceRow,
        parentTaskKey: parentTaskKey,
        manualOrder: itemManualOrder
      };
    });
}

