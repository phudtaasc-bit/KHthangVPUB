/**
 * KH tuần WBS - Carry và merge
 * Viết lại theo luồng đơn giản:
 * baseRows tháng + savedRows tuần hiện tại + carryRows tuần trước → finalRows
 */

function khTuanWbs_getCarryForwardRows_(sourceSheet, weekName, options) {
  const opts = options || {};
  const weekIndex = khTuanWbs_getWeekIndex_(weekName);

  if (weekIndex <= 1) return [];

  const thongTinThang = khTuanWbs_getMonthInfoFromSourceSheet_(sourceSheet);
  if (!thongTinThang) return [];

  const dsTuan = khTuanWbs_buildWorkingWeeks_(thongTinThang.nam, thongTinThang.thang);
  if (weekIndex > dsTuan.length) return [];

  const previousWeek = khTuanWbs_getPreviousWeekName_(weekName);
  if (!previousWeek || khTuanWbs_getWeekIndex_(previousWeek) < 1) return [];

  const config = KH_TUAN_WBS_CONFIG;
  const col = config.COL;

  const previousRows = khTuanWbs_loadSavedWeekItems_(sourceSheet, previousWeek);
  const currentRows = opts.ignoreCurrentRows === true
    ? []
    : khTuanWbs_loadSavedWeekItems_(sourceSheet, weekName);

  const currentTaskKeys = {};
  const currentCarryFromTaskKeys = {};

  currentRows.forEach(function(item) {
    const row = khTuanWbs_normalizeRowLength_(item.values);
    const taskKey = String(item.taskKey || row[col.TASK_KEY - 1] || '').trim();
    const carryFromTaskKey = String(
      item.carryFromTaskKey ||
      khTuanWbs_extractCarryFromTaskKey_(taskKey) ||
      ''
    ).trim();

    if (taskKey) currentTaskKeys[taskKey] = true;
    if (carryFromTaskKey) currentCarryFromTaskKeys[carryFromTaskKey] = true;
  });

  const result = [];
  const parentOrderCounter = {};

  previousRows.forEach(function(item) {
    const oldRow = khTuanWbs_normalizeRowLength_(item.values);

    const viecTuan = String(oldRow[col.VIEC_TUAN - 1] || '').trim();
    const chuTri = String(oldRow[col.CHU_TRI - 1] || '').trim();
    const pheDuyet = String(oldRow[col.PHE_DUYET - 1] || '').trim();

    /*
     * Quy tắc nghiệp vụ:
     * Chỉ chuyển dòng có đủ C/D/E.
     * Không chuyển dòng nền tháng/dự án/tổng.
     */
    if (!viecTuan || !chuTri || !pheDuyet) return;

    if (!khTuanWbs_isCarryForwardStatus_(oldRow[col.DANH_GIA - 1])) return;

    const oldTaskKey = String(item.taskKey || oldRow[col.TASK_KEY - 1] || '').trim();
    if (!oldTaskKey) return;

    if (currentCarryFromTaskKeys[oldTaskKey]) return;

    let parentSourceRow = String(item.parentSourceRow || '').trim();

    if (!parentSourceRow) {
      const oldParentTaskKey = String(item.parentTaskKey || '').trim();
      const matched = oldParentTaskKey.match(/\|R(\d+)\b/);
      if (matched && matched[1]) parentSourceRow = matched[1];
    }

    if (!parentSourceRow) return;

    const parentTaskKey = sourceSheet + '|' + weekName + '|R' + parentSourceRow;
    const parentKeyForOrder = khTuanWbs_buildParentSourceKey_(sourceSheet, parentSourceRow);

    parentOrderCounter[parentKeyForOrder] = (parentOrderCounter[parentKeyForOrder] || 0) + 1;
    const manualOrder = parentOrderCounter[parentKeyForOrder];

    const newTaskKey =
      sourceSheet +
      '|' +
      weekName +
      '|CARRY|PARENT:' +
      parentTaskKey +
      '|FROM:' +
      oldTaskKey;

    if (currentTaskKeys[newTaskKey]) return;

    const newRow = khTuanWbs_createBlankRow_();

    // Dòng carry là dòng con, không hiển thị A/B.
    newRow[col.WBS - 1] = '';
    newRow[col.CV_THANG - 1] = '';

    // Chỉ giữ C/D/E từ tuần trước.
    newRow[col.VIEC_TUAN - 1] = oldRow[col.VIEC_TUAN - 1];
    newRow[col.CHU_TRI - 1] = oldRow[col.CHU_TRI - 1];
    newRow[col.PHE_DUYET - 1] = oldRow[col.PHE_DUYET - 1];

    // Không chuyển deadline/kết quả/đánh giá sang tuần sau.
    newRow[col.DEADLINE - 1] = '';
    newRow[col.KET_QUA - 1] = '';
    newRow[col.TY_LE_HT - 1] = '';
    newRow[col.NGAY_HT - 1] = '';
    newRow[col.DANH_GIA - 1] = '';
    newRow[col.Y_KIEN_CHI_DAO - 1] = '';

    newRow[col.SOURCE_SHEET - 1] = sourceSheet;
    newRow[col.SOURCE_ROW - 1] = '';
    newRow[col.MA_DU_AN - 1] = oldRow[col.MA_DU_AN - 1] || '';
    newRow[col.TEN_DU_AN - 1] = oldRow[col.TEN_DU_AN - 1] || '';
    newRow[col.LOAI_DONG - 1] = 'Carry';
    newRow[col.TASK_KEY - 1] = newTaskKey;

    result.push({
      values: newRow,
      sourceSheet: sourceSheet,
      weekName: weekName,
      taskKey: newTaskKey,
      isManual: true,
      parentSourceSheet: sourceSheet,
      parentSourceRow: parentSourceRow,
      parentTaskKey: parentTaskKey,
      manualOrder: manualOrder,
      rowOrigin: 'CARRY',
      carryFromWeek: previousWeek,
      carryFromTaskKey: oldTaskKey,
      isUserEdited: '',
      editedAt: '',
      raw: null
    });
  });

  return result;
}

function khTuanWbs_mergeRows_(newRowsFromMonth, existingRows, carryForwardRows) {
  const config = KH_TUAN_WBS_CONFIG;
  const col = config.COL;

  const existingBySourceKey = {};
  const savedManualByParentKey = {};
  const savedCarryFromKeys = {};
  const savedTaskKeys = {};

  /*
   * existingRows là dữ liệu tuần hiện tại:
   * - có thể là savedRows từ KH_TUAN_DATA;
   * - hoặc dữ liệu đang có trên sheet tuần.
   * Dữ liệu tuần hiện tại phải thắng carry.
   */
  (existingRows || []).forEach(function(item) {
    const row = khTuanWbs_normalizeRowLength_(item.values || item);
    const sourceSheet = String(row[col.SOURCE_SHEET - 1] || item.sourceSheet || '').trim();
    const sourceRow = String(row[col.SOURCE_ROW - 1] || '').trim();
    const taskKey = String(item.taskKey || row[col.TASK_KEY - 1] || '').trim();
    const carryFromTaskKey = String(
      item.carryFromTaskKey ||
      khTuanWbs_extractCarryFromTaskKey_(taskKey) ||
      ''
    ).trim();

    if (taskKey) savedTaskKeys[taskKey] = true;
    if (carryFromTaskKey) savedCarryFromKeys[carryFromTaskKey] = true;

    if (sourceSheet && sourceRow) {
      existingBySourceKey[khTuanWbs_buildParentSourceKey_(sourceSheet, sourceRow)] = row;
      return;
    }

    if (!khTuanWbs_hasManualUserData_(row)) return;

    const parentSourceSheet = String(item.parentSourceSheet || sourceSheet || '').trim();
    const parentSourceRow = String(item.parentSourceRow || '').trim();
    const parentKey = khTuanWbs_buildParentSourceKey_(parentSourceSheet, parentSourceRow);

    if (!parentKey) return;

    if (!savedManualByParentKey[parentKey]) savedManualByParentKey[parentKey] = [];

    savedManualByParentKey[parentKey].push({
      values: row,
      manualOrder: Number(item.manualOrder || 0),
      taskKey: taskKey,
      carryFromTaskKey: carryFromTaskKey
    });
  });

  Object.keys(savedManualByParentKey).forEach(function(parentKey) {
    savedManualByParentKey[parentKey].sort(function(a, b) {
      return Number(a.manualOrder || 0) - Number(b.manualOrder || 0);
    });
  });

  const carryManualByParentKey = {};

  (carryForwardRows || []).forEach(function(item) {
    const row = khTuanWbs_normalizeRowLength_(item.values || item);
    const taskKey = String(item.taskKey || row[col.TASK_KEY - 1] || '').trim();
    const carryFromTaskKey = String(
      item.carryFromTaskKey ||
      khTuanWbs_extractCarryFromTaskKey_(taskKey) ||
      ''
    ).trim();

    if (!khTuanWbs_hasManualUserData_(row)) return;
    if (taskKey && savedTaskKeys[taskKey]) return;
    if (carryFromTaskKey && savedCarryFromKeys[carryFromTaskKey]) return;

    const parentKey = khTuanWbs_buildParentSourceKey_(item.parentSourceSheet || row[col.SOURCE_SHEET - 1], item.parentSourceRow);
    if (!parentKey) return;

    if (!carryManualByParentKey[parentKey]) carryManualByParentKey[parentKey] = [];

    carryManualByParentKey[parentKey].push({
      values: row,
      manualOrder: Number(item.manualOrder || 0),
      taskKey: taskKey,
      carryFromTaskKey: carryFromTaskKey
    });
  });

  Object.keys(carryManualByParentKey).forEach(function(parentKey) {
    carryManualByParentKey[parentKey].sort(function(a, b) {
      return Number(a.manualOrder || 0) - Number(b.manualOrder || 0);
    });
  });

  const merged = [];
  const addedTaskKeys = {};

  (newRowsFromMonth || []).forEach(function(newRow) {
    const baseRow = khTuanWbs_normalizeRowLength_(newRow);
    const sourceSheet = String(baseRow[col.SOURCE_SHEET - 1] || '').trim();
    const sourceRow = String(baseRow[col.SOURCE_ROW - 1] || '').trim();
    const parentKey = khTuanWbs_buildParentSourceKey_(sourceSheet, sourceRow);
    const taskKey = String(baseRow[col.TASK_KEY - 1] || '').trim();

    const oldRow = parentKey ? existingBySourceKey[parentKey] : null;
    const rowToPush = oldRow ? khTuanWbs_mergeOneSourceRow_(baseRow, oldRow) : baseRow;

    merged.push(rowToPush);

    if (taskKey) addedTaskKeys[taskKey] = true;
    khTuanWbs_markTaskKey_(addedTaskKeys, rowToPush);

    /*
     * Chỉ dòng Gốc mới được chèn manual/carry bên dưới.
     * Dòng Dự án/Tổng không chèn.
     */
    const loaiDong = String(baseRow[col.LOAI_DONG - 1] || '').trim();
    if (loaiDong !== 'Gốc') return;

    const savedManuals = savedManualByParentKey[parentKey] || [];
    savedManuals.forEach(function(manualItem) {
      const row = khTuanWbs_normalizeRowLength_(manualItem.values);
      const rowTaskKey = String(manualItem.taskKey || row[col.TASK_KEY - 1] || '').trim();

      if (rowTaskKey && addedTaskKeys[rowTaskKey]) return;

      merged.push(row);

      if (rowTaskKey) addedTaskKeys[rowTaskKey] = true;
      khTuanWbs_markTaskKey_(addedTaskKeys, row);
    });

    const carryManuals = carryManualByParentKey[parentKey] || [];
    carryManuals.forEach(function(carryItem) {
      const row = khTuanWbs_mergeOneCarryForwardRow_(khTuanWbs_createBlankRow_(), carryItem.values);
      const rowTaskKey = String(carryItem.taskKey || row[col.TASK_KEY - 1] || '').trim();
      const carryFromTaskKey = String(
        carryItem.carryFromTaskKey ||
        khTuanWbs_extractCarryFromTaskKey_(rowTaskKey) ||
        ''
      ).trim();

      if (rowTaskKey && addedTaskKeys[rowTaskKey]) return;
      if (carryFromTaskKey && savedCarryFromKeys[carryFromTaskKey]) return;

      merged.push(row);

      if (rowTaskKey) addedTaskKeys[rowTaskKey] = true;
      khTuanWbs_markTaskKey_(addedTaskKeys, row);
    });
  });

  return merged;
}

function khTuanWbs_isCarryForwardStatus_(value) {
  const text = khTuanWbs_normalizeText_(value);

  // Đánh giá trống nghĩa là chưa kết luận, phải chuyển.
  if (!text) return true;

  if (text.indexOf('không đạt') !== -1) return true;
  if (text.indexOf('khong dat') !== -1) return true;

  if (text.indexOf('hoàn thành một phần') !== -1) return true;
  if (text.indexOf('hoan thanh mot phan') !== -1) return true;

  if (text.indexOf('chuyển kỳ sau') !== -1) return true;
  if (text.indexOf('chuyen ky sau') !== -1) return true;

  if (text.indexOf('chưa bắt đầu') !== -1) return true;
  if (text.indexOf('chua bat dau') !== -1) return true;

  if (text.indexOf('chưa thực hiện') !== -1) return true;
  if (text.indexOf('chua thuc hien') !== -1) return true;

  // Phải đặt sau "không đạt".
  if (text.indexOf('đạt') !== -1) return false;
  if (text.indexOf('dat') !== -1) return false;

  return false;
}

function khTuanWbs_mergeOneSourceRow_(newRow, oldRow) {
  const config = KH_TUAN_WBS_CONFIG;
  const col = config.COL;
  const merged = khTuanWbs_normalizeRowLength_(newRow).slice();
  const normalizedOldRow = khTuanWbs_normalizeRowLength_(oldRow);

  /*
   * Dòng gốc tháng:
   * - Giữ cấu trúc A/B/D/E/F/M:R từ tháng mới.
   * - Giữ cập nhật thực hiện G/H/I/J/K/L nếu người dùng đã nhập trong tuần hiện tại.
   */

  const userCols = [
    col.VIEC_TUAN,
    col.KET_QUA,
    col.TY_LE_HT,
    col.NGAY_HT,
    col.DANH_GIA,
    col.CAP_NHAT,
    col.Y_KIEN_CHI_DAO
  ];

  userCols.forEach(function(c) {
    if (String(normalizedOldRow[c - 1] || '').trim() !== '') {
      merged[c - 1] = normalizedOldRow[c - 1];
    }
  });

  return merged;
}

function khTuanWbs_mergeOneCarryForwardRow_(newRow, carryRow) {
  const config = KH_TUAN_WBS_CONFIG;
  const col = config.COL;

  const merged = khTuanWbs_normalizeRowLength_(newRow).slice();
  const normalizedCarryRow = khTuanWbs_normalizeRowLength_(carryRow);

  /*
   * Carry chỉ hiển thị lại C/D/E.
   * Không hiển thị F/G/H/I/J/L của tuần trước.
   */
  merged[col.WBS - 1] = '';
  merged[col.CV_THANG - 1] = '';

  merged[col.VIEC_TUAN - 1] = normalizedCarryRow[col.VIEC_TUAN - 1];
  merged[col.CHU_TRI - 1] = normalizedCarryRow[col.CHU_TRI - 1];
  merged[col.PHE_DUYET - 1] = normalizedCarryRow[col.PHE_DUYET - 1];

  merged[col.DEADLINE - 1] = '';
  merged[col.KET_QUA - 1] = '';
  merged[col.TY_LE_HT - 1] = '';
  merged[col.NGAY_HT - 1] = '';
  merged[col.DANH_GIA - 1] = '';
  merged[col.Y_KIEN_CHI_DAO - 1] = '';

  merged[col.SOURCE_SHEET - 1] = normalizedCarryRow[col.SOURCE_SHEET - 1];
  merged[col.SOURCE_ROW - 1] = '';
  merged[col.MA_DU_AN - 1] = normalizedCarryRow[col.MA_DU_AN - 1];
  merged[col.TEN_DU_AN - 1] = normalizedCarryRow[col.TEN_DU_AN - 1];
  merged[col.LOAI_DONG - 1] = 'Carry';
  merged[col.TASK_KEY - 1] = normalizedCarryRow[col.TASK_KEY - 1];

  return merged;
}

function khTuanWbs_indexCarryForwardRowsByTaskKey_(carryForwardRows) {
  const config = KH_TUAN_WBS_CONFIG;
  const result = {};

  (carryForwardRows || []).forEach(function(item) {
    const row = khTuanWbs_normalizeRowLength_(item.values || item);
    const taskKey = String(row[config.COL.TASK_KEY - 1] || '').trim();
    if (taskKey) result[taskKey] = row;
  });

  return result;
}

function khTuanWbs_markTaskKey_(addedTaskKeys, row) {
  const config = KH_TUAN_WBS_CONFIG;
  const normalizedRow = khTuanWbs_normalizeRowLength_(row);
  const taskKey = String(normalizedRow[config.COL.TASK_KEY - 1] || '').trim();

  if (taskKey) {
    addedTaskKeys[taskKey] = true;
  }
}

function khTuanWbs_isCarryTaskKey_(taskKey) {
  return String(taskKey || '').indexOf('|CARRY|') !== -1;
}

function khTuanWbs_extractCarryFromTaskKey_(taskKey) {
  const text = String(taskKey || '');
  const marker = '|FROM:';
  const index = text.indexOf(marker);

  if (index === -1) return '';

  return text.substring(index + marker.length).trim();
}

function khTuanWbs_makeStableManualKey_(row) {
  const config = KH_TUAN_WBS_CONFIG;
  const normalizedRow = khTuanWbs_normalizeRowLength_(row);

  return [
    String(normalizedRow[config.COL.VIEC_TUAN - 1] || '').trim(),
    String(normalizedRow[config.COL.CHU_TRI - 1] || '').trim(),
    String(normalizedRow[config.COL.PHE_DUYET - 1] || '').trim()
  ].join('|');
}

function khTuanWbs_hasWeekTaskForCarry_(row) {
  const config = KH_TUAN_WBS_CONFIG;
  const col = config.COL;
  const normalizedRow = khTuanWbs_normalizeRowLength_(row);

  return !!(
    String(normalizedRow[col.VIEC_TUAN - 1] || '').trim() &&
    String(normalizedRow[col.CHU_TRI - 1] || '').trim() &&
    String(normalizedRow[col.PHE_DUYET - 1] || '').trim()
  );
}
