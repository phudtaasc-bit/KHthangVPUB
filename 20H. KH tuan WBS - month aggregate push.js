/**
 * KH tuần WBS - Cập nhật KH tháng theo dữ liệu tuần đã lưu
 *
 * v0.5:
 * - Quét KH_TUAN_DATA theo đúng sheet tháng đang chọn tại B2.
 * - Tháng hiện tại: quét từ Tuần 1 đến tuần chứa ngày hôm nay.
 * - Tháng cũ: quét toàn bộ tháng để hỗ trợ cập nhật/chốt hồi cứu.
 * - Tháng tương lai: chặn để tránh ghi sớm.
 * - Chỉ ghi các dòng có _SourceRow và Cập nhật vào KH tháng? = Có.
 */
function khTuanWbs_pushMonthResultFromWeekData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const config = KH_TUAN_WBS_CONFIG;
  const sheetTuan = ss.getSheetByName(config.SHEET_TUAN);

  if (!sheetTuan) {
    ui.alert('Không tìm thấy sheet KH_TUAN_WBS.');
    return;
  }

  const sourceSheetName = String(sheetTuan.getRange(config.SOURCE_CELL).getDisplayValue() || '').trim();
  if (!sourceSheetName) {
    ui.alert('Thiếu nguồn KH tháng tại ' + config.SHEET_TUAN + '!' + config.SOURCE_CELL + '.');
    return;
  }

  const monthSheet = ss.getSheetByName(sourceSheetName);
  if (!monthSheet) {
    ui.alert('Không tìm thấy sheet tháng: ' + sourceSheetName);
    return;
  }

  const scope = khTuanWbs_getMonthPushScope_(sourceSheetName, new Date());
  if (!scope.ok) {
    ui.alert('Không thể cập nhật KH tháng', scope.message, ui.ButtonSet.OK);
    return;
  }

  const dataSheet = ss.getSheetByName(config.SHEET_DATA);
  if (!dataSheet) {
    ui.alert('Không tìm thấy sheet dữ liệu tuần: ' + config.SHEET_DATA + '. Chưa có dữ liệu để cập nhật.');
    return;
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

  try {
    khTuanWbs_saveCurrentWeek_();
    SpreadsheetApp.flush();
  } catch (err) {
    ui.alert(
      'Không thể lưu dữ liệu tuần hiện tại trước khi cập nhật KH tháng',
      String(err && err.message ? err.message : err),
      ui.ButtonSet.OK
    );
    return;
  }

  const dataRows = khTuanWbs_readDataRows_(dataSheet);
  const candidates = khTuanWbs_collectMonthPushCandidatesFromWeekData_(
    dataRows,
    sourceSheetName,
    scope.maxWeekIndex,
    targetCols
  );
  const sourceRowsToUpdate = candidates.toUpdate
    .map(function(item) { return item.sourceRow; })
    .join(', ');

  const confirmMessage = [
    scope.message,
    '',
    'Phạm vi tuần: ' + (scope.weekNames.length ? scope.weekNames.join(', ') : 'Không có'),
    'Số dòng đủ điều kiện cập nhật: ' + candidates.toUpdate.length,
    'SourceRow sẽ cập nhật: ' + (sourceRowsToUpdate || 'Không có'),
    'Số dòng bỏ qua do khác tháng/ngoài phạm vi: ' + candidates.skipOutOfScope,
    'Số dòng bỏ qua do không có SourceRow: ' + candidates.skipNoSourceRow,
    'Số dòng bỏ qua do Cập nhật vào KH tháng? không phải Có: ' + candidates.skipNotMarked,
    'Số dòng bỏ qua do không có dữ liệu kết quả: ' + candidates.skipNoData,
    '',
    'Nguyên tắc chọn dữ liệu:',
    '- Nếu một việc có tuần Đánh giá = Đạt: lấy dòng Đạt đầu tiên.',
    '- Nếu chưa có Đạt: lấy dòng tuần mới nhất có dữ liệu.',
    '- Không cập nhật dòng Dự án/Tổng hoặc dòng không có SourceRow.',
    '',
    'Bạn có chắc chắn thực hiện không?'
  ].join('\n');

  const confirm = ui.alert(
    'Cập nhật KH tháng theo dữ liệu tuần',
    confirmMessage,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    ss.toast('Đã hủy cập nhật KH tháng theo dữ liệu tuần.', 'KH_TUAN_WBS', 5);
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
    mode: scope.mode,
    maxWeekIndex: scope.maxWeekIndex,
    updatedRows: result.updatedRows,
    updatedCells: result.updatedCells,
    skipOutOfScope: candidates.skipOutOfScope,
    skipNoSourceRow: candidates.skipNoSourceRow,
    skipNotMarked: candidates.skipNotMarked,
    skipNoData: candidates.skipNoData
  };
}

function khTuanWbs_collectMonthPushCandidatesFromWeekData_(dataRows, expectedSourceSheet, maxWeekIndex, targetCols) {
  const config = KH_TUAN_WBS_CONFIG;
  const col = config.COL;
  const selectedBySourceRow = {};

  const output = {
    toUpdate: [],
    skipOutOfScope: 0,
    skipNoSourceRow: 0,
    skipNotMarked: 0,
    skipNoData: 0
  };

  (dataRows || []).forEach(function(item) {
    if (item.sourceSheet !== expectedSourceSheet) {
      output.skipOutOfScope++;
      return;
    }

    const weekIndex = khTuanWbs_getWeekIndex_(item.weekName);
    if (weekIndex < 1 || weekIndex > maxWeekIndex) {
      output.skipOutOfScope++;
      return;
    }

    const row = khTuanWbs_normalizeRowLength_(item.values);
    const sourceRow = Number(row[col.SOURCE_ROW - 1] || 0);
    const loaiDong = khTuanWbs_normalizeText_(row[col.LOAI_DONG - 1]);

    if (!sourceRow || loaiDong === 'dự án' || loaiDong === 'du an' || loaiDong === 'tổng' || loaiDong === 'tong') {
      output.skipNoSourceRow++;
      return;
    }

    if (!khTuanWbs_isUpdateMonthFlag_(row[col.CAP_NHAT - 1])) {
      output.skipNotMarked++;
      return;
    }

    const payload = khTuanWbs_buildMonthPayloadFromWeekRow_(row, targetCols);
    const dataCount = Object.keys(payload).length;

    if (!dataCount) {
      output.skipNoData++;
      return;
    }

    const candidate = {
      sourceRow: sourceRow,
      payload: payload,
      weekIndex: weekIndex,
      isDone: khTuanWbs_isDoneStatus_(row[col.DANH_GIA - 1]),
      dataCount: dataCount
    };

    const current = selectedBySourceRow[sourceRow];
    if (!current || khTuanWbs_shouldReplaceMonthPushCandidate_(current, candidate)) {
      selectedBySourceRow[sourceRow] = candidate;
    }
  });

  output.toUpdate = Object.keys(selectedBySourceRow)
    .map(function(sourceRowText) {
      const item = selectedBySourceRow[sourceRowText];
      return {
        sourceRow: item.sourceRow,
        payload: item.payload
      };
    })
    .sort(function(a, b) {
      return Number(a.sourceRow || 0) - Number(b.sourceRow || 0);
    });

  return output;
}

function khTuanWbs_buildMonthPayloadFromWeekRow_(row, targetCols) {
  const config = KH_TUAN_WBS_CONFIG;
  const col = config.COL;
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

  return payload;
}

function khTuanWbs_shouldReplaceMonthPushCandidate_(current, candidate) {
  if (candidate.isDone && !current.isDone) return true;
  if (!candidate.isDone && current.isDone) return false;

  if (candidate.isDone && current.isDone) {
    if (candidate.weekIndex !== current.weekIndex) {
      return candidate.weekIndex < current.weekIndex;
    }
    return candidate.dataCount > current.dataCount;
  }

  if (candidate.weekIndex !== current.weekIndex) {
    return candidate.weekIndex > current.weekIndex;
  }

  return candidate.dataCount > current.dataCount;
}
