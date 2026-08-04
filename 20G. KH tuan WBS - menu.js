/**
 * KH tuần WBS - Menu cài đặt và vận hành
 *
 * Có 2 trạng thái:
 * 1. Khi mới nhân bản: hiện nhóm Cài đặt ban đầu + nhóm Vận hành.
 * 2. Sau khi ẩn cài đặt: chỉ hiện 2 lệnh vận hành.
 */

function khTuanWbs_addMenu_() {
  const ui = SpreadsheetApp.getUi();
  const showSetup = khTuanWbs_shouldShowSetupMenu_();

  const menu = ui.createMenu('📅 Quản lý KH tuần');

  if (showSetup) {
    menu.addSubMenu(
      ui.createMenu('1. Cài đặt ban đầu khi nhân bản')
        .addItem('1.1. Thiết lập / chuẩn hóa KH tuần WBS', 'khTuanWbs_setupInitialForClone')
        .addItem('1.2. Làm mới dữ liệu KH tuần khi nhân bản', 'khTuanWbs_resetWeeklyDataForClone')
        .addSeparator()
        .addItem('1.3. Ẩn nhóm cài đặt ban đầu', 'khTuanWbs_hideSetupMenu')
    );

    menu.addSeparator();

    menu.addSubMenu(
      ui.createMenu('2. Vận hành KH tuần')
        .addItem('2.1. Nạp / cập nhật KH tuần WBS', 'napKhTuanWbsTuSheetThang')
        .addItem('2.2. Cập nhật KH tháng theo dữ liệu tuần', 'khTuanWbs_pushMonthResultFromWeekData')
    );
  } else {
    menu
      .addItem('1. Nạp / cập nhật KH tuần WBS', 'napKhTuanWbsTuSheetThang')
      .addItem('2. Cập nhật KH tháng theo dữ liệu tuần', 'khTuanWbs_pushMonthResultFromWeekData');
  }

  menu.addToUi();
}

function khTuanWbs_shouldShowSetupMenu_() {
  const props = PropertiesService.getDocumentProperties();
  const value = props.getProperty('KH_TUAN_WBS_SHOW_SETUP_MENU');

  // Mặc định file mới nhân bản sẽ hiện nhóm cài đặt.
  if (value === null || value === '') return true;

  return value !== 'FALSE';
}

function khTuanWbs_hideSetupMenu() {
  PropertiesService
    .getDocumentProperties()
    .setProperty('KH_TUAN_WBS_SHOW_SETUP_MENU', 'FALSE');

  SpreadsheetApp.getActive().toast(
    'Đã ẩn nhóm Cài đặt ban đầu. Vui lòng tải lại Google Sheet để menu cập nhật.',
    'KH_TUAN_WBS',
    8
  );
}

function khTuanWbs_showSetupMenu() {
  PropertiesService
    .getDocumentProperties()
    .setProperty('KH_TUAN_WBS_SHOW_SETUP_MENU', 'TRUE');

  SpreadsheetApp.getActive().toast(
    'Đã bật lại nhóm Cài đặt ban đầu. Vui lòng tải lại Google Sheet để menu cập nhật.',
    'KH_TUAN_WBS',
    8
  );
}

function khTuanWbs_setupInitialForClone() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = KH_TUAN_WBS_CONFIG;
  const sheetTuan = ss.getSheetByName(config.SHEET_TUAN);

  if (!sheetTuan) {
    SpreadsheetApp.getUi().alert('Không tìm thấy sheet KH_TUAN_WBS. Cần tạo sheet này trước.');
    return;
  }

  const dataSheet = khTuanWbs_getOrCreateDataSheet_(ss);
  khTuanWbs_ensureDataSheetHeaders_(dataSheet);
  dataSheet.hideSheet();

  const logSheet = ss.getSheetByName('KH_TUAN_PUSH_LOG');
  if (logSheet) {
    logSheet.hideSheet();
  }

  khTuanWbs_applyFormats_(sheetTuan);

  const sourceSheetName = String(sheetTuan.getRange(config.SOURCE_CELL).getDisplayValue() || '').trim();
  if (sourceSheetName) {
    khTuanWbs_updateWeekDropdown_(sheetTuan, sourceSheetName);
  }

  SpreadsheetApp.getActive().toast(
    'Đã thiết lập / chuẩn hóa KH_TUAN_WBS và các sheet kỹ thuật cần thiết.',
    'KH_TUAN_WBS',
    8
  );
}

function khTuanWbs_resetWeeklyDataForClone() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    'Làm mới dữ liệu KH tuần khi nhân bản',
    [
      'Thao tác này sẽ xóa dữ liệu trong KH_TUAN_DATA để tránh mang dữ liệu từ file mẫu sang file mới.',
      '',
      'Không xóa dữ liệu KH tháng.',
      'Không xóa giao diện KH_TUAN_WBS.',
      '',
      'Bạn có chắc chắn thực hiện không?'
    ].join('\n'),
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    SpreadsheetApp.getActive().toast('Đã hủy làm mới dữ liệu KH tuần.', 'KH_TUAN_WBS', 5);
    return;
  }

  khTuanWbs_resetAllDataForTest();
}
