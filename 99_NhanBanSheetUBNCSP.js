function nhanBanGhiDe_UBNCSP_12Thang() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceSheet =
    ss.getSheetByName("UBNCSP - T01.2026") ||
    ss.getSheetByName("UBNCSP – T01.2026");

  if (!sourceSheet) {
    throw new Error("Không tìm thấy sheet nguồn: UBNCSP - T01.2026 hoặc UBNCSP – T01.2026");
  }

  const sourceName = sourceSheet.getName();
  const dash = sourceName.includes(" – ") ? " – " : " - ";
  const prefix = "UBNCSP" + dash + "T";

  const sourceIndex = sourceSheet.getIndex();
  const result = [];

  // Xóa sheet T02 đến T12 nếu đã tồn tại
  for (let month = 2; month <= 12; month++) {
    const monthText = String(month).padStart(2, "0");
    const targetName = `${prefix}${monthText}.2026`;

    const existingSheet = ss.getSheetByName(targetName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
      result.push(`Đã xóa sheet cũ: ${targetName}`);
    }
  }

  // Nhân bản lại từ T01
  for (let month = 2; month <= 12; month++) {
    const monthText = String(month).padStart(2, "0");
    const targetName = `${prefix}${monthText}.2026`;

    const newSheet = sourceSheet.copyTo(ss);
    newSheet.setName(targetName);

    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(sourceIndex + month - 1);

    result.push(`Đã tạo sheet mới: ${targetName}`);
  }

  SpreadsheetApp.getUi().alert(
    "Đã hoàn thành ghi đè và nhân bản sheet UBNCSP từ T02 đến T12.\n\n" +
    result.join("\n")
  );
}
