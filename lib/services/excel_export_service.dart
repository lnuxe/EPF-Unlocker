import 'dart:io';

import 'package:excel/excel.dart';

import '../models/excel_models.dart';

/// 将内部 ExcelSheetModel 导出为无保护的 .xlsx 文件
Future<void> exportToExcel(ExcelSheetModel sheet, String savePath) async {
  final excel = Excel.createExcel();
  final defaultSheetName = excel.getDefaultSheet() ?? 'Sheet1';
  final sheetObject = excel[defaultSheetName];
  if (sheet.name != defaultSheetName) {
    excel.rename(defaultSheetName, sheet.name);
  }

  for (final row in sheet.rows) {
    for (final cell in row.cells) {
      final value = cell.value;
      if (value == null || value.isEmpty) continue;
      final cellIndex = CellIndex.indexByColumnRow(
        columnIndex: cell.column,
        rowIndex: cell.row,
      );
      sheetObject.cell(cellIndex).value = TextCellValue(value);
    }
  }

  final bytes = excel.save();
  if (bytes == null) return;

  final file = File(savePath);
  await file.create(recursive: true);
  await file.writeAsBytes(bytes, flush: true);
}
