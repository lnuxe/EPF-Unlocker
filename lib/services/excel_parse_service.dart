import 'dart:io';

import 'package:excel/excel.dart';
import 'package:flutter/foundation.dart';

import '../models/excel_models.dart';

/// 负责把已解锁的 Excel 文件解析为内部数据模型
class ExcelParseService {
  Future<ExcelSheetModel> loadFirstSheet(File unlockedFile) async {
    final bytes = await unlockedFile.readAsBytes();

    try {
      debugPrint('[ExcelParseService] 尝试使用 excel 包解析...');
      final excel = Excel.decodeBytes(bytes);

      if (excel.tables.isEmpty) {
        return const ExcelSheetModel(name: 'Sheet1', rows: []);
      }

      final firstSheetName = excel.tables.keys.first;
      final sheet = excel.tables[firstSheetName]!;

      final List<ExcelRow> rows = [];
      for (var r = 0; r < sheet.maxRows; r++) {
        final List<ExcelCell> cells = [];
        for (var c = 0; c < sheet.maxColumns; c++) {
          final cell = sheet.cell(
            CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
          );
          final value = cell.value;

          cells.add(
            ExcelCell(
              row: r,
              column: c,
              value: value?.toString(),
            ),
          );
        }
        rows.add(ExcelRow(index: r, cells: cells));
      }

      return ExcelSheetModel(
        name: firstSheetName,
        rows: rows,
        mergeRanges: const [],
      );
    } catch (e, st) {
      // excel 包解析失败
      debugPrint('[ExcelParseService] excel.decodeBytes 失败: $e');
      debugPrint('[ExcelParseService] stack: $st');

      // 如果解析失败，返回空的工作表
      return const ExcelSheetModel(name: 'Sheet1', rows: []);
    }
  }
}
