import 'package:flutter/foundation.dart';
import 'package:excel/excel.dart';
import 'package:xml/xml.dart';
import 'package:string_similarity/string_similarity.dart';
import '../models/excel_models.dart';

/// 列结构识别服务：负责识别 Excel 文件中的列结构
/// 职责：识别 Item, Description, Unit, Qty, Rate, Amount 等列
class ColumnIdentifierService {
  /// 从 Excel 包识别目标文件列结构
  TargetColumns? identifyTargetColumnsFromExcel(Sheet sheet) {
    int? itemCol, descriptionCol, unitCol, qtyCol, unitRateCol, amountCol;

    // 尝试从第 6 行（索引 5）读取表头，如果找不到则尝试其他行
    List<dynamic>? headerRow;

    // 先尝试第 6 行（索引 5）
    for (var row in sheet.rows) {
      final rowIndex = row.first?.rowIndex ?? -1;
      if (rowIndex == 5) {
        headerRow = row;
        break;
      }
    }

    // 如果第 6 行找不到，尝试前 10 行查找包含 "Item" 的行
    if (headerRow == null) {
      for (var row in sheet.rows) {
        final rowIndex = row.first?.rowIndex ?? -1;
        if (rowIndex >= 10) break;

        for (var cell in row) {
          if (cell == null) continue;
          final cellText = (_getCellValueFromExcel(cell) ?? '').toLowerCase();
          if (matchesColumnName(
              cellText, ['item', 'item no', 'no.', 'description'])) {
            headerRow = row;
            break;
          }
        }
        if (headerRow != null) break;
      }
    }

    if (headerRow == null) return null;

    // 遍历表头行的所有单元格
    for (var cell in headerRow) {
      if (cell == null) continue;
      final colIndex = cell.columnIndex;
      final headerText = (_getCellValueFromExcel(cell) ?? '').toLowerCase();

      // 识别各列
      if (itemCol == null &&
          matchesColumnName(headerText, ['item', 'item no', 'no.', '序号'])) {
        itemCol = colIndex;
      }
      if (descriptionCol == null &&
          matchesColumnName(headerText,
              ['description', 'desc', '工作描述', 'description of work'])) {
        descriptionCol = colIndex;
      }
      if (unitCol == null && matchesColumnName(headerText, ['unit', 'u'])) {
        unitCol = colIndex;
      }
      if (qtyCol == null &&
          matchesColumnName(headerText, ['qty', 'quantity'])) {
        qtyCol = colIndex;
      }
      if (unitRateCol == null) {
        final similarity = getBestSimilarity(
            headerText, ['rate', 'unit rate', '(b)', 'unit rate (hk\$)']);
        if (similarity > 0.7) {
          unitRateCol = colIndex;
        }
      }
      if (amountCol == null) {
        final similarity = getBestSimilarity(
            headerText, ['amount', 'total', '(c)', 'amount (hk\$)']);
        if (similarity > 0.7) {
          amountCol = colIndex;
        }
      }
    }

    // 检查是否所有必需的列都已识别
    if (itemCol == null ||
        descriptionCol == null ||
        unitRateCol == null ||
        amountCol == null) {
      return null;
    }

    return TargetColumns(
      itemCol: itemCol,
      descriptionCol: descriptionCol,
      unitCol: unitCol ?? -1,
      qtyCol: qtyCol ?? -1,
      unitRateCol: unitRateCol,
      amountCol: amountCol,
    );
  }

  /// 从 XML 识别目标文件列结构（与草稿文件类似，但返回 TargetColumns）
  TargetColumns? identifyTargetColumnsFromXml(
      XmlElement firstRow, List<String> sharedStrings) {
    final cells = firstRow.findElements('c').toList();
    if (cells.isEmpty) {
      debugPrint('[ColumnIdentifierService] XML: 第一行没有单元格');
      return null;
    }

    int? itemCol, descriptionCol, unitCol, qtyCol, unitRateCol, amountCol;

    for (var cell in cells) {
      final cellRef = cell.getAttribute('r');
      if (cellRef == null) continue;

      final colIndex = _cellRefToColIndex(cellRef);
      if (colIndex < 0) continue;

      // 直接从当前单元格读取值
      final headerText = _getCellValueFromXmlElement(cell, sharedStrings) ?? '';
      final headerTextLower = headerText.toLowerCase();

      // 识别各列
      if (itemCol == null &&
          matchesColumnName(headerTextLower,
              ['item', 'item no', 'no.', 'description', '序号'])) {
        itemCol = colIndex;
      }
      if (descriptionCol == null &&
          matchesColumnName(headerTextLower,
              ['description', 'desc', '工作描述', 'description of work'])) {
        descriptionCol = colIndex;
      }
      if (unitCol == null &&
          matchesColumnName(headerTextLower, ['unit', 'u'])) {
        unitCol = colIndex;
      }
      if (qtyCol == null &&
          matchesColumnName(headerTextLower, ['qty', 'quantity'])) {
        qtyCol = colIndex;
      }
      if (unitRateCol == null) {
        final similarity = getBestSimilarity(
            headerTextLower, ['rate', 'unit rate', '(b)', 'unit rate (hk\$)']);
        if (similarity > 0.7) {
          unitRateCol = colIndex;
        }
      }
      if (amountCol == null) {
        final similarity = getBestSimilarity(
            headerTextLower, ['amount', 'total', '(c)', 'amount (hk\$)']);
        if (similarity > 0.7) {
          amountCol = colIndex;
        }
      }
    }

    // 检查是否所有必需的列都已识别
    if (itemCol == null ||
        descriptionCol == null ||
        unitRateCol == null ||
        amountCol == null) {
      return null;
    }

    return TargetColumns(
      itemCol: itemCol,
      descriptionCol: descriptionCol,
      unitCol: unitCol ?? -1,
      qtyCol: qtyCol ?? -1,
      unitRateCol: unitRateCol,
      amountCol: amountCol,
    );
  }

  /// 从 XML 识别草稿文件列结构
  DraftColumns? identifyDraftColumnsFromXml(
      XmlElement firstRow, List<String> sharedStrings) {
    final cells = firstRow.findElements('c').toList();
    if (cells.isEmpty) {
      debugPrint('[ColumnIdentifierService] XML: 第一行没有单元格');
      return null;
    }

    debugPrint('[ColumnIdentifierService] XML: 第一行有 ${cells.length} 个单元格');
    debugPrint(
        '[ColumnIdentifierService] XML: 共享字符串数量: ${sharedStrings.length}');

    int? itemCol, descriptionCol, unitCol, qtyCol, rateCol, amountCol;

    for (var cell in cells) {
      final cellRef = cell.getAttribute('r');
      if (cellRef == null) continue;

      final colIndex = _cellRefToColIndex(cellRef);
      if (colIndex < 0) continue;

      // 直接从当前单元格读取值
      final headerText = _getCellValueFromXmlElement(cell, sharedStrings) ?? '';
      final headerTextLower = headerText.toLowerCase();

      // 记录所有单元格，包括空值，便于调试
      final cellType = cell.getAttribute('t') ?? 'normal';
      if (headerText.isNotEmpty) {
        debugPrint(
            '[ColumnIdentifierService] XML: 列 ${colToLetter(colIndex)} ($colIndex) [$cellType]: "$headerText"');
      } else {
        debugPrint(
            '[ColumnIdentifierService] XML: 列 ${colToLetter(colIndex)} ($colIndex) [$cellType]: (空)');
      }

      // 识别各列
      if (itemCol == null &&
          matchesColumnName(headerTextLower,
              ['item', 'item no', 'no.', 'description', '序号'])) {
        itemCol = colIndex;
        debugPrint(
            '[ColumnIdentifierService] XML: 识别到 Item 列: ${colToLetter(colIndex)}');
      }
      if (descriptionCol == null &&
          matchesColumnName(headerTextLower,
              ['description', 'desc', '工作描述', 'description of work'])) {
        descriptionCol = colIndex;
        debugPrint(
            '[ColumnIdentifierService] XML: 识别到 Description 列: ${colToLetter(colIndex)}');
      }
      if (unitCol == null &&
          matchesColumnName(headerTextLower, ['unit', 'u'])) {
        unitCol = colIndex;
        debugPrint(
            '[ColumnIdentifierService] XML: 识别到 Unit 列: ${colToLetter(colIndex)}');
      }
      if (qtyCol == null &&
          matchesColumnName(headerTextLower, ['qty', 'quantity'])) {
        qtyCol = colIndex;
        debugPrint(
            '[ColumnIdentifierService] XML: 识别到 Qty 列: ${colToLetter(colIndex)}');
      }
      if (rateCol == null) {
        final similarity = getBestSimilarity(
            headerTextLower, ['rate', 'unit rate', '(b)', 'unit rate (hk\$)']);
        if (similarity > 0.7) {
          rateCol = colIndex;
          debugPrint(
              '[ColumnIdentifierService] XML: 识别到 Rate 列: ${colToLetter(colIndex)} (相似度: $similarity)');
        }
      }
      if (amountCol == null) {
        final similarity = getBestSimilarity(
            headerTextLower, ['amount', 'total', '(c)', 'amount (hk\$)']);
        if (similarity > 0.7) {
          amountCol = colIndex;
          debugPrint(
              '[ColumnIdentifierService] XML: 识别到 Amount 列: ${colToLetter(colIndex)} (相似度: $similarity)');
        }
      }
    }

    // 检查必需列
    if (itemCol == null ||
        descriptionCol == null ||
        rateCol == null ||
        amountCol == null) {
      return null;
    }

    return DraftColumns(
      itemCol: itemCol,
      descriptionCol: descriptionCol,
      unitCol: unitCol ?? -1,
      qtyCol: qtyCol ?? -1,
      rateCol: rateCol,
      amountCol: amountCol,
    );
  }

  /// 判断列名是否匹配（公共方法）
  bool matchesColumnName(String headerText, List<String> patterns) {
    final normalized = _normalizeText(headerText);
    for (final pattern in patterns) {
      if (normalized.contains(_normalizeText(pattern))) {
        return true;
      }
    }
    return false;
  }

  /// 获取最佳相似度（公共方法）
  double getBestSimilarity(String text, List<String> patterns) {
    double best = 0.0;
    for (final pattern in patterns) {
      final similarity = StringSimilarity.compareTwoStrings(
        _normalizeText(text),
        _normalizeText(pattern),
      );
      if (similarity > best) {
        best = similarity;
      }
    }
    return best;
  }

  /// 文本预处理
  String _normalizeText(String text) {
    return text
        .toLowerCase()
        .replaceAll(RegExp(r'\s+'), '')
        .replaceAll(RegExp(r'[()（）]'), '');
  }

  /// 列索引转字母（公共方法）
  String colToLetter(int colIndex) {
    if (colIndex < 0) return 'A';
    var n = colIndex + 1;
    final buffer = StringBuffer();
    while (n > 0) {
      n -= 1;
      buffer.writeCharCode('A'.codeUnitAt(0) + (n % 26));
      n ~/= 26;
    }
    final str = buffer.toString();
    return String.fromCharCodes(str.codeUnits.reversed);
  }

  /// 将单元格引用转换为列索引
  int _cellRefToColIndex(String cellRef) {
    final match = RegExp(r'^([A-Z]+)').firstMatch(cellRef);
    if (match == null) return -1;

    final colLetters = match.group(1)!;
    int colIndex = 0;
    for (var i = 0; i < colLetters.length; i++) {
      colIndex = colIndex * 26 + (colLetters.codeUnitAt(i) - 64);
    }
    return colIndex - 1;
  }

  /// 从 XML 单元格元素读取值
  String? _getCellValueFromXmlElement(
      XmlElement cell, List<String> sharedStrings) {
    // 检查单元格类型
    final t = cell.getAttribute('t');

    // 处理内联字符串（inlineStr）
    if (t == 'inlineStr') {
      final inlineStrElement = cell.findElements('is').firstOrNull;
      if (inlineStrElement != null) {
        final textNodes = inlineStrElement.findAllElements('t');
        if (textNodes.isNotEmpty) {
          final buffer = StringBuffer();
          for (final node in textNodes) {
            buffer.write(node.innerText);
          }
          final text = buffer.toString().trim();
          if (text.isNotEmpty) {
            return text;
          }
        } else {
          final raw = inlineStrElement.innerText.trim();
          if (raw.isNotEmpty) {
            return raw;
          }
        }
      }
    }

    // 处理共享字符串（s）
    if (t == 's') {
      final v = cell.findElements('v').firstOrNull;
      if (v != null) {
        final value = v.innerText.trim();
        if (value.isNotEmpty && sharedStrings.isNotEmpty) {
          final index = int.tryParse(value);
          if (index != null && index < sharedStrings.length) {
            return sharedStrings[index];
          }
        }
      }
      return null;
    }

    // 处理普通值（v）
    final v = cell.findElements('v').firstOrNull;
    if (v != null) {
      final value = v.innerText.trim();
      return value.isEmpty ? null : value;
    }

    return null;
  }

  /// 获取单元格值（excel 包）
  String? _getCellValueFromExcel(dynamic cell) {
    if (cell == null) return null;

    try {
      final value = cell.value;
      if (value == null) return null;

      try {
        return switch (value) {
          TextCellValue() => value.value.toString(),
          IntCellValue() => value.value.toString(),
          DoubleCellValue() => value.value.toString(),
          BoolCellValue() => value.value.toString(),
          DateCellValue() => '${value.year}-${value.month}-${value.day}',
          DateTimeCellValue() => value.asDateTimeLocal().toString(),
          TimeCellValue() => value.asDuration().toString(),
          FormulaCellValue() => value.formula.toString(),
          _ => null,
        };
      } catch (e) {
        try {
          final str = value.toString();
          return str.isEmpty ? null : str;
        } catch (e2) {
          return null;
        }
      }
    } catch (e) {
      return null;
    }
  }
}
