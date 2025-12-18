import 'dart:io';
import 'dart:convert';
import 'package:flutter/foundation.dart';
import 'package:excel/excel.dart';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';
import '../models/excel_models.dart';
import 'column_identifier_service.dart';

/// Excel 解析服务：负责解析 Excel 文件（XML 和 excel 包）
/// 职责：文件解析、数据提取、工作表查找
class ExcelParserService {
  final _columnIdentifier = ColumnIdentifierService();

  /// 使用 excel 包解析 Excel 文件
  Excel? parseWithExcelPackage(List<int> bytes) {
    try {
      return Excel.decodeBytes(bytes);
    } catch (e) {
      debugPrint('[ExcelParserService] excel 包解析失败: $e');
      return null;
    }
  }

  /// 使用 XML 方式解析草稿文件（避免 numFmtId 错误）
  Future<Map<String, dynamic>> parseDraftFileWithXml(
    File draftFile, {
    String? targetSheetName,
  }) async {
    final draftMap = <String, DraftRow>{};

    // 先检查文件是否存在
    if (!await draftFile.exists()) {
      return {
        'map': draftMap,
        'columns': null,
        'matchedSheet': null,
        'error': '文件不存在: ${draftFile.path}',
      };
    }

    // 尝试读取文件，捕获权限错误
    List<int> bytes;
    try {
      bytes = await draftFile.readAsBytes();
    } on FileSystemException catch (e) {
      String errorMessage;
      if (e.osError?.errorCode == 1) {
        // Operation not permitted
        errorMessage = '无法访问文件（权限不足）: ${draftFile.path}\n'
            '请检查：\n'
            '1. 文件是否被其他程序打开（如 Excel）\n'
            '2. 是否授予了应用文件访问权限\n'
            '3. 文件是否在受保护的目录中';
      } else if (e.osError?.errorCode == 13) {
        // Permission denied
        errorMessage = '文件访问被拒绝: ${draftFile.path}\n'
            '请检查文件权限设置';
      } else {
        errorMessage = '无法读取文件: ${draftFile.path}\n'
            '错误代码: ${e.osError?.errorCode ?? "未知"}\n'
            '错误信息: ${e.message}';
      }
      debugPrint('[ExcelParserService] 文件读取失败: $errorMessage');
      return {
        'map': draftMap,
        'columns': null,
        'matchedSheet': null,
        'error': errorMessage,
      };
    } catch (e, st) {
      debugPrint('[ExcelParserService] 文件读取异常: $e');
      debugPrint('[ExcelParserService] 堆栈: $st');
      return {
        'map': draftMap,
        'columns': null,
        'matchedSheet': null,
        'error': '无法读取文件: ${draftFile.path}\n错误: $e',
      };
    }

    debugPrint('[ExcelParserService] 使用 XML 直接解析草稿文件...');
    try {
      return await _parseDraftMapWithXml(
        bytes,
        draftMap,
        targetSheetName: targetSheetName,
      );
    } catch (xmlError, xmlSt) {
      debugPrint('[ExcelParserService] XML 解析失败: $xmlError');
      debugPrint('[ExcelParserService] 堆栈: $xmlSt');
      return {
        'map': draftMap,
        'columns': null,
        'matchedSheet': null,
        'error': 'XML 解析失败: $xmlError',
      };
    }
  }

  /// 从 XML 解析草稿文件数据
  Future<Map<String, dynamic>> _parseDraftMapWithXml(
    List<int> bytes,
    Map<String, DraftRow> draftMap, {
    String? targetSheetName,
  }) async {
    try {
      // 解压 ZIP 归档
      final archive = ZipDecoder().decodeBytes(bytes);

      // 读取 workbook.xml 获取工作表列表
      final workbookFile = archive.findFile('xl/workbook.xml');
      if (workbookFile == null) {
        throw Exception('无法找到 workbook.xml，文件可能已损坏');
      }

      final workbookXml = utf8.decode(workbookFile.content as List<int>);
      final workbookDoc = XmlDocument.parse(workbookXml);

      // 获取所有工作表信息
      final sheetElements = workbookDoc.findAllElements('sheet');
      if (sheetElements.isEmpty) {
        throw Exception('工作簿中没有工作表');
      }

      // 构建工作表名称到 sheetId 的映射
      final sheetMap = <String, String>{};
      for (var sheet in sheetElements) {
        final name = sheet.getAttribute('name');
        final sheetId = sheet.getAttribute('sheetId');
        if (name != null && sheetId != null) {
          sheetMap[name] = sheetId;
        }
      }

      final allSheetNames = sheetMap.keys.toList();
      debugPrint(
          '[ExcelParserService] XML: 草稿文件包含的工作表: ${allSheetNames.join(", ")}');

      // 优先查找同名工作表
      String? matchedSheetName;
      List<String> sheetsToProcess = [];

      if (targetSheetName != null) {
        if (sheetMap.containsKey(targetSheetName)) {
          matchedSheetName = targetSheetName;
          sheetsToProcess = [targetSheetName];
          debugPrint('[ExcelParserService] XML: 找到同名工作表: $targetSheetName');
        } else {
          // 尝试模糊匹配
          final normalizedTarget = _normalizeSheetName(targetSheetName);
          for (var sheetName in allSheetNames) {
            if (_normalizeSheetName(sheetName) == normalizedTarget) {
              matchedSheetName = sheetName;
              sheetsToProcess = [sheetName];
              debugPrint(
                  '[ExcelParserService] XML: 找到相似工作表: $sheetName (目标: $targetSheetName)');
              break;
            }
          }
        }
      }

      // 如果没找到同名工作表，处理所有工作表
      if (sheetsToProcess.isEmpty) {
        sheetsToProcess = allSheetNames;
        debugPrint('[ExcelParserService] XML: 未找到同名工作表，将在所有工作表中匹配');
      }

      // 读取 sharedStrings.xml（如果存在）
      final sharedStringsFile = archive.findFile('xl/sharedStrings.xml');
      final sharedStrings = <String>[];
      if (sharedStringsFile != null) {
        final sharedStringsXml =
            utf8.decode(sharedStringsFile.content as List<int>);
        final sharedStringsDoc = XmlDocument.parse(sharedStringsXml);
        for (var si in sharedStringsDoc.findAllElements('si')) {
          sharedStrings.add(_extractTextFromSharedString(si));
        }
      }

      // 遍历要处理的工作表
      DraftColumns? draftColumns;

      // 按 sheetId 找到对应的 sheetX.xml 文件，而不是按文件名顺序
      for (var sheetName in sheetsToProcess) {
        // 获取该工作表的 sheetId
        final sheetId = sheetMap[sheetName];
        if (sheetId == null) {
          debugPrint(
              '[ExcelParserService] XML: 警告：工作表 "$sheetName" 在 sheetMap 中找不到 sheetId，跳过');
          continue;
        }

        // 方法1：尝试使用 sheetId 直接查找
        ArchiveFile? sheetFile;
        final expectedSheetFileName = 'xl/worksheets/sheet$sheetId.xml';
        for (final file in archive.files) {
          if (file.name == expectedSheetFileName) {
            sheetFile = file;
            debugPrint(
                '[ExcelParserService] XML: 通过 sheetId 找到工作表文件: $sheetName -> ${sheetFile.name} (sheetId: $sheetId)');
            break;
          }
        }

        // 方法2：如果精确匹配失败，根据工作表在 workbook.xml 中的出现顺序查找
        if (sheetFile == null) {
          final sheetFiles = archive.files
              .where((f) =>
                  f.name.contains('xl/worksheets/sheet') &&
                  f.name.endsWith('.xml'))
              .toList()
            ..sort((a, b) {
              // 提取数字部分进行排序
              final numA = _extractSheetNumber(a.name);
              final numB = _extractSheetNumber(b.name);
              return numA.compareTo(numB);
            });

          // 按照 sheet 元素在 workbook.xml 中出现的顺序查找
          final sheetElementsList = sheetElements.toList();
          int targetIndex = -1;
          for (int i = 0; i < sheetElementsList.length; i++) {
            final sheet = sheetElementsList[i];
            final id = sheet.getAttribute('sheetId');
            if (id == sheetId) {
              targetIndex = i;
              break;
            }
          }

          if (targetIndex >= 0 && targetIndex < sheetFiles.length) {
            sheetFile = sheetFiles[targetIndex];
            debugPrint(
                '[ExcelParserService] XML: 通过工作表顺序找到文件: $sheetName -> ${sheetFile.name} (工作表索引: $targetIndex, sheetId: $sheetId)');
          }
        }

        if (sheetFile == null) {
          debugPrint(
              '[ExcelParserService] XML: 无法找到工作表 "$sheetName" 对应的 XML 文件，跳过');
          continue;
        }

        debugPrint(
            '[ExcelParserService] XML: 正在处理工作表: $sheetName (文件: ${sheetFile.name})');

        // 解析工作表 XML
        final sheetXml = utf8.decode(sheetFile.content as List<int>);
        final sheetDoc = XmlDocument.parse(sheetXml);
        final sheetData = sheetDoc.findAllElements('sheetData').firstOrNull;
        if (sheetData == null) continue;

        final rows = sheetData.findElements('row').toList();

        // 识别列结构（尝试前30行来找到表头）
        int headerRowIndex = -1;
        DraftColumns? currentSheetColumns;
        // 尝试前30行来找到表头（有些Excel文件表头可能在后面）
        for (int i = 0; i < rows.length && i < 30; i++) {
          final row = rows[i];
          debugPrint('[ExcelParserService] XML: 尝试第 ${i + 1} 行作为表头...');
          currentSheetColumns =
              _columnIdentifier.identifyDraftColumnsFromXml(row, sharedStrings);
          if (currentSheetColumns != null) {
            draftColumns = currentSheetColumns; // 保存到外层变量
            headerRowIndex = i;
            debugPrint('[ExcelParserService] XML: 在第 ${i + 1} 行找到表头！');
            debugPrint(
                '[ExcelParserService] XML: 已识别列结构: Item=${_colToLetter(draftColumns.itemCol)}, '
                'Description=${_colToLetter(draftColumns.descriptionCol)}, '
                'Rate=${_colToLetter(draftColumns.rateCol)}, '
                'Amount=${_colToLetter(draftColumns.amountCol)}');
            debugPrint(
                '[ExcelParserService] XML: 表头在第 ${headerRowIndex + 1} 行，将从第 ${headerRowIndex + 2} 行开始读取数据');
            break;
          }
        }

        if (draftColumns == null) {
          debugPrint(
              '[ExcelParserService] XML: 在前20行中无法识别工作表 "$sheetName" 的列结构，跳过');
          debugPrint(
              '[ExcelParserService] XML: 提示：请检查Excel文件，确保表头行包含 Item、Description、Rate、Amount 等列名');
          continue;
        }

        // 从表头之后开始读取数据
        int rowCount = 0;
        for (int i = 0; i < rows.length; i++) {
          if (headerRowIndex >= 0 && i <= headerRowIndex) {
            continue; // 跳过表头行
          }

          final row = rows[i];
          final cells = row.findElements('c').toList();
          if (cells.isEmpty) continue;

          // 提取字段
          final item =
              getCellValueFromXml(cells, draftColumns.itemCol, sharedStrings) ??
                  '';
          if (item.isEmpty) continue;

          final description = getCellValueFromXml(
                  cells, draftColumns.descriptionCol, sharedStrings) ??
              '';

          // 跳过包含 "Remark" 的行（这些是备注行，不应该匹配）
          final descriptionLower = description.toLowerCase();
          final itemTrimmed = item.trim();
          if (descriptionLower.contains('remark') ||
              itemTrimmed.toLowerCase().contains('remark')) {
            debugPrint(
                '[ExcelParserService] XML草稿文件：跳过备注行: Item=$item, Description=$description');
            continue;
          }

          String? unit;
          if (draftColumns.unitCol >= 0) {
            unit =
                getCellValueFromXml(cells, draftColumns.unitCol, sharedStrings);
          }

          double? qty;
          if (draftColumns.qtyCol >= 0) {
            final qtyStr =
                getCellValueFromXml(cells, draftColumns.qtyCol, sharedStrings);
            if (qtyStr != null && qtyStr.isNotEmpty) {
              qty = double.tryParse(qtyStr);
            }
          }

          double? rate;
          if (draftColumns.rateCol >= 0) {
            final rateStr =
                getCellValueFromXml(cells, draftColumns.rateCol, sharedStrings);
            if (rateStr != null && rateStr.isNotEmpty) {
              rate = double.tryParse(rateStr);
            }
          }

          double? amount;
          if (draftColumns.amountCol >= 0) {
            final amountStr = getCellValueFromXml(
                cells, draftColumns.amountCol, sharedStrings);
            if (amountStr != null && amountStr.isNotEmpty) {
              amount = double.tryParse(amountStr);
            }
          }

          // 构建 key
          final key = _normalizeText('$item|$description');

          // 如果已存在，保留 rate 和 amount 不为 null 的版本
          if (draftMap.containsKey(key)) {
            final existing = draftMap[key]!;
            if (existing.rate == null && rate != null) {
              draftMap[key] = existing.copyWith(rate: rate);
            }
            if (existing.amount == null && amount != null) {
              draftMap[key] = existing.copyWith(amount: amount);
            }
            if (existing.qty == null && qty != null) {
              draftMap[key] = existing.copyWith(qty: qty);
            }
          } else {
            draftMap[key] = DraftRow(
              item: item,
              description: description,
              unit: unit,
              qty: qty,
              rate: rate,
              amount: amount,
            );
            rowCount++;
          }
        }
        debugPrint(
            '[ExcelParserService] XML: 从工作表 "$sheetName" 读取了 $rowCount 条数据');
      }

      return {
        'map': draftMap,
        'columns': draftColumns,
        'matchedSheet': matchedSheetName,
      };
    } catch (e, st) {
      debugPrint('[ExcelParserService] XML 解析失败: $e');
      debugPrint('[ExcelParserService] XML 堆栈: $st');
      rethrow;
    }
  }

  /// 文本预处理
  /// 确保与 excel_match_service.dart 中的标准化逻辑完全一致
  String _normalizeText(String text) {
    if (text.isEmpty) return '';
    return text
        .toLowerCase()
        .replaceAll(RegExp(r'\s+'), '') // 去除所有空格（包括换行、制表符等）
        .replaceAll(RegExp(r'[()（）]'), ''); // 去除括号
  }

  /// 列索引转字母
  String _colToLetter(int colIndex) {
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

  /// 从 XML 单元格元素读取值
  String? getCellValueFromXmlElement(
      XmlElement cell, List<String> sharedStrings) {
    // 检查单元格类型
    final t = cell.getAttribute('t');

    // 处理内联字符串（inlineStr）
    if (t == 'inlineStr') {
      final inlineStrElement = cell.findElements('is').firstOrNull;
      if (inlineStrElement != null) {
        return _extractTextFromInlineString(inlineStrElement);
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

  /// 从 XML 单元格列表获取指定列的值
  String? getCellValueFromXml(
      List<XmlElement> cells, int colIndex, List<String> sharedStrings) {
    for (var cell in cells) {
      final cellRef = cell.getAttribute('r');
      if (cellRef == null) continue;

      final cellColIndex = _cellRefToColIndex(cellRef);
      if (cellColIndex != colIndex) continue;

      return getCellValueFromXmlElement(cell, sharedStrings);
    }
    return null;
  }

  /// 将单元格引用（如 "A1"）转换为列索引（0-based）
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

  /// 规范化工作表名称（用于匹配）
  String _normalizeSheetName(String sheetName) {
    return sheetName
        .toLowerCase()
        .replaceAll(RegExp(r'\s+'), '')
        .replaceAll(RegExp(r'[_-]'), '');
  }

  /// 从文件名中提取工作表编号（如 "sheet1.xml" -> 1）
  int _extractSheetNumber(String fileName) {
    final match = RegExp(r'sheet(\d+)\.xml').firstMatch(fileName);
    if (match != null) {
      return int.tryParse(match.group(1) ?? '0') ?? 0;
    }
    return 0;
  }

  /// 提取 sharedStrings.xml 中 <si> 节点的文本（支持富文本）
  String _extractTextFromSharedString(XmlElement siElement) {
    final buffer = StringBuffer();
    final textNodes = siElement.findAllElements('t');
    if (textNodes.isEmpty) {
      final raw = siElement.innerText.trim();
      return raw;
    }
    for (final t in textNodes) {
      buffer.write(t.innerText);
    }
    return buffer.toString();
  }

  /// 提取 inlineStr 节点文本（支持多个 <t> 富文本片段）
  String? _extractTextFromInlineString(XmlElement inlineStrElement) {
    final buffer = StringBuffer();
    final textNodes = inlineStrElement.findAllElements('t');
    if (textNodes.isEmpty) {
      final raw = inlineStrElement.innerText.trim();
      return raw.isEmpty ? null : raw;
    }
    for (final t in textNodes) {
      buffer.write(t.innerText);
    }
    final text = buffer.toString().trim();
    return text.isEmpty ? null : text;
  }

  /// 使用 XML 方式解析目标文件（避免 numFmtId 错误）
  /// 返回：{ 'sheetName': String, 'columns': TargetColumns, 'rows': List<Map> }
  Future<Map<String, dynamic>> parseTargetFileWithXml(File targetFile) async {
    try {
      final bytes = await targetFile.readAsBytes();
      final archive = ZipDecoder().decodeBytes(bytes);

      // 读取 workbook.xml 获取工作表列表
      final workbookFile = archive.findFile('xl/workbook.xml');
      if (workbookFile == null) {
        throw Exception('无法找到 workbook.xml，文件可能已损坏');
      }

      final workbookXml = utf8.decode(workbookFile.content as List<int>);
      final workbookDoc = XmlDocument.parse(workbookXml);

      // 获取第一个工作表
      final sheetElements = workbookDoc.findAllElements('sheet').toList();
      if (sheetElements.isEmpty) {
        throw Exception('工作簿中没有工作表');
      }

      final firstSheet = sheetElements.first;
      final sheetName = firstSheet.getAttribute('name') ?? '';
      final sheetId = firstSheet.getAttribute('sheetId') ?? '1';

      // 读取 sharedStrings.xml
      final sharedStringsFile = archive.findFile('xl/sharedStrings.xml');
      final sharedStrings = <String>[];
      if (sharedStringsFile != null) {
        final sharedStringsXml =
            utf8.decode(sharedStringsFile.content as List<int>);
        final sharedStringsDoc = XmlDocument.parse(sharedStringsXml);
        for (var si in sharedStringsDoc.findAllElements('si')) {
          sharedStrings.add(_extractTextFromSharedString(si));
        }
      }

      // 查找工作表文件
      ArchiveFile? sheetFile;
      final expectedSheetFileName = 'xl/worksheets/sheet$sheetId.xml';
      for (final file in archive.files) {
        if (file.name == expectedSheetFileName) {
          sheetFile = file;
          break;
        }
      }

      // 如果精确匹配失败，根据工作表顺序查找
      if (sheetFile == null) {
        final sheetFiles = archive.files
            .where((f) =>
                f.name.contains('xl/worksheets/sheet') &&
                f.name.endsWith('.xml'))
            .toList()
          ..sort((a, b) {
            final numA = _extractSheetNumber(a.name);
            final numB = _extractSheetNumber(b.name);
            return numA.compareTo(numB);
          });
        if (sheetFiles.isNotEmpty) {
          sheetFile = sheetFiles[0];
        }
      }

      if (sheetFile == null) {
        throw Exception('无法找到工作表文件');
      }

      // 解析工作表 XML
      final sheetXml = utf8.decode(sheetFile.content as List<int>);
      final sheetDoc = XmlDocument.parse(sheetXml);
      final sheetData = sheetDoc.findAllElements('sheetData').firstOrNull;
      if (sheetData == null) {
        throw Exception('工作表中没有数据');
      }

      final rows = sheetData.findElements('row').toList();

      // 识别列结构（尝试前30行来找到表头）
      TargetColumns? targetColumns;
      int headerRowIndex = -1;
      for (int i = 0; i < rows.length && i < 30; i++) {
        final row = rows[i];
        final columns =
            _columnIdentifier.identifyTargetColumnsFromXml(row, sharedStrings);
        if (columns != null) {
          targetColumns = columns;
          headerRowIndex = i;
          break;
        }
      }

      if (targetColumns == null) {
        throw Exception('无法识别目标文件的列结构');
      }

      return {
        'sheetName': sheetName,
        'columns': targetColumns,
        'rows': rows,
        'sharedStrings': sharedStrings,
        'headerRowIndex': headerRowIndex,
      };
    } catch (e, st) {
      debugPrint('[ExcelParserService] XML 解析目标文件失败: $e');
      debugPrint('[ExcelParserService] 堆栈: $st');
      rethrow;
    }
  }
}
