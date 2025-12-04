import 'dart:io';
import 'dart:convert';
import 'package:flutter/foundation.dart';
import 'package:excel/excel.dart';
import 'package:string_similarity/string_similarity.dart';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';
import '../models/excel_models.dart';

/// Excel 匹配服务：根据 ss 文档设计，字段驱动匹配
class ExcelMatchService {
  /// 主入口：匹配 Excel 文件
  /// 不再需要 targetColumns 参数，自动识别列结构
  Future<MatchServiceResult> matchExcelFiles({
    required File draftFile,
    required File targetFile,
    required String outputPath,
  }) async {
    final logs = <String>[];
    int matchedCount = 0;
    int totalCount = 0;

    try {
      logs.add('开始匹配流程...');
      logs.add('草稿文件: ${draftFile.path}');
      logs.add('目标文件: ${targetFile.path}');

      // Step 1: 使用 excel 包读取目标文件（尝试读取背景色）
      logs.add('正在加载目标文件...');
      final targetBytes = await targetFile.readAsBytes();
      Excel targetExcel;

      try {
        targetExcel = Excel.decodeBytes(targetBytes);
      } catch (e) {
        debugPrint('[ExcelMatchService] 解析目标文件失败: $e');
        return MatchServiceResult(
          success: false,
          message: '无法解析目标文件，请检查文件格式是否正确',
          matchedCount: 0,
          totalCount: 0,
          logs: logs,
        );
      }

      if (targetExcel.tables.isEmpty) {
        return MatchServiceResult(
          success: false,
          message: '目标文件中没有工作表',
          matchedCount: 0,
          totalCount: 0,
          logs: logs,
        );
      }

      final targetSheetName = targetExcel.tables.keys.first;
      final targetSheet = targetExcel.tables[targetSheetName]!;
      logs.add('目标文件已加载，工作表: $targetSheetName');

      // Step 1.1: 自动识别目标文件列结构（从第一行）
      final targetColumns = _identifyTargetColumnsFromExcel(targetSheet);
      if (targetColumns == null) {
        return MatchServiceResult(
          success: false,
          message:
              '无法识别目标文件的列结构，请确保第一行包含 Item, Description, Unit, Qty, Unit Rate, Amount 等列',
          matchedCount: 0,
          totalCount: 0,
          logs: logs,
        );
      }
      logs.add('目标列结构已识别: Item=${_colToLetter(targetColumns.itemCol)}, '
          'Description=${_colToLetter(targetColumns.descriptionCol)}, '
          'Unit=${_colToLetter(targetColumns.unitCol)}, '
          'Qty=${_colToLetter(targetColumns.qtyCol)}, '
          'Unit Rate=${_colToLetter(targetColumns.unitRateCol)}, '
          'Amount=${_colToLetter(targetColumns.amountCol)}');

      // Step 2: 扫描目标文件的所有子项行（根据 Item 和 Description 匹配）
      logs.add('正在扫描目标文件子项行...');
      final targetRows = _scanTargetSubItemRows(targetSheet, targetColumns);
      totalCount = targetRows.length;
      logs.add('找到 $totalCount 个子项行需要匹配');

      if (targetRows.isEmpty) {
        return MatchServiceResult(
          success: false,
          message: '目标文件中没有找到需要匹配的子项行（Item 应包含小数点，如 1.1, 1.2）',
          matchedCount: 0,
          totalCount: 0,
          logs: logs,
        );
      }

      // Step 3: 构建草稿文件的匹配 Map（自动识别列结构，优先匹配同名工作表）
      logs.add('正在构建草稿文件数据...');
      logs.add('目标文件工作表: $targetSheetName');
      final draftResult =
          await _buildDraftMap(draftFile, targetSheetName: targetSheetName);
      final draftMap = draftResult['map'] as Map<String, DraftRow>?;
      final matchedSheetName = draftResult['matchedSheet'] as String?;
      final parseError = draftResult['error'] as String?;

      // 如果有解析错误，优先显示解析错误
      if (parseError != null) {
        logs.add('❌ 草稿文件解析失败');
        logs.add(
            '错误类型：${parseError.contains('numFmtId') ? 'numFmtId 格式错误' : '未知解析错误'}');
        logs.add('提示：请查看下方详细错误信息');
        return MatchServiceResult(
          success: false,
          message: parseError,
          matchedCount: 0,
          totalCount: totalCount,
          logs: logs,
        );
      }

      if (draftMap == null || draftMap.isEmpty) {
        final sheetInfo = matchedSheetName != null
            ? '未在草稿文件中找到工作表 "$matchedSheetName"'
            : '草稿文件中没有有效数据';
        return MatchServiceResult(
          success: false,
          message: '$sheetInfo，请检查文件格式是否正确。\n\n'
              '提示：\n'
              '1. 如果草稿文件包含多个工作表（如 Schedule_A, Schedule_B），\n'
              '   请确保目标文件使用的工作表名称在草稿文件中存在。\n'
              '2. 如果文件格式不兼容，请尝试用 Microsoft Excel 打开并另存为标准的 .xlsx 格式。\n'
              '3. 如果文件已解锁但仍无法解析，请检查文件是否损坏。',
          matchedCount: 0,
          totalCount: totalCount,
          logs: logs,
        );
      }

      if (matchedSheetName != null) {
        logs.add('✓ 在草稿文件的 "$matchedSheetName" 工作表中找到数据');
      } else {
        logs.add('⚠ 未找到同名工作表，已在所有工作表中匹配');
      }
      logs.add('草稿文件已加载，共 ${draftMap.length} 条数据');

      // Step 4: 字段驱动的匹配算法
      logs.add('开始匹配...');
      final matchResults = <MatchResult>[];
      for (final targetRow in targetRows) {
        final matchResult = _matchTargetRow(targetRow, draftMap);
        matchResults.add(matchResult);

        if (matchResult.matched) {
          matchedCount++;
          logs.add(
            '✓ 匹配成功 [${matchResult.matchType}]: ${targetRow.item} | ${targetRow.description} → Rate: ${matchResult.rate}, Amount: ${matchResult.amount}',
          );
        } else {
          logs.add(
            '✗ 匹配失败: ${targetRow.item} | ${targetRow.description}',
          );
        }
      }

      // Step 5: 写入匹配值到目标列
      logs.add('正在写入匹配值...');
      _writeMatchedValuesToExcel(targetExcel, targetSheetName, targetSheet,
          matchResults, targetColumns);
      logs.add('匹配值已写入');

      // Step 6: 保存处理后的目标文件
      logs.add('正在保存文件...');
      final outputBytes = targetExcel.save();
      if (outputBytes == null) {
        return MatchServiceResult(
          success: false,
          message: '保存文件失败',
          matchedCount: matchedCount,
          totalCount: totalCount,
          logs: logs,
        );
      }

      // 确保输出目录存在
      final outputFile = File(outputPath);
      final outputDir = outputFile.parent;
      if (!await outputDir.exists()) {
        await outputDir.create(recursive: true);
        logs.add('创建输出目录: ${outputDir.path}');
      }

      await outputFile.writeAsBytes(outputBytes);
      logs.add('文件已保存: $outputPath');

      // Step 7: 清除匹配列的黄色背景（使用 XML 直接操作）
      logs.add('正在清除匹配列的黄色背景...');
      await _clearMatchedCellBackgrounds(
        outputFile,
        targetSheetName,
        matchResults,
        targetColumns,
      );
      logs.add('黄色背景已清除');

      return MatchServiceResult(
        success: true,
        message:
            '匹配完成：成功 $matchedCount/$totalCount，输出文件：${outputFile.uri.pathSegments.last}',
        matchedCount: matchedCount,
        totalCount: totalCount,
        logs: logs,
      );
    } catch (e, st) {
      debugPrint('[ExcelMatchService] 匹配失败: $e');
      debugPrint('[ExcelMatchService] 堆栈: $st');
      logs.add('错误: $e');
      return MatchServiceResult(
        success: false,
        message: '匹配失败：$e',
        matchedCount: matchedCount,
        totalCount: totalCount,
        logs: logs,
      );
    }
  }

  /// 自动识别目标文件列结构（从表头行，使用 excel 包）
  /// 根据实际表格结构：表头在第 6 行（索引 5）
  TargetColumns? _identifyTargetColumnsFromExcel(Sheet sheet) {
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
        if (rowIndex >= 10) break; // 只检查前 10 行

        for (var cell in row) {
          if (cell == null) continue;
          final cellText = (_getCellValueFromExcel(cell) ?? '').toLowerCase();
          if (_matchesColumnName(cellText, ['item', 'item no', 'no.', '序号'])) {
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
          _matchesColumnName(headerText, ['item', 'item no', 'no.', '序号'])) {
        itemCol = colIndex;
      }
      if (descriptionCol == null &&
          _matchesColumnName(headerText,
              ['description', 'desc', '工作描述', 'description of work'])) {
        descriptionCol = colIndex;
      }
      if (unitCol == null && _matchesColumnName(headerText, ['unit', 'u'])) {
        unitCol = colIndex;
      }
      if (qtyCol == null &&
          _matchesColumnName(headerText, ['qty', 'quantity'])) {
        qtyCol = colIndex;
      }
      if (unitRateCol == null) {
        final similarity = _getBestSimilarity(
            headerText, ['rate', 'unit rate', '(b)', 'unit rate (hk\$)']);
        if (similarity > 0.7) {
          unitRateCol = colIndex;
        }
      }
      if (amountCol == null) {
        final similarity = _getBestSimilarity(
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

  /// 扫描目标文件的所有子项行（不再依赖黄色单元格）
  /// 根据实际表格结构：数据从第 8 行开始（索引 7），跳过主项行
  /// 子项行的 Item 包含小数点（如 "1.1", "1.2"），主项行的 Item 是纯数字（如 "1"）
  List<TargetRow> _scanTargetSubItemRows(
    Sheet sheet,
    TargetColumns columns,
  ) {
    final targetRows = <TargetRow>[];

    // 从第 8 行开始扫描（索引 7，excel 包是 0-based）
    // 跳过表头（前 7 行，索引 0-6）
    for (var row in sheet.rows) {
      final rowIndex = row.first?.rowIndex ?? -1;
      if (rowIndex < 7) continue; // 跳过表头（前 7 行）

      // 提取 Item
      final itemCell = _getCellByIndex(sheet, rowIndex, columns.itemCol);
      final item = _getCellValueFromExcel(itemCell) ?? '';
      if (item.isEmpty) continue;

      // 只处理子项行：Item 包含小数点（如 "1.1", "1.2"）
      // 跳过主项行：Item 是纯数字（如 "1", "2"）
      final itemTrimmed = item.trim();
      if (RegExp(r'^\d+$').hasMatch(itemTrimmed)) {
        // 纯数字，是主项，跳过
        continue;
      }

      // 提取 Description
      final descCell = _getCellByIndex(sheet, rowIndex, columns.descriptionCol);
      final description = _getCellValueFromExcel(descCell) ?? '';

      // 跳过包含 "Remark" 的行（这些是备注行，不应该匹配）
      final descriptionLower = description.toLowerCase();
      if (descriptionLower.contains('remark') ||
          itemTrimmed.toLowerCase().contains('remark')) {
        debugPrint(
            '[ExcelMatchService] 跳过备注行: Item=$item, Description=$description');
        continue;
      }

      // 提取 Unit（可选）
      String unit = '';
      if (columns.unitCol >= 0) {
        final unitCell = _getCellByIndex(sheet, rowIndex, columns.unitCol);
        unit = _getCellValueFromExcel(unitCell) ?? '';
      }

      // 提取 Qty（可选）
      double qty = 0.0;
      if (columns.qtyCol >= 0) {
        final qtyCell = _getCellByIndex(sheet, rowIndex, columns.qtyCol);
        final qtyStr = _getCellValueFromExcel(qtyCell) ?? '';
        qty = double.tryParse(qtyStr) ?? 0.0;
      }

      // 检查 Unit Rate 和 Amount 列是否为空（需要填充）
      final rateCell = _getCellByIndex(sheet, rowIndex, columns.unitRateCol);
      final amountCell = _getCellByIndex(sheet, rowIndex, columns.amountCol);
      final rateValue = _getCellValueFromExcel(rateCell);
      final amountValue = _getCellValueFromExcel(amountCell);

      // 如果 Unit Rate 或 Amount 列为空，需要匹配填充
      final needsRate = rateValue == null || rateValue.trim().isEmpty;
      final needsAmount = amountValue == null || amountValue.trim().isEmpty;

      if (needsRate || needsAmount) {
        targetRows.add(
          TargetRow(
            item: item,
            description: description,
            unit: unit,
            qty: qty,
            rowIndex: rowIndex,
            rateColumn: columns.unitRateCol,
            amountColumn: columns.amountCol,
          ),
        );
      }
    }

    return targetRows;
  }

  /// 构建草稿文件的匹配 Map（自动识别列结构，优先匹配同名工作表）
  /// [targetSheetName] 目标文件的工作表名称，用于优先匹配同名工作表
  Future<Map<String, dynamic>> _buildDraftMap(File draftFile,
      {String? targetSheetName}) async {
    final draftMap = <String, DraftRow>{};
    final bytes = await draftFile.readAsBytes();

    // 尝试使用 excel 包解析
    Excel? excel;
    String? parseError;

    try {
      excel = Excel.decodeBytes(bytes);
      debugPrint('[ExcelMatchService] 使用 excel 包成功解析草稿文件');
    } catch (e) {
      debugPrint('[ExcelMatchService] excel 包解析失败，尝试使用 XML 直接解析: $e');

      // 尝试使用 XML 直接解析作为备用方案
      try {
        return await _buildDraftMapWithSyncfusion(
          bytes,
          draftMap,
          targetSheetName: targetSheetName,
        );
      } catch (xmlError) {
        debugPrint('[ExcelMatchService] XML 直接解析也失败: $xmlError');

        // 两种方法都失败了
        parseError = e.toString();
        final errorStr = e.toString();
        if (errorStr.contains('numFmtId')) {
          parseError = '';
        } else {
          parseError = '';
        }

        return {
          'map': draftMap,
          'columns': null,
          'matchedSheet': null,
          'error': parseError,
        };
      }
    }

    if (excel.tables.isEmpty) {
      debugPrint('[ExcelMatchService] 草稿文件中没有工作表');
      return {
        'map': draftMap,
        'columns': null,
        'matchedSheet': null,
        'error': '草稿文件中没有工作表，请检查文件是否正确',
      };
    }

    // 列出所有工作表名称
    final allSheetNames = excel.tables.keys.toList();
    debugPrint('[ExcelMatchService] 草稿文件包含的工作表: ${allSheetNames.join(", ")}');

    // 优先查找同名工作表
    String? matchedSheetName;
    List<String> sheetsToProcess = [];

    if (targetSheetName != null) {
      // 尝试精确匹配
      if (excel.tables.containsKey(targetSheetName)) {
        matchedSheetName = targetSheetName;
        sheetsToProcess = [targetSheetName];
        debugPrint('[ExcelMatchService] 找到同名工作表: $targetSheetName');
      } else {
        // 尝试模糊匹配（不区分大小写，忽略空格）
        final normalizedTarget = _normalizeSheetName(targetSheetName);
        for (var sheetName in allSheetNames) {
          if (_normalizeSheetName(sheetName) == normalizedTarget) {
            matchedSheetName = sheetName;
            sheetsToProcess = [sheetName];
            debugPrint(
                '[ExcelMatchService] 找到相似工作表: $sheetName (目标: $targetSheetName)');
            break;
          }
        }
      }
    }

    // 如果没找到同名工作表，处理所有工作表
    if (sheetsToProcess.isEmpty) {
      sheetsToProcess = allSheetNames;
      debugPrint('[ExcelMatchService] 未找到同名工作表，将在所有工作表中匹配');
    }

    // 遍历要处理的工作表
    DraftColumns? draftColumns;
    for (var tableName in sheetsToProcess) {
      final sheet = excel.tables[tableName]!;
      debugPrint('[ExcelMatchService] 正在处理工作表: $tableName');

      // 识别草稿文件的列结构（从第一行）
      if (draftColumns == null) {
        draftColumns = _identifyDraftColumns(sheet);
        if (draftColumns == null) {
          debugPrint('[ExcelMatchService] 无法识别工作表 "$tableName" 的列结构，跳过');
          continue;
        }
        debugPrint(
            '[ExcelMatchService] 已识别列结构: Item=${_colToLetter(draftColumns.itemCol)}, '
            'Description=${_colToLetter(draftColumns.descriptionCol)}, '
            'Rate=${_colToLetter(draftColumns.rateCol)}, '
            'Amount=${_colToLetter(draftColumns.amountCol)}');
      }

      // 从第2行开始读取数据（索引 1，excel 包是 0-based）
      int rowCount = 0;
      for (var row in sheet.rows) {
        final rowIndex = row.first?.rowIndex ?? -1;
        if (rowIndex < 1) continue; // 跳过表头

        // 提取字段
        final itemCell = _getCellByIndex(sheet, rowIndex, draftColumns.itemCol);
        final item = _getCellValueFromExcel(itemCell) ?? '';
        if (item.isEmpty) continue;

        final descCell =
            _getCellByIndex(sheet, rowIndex, draftColumns.descriptionCol);
        final description = _getCellValueFromExcel(descCell) ?? '';

        // 跳过包含 "Remark" 的行（这些是备注行，不应该匹配）
        final descriptionLower = description.toLowerCase();
        final itemTrimmed = item.trim();
        if (descriptionLower.contains('remark') ||
            itemTrimmed.toLowerCase().contains('remark')) {
          debugPrint(
              '[ExcelMatchService] 草稿文件：跳过备注行: Item=$item, Description=$description');
          continue;
        }

        String? unit;
        if (draftColumns.unitCol >= 0) {
          final unitCell =
              _getCellByIndex(sheet, rowIndex, draftColumns.unitCol);
          unit = _getCellValueFromExcel(unitCell);
        }

        double? qty;
        if (draftColumns.qtyCol >= 0) {
          final qtyCell = _getCellByIndex(sheet, rowIndex, draftColumns.qtyCol);
          final qtyStr = _getCellValueFromExcel(qtyCell);
          if (qtyStr != null && qtyStr.isNotEmpty) {
            qty = double.tryParse(qtyStr);
          }
        }

        double? rate;
        if (draftColumns.rateCol >= 0) {
          final rateCell =
              _getCellByIndex(sheet, rowIndex, draftColumns.rateCol);
          final rateStr = _getCellValueFromExcel(rateCell);
          if (rateStr != null && rateStr.isNotEmpty) {
            rate = double.tryParse(rateStr);
          }
        }

        double? amount;
        if (draftColumns.amountCol >= 0) {
          final amountCell =
              _getCellByIndex(sheet, rowIndex, draftColumns.amountCol);
          final amountStr = _getCellValueFromExcel(amountCell);
          if (amountStr != null && amountStr.isNotEmpty) {
            amount = double.tryParse(amountStr);
          }
        }

        // 构建 key（不再使用固定键，而是使用字段组合）
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
      debugPrint('[ExcelMatchService] 从工作表 "$tableName" 读取了 $rowCount 条数据');
    }

    return {
      'map': draftMap,
      'columns': draftColumns,
      'matchedSheet': matchedSheetName,
    };
  }

  /// 规范化工作表名称（用于匹配）
  String _normalizeSheetName(String sheetName) {
    return sheetName
        .toLowerCase()
        .replaceAll(RegExp(r'\s+'), '') // 去除所有空格
        .replaceAll(RegExp(r'[_-]'), ''); // 去除下划线和连字符
  }

  /// 使用直接解析 XML 的方法解析草稿文件（备用方案）
  /// .xlsx 文件本质上是 ZIP 压缩的 XML 文件
  Future<Map<String, dynamic>> _buildDraftMapWithSyncfusion(
    List<int> bytes,
    Map<String, DraftRow> draftMap, {
    String? targetSheetName,
  }) async {
    debugPrint('[ExcelMatchService] 使用 XML 直接解析草稿文件（备用方案）...');

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
          '[ExcelMatchService] XML: 草稿文件包含的工作表: ${allSheetNames.join(", ")}');

      // 优先查找同名工作表
      String? matchedSheetName;
      List<String> sheetsToProcess = [];

      if (targetSheetName != null) {
        if (sheetMap.containsKey(targetSheetName)) {
          matchedSheetName = targetSheetName;
          sheetsToProcess = [targetSheetName];
          debugPrint('[ExcelMatchService] XML: 找到同名工作表: $targetSheetName');
        } else {
          // 尝试模糊匹配
          final normalizedTarget = _normalizeSheetName(targetSheetName);
          for (var sheetName in allSheetNames) {
            if (_normalizeSheetName(sheetName) == normalizedTarget) {
              matchedSheetName = sheetName;
              sheetsToProcess = [sheetName];
              debugPrint(
                  '[ExcelMatchService] XML: 找到相似工作表: $sheetName (目标: $targetSheetName)');
              break;
            }
          }
        }
      }

      // 如果没找到同名工作表，处理所有工作表
      if (sheetsToProcess.isEmpty) {
        sheetsToProcess = allSheetNames;
        debugPrint('[ExcelMatchService] XML: 未找到同名工作表，将在所有工作表中匹配');
      }

      // 读取 sharedStrings.xml（如果存在）
      final sharedStringsFile = archive.findFile('xl/sharedStrings.xml');
      final sharedStrings = <String>[];
      if (sharedStringsFile != null) {
        final sharedStringsXml =
            utf8.decode(sharedStringsFile.content as List<int>);
        final sharedStringsDoc = XmlDocument.parse(sharedStringsXml);
        for (var si in sharedStringsDoc.findAllElements('si')) {
          final t = si.findElements('t').firstOrNull;
          if (t != null) {
            sharedStrings.add(t.innerText);
          } else {
            sharedStrings.add('');
          }
        }
      }

      // 遍历要处理的工作表
      DraftColumns? draftColumns;
      final sheetFiles = archive.files
          .where((f) =>
              f.name.contains('xl/worksheets/sheet') && f.name.endsWith('.xml'))
          .toList()
        ..sort((a, b) => a.name.compareTo(b.name));

      int sheetIndex = 0;
      for (var sheetName in sheetsToProcess) {
        if (sheetIndex >= sheetFiles.length) break;

        final sheetFile = sheetFiles[sheetIndex];
        sheetIndex++;

        debugPrint(
            '[ExcelMatchService] XML: 正在处理工作表: $sheetName (文件: ${sheetFile.name})');

        // 解析工作表 XML
        final sheetXml = utf8.decode(sheetFile.content as List<int>);
        final sheetDoc = XmlDocument.parse(sheetXml);
        final sheetData = sheetDoc.findAllElements('sheetData').firstOrNull;
        if (sheetData == null) continue;

        final rows = sheetData.findElements('row').toList();

        // 识别列结构（尝试前20行来找到表头）
        int headerRowIndex = -1;
        if (draftColumns == null) {
          // 尝试前20行来找到表头（有些Excel文件表头可能在后面）
          for (int i = 0; i < rows.length && i < 20; i++) {
            final row = rows[i];
            debugPrint('[ExcelMatchService] XML: 尝试第 ${i + 1} 行作为表头...');
            draftColumns = _identifyDraftColumnsFromXml(row, sharedStrings);
            if (draftColumns != null) {
              headerRowIndex = i;
              debugPrint('[ExcelMatchService] XML: 在第 ${i + 1} 行找到表头！');
              debugPrint(
                  '[ExcelMatchService] XML: 已识别列结构: Item=${_colToLetter(draftColumns.itemCol)}, '
                  'Description=${_colToLetter(draftColumns.descriptionCol)}, '
                  'Rate=${_colToLetter(draftColumns.rateCol)}, '
                  'Amount=${_colToLetter(draftColumns.amountCol)}');
              debugPrint(
                  '[ExcelMatchService] XML: 表头在第 ${headerRowIndex + 1} 行，将从第 ${headerRowIndex + 2} 行开始读取数据');
              break;
            }
          }

          if (draftColumns == null) {
            debugPrint(
                '[ExcelMatchService] XML: 在前20行中无法识别工作表 "$sheetName" 的列结构，跳过');
            debugPrint(
                '[ExcelMatchService] XML: 提示：请检查Excel文件，确保表头行包含 Item、Description、Rate、Amount 等列名');
            continue;
          }
        }

        // 从表头之后开始读取数据
        int rowCount = 0;

        // 从表头之后开始读取
        for (int i = 0; i < rows.length; i++) {
          if (headerRowIndex >= 0 && i <= headerRowIndex) {
            continue; // 跳过表头行
          }

          final row = rows[i];

          final cells = row.findElements('c').toList();
          if (cells.isEmpty) continue;

          // 提取字段
          final item = _getCellValueFromXml(
                  cells, draftColumns.itemCol, sharedStrings) ??
              '';
          if (item.isEmpty) continue;

          final description = _getCellValueFromXml(
                  cells, draftColumns.descriptionCol, sharedStrings) ??
              '';

          // 跳过包含 "Remark" 的行（这些是备注行，不应该匹配）
          final descriptionLower = description.toLowerCase();
          final itemTrimmed = item.trim();
          if (descriptionLower.contains('remark') ||
              itemTrimmed.toLowerCase().contains('remark')) {
            debugPrint(
                '[ExcelMatchService] XML草稿文件：跳过备注行: Item=$item, Description=$description');
            continue;
          }

          String? unit;
          if (draftColumns.unitCol >= 0) {
            unit = _getCellValueFromXml(
                cells, draftColumns.unitCol, sharedStrings);
          }

          double? qty;
          if (draftColumns.qtyCol >= 0) {
            final qtyStr =
                _getCellValueFromXml(cells, draftColumns.qtyCol, sharedStrings);
            if (qtyStr != null && qtyStr.isNotEmpty) {
              qty = double.tryParse(qtyStr);
            }
          }

          double? rate;
          if (draftColumns.rateCol >= 0) {
            final rateStr = _getCellValueFromXml(
                cells, draftColumns.rateCol, sharedStrings);
            if (rateStr != null && rateStr.isNotEmpty) {
              rate = double.tryParse(rateStr);
            }
          }

          double? amount;
          if (draftColumns.amountCol >= 0) {
            final amountStr = _getCellValueFromXml(
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
            '[ExcelMatchService] XML: 从工作表 "$sheetName" 读取了 $rowCount 条数据');
      }

      return {
        'map': draftMap,
        'columns': draftColumns,
        'matchedSheet': matchedSheetName,
      };
    } catch (e, st) {
      debugPrint('[ExcelMatchService] XML 解析失败: $e');
      debugPrint('[ExcelMatchService] XML 堆栈: $st');
      rethrow;
    }
  }

  /// 从 XML 行识别列结构
  DraftColumns? _identifyDraftColumnsFromXml(
      XmlElement firstRow, List<String> sharedStrings) {
    final cells = firstRow.findElements('c').toList();
    if (cells.isEmpty) {
      debugPrint('[ExcelMatchService] XML: 第一行没有单元格');
      return null;
    }

    debugPrint('[ExcelMatchService] XML: 第一行有 ${cells.length} 个单元格');
    debugPrint('[ExcelMatchService] XML: 共享字符串数量: ${sharedStrings.length}');

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
            '[ExcelMatchService] XML: 列 ${_colToLetter(colIndex)} ($colIndex) [$cellType]: "$headerText"');
      } else {
        debugPrint(
            '[ExcelMatchService] XML: 列 ${_colToLetter(colIndex)} ($colIndex) [$cellType]: (空)');
      }

      // 识别各列
      if (itemCol == null &&
          _matchesColumnName(
              headerTextLower, ['item', 'item no', 'no.', '序号'])) {
        itemCol = colIndex;
        debugPrint(
            '[ExcelMatchService] XML: 识别到 Item 列: ${_colToLetter(colIndex)}');
      }
      if (descriptionCol == null &&
          _matchesColumnName(headerTextLower,
              ['description', 'desc', '工作描述', 'description of work'])) {
        descriptionCol = colIndex;
        debugPrint(
            '[ExcelMatchService] XML: 识别到 Description 列: ${_colToLetter(colIndex)}');
      }
      if (unitCol == null &&
          _matchesColumnName(headerTextLower, ['unit', 'u'])) {
        unitCol = colIndex;
        debugPrint(
            '[ExcelMatchService] XML: 识别到 Unit 列: ${_colToLetter(colIndex)}');
      }
      if (qtyCol == null &&
          _matchesColumnName(headerTextLower, ['qty', 'quantity'])) {
        qtyCol = colIndex;
        debugPrint(
            '[ExcelMatchService] XML: 识别到 Qty 列: ${_colToLetter(colIndex)}');
      }
      if (rateCol == null) {
        final similarity = _getBestSimilarity(
            headerTextLower, ['rate', 'unit rate', '(b)', 'unit rate (hk\$)']);
        if (similarity > 0.7) {
          rateCol = colIndex;
          debugPrint(
              '[ExcelMatchService] XML: 识别到 Rate 列: ${_colToLetter(colIndex)} (相似度: $similarity)');
        }
      }
      if (amountCol == null) {
        final similarity = _getBestSimilarity(
            headerTextLower, ['amount', 'total', '(c)', 'amount (hk\$)']);
        if (similarity > 0.7) {
          amountCol = colIndex;
          debugPrint(
              '[ExcelMatchService] XML: 识别到 Amount 列: ${_colToLetter(colIndex)} (相似度: $similarity)');
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

  /// 从 XML 单元格元素直接读取值
  String? _getCellValueFromXmlElement(
      XmlElement cell, List<String> sharedStrings) {
    // 检查单元格类型
    final t = cell.getAttribute('t');

    // 处理内联字符串（inlineStr）- 直接包含文本，不使用共享字符串
    if (t == 'inlineStr') {
      final inlineStrElement = cell.findElements('is').firstOrNull;
      if (inlineStrElement != null) {
        final tElement = inlineStrElement.findElements('t').firstOrNull;
        if (tElement != null) {
          final value = tElement.innerText.trim();
          return value.isEmpty ? null : value;
        }
      }
    }

    // 处理共享字符串（s）- 值是指向共享字符串表的索引
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

    // 处理普通值（v）- 数字、日期等
    final v = cell.findElements('v').firstOrNull;
    if (v != null) {
      final value = v.innerText.trim();
      return value.isEmpty ? null : value;
    }

    return null;
  }

  /// 从 XML 单元格列表获取指定列的值
  String? _getCellValueFromXml(
      List<XmlElement> cells, int colIndex, List<String> sharedStrings) {
    for (var cell in cells) {
      final cellRef = cell.getAttribute('r');
      if (cellRef == null) continue;

      final cellColIndex = _cellRefToColIndex(cellRef);
      if (cellColIndex != colIndex) continue;

      return _getCellValueFromXmlElement(cell, sharedStrings);
    }
    return null;
  }

  /// 将单元格引用（如 "A1"）转换为列索引（0-based）
  int _cellRefToColIndex(String cellRef) {
    // 提取列字母部分（如 "A1" -> "A"）
    final match = RegExp(r'^([A-Z]+)').firstMatch(cellRef);
    if (match == null) return -1;

    final colLetters = match.group(1)!;
    int colIndex = 0;
    for (var i = 0; i < colLetters.length; i++) {
      colIndex = colIndex * 26 + (colLetters.codeUnitAt(i) - 64);
    }
    return colIndex - 1; // 转换为 0-based
  }

  /// 自动识别草稿文件列结构（从第一行标题）
  DraftColumns? _identifyDraftColumns(Sheet sheet) {
    // 读取第一行（索引 0，excel 包是 0-based）
    final headerRow = sheet.rows.firstOrNull;
    if (headerRow == null) return null;

    int? itemCol, descriptionCol, unitCol, qtyCol, rateCol, amountCol;

    // 遍历第一行的所有单元格
    for (var cell in headerRow) {
      if (cell == null) continue;
      final colIndex = cell.columnIndex;
      final headerText = (_getCellValueFromExcel(cell) ?? '').toLowerCase();

      // 识别各列
      if (itemCol == null &&
          _matchesColumnName(headerText, ['item', 'item no', 'no.', '序号'])) {
        itemCol = colIndex;
      }
      if (descriptionCol == null &&
          _matchesColumnName(headerText,
              ['description', 'desc', '工作描述', 'description of work'])) {
        descriptionCol = colIndex;
      }
      if (unitCol == null && _matchesColumnName(headerText, ['unit', 'u'])) {
        unitCol = colIndex;
      }
      if (qtyCol == null &&
          _matchesColumnName(headerText, ['qty', 'quantity'])) {
        qtyCol = colIndex;
      }
      if (rateCol == null) {
        final similarity = _getBestSimilarity(
            headerText, ['rate', 'unit rate', '(b)', 'unit rate (hk\$)']);
        if (similarity > 0.7) {
          rateCol = colIndex;
        }
      }
      if (amountCol == null) {
        final similarity = _getBestSimilarity(
            headerText, ['amount', 'total', '(c)', 'amount (hk\$)']);
        if (similarity > 0.7) {
          amountCol = colIndex;
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

  /// 字段驱动的匹配算法（根据 ss 文档）
  MatchResult _matchTargetRow(
    TargetRow target,
    Map<String, DraftRow> draftMap,
  ) {
    final normalizedKey =
        _normalizeText('${target.item}|${target.description}');

    // ① 强匹配（优先）：Item 完全一致
    if (draftMap.containsKey(normalizedKey)) {
      final draft = draftMap[normalizedKey]!;
      return MatchResult(
        target: target,
        draft: draft,
        matched: true,
        rate: draft.rate,
        amount: draft.amount,
        matchType: 'strong',
      );
    }

    // ② 次匹配：Description 相似度 ≥ 0.8，且 Unit 匹配（可选）
    double bestSimilarity = 0.0;
    DraftRow? bestMatch;

    for (final draft in draftMap.values) {
      // 检查 Item 是否相似（允许部分匹配）
      final targetItemNorm = _normalizeText(target.item);
      final draftItemNorm = _normalizeText(draft.item);

      // 如果 Item 完全不匹配，计算相似度
      if (targetItemNorm != draftItemNorm &&
          !targetItemNorm.contains(draftItemNorm) &&
          !draftItemNorm.contains(targetItemNorm)) {
        final itemSimilarity =
            StringSimilarity.compareTwoStrings(targetItemNorm, draftItemNorm);
        if (itemSimilarity < 0.5) {
          continue; // Item 相似度太低，跳过
        }
      }

      // 计算 Description 相似度
      final similarity = StringSimilarity.compareTwoStrings(
        _normalizeText(target.description),
        _normalizeText(draft.description),
      );

      if (similarity >= 0.8) {
        // 检查 Unit 是否匹配（可选）
        bool unitMatch = true;
        if (target.unit.isNotEmpty && draft.unit != null) {
          unitMatch =
              _normalizeText(target.unit) == _normalizeText(draft.unit!);
        }

        // 检查 Qty 是否匹配（可选）
        bool qtyMatch = true;
        if (target.qty > 0 && draft.qty != null) {
          qtyMatch = (target.qty - draft.qty!).abs() < 0.01;
        }

        // 如果相似度更高，且 Unit 或 Qty 匹配，更新最佳匹配
        if (similarity > bestSimilarity && (unitMatch || qtyMatch)) {
          bestSimilarity = similarity;
          bestMatch = draft;
        }
      }
    }

    if (bestMatch != null) {
      return MatchResult(
        target: target,
        draft: bestMatch,
        matched: true,
        rate: bestMatch.rate,
        amount: bestMatch.amount,
        matchType: 'medium',
      );
    }

    // ③ 兜底匹配：Description 包含匹配
    for (final draft in draftMap.values) {
      final targetDescNorm = _normalizeText(target.description);
      final draftDescNorm = _normalizeText(draft.description);

      if (targetDescNorm.contains(draftDescNorm) ||
          draftDescNorm.contains(targetDescNorm)) {
        return MatchResult(
          target: target,
          draft: draft,
          matched: true,
          rate: draft.rate,
          amount: draft.amount,
          matchType: 'weak',
        );
      }
    }

    // 没有匹配成功
    return MatchResult(
      target: target,
      matched: false,
    );
  }

  /// 写入匹配值到目标列（使用 excel 包）
  /// 只修改单元格的值，保持原有格式不变
  /// 如果目标文件有公式，保留公式；否则写入值
  void _writeMatchedValuesToExcel(
    Excel excel,
    String sheetName,
    Sheet sheet,
    List<MatchResult> matchResults,
    TargetColumns targetColumns,
  ) {
    // 设置列宽，避免显示 #######
    // Rate 列
    try {
      sheet.setColumnWidth(targetColumns.unitRateCol, 15.0); // Unit Rate 列
    } catch (e) {
      debugPrint('[ExcelMatchService] 设置 Rate 列宽失败: $e');
    }

    // Amount 列
    try {
      sheet.setColumnWidth(targetColumns.amountCol, 15.0); // Amount 列
    } catch (e) {
      debugPrint('[ExcelMatchService] 设置 Amount 列宽失败: $e');
    }

    for (final result in matchResults) {
      if (!result.matched) continue;

      final target = result.target;
      final rowIndex = target.rowIndex; // excel 包是 0-based

      // 写入 Rate（只修改 Rate 列，不影响其他列）
      if (result.rate != null) {
        final rateCell = _getCellByIndex(sheet, rowIndex, target.rateColumn);
        if (rateCell != null) {
          // 只修改值，不修改样式（让 Excel 使用默认样式，清除黄色背景）
          rateCell.value = DoubleCellValue(result.rate!);
          // 不设置 cellStyle，让单元格使用默认样式（无背景色）
          // 这样可以清除黄色背景
        }
      }

      // 处理 Amount：检查是否有公式，如果有则设置公式，否则写入值
      if (result.amount != null) {
        final amountCell =
            _getCellByIndex(sheet, rowIndex, target.amountColumn);
        if (amountCell != null) {
          // 检查目标文件是否已经有公式
          final existingValue = amountCell.value;
          bool hasFormula = existingValue is FormulaCellValue;

          // 如果目标文件有公式，或者 Qty 列有值，尝试设置公式（Amount = Qty × Rate）
          if (hasFormula || _shouldUseFormula(target, targetColumns)) {
            // 构建公式：Amount = Qty × Rate
            // 例如：=D8*E8（假设 Qty 在 D 列，Rate 在 E 列）
            final qtyCol = targetColumns.qtyCol >= 0
                ? targetColumns.qtyCol
                : target.rateColumn - 1;
            final rateCol = target.rateColumn;
            final qtyColLetter = _colToLetter(qtyCol);
            final rateColLetter = _colToLetter(rateCol);
            final rowNum = rowIndex + 1; // Excel 行号从 1 开始

            final formula = '=$qtyColLetter$rowNum*$rateColLetter$rowNum';

            try {
              // 设置公式（公式会自动计算，不需要点击）
              amountCell.value = FormulaCellValue(formula);
              // 不设置 cellStyle，让单元格使用默认样式（无背景色）
              debugPrint(
                  '[ExcelMatchService] 设置 Amount 公式: $formula (行 $rowNum)');
            } catch (e) {
              debugPrint('[ExcelMatchService] 设置公式失败，改为写入值: $e');
              // 如果设置公式失败，回退到写入值
              amountCell.value = DoubleCellValue(result.amount!);
            }
          } else {
            // 没有公式，直接写入值
            amountCell.value = DoubleCellValue(result.amount!);
          }
          // 不设置 cellStyle，让单元格使用默认样式（无背景色）
        }
      }
    }
  }

  /// 判断是否应该使用公式（根据目标行的 Qty 列是否有值）
  bool _shouldUseFormula(TargetRow target, TargetColumns columns) {
    // 如果 Qty 列有值，应该使用公式
    return columns.qtyCol >= 0 && target.qty > 0;
  }

  /// 清除匹配单元格的黄色背景（使用 XML 直接操作）
  /// 在文件保存后，通过 XML 操作清除 Rate 和 Amount 列的黄色背景
  Future<void> _clearMatchedCellBackgrounds(
    File excelFile,
    String sheetName,
    List<MatchResult> matchResults,
    TargetColumns targetColumns,
  ) async {
    try {
      final bytes = await excelFile.readAsBytes();
      final archive = ZipDecoder().decodeBytes(bytes);

      // 查找目标工作表文件
      ArchiveFile? targetSheetFile;
      for (final file in archive.files) {
        if (file.name.contains('xl/worksheets/sheet') &&
            file.name.endsWith('.xml')) {
          // 简化处理：使用第一个工作表文件（通常 sheet1.xml 对应第一个工作表）
          targetSheetFile ??= file;
        }
      }

      if (targetSheetFile == null) {
        debugPrint('[ExcelMatchService] 未找到工作表文件，跳过清除背景色');
        return;
      }

      // 解析工作表 XML
      final sheetXml = utf8.decode(targetSheetFile.content as List<int>);
      final sheetDoc = XmlDocument.parse(sheetXml);
      final sheetData = sheetDoc.findAllElements('sheetData').firstOrNull;
      if (sheetData == null) return;

      // 收集所有需要清除背景色的单元格引用
      final cellsToClear = <String>{};
      for (final result in matchResults) {
        if (!result.matched) continue;
        final target = result.target;
        final rowNum = target.rowIndex + 1; // Excel 行号从 1 开始

        // Rate 列
        if (result.rate != null) {
          final rateColLetter = _colToLetter(target.rateColumn);
          cellsToClear.add('$rateColLetter$rowNum');
        }

        // Amount 列
        if (result.amount != null) {
          final amountColLetter = _colToLetter(target.amountColumn);
          cellsToClear.add('$amountColLetter$rowNum');
        }
      }

      // 清除这些单元格的背景色（只清除 Rate 和 Amount 列，不影响其他列如 C 列）
      int clearedCount = 0;
      final clearedCells = <String>[];
      for (final cell in sheetData.findAllElements('c')) {
        final cellRef = cell.getAttribute('r');
        if (cellRef == null || !cellsToClear.contains(cellRef)) continue;

        // 双重验证：确保单元格引用在 cellsToClear 集合中
        // 并且只处理 Rate 和 Amount 列的单元格
        final cellCol = _cellRefToColIndex(cellRef);
        final isRateCol = cellCol == targetColumns.unitRateCol;
        final isAmountCol = cellCol == targetColumns.amountCol;

        if (!isRateCol && !isAmountCol) {
          // 不是 Rate 或 Amount 列，跳过（保护 C 列等其他列）
          debugPrint('[ExcelMatchService] 跳过非匹配列: $cellRef (列索引: $cellCol)');
          continue;
        }

        // Excel 的背景色通过 s 属性引用样式表中的 fill
        // 清除 s 属性，让单元格使用默认样式（无背景色）
        final sAttr = cell.getAttribute('s');
        if (sAttr != null) {
          // 清除样式引用，单元格将使用默认样式（无背景色）
          cell.attributes.removeWhere((attr) => attr.name.local == 's');
          clearedCount++;
          clearedCells.add(cellRef);
        }
      }

      debugPrint(
          '[ExcelMatchService] 清除了 $clearedCount 个单元格的背景色引用（仅Rate和Amount列）');
      if (clearedCells.isNotEmpty) {
        debugPrint(
            '[ExcelMatchService] 清除的单元格: ${clearedCells.take(10).join(", ")}${clearedCells.length > 10 ? "..." : ""}');
      }

      // 确保计算模式为自动（在 workbook.xml 中设置）
      final workbookFile = archive.findFile('xl/workbook.xml');
      if (workbookFile != null) {
        final workbookXml = utf8.decode(workbookFile.content as List<int>);
        final workbookDoc = XmlDocument.parse(workbookXml);

        // 查找 calcPr 元素（计算属性）
        final calcPrElements = workbookDoc.findAllElements('calcPr');
        if (calcPrElements.isNotEmpty) {
          for (final calcPr in calcPrElements) {
            // 设置计算模式为自动
            calcPr.setAttribute('calcMode', 'auto');
            calcPr.setAttribute('fullCalcOnLoad', '1'); // 加载时完全计算
          }
        } else {
          // 如果没有 calcPr，创建一个
          final workbook = workbookDoc.findAllElements('workbook').firstOrNull;
          if (workbook != null) {
            final calcPr = XmlElement(XmlName('calcPr'), [
              XmlAttribute(XmlName('calcMode'), 'auto'),
              XmlAttribute(XmlName('fullCalcOnLoad'), '1'),
            ]);
            workbook.children.add(calcPr);
          }
        }

        // 更新 workbook.xml
        final updatedWorkbookXml = workbookDoc.toXmlString();
        final updatedWorkbookBytes = utf8.encode(updatedWorkbookXml);

        // 在重新打包时更新 workbook.xml
        final updatedArchive = Archive();
        for (final file in archive.files) {
          if (file.name == 'xl/workbook.xml') {
            updatedArchive.addFile(
              ArchiveFile(
                  file.name, updatedWorkbookBytes.length, updatedWorkbookBytes),
            );
          } else if (file.name == targetSheetFile.name) {
            // 工作表文件会在下面更新
            continue;
          } else {
            updatedArchive.addFile(file);
          }
        }

        // 更新工作表文件
        final updatedXml = sheetDoc.toXmlString();
        final updatedBytes = utf8.encode(updatedXml);
        updatedArchive.addFile(
          ArchiveFile(targetSheetFile.name, updatedBytes.length, updatedBytes),
        );

        final encoder = ZipEncoder();
        final outBytes = encoder.encode(updatedArchive) ?? <int>[];
        await excelFile.writeAsBytes(outBytes, flush: true);
        return;
      }

      // 重新打包
      final updatedXml = sheetDoc.toXmlString();
      final updatedBytes = utf8.encode(updatedXml);

      final updatedArchive = Archive();
      for (final file in archive.files) {
        if (file.name == targetSheetFile.name) {
          updatedArchive.addFile(
            ArchiveFile(file.name, updatedBytes.length, updatedBytes),
          );
        } else {
          updatedArchive.addFile(file);
        }
      }

      final encoder = ZipEncoder();
      final outBytes = encoder.encode(updatedArchive) ?? <int>[];
      await excelFile.writeAsBytes(outBytes, flush: true);
    } catch (e, st) {
      debugPrint('[ExcelMatchService] 清除背景色失败: $e');
      debugPrint('[ExcelMatchService] 堆栈: $st');
      // 不清除背景色不影响主要功能，继续执行
    }
  }

  /// 辅助方法：判断列名是否匹配
  bool _matchesColumnName(String headerText, List<String> patterns) {
    final normalized = _normalizeText(headerText);
    for (final pattern in patterns) {
      if (normalized.contains(_normalizeText(pattern))) {
        return true;
      }
    }
    return false;
  }

  /// 辅助方法：获取最佳相似度
  double _getBestSimilarity(String text, List<String> patterns) {
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

  /// 文本预处理（去除空格、转小写等）
  String _normalizeText(String text) {
    return text
        .toLowerCase()
        .replaceAll(RegExp(r'\s+'), '') // 去除所有空格
        .replaceAll(RegExp(r'[()（）]'), ''); // 去除括号
  }

  /// 列索引转字母（0 -> A, 1 -> B, ..., 26 -> AA, ...）
  String _colToLetter(int colIndex) {
    if (colIndex < 0) return 'A';
    var n = colIndex + 1; // 列从 1 开始
    final buffer = StringBuffer();
    while (n > 0) {
      n -= 1;
      buffer.writeCharCode('A'.codeUnitAt(0) + (n % 26));
      n ~/= 26;
    }
    final str = buffer.toString();
    return String.fromCharCodes(str.codeUnits.reversed);
  }

  /// 根据行列索引获取单元格（excel 包）
  dynamic _getCellByIndex(Sheet sheet, int rowIndex, int colIndex) {
    try {
      return sheet.cell(CellIndex.indexByColumnRow(
        columnIndex: colIndex,
        rowIndex: rowIndex,
      ));
    } catch (e) {
      return null;
    }
  }

  /// 获取单元格值（excel 包）
  String? _getCellValueFromExcel(dynamic cell) {
    if (cell == null) return null;

    try {
      final value = cell.value;
      if (value == null) return null;

      try {
        return switch (value) {
          TextCellValue() => _extractTextFromTextCellValue(value),
          IntCellValue() => value.value.toString(),
          DoubleCellValue() => value.value.toString(),
          BoolCellValue() => value.value.toString(),
          DateCellValue() => '${value.year}-${value.month}-${value.day}',
          DateTimeCellValue() => value.asDateTimeLocal().toString(),
          TimeCellValue() => value.asDuration().toString(),
          FormulaCellValue() => _extractTextFromFormulaCellValue(value),
          _ => null,
        };
      } catch (e) {
        debugPrint('[ExcelMatchService] switch 失败，尝试 toString: $e');
        try {
          final str = value.toString();
          return str.isEmpty ? null : str;
        } catch (e2) {
          debugPrint('[ExcelMatchService] toString 也失败: $e2');
          return null;
        }
      }
    } catch (e) {
      debugPrint('[ExcelMatchService] 获取单元格值失败: $e');
      return null;
    }
  }

  /// 从 TextCellValue 中提取文本
  String? _extractTextFromTextCellValue(TextCellValue textValue) {
    try {
      final value = textValue.value;
      if (value is String) {
        return value as String;
      }
      // 如果 value 不是 String，尝试提取文本
      try {
        final dynamicValue = value as dynamic;
        final text = dynamicValue.text;
        if (text != null && text is String && text.isNotEmpty) {
          return text;
        }
        final children = dynamicValue.children;
        if (children != null) {
          final buffer = StringBuffer();
          _extractTextFromTextSpanChildren(children, buffer);
          final result = buffer.toString();
          if (result.isNotEmpty) {
            return result;
          }
        }
      } catch (e) {
        debugPrint('[ExcelMatchService] 提取 TextSpan 文本失败: $e');
      }

      // 最后尝试 toString()
      try {
        final str = value.toString();
        if (str.contains('TextSpan') || str.contains('Instance of')) {
          return null;
        }
        return str.isNotEmpty ? str : null;
      } catch (_) {
        return null;
      }
    } catch (e) {
      debugPrint('[ExcelMatchService] 提取 TextCellValue 文本失败: $e');
      return null;
    }
  }

  /// 递归提取 TextSpan children 中的文本
  void _extractTextFromTextSpanChildren(dynamic children, StringBuffer buffer) {
    try {
      if (children is List) {
        for (final child in children) {
          if (child is String) {
            buffer.write(child);
          } else {
            final dynamicChild = child as dynamic;
            final text = dynamicChild.text;
            if (text != null && text is String) {
              buffer.write(text);
            }
            final childChildren = dynamicChild.children;
            if (childChildren != null) {
              _extractTextFromTextSpanChildren(childChildren, buffer);
            }
          }
        }
      }
    } catch (e) {
      debugPrint('[ExcelMatchService] 提取 TextSpan children 失败: $e');
    }
  }

  /// 从 FormulaCellValue 中提取文本
  String? _extractTextFromFormulaCellValue(FormulaCellValue formulaValue) {
    try {
      final formula = formulaValue.formula;
      return formula.toString();
    } catch (e) {
      debugPrint('[ExcelMatchService] 提取 FormulaCellValue 文本失败: $e');
      return null;
    }
  }
}
