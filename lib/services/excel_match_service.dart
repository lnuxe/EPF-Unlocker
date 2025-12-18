import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:excel/excel.dart';
import 'package:xml/xml.dart';
import '../models/excel_models.dart';
import 'excel_parser_service.dart';
import 'column_identifier_service.dart';
import 'excel_writer_service.dart';
import 'match_report_service.dart';

/// Excel 匹配服务：协调整个匹配流程
/// 职责：协调解析、识别、匹配、写入等各个服务
class ExcelMatchService {
  final _parserService = ExcelParserService();
  final _columnIdentifier = ColumnIdentifierService();
  final _writerService = ExcelWriterService();
  final _reportService = MatchReportService();

  /// 主入口：匹配 Excel 文件
  /// 不再需要 targetColumns 参数，自动识别列结构
  /// 简化错误处理：错误信息写入报告，只执行成功匹配的逻辑
  Future<MatchServiceResult> matchExcelFiles({
    required File draftFile,
    required File targetFile,
    required String outputPath,
  }) async {
    final logs = <String>[];
    int matchedCount = 0;
    int totalCount = 0;

    logs.add('开始匹配流程...');
    logs.add('草稿文件: ${draftFile.path}');
    logs.add('目标文件: ${targetFile.path}');

    // Step 1: 使用 XML 方式读取目标文件（避免 numFmtId 错误）
    logs.add('正在加载目标文件（使用 XML 方式）...');
    Map<String, dynamic>? targetResult;
    final targetParseError = await _safeParse(() async {
      targetResult = await _parserService.parseTargetFileWithXml(targetFile);
    });
    if (targetParseError != null) {
      logs.add('❌ 错误: 无法解析目标文件: $targetParseError');
      logs.add('提示：文件格式可能有问题，或者文件已损坏');
      return MatchServiceResult(
        success: false,
        message: '无法解析目标文件: $targetParseError',
        matchedCount: 0,
        totalCount: 0,
        logs: logs,
      );
    }

    if (targetResult == null) {
      logs.add('❌ 错误: 目标文件解析结果为空');
      return MatchServiceResult(
        success: false,
        message: '目标文件解析结果为空',
        matchedCount: 0,
        totalCount: 0,
        logs: logs,
      );
    }

    final targetSheetName = targetResult!['sheetName'] as String;
    final targetColumns = targetResult!['columns'] as TargetColumns;
    final targetRowsXml = targetResult!['rows'] as List<XmlElement>;
    final sharedStrings = targetResult!['sharedStrings'] as List<String>;
    final headerRowIndex = targetResult!['headerRowIndex'] as int;

    logs.add('目标文件已加载（XML 方式），工作表: $targetSheetName');
    logs.add(
        '目标列结构已识别: Item=${_columnIdentifier.colToLetter(targetColumns.itemCol)}, '
        'Description=${_columnIdentifier.colToLetter(targetColumns.descriptionCol)}, '
        'Unit=${_columnIdentifier.colToLetter(targetColumns.unitCol)}, '
        'Qty=${_columnIdentifier.colToLetter(targetColumns.qtyCol)}, '
        'Unit Rate=${_columnIdentifier.colToLetter(targetColumns.unitRateCol)}, '
        'Amount=${_columnIdentifier.colToLetter(targetColumns.amountCol)}');

    // Step 2: 扫描目标文件的所有子项行（使用 XML 方式）
    logs.add('正在扫描目标文件子项行（XML 方式）...');
    final targetRows = _scanTargetRowsFromXml(
      targetRowsXml,
      targetColumns,
      sharedStrings,
      headerRowIndex,
    );
    totalCount = targetRows.length;
    logs.add('找到 $totalCount 个行需要匹配');

    if (targetRows.isEmpty) {
      return MatchServiceResult(
        success: false,
        message: '目标文件中没有找到需要匹配的行（Rate 或 Amount 列为空）',
        matchedCount: 0,
        totalCount: 0,
        logs: logs,
      );
    }

    // Step 3: 构建草稿文件的匹配 Map（自动识别列结构，优先匹配同名工作表）
    logs.add('正在构建草稿文件数据...');
    logs.add('目标文件工作表: $targetSheetName');
    Map<String, dynamic>? draftResult;
    final draftParseError = await _safeParse(() async {
      draftResult = await _parserService.parseDraftFileWithXml(draftFile,
          targetSheetName: targetSheetName);
    });
    if (draftParseError != null) {
      logs.add('❌ 错误: 无法解析草稿文件: $draftParseError');
      logs.add('提示：文件格式可能有问题，或者文件已损坏');
      return MatchServiceResult(
        success: false,
        message: '无法解析草稿文件: $draftParseError',
        matchedCount: 0,
        totalCount: totalCount,
        logs: logs,
      );
    }

    if (draftResult == null) {
      logs.add('❌ 错误: 草稿文件解析结果为空');
      return MatchServiceResult(
        success: false,
        message: '草稿文件解析结果为空',
        matchedCount: 0,
        totalCount: totalCount,
        logs: logs,
      );
    }

    final draftMap = draftResult!['map'] as Map<String, DraftRow>?;
    final matchedSheetName = draftResult!['matchedSheet'] as String?;
    final parseError = draftResult!['error'] as String?;
    final draftSheetName = matchedSheetName; // 保存草稿工作表名称用于PDF报告

    // 如果有解析错误，写入日志
    if (parseError != null) {
      logs.add('❌ 草稿文件解析错误: $parseError');
      logs.add(
          '错误类型：${parseError.contains('numFmtId') ? 'numFmtId 格式错误' : '未知解析错误'}');
    }

    if (draftMap == null || draftMap.isEmpty) {
      final sheetInfo = matchedSheetName != null
          ? '未在草稿文件中找到工作表 "$matchedSheetName"'
          : '草稿文件中没有有效数据';
      logs.add('❌ 错误: $sheetInfo');
      logs.add('提示：请检查文件格式是否正确');
      return MatchServiceResult(
        success: false,
        message: '$sheetInfo，请检查文件格式是否正确',
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

    // 添加调试信息：显示草稿文件中的 Item 列表（前30个）
    final draftItems = draftMap.values.map((d) => d.item).toSet().toList()
      ..sort();
    logs.add(
        '草稿文件中的 Item 列表（前30个）: ${draftItems.take(30).join(", ")}${draftItems.length > 30 ? "..." : ""}');
    debugPrint(
        '[ExcelMatchService] 草稿文件包含 ${draftMap.length} 条数据，Item 范围: ${draftItems.isNotEmpty ? "${draftItems.first} - ${draftItems.last}" : "无"}');

    // 检查是否有子项（如 6.1, 6.2, 8.1 等）
    final subItems = draftItems.where((item) => item.contains('.')).toList();
    if (subItems.isNotEmpty) {
      debugPrint(
          '[ExcelMatchService] 草稿文件中的子项（包含小数点）: ${subItems.take(20).join(", ")}${subItems.length > 20 ? "..." : ""}');
    }

    // Step 4: 构建按Item和Description分别索引的映射（用于强匹配）
    logs.add('正在构建草稿文件索引（用于强匹配）...');
    final draftByItem = <String, List<DraftRow>>{}; // Item -> List<DraftRow>
    final draftByDescription =
        <String, List<DraftRow>>{}; // Description -> List<DraftRow>
    for (final draft in draftMap.values) {
      // Item索引（支持模糊匹配）
      final itemKey = _normalizeItemForMatching(draft.item);
      if (!draftByItem.containsKey(itemKey)) {
        draftByItem[itemKey] = [];
      }
      draftByItem[itemKey]!.add(draft);

      // Description索引
      final descKey = _normalizeText(draft.description);
      if (descKey.isNotEmpty) {
        if (!draftByDescription.containsKey(descKey)) {
          draftByDescription[descKey] = [];
        }
        draftByDescription[descKey]!.add(draft);
      }
    }
    logs.add(
        '草稿文件索引构建完成: ${draftByItem.length} 个Item, ${draftByDescription.length} 个Description');

    // Step 5: 执行匹配（强匹配算法：Item或Description任一匹配即可）
    logs.add('开始匹配（强匹配算法：Item或Description任一匹配即可）...');
    var matchResults = <MatchResult>[];
    final failedMatches = <String>[]; // 记录匹配失败的项目

    for (final targetRow in targetRows) {
      final matchResult = _matchTargetRowStrong(
        target: targetRow,
        draftMap: draftMap,
        draftByItem: draftByItem,
        draftByDescription: draftByDescription,
      );
      matchResults.add(matchResult);

      if (matchResult.matched) {
        matchedCount++;
        final matchTypeText = matchResult.matchType == 'exact'
            ? '完全匹配'
            : matchResult.matchType == 'item'
                ? 'Item匹配'
                : 'Description匹配';
        final qtyInfo = matchResult.qty != null
            ? matchResult.qty!.toStringAsFixed(1)
            : 'N/A';
        final rateInfo = matchResult.rate != null
            ? matchResult.rate!.toStringAsFixed(2)
            : 'N/A';
        final amountInfo = matchResult.amount != null
            ? matchResult.amount!.toStringAsFixed(2)
            : 'N/A';
        logs.add(
          '✓ $matchTypeText: Item=${targetRow.item} | Desc=${targetRow.description.substring(0, targetRow.description.length > 30 ? 30 : targetRow.description.length)}${targetRow.description.length > 30 ? "..." : ""} | '
          'Qty=$qtyInfo, Rate=$rateInfo, Amount=$amountInfo',
        );
      } else {
        final failInfo =
            'Item=${targetRow.item}, Description=${targetRow.description}';
        failedMatches.add(failInfo);
        logs.add('✗ 匹配失败: $failInfo');
      }
    }

    // 输出匹配失败统计
    if (failedMatches.isNotEmpty) {
      logs.add('');
      logs.add('⚠ 匹配失败统计: 共 ${failedMatches.length} 个项目未找到匹配');
      if (failedMatches.isNotEmpty) {
        logs.add('匹配失败的项目列表（前20个）:');
        for (final failInfo in failedMatches.take(20)) {
          logs.add('  - $failInfo');
        }
        if (failedMatches.length > 20) {
          logs.add('  ... 还有 ${failedMatches.length - 20} 个项目未匹配');
        }
      }
      logs.add('');
      logs.add('提示: 强匹配策略（Item或Description任一匹配即可），支持Item模糊匹配（如1.9和1.90）');
    }

    // Step 6: 直接使用原始目标文件，通过 XML 方式更新值和清除背景色（完全保留格式）
    // 不使用 targetExcel.save()，避免破坏原始格式
    logs.add('准备使用 XML 方式写入匹配值并清除背景色...');

    // 确保输出目录存在
    final outputFile = File(outputPath);
    final outputDir = outputFile.parent;
    if (!await outputDir.exists()) {
      await outputDir.create(recursive: true);
      logs.add('创建输出目录: ${outputDir.path}');
    }

    // 复制原始目标文件到输出路径（保留原始格式）
    await targetFile.copy(outputFile.path);
    logs.add('已复制原始文件到输出路径: $outputPath');

    // Step 7: 使用 XML 方式写入匹配值并清除背景色（只写入成功匹配的数据）
    logs.add('正在使用 XML 方式写入匹配值并清除背景色（只写入成功匹配的数据）...');
    List<MatchResult>? updatedMatchResults;
    final writeError = await _safeParse(() async {
      updatedMatchResults =
          await _writerService.writeMatchedValuesAndClearBackgrounds(
        outputFile,
        targetSheetName,
        matchResults,
        targetColumns,
      );
    });
    if (writeError != null) {
      logs.add('❌ 错误: 写入匹配值失败: $writeError');
      updatedMatchResults = matchResults; // 使用原始结果
    } else {
      logs.add('✓ 匹配值已写入，背景色已清除，格式已保留（仅成功匹配的数据）');
      // 使用更新后的结果（包含实际写入的值）
      if (updatedMatchResults != null) {
        matchResults = updatedMatchResults!;

        // 输出写入值统计
        int writtenCount = 0;
        int formulaCount = 0;
        for (final result in matchResults) {
          if (result.matched) {
            if (result.writtenRate != null || result.writtenAmount != null) {
              writtenCount++;
            }
            if (result.amountFormula != null) {
              formulaCount++;
            }
          }
        }
        logs.add('写入统计: 共写入 $writtenCount 行数据，其中 $formulaCount 行使用公式计算Amount');
      }
    }

    // Step 8: 生成PDF比对报告（包含所有结果，包括失败的）
    final pdfOutputPath = outputPath.replaceAll('.xlsx', '_匹配报告.pdf');
    logs.add('正在生成PDF比对报告（包含所有匹配结果）...');
    final reportError = await _safeParse(() async {
      await _reportService.generateMatchReport(
        outputPath: pdfOutputPath,
        matchResults: matchResults,
        logs: logs,
        targetSheetName: targetSheetName,
        draftSheetName: draftSheetName,
      );
    });
    if (reportError != null) {
      logs.add('❌ 错误: PDF报告生成失败: $reportError');
    } else {
      logs.add('✓ PDF比对报告已生成: ${pdfOutputPath.split('/').last}');
    }

    // 返回结果（即使有部分错误，只要匹配流程完成就返回成功）
    final hasErrors = writeError != null || reportError != null;
    return MatchServiceResult(
      success: !hasErrors,
      message: hasErrors
          ? '匹配完成（有错误）：成功 $matchedCount/$totalCount，输出文件：${outputFile.uri.pathSegments.last}，详情请查看PDF报告'
          : '匹配完成：成功 $matchedCount/$totalCount，输出文件：${outputFile.uri.pathSegments.last}',
      matchedCount: matchedCount,
      totalCount: totalCount,
      logs: logs,
    );
  }

  /// 安全执行异步操作，捕获错误并返回错误信息（不抛出异常）
  Future<String?> _safeParse(Future<void> Function() action) async {
    try {
      await action();
      return null;
    } on FileSystemException catch (e, st) {
      String errorMessage;
      if (e.osError?.errorCode == 1) {
        // Operation not permitted
        errorMessage = '无法访问文件（权限不足）\n'
            '请检查：\n'
            '1. 文件是否被其他程序打开（如 Excel）\n'
            '2. 是否授予了应用文件访问权限\n'
            '3. 文件是否在受保护的目录中';
      } else if (e.osError?.errorCode == 13) {
        // Permission denied
        errorMessage = '文件访问被拒绝\n'
            '请检查文件权限设置';
      } else {
        errorMessage = '文件系统错误: ${e.message}';
      }
      debugPrint('[ExcelMatchService] 操作失败: $errorMessage');
      debugPrint('[ExcelMatchService] 堆栈: $st');
      return errorMessage;
    } catch (e, st) {
      debugPrint('[ExcelMatchService] 操作失败: $e');
      debugPrint('[ExcelMatchService] 堆栈: $st');
      // 检查是否是 PathAccessException
      final errorString = e.toString();
      if (errorString.contains('PathAccessException') ||
          errorString.contains('Operation not permitted')) {
        return '无法访问文件（权限不足）\n'
            '请检查：\n'
            '1. 文件是否被其他程序打开（如 Excel）\n'
            '2. 是否授予了应用文件访问权限\n'
            '3. 文件是否在受保护的目录中';
      }
      return e.toString();
    }
  }

  /// 扫描目标文件的所有子项行（已废弃，改用XML方式）
  /// 保留此方法用于向后兼容，但实际已不再使用
  @Deprecated('使用 _scanTargetRowsFromXml 代替')
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
      if (item.isEmpty) {
        // 如果 Item 为空，但可能还有数据行，继续检查
        // 添加调试日志
        if (rowIndex < 100) {
          debugPrint('[ExcelMatchService] 跳过 Item 为空的第 ${rowIndex + 1} 行');
        }
        continue;
      }

      // 处理所有行（包括主项和子项），只要 Rate 或 Amount 为空就需要匹配
      // 不再跳过主项行，因为主项行也可能需要填充数据
      final itemTrimmed = item.trim();

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
      // 注意：即使 Rate 有值但 Amount 为空，或者反过来，也应该匹配并填充
      final needsRate = rateValue == null || rateValue.trim().isEmpty;
      final needsAmount = amountValue == null || amountValue.trim().isEmpty;

      // 只要有一个列为空，就需要匹配填充
      // 这样可以确保所有需要填充的行都被扫描到
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
        debugPrint(
            '[ExcelMatchService] 扫描到需要匹配的行: Item=$item, Description=$description, '
            'needsRate=$needsRate, needsAmount=$needsAmount (行号: ${rowIndex + 1})');
      } else {
        // 添加调试日志，显示为什么某些行被跳过
        // 在这个分支中，needsRate 和 needsAmount 都是 false，说明 rateValue 和 amountValue 都有值
        if (rowIndex < 100) {
          debugPrint(
              '[ExcelMatchService] 跳过已填充的行: Item=$item, Description=$description (行号: ${rowIndex + 1}, Rate=$rateValue, Amount=$amountValue)');
        }
      }
    }

    return targetRows;
  }

  /// 根据上一行的Item值推断当前行的Item值
  /// 规则：
  /// - 上一行Item=5 → 当前行Item=5.1
  /// - 上一行Item=5.1 → 当前行Item=5.2
  /// - 上一行Item=5.2.1 → 当前行Item=5.2.2
  String _inferItemFromPrevious(String previousItem) {
    final trimmed = previousItem.trim();
    if (trimmed.isEmpty) return '';

    // 按小数点分割Item值
    final parts = trimmed.split('.');
    if (parts.isEmpty) return '';

    // 解析最后一部分为数字
    final lastPart = parts.last;
    final lastNumber = int.tryParse(lastPart);

    if (lastNumber != null) {
      // 如果最后一部分是数字，递增它
      final newLastPart = (lastNumber + 1).toString();
      if (parts.length == 1) {
        // 如果只有一部分（如 "5"），添加 ".1"
        return '$trimmed.1';
      } else {
        // 如果有多部分（如 "5.1" 或 "5.2.1"），替换最后一部分
        parts[parts.length - 1] = newLastPart;
        return parts.join('.');
      }
    } else {
      // 如果最后一部分不是数字，直接添加 ".1"
      return '$trimmed.1';
    }
  }

  /// 从 XML 扫描目标文件的所有行（避免 numFmtId 错误）
  /// 如果Item列为空，根据上一行的Item值推断当前行的Item值
  List<TargetRow> _scanTargetRowsFromXml(
    List<XmlElement> rows,
    TargetColumns columns,
    List<String> sharedStrings,
    int headerRowIndex,
  ) {
    final targetRows = <TargetRow>[];
    String? lastItem; // 记录上一行的Item值，用于推断空白行的Item

    // 从表头之后开始扫描（跳过表头行）
    for (int i = headerRowIndex + 1; i < rows.length; i++) {
      final row = rows[i];
      final rowNumAttr = row.getAttribute('r');
      final rowNum =
          rowNumAttr != null ? int.tryParse(rowNumAttr) ?? (i + 1) : (i + 1);
      final rowIndex = rowNum - 1; // 转换为 0-based

      // 提取所有单元格
      final cells = row.findElements('c').toList();
      if (cells.isEmpty) continue;

      // 提取 Item
      String item =
          _getCellValueFromXml(cells, columns.itemCol, sharedStrings) ?? '';
      bool isInferredItem = false; // 标记Item是否为推断的

      // 如果Item为空，根据上一行的Item值推断
      if (item.isEmpty || item.trim().isEmpty) {
        if (lastItem != null && lastItem.isNotEmpty) {
          item = _inferItemFromPrevious(lastItem);
          isInferredItem = true;
          debugPrint(
              '[ExcelMatchService] Item为空，根据上一行推断: 上一行Item=$lastItem → 当前行Item=$item (行号: ${rowIndex + 1})');
        } else {
          // 如果上一行也没有Item，跳过该行
          if (rowIndex < 100) {
            debugPrint(
                '[ExcelMatchService] 跳过 Item 为空的第 ${rowIndex + 1} 行（上一行也没有Item）');
          }
          continue;
        }
      } else {
        // 如果Item不为空，更新lastItem
        lastItem = item.trim();
      }

      final itemTrimmed = item.trim();

      // 提取 Description
      final description =
          _getCellValueFromXml(cells, columns.descriptionCol, sharedStrings) ??
              '';

      // 跳过包含 "Remark" 的行
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
        unit =
            _getCellValueFromXml(cells, columns.unitCol, sharedStrings) ?? '';
      }

      // 提取 Qty（可选）
      double qty = 0.0;
      if (columns.qtyCol >= 0) {
        final qtyStr =
            _getCellValueFromXml(cells, columns.qtyCol, sharedStrings) ?? '';
        qty = double.tryParse(qtyStr) ?? 0.0;
      }

      // 检查 Unit Rate 和 Amount 列是否为空（需要填充）
      final rateValue =
          _getCellValueFromXml(cells, columns.unitRateCol, sharedStrings);
      final amountValue =
          _getCellValueFromXml(cells, columns.amountCol, sharedStrings);

      final needsRate = rateValue == null || rateValue.trim().isEmpty;
      final needsAmount = amountValue == null || amountValue.trim().isEmpty;

      if (needsRate || needsAmount) {
        targetRows.add(
          TargetRow(
            item: item, // 使用推断后的Item值
            description: description,
            unit: unit,
            qty: qty,
            rowIndex: rowIndex,
            rateColumn: columns.unitRateCol,
            amountColumn: columns.amountCol,
          ),
        );
        final itemInfo =
            isInferredItem ? 'Item=$item (推断自上一行: $lastItem)' : 'Item=$item';
        debugPrint(
            '[ExcelMatchService] 扫描到需要匹配的行: $itemInfo, Description=$description, '
            'needsRate=$needsRate, needsAmount=$needsAmount (行号: ${rowIndex + 1})');
      } else {
        if (rowIndex < 100) {
          debugPrint(
              '[ExcelMatchService] 跳过已填充的行: Item=$item, Description=$description (行号: ${rowIndex + 1}, Rate=$rateValue, Amount=$amountValue)');
        }
      }
    }

    return targetRows;
  }

  /// 从 XML 单元格列表中获取指定列的值
  String? _getCellValueFromXml(
    List<XmlElement> cells,
    int colIndex,
    List<String> sharedStrings,
  ) {
    for (var cell in cells) {
      final cellRef = cell.getAttribute('r');
      if (cellRef == null) continue;

      final cellCol = _cellRefToColIndex(cellRef);
      if (cellCol == colIndex) {
        return _getCellValueFromXmlElement(cell, sharedStrings);
      }
    }
    return null;
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
        final inlineText = _extractInlineText(inlineStrElement);
        if (inlineText != null) {
          return inlineText;
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

  /// 解析 inlineStr 节点（支持富文本）
  String? _extractInlineText(XmlElement inlineStrElement) {
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

  /// 强匹配算法：Item或Description任一匹配即可
  /// 支持Item模糊匹配（如1.9和1.90视为同一项）
  MatchResult _matchTargetRowStrong({
    required TargetRow target,
    required Map<String, DraftRow> draftMap,
    required Map<String, List<DraftRow>> draftByItem,
    required Map<String, List<DraftRow>> draftByDescription,
  }) {
    // 检测是否为Total行（检查item或description是否包含"total"）
    final isTotalRow = target.item.toLowerCase().contains('total') ||
        target.description.toLowerCase().contains('total');

    // 首先尝试完全匹配（Item和Description都匹配）
    final normalizedKey =
        _normalizeText('${target.item}|${target.description}');
    if (draftMap.containsKey(normalizedKey)) {
      final d = draftMap[normalizedKey]!;
      debugPrint(
          '[ExcelMatchService] ✓ 完全匹配: Item=${target.item}, Description=${target.description}');
      return MatchResult(
        target: target,
        draft: d,
        matched: true,
        qty: d.qty,
        rate: d.rate,
        amount: d.amount,
        matchType: 'exact',
        isTotalRow: isTotalRow,
        draftTotalAmount: isTotalRow ? d.amount : null,
      );
    }

    // 尝试Item匹配（支持模糊匹配，如1.9和1.90）
    final targetItemKey = _normalizeItemForMatching(target.item);
    if (draftByItem.containsKey(targetItemKey)) {
      final candidates = draftByItem[targetItemKey]!;
      if (candidates.isNotEmpty) {
        // 优先选择Description也匹配的，否则选择第一个
        DraftRow? bestMatch;
        for (final candidate in candidates) {
          final candidateDescNorm = _normalizeText(candidate.description);
          final targetDescNorm = _normalizeText(target.description);
          if (candidateDescNorm == targetDescNorm) {
            bestMatch = candidate;
            break;
          }
        }
        bestMatch ??= candidates.first;

        debugPrint(
            '[ExcelMatchService] ✓ Item匹配: Item=${target.item} (模糊匹配: $targetItemKey), Description=${target.description}');
        return MatchResult(
          target: target,
          draft: bestMatch,
          matched: true,
          qty: bestMatch.qty,
          rate: bestMatch.rate,
          amount: bestMatch.amount,
          matchType: 'item',
          isTotalRow: isTotalRow,
          draftTotalAmount: isTotalRow ? bestMatch.amount : null,
        );
      }
    }

    // 尝试Description匹配
    final targetDescKey = _normalizeText(target.description);
    if (targetDescKey.isNotEmpty &&
        draftByDescription.containsKey(targetDescKey)) {
      final candidates = draftByDescription[targetDescKey]!;
      if (candidates.isNotEmpty) {
        // 优先选择Item也匹配的，否则选择第一个
        DraftRow? bestMatch;
        final targetItemKey = _normalizeItemForMatching(target.item);
        for (final candidate in candidates) {
          final candidateItemKey = _normalizeItemForMatching(candidate.item);
          if (candidateItemKey == targetItemKey) {
            bestMatch = candidate;
            break;
          }
        }
        bestMatch ??= candidates.first;

        debugPrint(
            '[ExcelMatchService] ✓ Description匹配: Item=${target.item}, Description=${target.description}');
        return MatchResult(
          target: target,
          draft: bestMatch,
          matched: true,
          qty: bestMatch.qty,
          rate: bestMatch.rate,
          amount: bestMatch.amount,
          matchType: 'description',
          isTotalRow: isTotalRow,
          draftTotalAmount: isTotalRow ? bestMatch.amount : null,
        );
      }
    }

    // 未找到匹配
    debugPrint(
        '[ExcelMatchService] ✗ 未找到匹配: Item=${target.item}, Description=${target.description}');
    return MatchResult(
      target: target,
      matched: false,
      isTotalRow: isTotalRow,
    );
  }

  /// Item模糊匹配标准化：将1.9和1.90视为同一项
  /// 规则：去除末尾的0，如1.90 -> 1.9, 1.9 -> 1.9
  String _normalizeItemForMatching(String item) {
    if (item.isEmpty) return '';

    // 先进行基本标准化
    final normalized = _normalizeText(item);
    if (normalized.isEmpty) return '';

    // 处理数字格式的Item（如1.9, 1.90, 1.900等）
    // 尝试解析为数字，如果成功则标准化
    final parts = normalized.split('.');
    if (parts.length >= 2) {
      // 尝试解析为数字
      final firstPart = int.tryParse(parts[0]);
      if (firstPart != null) {
        // 尝试将整个Item解析为数字
        final fullNumber = double.tryParse(normalized);
        if (fullNumber != null) {
          // 转换为字符串，去除末尾的0和小数点
          final normalizedNumber = fullNumber.toString();
          return normalizedNumber.replaceAll(RegExp(r'\.?0+$'), '');
        }
      }
    }

    return normalized;
  }

  /// 文本预处理（去除空格、转小写等）
  /// 确保与 excel_parser_service.dart 中的标准化逻辑完全一致
  String _normalizeText(String text) {
    if (text.isEmpty) return '';
    return text
        .toLowerCase()
        .replaceAll(RegExp(r'\s+'), '') // 去除所有空格（包括换行、制表符等）
        .replaceAll(RegExp(r'[()（）]'), ''); // 去除括号
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
