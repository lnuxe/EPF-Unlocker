import 'dart:io';
import 'dart:convert';
import 'package:flutter/foundation.dart';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';
import '../models/excel_models.dart';

/// Excel 写入服务：负责写入匹配值并处理格式
/// 职责：更新单元格值、保留格式、清除背景色、设置公式
class ExcelWriterService {
  /// 使用 XML 方式写入匹配值并清除背景色（完全保留格式）
  /// 只更新单元格的 <v> 节点值，保留所有其他格式
  /// 返回更新后的 MatchResult 列表，包含实际写入的值
  Future<List<MatchResult>> writeMatchedValuesAndClearBackgrounds(
    File excelFile,
    String sheetName,
    List<MatchResult> matchResults,
    TargetColumns targetColumns,
  ) async {
    try {
      final bytes = await excelFile.readAsBytes();
      final archive = ZipDecoder().decodeBytes(bytes);

      // 读取 workbook.xml 获取工作表列表，根据 sheetName 精确查找
      final workbookFile = archive.findFile('xl/workbook.xml');
      if (workbookFile == null) {
        debugPrint('[ExcelWriterService] 未找到 workbook.xml，跳过写入');
        return matchResults; // 返回原始结果
      }

      final workbookXml = utf8.decode(workbookFile.content as List<int>);
      final workbookDoc = XmlDocument.parse(workbookXml);

      // 获取所有工作表信息
      final sheetElements = workbookDoc.findAllElements('sheet').toList();
      if (sheetElements.isEmpty) {
        debugPrint('[ExcelWriterService] 工作簿中没有工作表，跳过写入');
        return matchResults; // 返回原始结果
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

      // 根据 sheetName 查找对应的工作表文件
      String? targetSheetId;
      if (sheetName.isNotEmpty && sheetMap.containsKey(sheetName)) {
        targetSheetId = sheetMap[sheetName];
        debugPrint(
            '[ExcelWriterService] 找到目标工作表: $sheetName (sheetId: $targetSheetId)');
      } else {
        // 如果没找到，尝试模糊匹配
        final normalizedTarget = _normalizeSheetName(sheetName);
        for (var entry in sheetMap.entries) {
          if (_normalizeSheetName(entry.key) == normalizedTarget) {
            targetSheetId = entry.value;
            debugPrint(
                '[ExcelWriterService] 找到相似工作表: ${entry.key} (sheetId: $targetSheetId)');
            break;
          }
        }
      }

      // 如果还是没找到，使用第一个工作表
      if (targetSheetId == null && sheetMap.isNotEmpty) {
        targetSheetId = sheetMap.values.first;
        debugPrint(
            '[ExcelWriterService] 未找到匹配工作表，使用第一个工作表 (sheetId: $targetSheetId)');
      }

      // 查找对应的工作表文件
      // 注意：sheetId 和实际的 sheetX.xml 文件名可能不是一一对应的
      // 需要根据 sheetMap 中工作表的位置来查找
      ArchiveFile? targetSheetFile;

      // 方法1：尝试使用 sheetId 直接查找
      final expectedSheetFileName = 'xl/worksheets/sheet$targetSheetId.xml';
      for (final file in archive.files) {
        if (file.name == expectedSheetFileName) {
          targetSheetFile = file;
          debugPrint(
              '[ExcelWriterService] 通过 sheetId 找到工作表文件: ${targetSheetFile.name}');
          break;
        }
      }

      // 方法2：如果精确匹配失败，根据工作表在 workbook.xml 中的出现顺序查找
      // 注意：sheetX.xml 的文件名是按照工作表在 workbook.xml 中出现的顺序命名的
      // 第一个 sheet 元素对应 sheet1.xml，第二个对应 sheet2.xml，以此类推
      if (targetSheetFile == null && targetSheetId != null) {
        // 获取所有工作表文件并按名称排序（sheet1.xml, sheet2.xml, ...）
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

        // 按照 sheet 元素在 workbook.xml 中出现的顺序查找目标工作表
        // 找到 sheetId 匹配的工作表在列表中的索引（从0开始）
        int targetIndex = -1;
        for (int i = 0; i < sheetElements.length; i++) {
          final sheet = sheetElements[i];
          final sheetId = sheet.getAttribute('sheetId');
          if (sheetId == targetSheetId) {
            targetIndex = i;
            break;
          }
        }

        // 如果找到了索引，使用对应位置的工作表文件（索引0对应sheet1.xml）
        if (targetIndex >= 0 && targetIndex < sheetFiles.length) {
          targetSheetFile = sheetFiles[targetIndex];
          final foundSheetName =
              sheetElements[targetIndex].getAttribute('name') ?? '未知';
          debugPrint(
              '[ExcelWriterService] 通过工作表顺序找到文件: ${targetSheetFile.name} (工作表索引: $targetIndex, sheetName: $foundSheetName, sheetId: $targetSheetId)');
        }
      }

      if (targetSheetFile == null) {
        debugPrint('[ExcelWriterService] 未找到工作表文件，跳过写入');
        return matchResults; // 返回原始结果
      }

      // 解析工作表 XML
      final sheetXml = utf8.decode(targetSheetFile.content as List<int>);
      final sheetDoc = XmlDocument.parse(sheetXml);
      final sheetData = sheetDoc.findAllElements('sheetData').firstOrNull;
      if (sheetData == null) return matchResults; // 返回原始结果

      // 获取或创建 cols 元素（用于设置列宽）
      var colsElement = sheetDoc.findAllElements('cols').firstOrNull;
      if (colsElement == null) {
        // 如果不存在，在 sheetData 之前创建
        final worksheet = sheetDoc.findAllElements('worksheet').firstOrNull;
        if (worksheet != null) {
          colsElement = XmlElement(XmlName('cols'), const []);
          final sheetDataIndex = worksheet.children.indexOf(sheetData);
          worksheet.children.insert(sheetDataIndex, colsElement);
        }
      }

      // 收集所有需要更新的单元格（值和背景色）
      // 同时记录每个MatchResult对应的实际写入值
      final cellsToUpdate = <String, Map<String, dynamic>>{};
      final resultWrittenValues =
          <MatchResult, Map<String, dynamic>>{}; // 记录每个结果的实际写入值
      int matchedResultsCount = 0;
      int cellsAddedCount = 0;

      // 跟踪Total行的位置，用于确定每个Total行的累加范围
      final totalRowPositions = <int>[]; // 存储Total行的行号（Excel行号，从1开始，已排序）
      final totalRowResults = <MatchResult>[]; // 存储Total行的MatchResult
      // 存储非Total行的Amount值（行号 -> Amount值），用于计算Total行的总和
      final nonTotalRowAmounts = <int, double>{}; // Excel行号 -> Amount值

      // 第一遍遍历：识别所有Total行并记录位置，同时收集非Total行的Amount值
      for (final result in matchResults) {
        if (result.isTotalRow && result.matched) {
          final rowNum = result.target.rowIndex + 1;
          totalRowPositions.add(rowNum);
          totalRowResults.add(result);
          debugPrint(
              '[ExcelWriterService] 检测到Total行: 行号=$rowNum, Item="${result.target.item}", '
              '草稿Total值=${result.draftTotalAmount}');
        } else if (result.matched && result.amount != null) {
          // 收集非Total行的Amount值（用于后续计算Total行的总和）
          final rowNum = result.target.rowIndex + 1;
          nonTotalRowAmounts[rowNum] = result.amount!;
        }
      }
      // 对Total行位置排序
      totalRowPositions.sort();

      for (final result in matchResults) {
        if (!result.matched) continue;
        matchedResultsCount++;
        final target = result.target;
        final draft = result.draft;
        if (draft == null) continue; // 如果没有草稿数据，跳过

        final rowNum = target.rowIndex + 1; // Excel 行号从 1 开始
        final writtenValues = <String, dynamic>{}; // 记录该结果的实际写入值

        // 优化空白行填充：检查目标行中哪些列是空的，从草稿文件中填充所有对应的值

        // Item 列（只在目标文件的Item为空时才写入，避免出现多余小数点）
        // 检查目标文件的Item是否为空（需要从XML中读取实际值来判断）
        // 空白行（Item被推断）是例外，需要写入草稿文件的Item值
        if (draft.item.isNotEmpty && targetColumns.itemCol >= 0) {
          // 从XML中读取目标文件的Item值，判断是否为空
          final itemColLetter = _colToLetter(targetColumns.itemCol);
          final cellRef = '$itemColLetter$rowNum';

          // 检查目标文件的Item是否为空
          // 如果target.item为空或只包含空白，说明是空白行，需要写入
          // 如果target.item不为空，说明目标文件已有Item值，不应该覆盖（避免出现多余小数点）
          final bool isTargetItemEmpty = target.item.trim().isEmpty;

          if (isTargetItemEmpty) {
            // 只在目标文件的Item为空时才写入（空白行例外）
            cellsToUpdate[cellRef] = {
              'value': draft.item,
              'col': targetColumns.itemCol,
              'type': 'item',
            };
            cellsAddedCount++;
            debugPrint('[ExcelWriterService] ✓ 添加Item到更新列表: 行号=$rowNum, '
                '单元格=$cellRef, 目标Item="${target.item}"(空), 草稿Item="${draft.item}"');
          } else {
            // 目标文件已有Item值，跳过写入（避免出现多余小数点）
            debugPrint('[ExcelWriterService] ⚠ 跳过Item写入: 行号=$rowNum, '
                '原因=目标文件已有Item值"${target.item}", 草稿Item="${draft.item}"');
          }
        }

        // Description 列（如果草稿文件有值，总是写入）
        // 对于空白行（Item被推断），Description应该总是写入草稿文件的值
        // 简化逻辑：如果草稿文件有Description值，就写入（无论目标文件是否有值）
        if (draft.description.isNotEmpty && targetColumns.descriptionCol >= 0) {
          final descColLetter = _colToLetter(targetColumns.descriptionCol);
          final cellRef = '$descColLetter$rowNum';
          cellsToUpdate[cellRef] = {
            'value': draft.description,
            'col': targetColumns.descriptionCol,
            'type': 'description',
          };
          cellsAddedCount++;
          debugPrint('[ExcelWriterService] ✓ 添加Description到更新列表: 行号=$rowNum, '
              '单元格=$cellRef, 目标Description="${target.description}", '
              '草稿Description="${draft.description}", '
              '目标Item="${target.item}", 草稿Item="${draft.item}"');
        } else {
          // 记录为什么没有添加Description
          if (draft.description.isEmpty) {
            debugPrint('[ExcelWriterService] ⚠ 跳过Description写入: 行号=$rowNum, '
                '原因=草稿文件Description为空, 目标Item="${target.item}", 草稿Item="${draft.item}"');
          } else if (targetColumns.descriptionCol < 0) {
            debugPrint('[ExcelWriterService] ⚠ 跳过Description写入: 行号=$rowNum, '
                '原因=Description列索引无效(${targetColumns.descriptionCol})');
          }
        }

        // Unit 列（如果目标文件为空且草稿文件有值）
        if (target.unit.isEmpty &&
            draft.unit != null &&
            draft.unit!.isNotEmpty &&
            targetColumns.unitCol >= 0) {
          final unitColLetter = _colToLetter(targetColumns.unitCol);
          final cellRef = '$unitColLetter$rowNum';
          cellsToUpdate[cellRef] = {
            'value': draft.unit!,
            'col': targetColumns.unitCol,
            'type': 'unit',
          };
          cellsAddedCount++;
        }

        // Qty 列（如果草稿文件有值且目标文件为空）
        if (result.qty != null && targetColumns.qtyCol >= 0) {
          final qtyColLetter = _colToLetter(targetColumns.qtyCol);
          final cellRef = '$qtyColLetter$rowNum';
          final qtyValue = result.qty!;
          cellsToUpdate[cellRef] = {
            'value': qtyValue.toString(),
            'col': targetColumns.qtyCol,
            'type': 'qty',
          };
          writtenValues['qty'] = qtyValue;
          cellsAddedCount++;
        }

        // Rate 列（只要草稿文件有值就写入，即使目标文件已有值）
        if (result.rate != null) {
          final rateColLetter = _colToLetter(target.rateColumn);
          final cellRef = '$rateColLetter$rowNum';
          final rateValue = result.rate!;
          cellsToUpdate[cellRef] = {
            'value': rateValue.toString(),
            'col': target.rateColumn,
            'type': 'rate',
          };
          writtenValues['rate'] = rateValue;
          cellsAddedCount++;
        }

        // Amount 列（只要草稿文件有值就写入，即使目标文件已有值）
        // 如果是Total行，需要生成SUM公式
        if (result.isTotalRow && result.matched) {
          // Total行：生成SUM公式
          final amountColLetter = _colToLetter(target.amountColumn);
          final cellRef = '$amountColLetter$rowNum';

          // 确定SUM公式的范围
          // 找到当前Total行之前的上一个Total行位置
          int? previousTotalRow;
          for (final totalRow in totalRowPositions) {
            if (totalRow < rowNum) {
              previousTotalRow = totalRow;
            } else {
              break; // 已排序，可以提前退出
            }
          }

          // 确定起始行：如果存在上一个Total行，从上一個Total行之后开始；否则从表头之后的第一行开始
          // 需要找到第一个匹配结果的行号作为起始行
          int startRow = rowNum; // 默认从当前行开始（需要调整）
          if (previousTotalRow != null) {
            startRow = previousTotalRow + 1; // 从上一个Total行之后开始
          } else {
            // 找到第一个匹配结果的行号
            for (final r in matchResults) {
              if (r.matched && !r.isTotalRow) {
                startRow = r.target.rowIndex + 1;
                break;
              }
            }
          }

          // 结束行是当前Total行之前一行
          final endRow = rowNum - 1;

          // 生成SUM公式
          String? totalFormula;
          double? calculatedTotal;
          if (startRow <= endRow) {
            final startCellRef = '$amountColLetter$startRow';
            final endCellRef = '$amountColLetter$endRow';
            totalFormula = '=SUM($startCellRef:$endCellRef)';

            // 计算总和（累加范围内所有非Total行的Amount值）
            calculatedTotal = 0.0;
            for (final entry in nonTotalRowAmounts.entries) {
              final rRowNum = entry.key;
              if (rRowNum >= startRow && rRowNum <= endRow) {
                calculatedTotal = (calculatedTotal ?? 0.0) + entry.value;
              }
            }

            debugPrint('[ExcelWriterService] Total行公式: 行号=$rowNum, '
                '范围=$startRow:$endRow, 公式=$totalFormula, '
                '计算值=$calculatedTotal, 草稿值=${result.draftTotalAmount}');
          } else {
            debugPrint(
                '[ExcelWriterService] ⚠ Total行范围无效: 行号=$rowNum, startRow=$startRow, endRow=$endRow');
          }

          // 写入Total行的Amount列（使用公式）
          cellsToUpdate[cellRef] = {
            'value': calculatedTotal?.toString() ?? '0',
            'formula': totalFormula,
            'col': target.amountColumn,
            'type': 'amount',
          };
          writtenValues['amount'] = calculatedTotal;
          writtenValues['amountFormula'] = totalFormula;
          writtenValues['totalFormula'] = totalFormula;
          writtenValues['calculatedTotal'] = calculatedTotal;
          cellsAddedCount++;
        } else if (result.amount != null) {
          // 普通行：使用原有的逻辑
          final amountColLetter = _colToLetter(target.amountColumn);
          final cellRef = '$amountColLetter$rowNum';

          // 检查是否应该使用公式
          bool useFormula = _shouldUseFormula(target, targetColumns);
          String? formula;
          if (useFormula) {
            final qtyCol = targetColumns.qtyCol >= 0
                ? targetColumns.qtyCol
                : target.rateColumn - 1;
            final rateCol = target.rateColumn;
            final qtyColLetter = _colToLetter(qtyCol);
            final rateColLetter = _colToLetter(rateCol);
            formula = '=$qtyColLetter$rowNum*$rateColLetter$rowNum';
          }

          final amountValue = result.amount!;
          cellsToUpdate[cellRef] = {
            'value': amountValue.toString(),
            'formula': formula,
            'col': target.amountColumn,
            'type': 'amount',
          };
          writtenValues['amount'] = amountValue;
          writtenValues['amountFormula'] = formula;
          cellsAddedCount++;
        }

        // 记录该结果的实际写入值
        if (writtenValues.isNotEmpty) {
          resultWrittenValues[result] = writtenValues;
        }
      }

      debugPrint(
          '[ExcelWriterService] 匹配结果统计: 匹配成功 $matchedResultsCount 行, 共添加 $cellsAddedCount 个单元格到更新列表');

      // 输出详细的写入信息（仅前10条，避免日志过长）
      int logCount = 0;
      for (final entry in resultWrittenValues.entries) {
        if (logCount >= 10) break;
        final result = entry.key;
        final written = entry.value;
        final target = result.target;
        final rowNum = target.rowIndex + 1;
        final info = <String>[];
        if (written['qty'] != null) {
          info.add('Qty=${written['qty']}');
        }
        if (written['rate'] != null) {
          info.add('Rate=${written['rate']}');
        }
        if (written['amount'] != null) {
          if (written['amountFormula'] != null) {
            info.add(
                'Amount=${written['amount']} (公式: ${written['amountFormula']})');
          } else {
            info.add('Amount=${written['amount']}');
          }
        }
        if (info.isNotEmpty) {
          debugPrint(
              '[ExcelWriterService] 行 $rowNum (Item=${target.item}): ${info.join(", ")}');
          logCount++;
        }
      }
      if (resultWrittenValues.length > 10) {
        debugPrint(
            '[ExcelWriterService] ... 还有 ${resultWrittenValues.length - 10} 行数据已写入（详细信息请查看PDF报告）');
      }

      // 更新单元格值和清除背景色
      int updatedCount = 0;
      int createdCount = 0;
      final updatedCells = <String>[];

      // 先把所有 <c> 元素收集起来，加速查找
      // 同时构建单元格引用到单元格元素的映射
      final List<XmlElement> cellElements =
          sheetData.findAllElements('c').toList(growable: false);
      final cellRefMap = <String, XmlElement>{};
      final rowMap = <int, XmlElement>{}; // 行号到 <row> 元素的映射

      for (final cell in cellElements) {
        final cellRef = cell.getAttribute('r');
        if (cellRef != null) {
          cellRefMap[cellRef] = cell;
        }
      }

      // 构建行映射，用于创建新单元格时找到对应的行元素
      final rows = sheetData.findElements('row').toList();
      for (final row in rows) {
        final rAttr = row.getAttribute('r');
        if (rAttr != null) {
          final rowNum = int.tryParse(rAttr);
          if (rowNum != null) {
            rowMap[rowNum] = row;
          }
        }
      }

      // 收集需要调整列宽的列
      final columnsToAdjust = <int, double>{};

      // 首先，扫描所有Rate和Amount列的现有值，确保列宽足够
      // 这样可以避免只更新部分单元格时，其他大数字单元格显示为 ####
      final rateCol = targetColumns.unitRateCol;
      final amountCol = targetColumns.amountCol;

      // 扫描所有行，查找Rate和Amount列的最大值
      for (final row in rows) {
        final cells = row.findElements('c').toList();
        for (final cell in cells) {
          final cellRef = cell.getAttribute('r');
          if (cellRef == null) continue;

          final colIndex = _cellRefToColIndex(cellRef);
          if (colIndex == rateCol || colIndex == amountCol) {
            // 获取单元格值
            final vElement = cell.findElements('v').firstOrNull;
            if (vElement != null) {
              final cellValue = vElement.innerText.trim();
              if (cellValue.isNotEmpty) {
                final cellType = colIndex == rateCol ? 'rate' : 'amount';
                final displayWidth =
                    _calculateCellDisplayWidth(cellValue, cellType);

                // 记录每列的最大宽度
                if (!columnsToAdjust.containsKey(colIndex) ||
                    displayWidth > columnsToAdjust[colIndex]!) {
                  columnsToAdjust[colIndex] = displayWidth;
                }
              }
            }
          }
        }
      }

      // 处理所有需要更新的单元格
      for (final entry in cellsToUpdate.entries) {
        final cellRef = entry.key;
        final cellInfo = entry.value;
        final cellType = cellInfo['type'] as String;
        final cellValue = cellInfo['value'] as String;

        // 计算单元格值的显示宽度（考虑数字格式）
        final displayWidth = _calculateCellDisplayWidth(cellValue, cellType);
        final colIndex = cellInfo['col'] as int;

        // 记录每列的最大宽度（包括新写入的值）
        if (!columnsToAdjust.containsKey(colIndex) ||
            displayWidth > columnsToAdjust[colIndex]!) {
          columnsToAdjust[colIndex] = displayWidth;
        }

        // 检查单元格是否已存在
        if (cellRefMap.containsKey(cellRef)) {
          // 单元格已存在，更新它
          final cell = cellRefMap[cellRef]!;
          final cellCol = _cellRefToColIndex(cellRef);
          final expectedCol = cellInfo['col'] as int;

          // 验证列索引
          if (cellCol != expectedCol) {
            debugPrint(
                '[ExcelWriterService] 跳过列索引不匹配的单元格: $cellRef (实际列: $cellCol, 期望列: $expectedCol, 类型: $cellType)');
            continue;
          }

          // 更新单元格值（只更新 <v> 节点，保留所有格式）
          _updateCellValue(cell, cellInfo);
          updatedCount++;
          updatedCells.add(cellRef);

          // 特别记录Description单元格的更新
          if (cellType == 'description') {
            final cellRowNum = _extractRowNumber(cellRef);
            final hasInlineStr = cell.findElements('is').isNotEmpty;
            final hasValue = cell.findElements('v').isNotEmpty;
            debugPrint('[ExcelWriterService] ✓ Description单元格已更新: $cellRef, '
                '行号=$cellRowNum, hasInlineStr=$hasInlineStr, hasValue=$hasValue, '
                't属性=${cell.getAttribute('t')}, 值="${cellInfo['value']}"');
          }
        } else {
          // 单元格不存在，需要创建
          final rowNum = _extractRowNumber(cellRef);
          final colIndex = cellInfo['col'] as int;

          // 找到或创建对应的行元素
          XmlElement? rowElement = rowMap[rowNum];
          if (rowElement == null) {
            // 行不存在，创建新行
            rowElement = XmlElement(XmlName('row'), [
              XmlAttribute(XmlName('r'), rowNum.toString()),
            ]);
            sheetData.children.add(rowElement);
            rowMap[rowNum] = rowElement;
            debugPrint('[ExcelWriterService] 创建新行: 行号 $rowNum');
          }

          // 创建新单元格
          final newCell = XmlElement(XmlName('c'), [
            XmlAttribute(XmlName('r'), cellRef),
          ]);

          // 更新单元格值
          _updateCellValue(newCell, cellInfo);

          // 特别检查Description单元格的创建
          if (cellType == 'description') {
            final hasInlineStr = newCell.findElements('is').isNotEmpty;
            final hasValue = newCell.findElements('v').isNotEmpty;
            debugPrint('[ExcelWriterService] Description单元格创建后检查: $cellRef, '
                'hasInlineStr=$hasInlineStr, hasValue=$hasValue, '
                't属性=${newCell.getAttribute('t')}, '
                '值="${cellInfo['value']}"');
          }

          // 将新单元格添加到行中（按列索引排序）
          bool inserted = false;
          for (int i = 0; i < rowElement.children.length; i++) {
            final child = rowElement.children[i];
            if (child is XmlElement && child.name.local == 'c') {
              final childRef = child.getAttribute('r');
              if (childRef != null) {
                final childCol = _cellRefToColIndex(childRef);
                if (colIndex < childCol) {
                  rowElement.children.insert(i, newCell);
                  inserted = true;
                  break;
                }
              }
            }
          }
          if (!inserted) {
            rowElement.children.add(newCell);
          }

          createdCount++;
          updatedCells.add(cellRef);
          final cellValuePreview = (cellInfo['value'] as String).length > 30
              ? '${(cellInfo['value'] as String).substring(0, 27)}...'
              : cellInfo['value'] as String;
          debugPrint(
              '[ExcelWriterService] 创建新单元格: $cellRef (类型: $cellType, 值: "$cellValuePreview")');

          // 特别记录Description单元格的创建
          if (cellType == 'description') {
            debugPrint('[ExcelWriterService] ✓ Description单元格已创建: $cellRef, '
                '行号=$rowNum, 完整值="${cellInfo['value']}"');
          }
        }
      }

      debugPrint(
          '[ExcelWriterService] 单元格更新统计: 更新了 $updatedCount 个现有单元格, 创建了 $createdCount 个新单元格, 总计 ${updatedCount + createdCount} 个单元格');
      if (updatedCells.isNotEmpty) {
        debugPrint(
            '[ExcelWriterService] 处理的单元格: ${updatedCells.take(10).join(", ")}${updatedCells.length > 10 ? "..." : ""}');
      }

      // 调整列宽（如果值过大）
      if (colsElement != null && columnsToAdjust.isNotEmpty) {
        _adjustColumnWidths(colsElement, columnsToAdjust);
        debugPrint('[ExcelWriterService] 已调整 ${columnsToAdjust.length} 列的宽度');
      }

      // 确保计算模式为自动（在 workbook.xml 中设置）
      // 查找 calcPr 元素（计算属性）
      final calcPrElements = workbookDoc.findAllElements('calcPr');
      if (calcPrElements.isNotEmpty) {
        for (final calcPr in calcPrElements) {
          calcPr.setAttribute('calcMode', 'auto');
          calcPr.setAttribute('fullCalcOnLoad', '1');
        }
      } else {
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

      // 在重新打包时更新 workbook.xml 和工作表文件
      final updatedArchive = Archive();
      for (final file in archive.files) {
        if (file.name == 'xl/workbook.xml') {
          updatedArchive.addFile(
            ArchiveFile(
                file.name, updatedWorkbookBytes.length, updatedWorkbookBytes),
          );
        } else if (file.name == targetSheetFile.name) {
          // 跳过，稍后添加更新后的版本
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

      // 构建更新后的MatchResult列表，包含实际写入的值
      final updatedResults = <MatchResult>[];
      for (final result in matchResults) {
        if (!result.matched) {
          updatedResults.add(result);
          continue;
        }

        final writtenValues = resultWrittenValues[result];
        if (writtenValues != null) {
          updatedResults.add(result.copyWith(
            writtenQty: writtenValues['qty'] as double?,
            writtenRate: writtenValues['rate'] as double?,
            writtenAmount: writtenValues['amount'] as double?,
            amountFormula: writtenValues['amountFormula'] as String?,
            totalFormula: writtenValues['totalFormula'] as String?,
            calculatedTotal: writtenValues['calculatedTotal'] as double?,
          ));
        } else {
          updatedResults.add(result);
        }
      }

      return updatedResults;
    } catch (e, st) {
      debugPrint('[ExcelWriterService] 写入失败: $e');
      debugPrint('[ExcelWriterService] 堆栈: $st');
      rethrow;
    }
  }

  /// 更新单元格值（保留格式）
  void _updateCellValue(XmlElement cell, Map<String, dynamic> cellInfo) {
    final originalTAttr = cell.getAttribute('t');
    final cellType = cellInfo['type'] as String;
    final isTextType =
        cellType == 'item' || cellType == 'description' || cellType == 'unit';

    final vElement = cell.findElements('v').firstOrNull;
    if (vElement != null) {
      if (cellInfo['formula'] != null) {
        // 设置公式
        final fElement = cell.findElements('f').firstOrNull;
        if (fElement != null) {
          fElement.children
            ..clear()
            ..add(XmlText(cellInfo['formula'] as String));
        } else {
          final newFElement = XmlElement(XmlName('f'), const [], [
            XmlText(cellInfo['formula'] as String),
          ]);
          final vIndex = cell.children.indexOf(vElement);
          cell.children.insert(vIndex, newFElement);
        }
        vElement.children
          ..clear()
          ..add(XmlText(cellInfo['value'] as String));
      } else if (isTextType) {
        // 文本类型：需要设置为共享字符串类型（t="s"）
        // 先检查 sharedStrings.xml 中是否已有该字符串
        // 这里简化处理：直接设置 t="inlineStr" 或使用共享字符串
        // 为了简化，我们使用 inlineStr 方式
        cell.setAttribute('t', 'inlineStr');
        // 删除旧的 <v> 节点
        cell.children.remove(vElement);
        // 创建 inlineStr 结构
        final inlineStrElement = XmlElement(XmlName('is'), const [], [
          XmlElement(XmlName('t'), const [], [
            XmlText(cellInfo['value'] as String),
          ]),
        ]);
        cell.children.add(inlineStrElement);

        // 删除公式节点（如果有）
        final fElement = cell.findElements('f').firstOrNull;
        if (fElement != null) {
          cell.children.remove(fElement);
        }
      } else {
        // 普通数值：只更新 <v> 节点
        vElement.children
          ..clear()
          ..add(XmlText(cellInfo['value'] as String));

        // 确保数值单元格有正确的 t 属性（不是字符串类型）
        if (originalTAttr == 's') {
          cell.attributes.removeWhere((attr) => attr.name.local == 't');
        }

        // 如果有 <f> 节点（公式），删除它
        final fElement = cell.findElements('f').firstOrNull;
        if (fElement != null) {
          cell.children.remove(fElement);
        }
      }
    } else {
      // 如果单元格没有 <v> 节点，创建一个
      if (isTextType) {
        // 文本类型：使用 inlineStr
        cell.setAttribute('t', 'inlineStr');
        final inlineStrElement = XmlElement(XmlName('is'), const [], [
          XmlElement(XmlName('t'), const [], [
            XmlText(cellInfo['value'] as String),
          ]),
        ]);
        cell.children.add(inlineStrElement);
      } else {
        // 数值类型：创建 <v> 节点
        final newVElement = XmlElement(XmlName('v'), const [], [
          XmlText(cellInfo['value'] as String),
        ]);
        if (cellInfo['formula'] != null) {
          final newFElement = XmlElement(XmlName('f'), const [], [
            XmlText(cellInfo['formula'] as String),
          ]);
          cell.children.add(newFElement);
        }
        cell.children.add(newVElement);
      }

      // 确保数值单元格的 t 属性正确
      if (!isTextType && cellInfo['formula'] == null) {
        if (cell.getAttribute('t') == 's') {
          cell.attributes.removeWhere((attr) => attr.name.local == 't');
        }
      }
    }
  }

  /// 判断是否应该使用公式（根据目标行的 Qty 列是否有值）
  bool _shouldUseFormula(TargetRow target, TargetColumns columns) {
    return columns.qtyCol >= 0 && target.qty > 0;
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

  /// 从单元格引用中提取行号（如 "A8" -> 8）
  int _extractRowNumber(String cellRef) {
    final match = RegExp(r'^[A-Z]+(\d+)$').firstMatch(cellRef);
    if (match != null) {
      return int.tryParse(match.group(1) ?? '1') ?? 1;
    }
    return 1;
  }

  /// 计算单元格值的显示宽度（字符数）
  /// Excel 列宽单位：1 单位 ≈ 7 像素 ≈ 0.71 字符宽度
  /// 对于货币格式，需要考虑 $ 符号和千位分隔符
  double _calculateCellDisplayWidth(String value, String cellType) {
    if (value.isEmpty) return 8.0; // 默认最小宽度

    // 对于数字，考虑小数点和千位分隔符
    if (cellType == 'rate' || cellType == 'amount' || cellType == 'qty') {
      final numValue = double.tryParse(value);
      if (numValue != null) {
        // 计算货币格式的宽度：$ + 数字 + 千位分隔符 + 小数点 + 2位小数
        // 例如：$160,287.00 = 11 个字符
        final absValue = numValue.abs();
        final integerPart = absValue.floor();
        final integerStr = integerPart.toString();

        // 计算千位分隔符的数量（每3位一个逗号）
        final thousandsCount = (integerStr.length - 1) ~/ 3;

        // 总宽度 = $符号(1) + 整数部分 + 千位分隔符 + 小数点(1) + 2位小数 + 余量(2)
        final totalWidth = 1 + integerStr.length + thousandsCount + 1 + 2 + 2;

        // Excel列宽单位转换：字符数 * 1.2（考虑字体宽度）
        // 对于大数字，确保有足够的宽度
        final width = (totalWidth * 1.2).clamp(12.0, 30.0);

        debugPrint('[ExcelWriterService] 计算列宽: 值=$value, 类型=$cellType, '
            '整数部分长度=${integerStr.length}, 千位分隔符=$thousandsCount, 计算宽度=$width');

        return width;
      }
    }

    // 对于文本，使用字符长度
    return (value.length * 1.2).clamp(8.0, 50.0);
  }

  /// 调整列宽
  void _adjustColumnWidths(
      XmlElement colsElement, Map<int, double> columnWidths) {
    // 获取现有的列宽设置（支持范围列，如 min=1, max=5）
    final existingCols = <int, XmlElement>{};
    for (var col in colsElement.findElements('col')) {
      final minAttr = col.getAttribute('min');
      final maxAttr = col.getAttribute('max');
      if (minAttr != null && maxAttr != null) {
        final min = int.tryParse(minAttr) ?? -1;
        final max = int.tryParse(maxAttr) ?? -1;
        // 处理范围列：如果列索引在范围内，更新该列
        if (min >= 0 && max >= min) {
          for (int i = min; i <= max; i++) {
            if (columnWidths.containsKey(i - 1)) {
              // 转换为0-based索引
              existingCols[i - 1] = col;
            }
          }
        }
      }
    }

    // 更新或创建列宽设置
    for (final entry in columnWidths.entries) {
      final colIndex = entry.key;
      final width = entry.value;
      final colNum = colIndex + 1; // Excel 列号从 1 开始

      if (existingCols.containsKey(colIndex)) {
        // 更新现有列宽
        final col = existingCols[colIndex]!;
        final currentWidth =
            double.tryParse(col.getAttribute('width') ?? '0') ?? 0;
        // 如果新宽度更大，更新它
        if (width > currentWidth) {
          col.setAttribute('width', width.toStringAsFixed(2));
          col.setAttribute('customWidth', '1');
          debugPrint(
              '[ExcelWriterService] 更新列宽: 列${_colToLetter(colIndex)} ($colNum), '
              '从 $currentWidth 更新到 ${width.toStringAsFixed(2)}');
        }
      } else {
        // 创建新的列宽设置
        final newCol = XmlElement(XmlName('col'), [
          XmlAttribute(XmlName('min'), colNum.toString()),
          XmlAttribute(XmlName('max'), colNum.toString()),
          XmlAttribute(XmlName('width'), width.toStringAsFixed(2)),
          XmlAttribute(XmlName('customWidth'), '1'),
        ]);
        colsElement.children.add(newCol);
        debugPrint(
            '[ExcelWriterService] 创建列宽: 列${_colToLetter(colIndex)} ($colNum), '
            '宽度=${width.toStringAsFixed(2)}');
      }
    }

    // 更新 cols count
    final colsCount = colsElement.findElements('col').length;
    colsElement.setAttribute('count', colsCount.toString());
  }
}
