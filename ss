除了值匹配我还需要确保目标文件原有的单元格式不变除了背景颜色，
根据下面的算法和匹配插件优化代码：
# Excel 自动匹配系统 - 数学版算法设计与完整 Dart 实现

本文给出一套 **基于向量特征与加权评分的匹配算法** 的完整 Dart 实现，包含：

1. `TargetRow / DraftRow / MatchResult` 数据模型  
2. 行向量化（item / description / unit / qty → vector）  
3. 匹配评分函数（Score）  
4. 完整 `ExcelMatchService`（数学版），可直接集成到 Flutter 项目中  

> 说明：  
> - 目标文件使用 **Syncfusion XlsIO** 读取（可读背景色）  
> - 草稿文件使用 **excel 包** 读取（只读数据）  
> - 文本相似度基于 `string_similarity`（Levenshtein-based）  

---

## 一、模型定义：`excel_models.dart`

```dart
// lib/models/excel_models.dart

/// 目标文件中的一行（需要匹配并写回）
class TargetRow {
  final String item;
  final String description;
  final String unit;
  final double qty;

  /// 行索引（0-based，内部使用）
  final int rowIndex;

  /// Rate / Amount 列索引（0-based）
  final int rateColumn;
  final int amountColumn;

  const TargetRow({
    required this.item,
    required this.description,
    required this.unit,
    required this.qty,
    required this.rowIndex,
    required this.rateColumn,
    required this.amountColumn,
  });
}

/// 草稿文件中的一行（提供单价与总价）
class DraftRow {
  final String item;
  final String description;
  final String? unit;
  final double? qty;
  final double? rate;
  final double? amount;

  const DraftRow({
    required this.item,
    required this.description,
    this.unit,
    this.qty,
    this.rate,
    this.amount,
  });

  DraftRow copyWith({
    String? item,
    String? description,
    String? unit,
    double? qty,
    double? rate,
    double? amount,
  }) {
    return DraftRow(
      item: item ?? this.item,
      description: description ?? this.description,
      unit: unit ?? this.unit,
      qty: qty ?? this.qty,
      rate: rate ?? this.rate,
      amount: amount ?? this.amount,
    );
  }
}

/// 匹配结果类型
class MatchResult {
  final TargetRow target;
  final DraftRow? draft;
  final bool matched;
  final double? rate;
  final double? amount;

  /// strong / medium / weak / none
  final String matchType;

  const MatchResult({
    required this.target,
    this.draft,
    required this.matched,
    this.rate,
    this.amount,
    this.matchType = 'none',
  });
}
二、行向量化与评分函数：vector_math_utils.dart
dart
Copy code
// lib/services/vector_math_utils.dart

import 'package:string_similarity/string_similarity.dart';

/// 行特征向量
class RowFeatureVector {
  /// [itemLevel1Norm, itemLevel2Norm, unitCode, qtyNorm]
  final List<double> features;

  const RowFeatureVector(this.features);
}

/// 向量化 & 评分工具类
class MatchVectorUtils {
  /// Item 字段转 [level1, level2]，再归一化
  /// 例如 "1.14" -> [1, 14] / [maxL1, maxL2]
  /// 这里为了简单，假定 level1 ∈ [0, 20]，level2 ∈ [0, 200]
  static List<double> itemToVector(String item) {
    final norm = _normalizeText(item);
    if (norm.isEmpty) return [0.0, 0.0];

    final parts = norm.split('.');
    int l1 = 0;
    int l2 = 0;
    if (parts.isNotEmpty) {
      l1 = int.tryParse(parts[0]) ?? 0;
    }
    if (parts.length > 1) {
      l2 = int.tryParse(parts[1]) ?? 0;
    }

    // 简单归一化
    final l1Norm = l1 / 20.0; // 假设最多 20 大项
    final l2Norm = l2 / 200.0; // 假设最多 200 小项

    return [l1Norm.clamp(0.0, 1.0), l2Norm.clamp(0.0, 1.0)];
  }

  /// Unit 转类别编码并归一化
  /// Sum -> 1, Number -> 2, M2 -> 3, 其他 -> 4
  static double unitToCode(String? unit) {
    if (unit == null || unit.trim().isEmpty) return 0.0;
    final u = _normalizeText(unit);
    int code;
    if (u == 'sum') {
      code = 1;
    } else if (u == 'number' || u == 'no') {
      code = 2;
    } else if (u == 'm2' || u == 'm²') {
      code = 3;
    } else {
      code = 4;
    }
    return code / 10.0; // 归一化到 0~0.4 区间
  }

  /// Qty 归一化（假定 0~1000）
  static double qtyToNorm(double? qty) {
    if (qty == null || qty <= 0) return 0.0;
    return (qty / 1000.0).clamp(0.0, 1.0);
  }

  /// 生成行特征向量（不含 description，相似度是成对计算的）
  static RowFeatureVector buildVector({
    required String item,
    required String? unit,
    required double? qty,
  }) {
    final itemVec = itemToVector(item);
    final unitCode = unitToCode(unit);
    final qtyNorm = qtyToNorm(qty);

    return RowFeatureVector([
      itemVec[0],
      itemVec[1],
      unitCode,
      qtyNorm,
    ]);
  }

  /// 描述文本相似度（0~1）
  static double descriptionSimilarity(String a, String b) {
    final na = _normalizeText(a);
    final nb = _normalizeText(b);
    if (na.isEmpty || nb.isEmpty) return 0.0;

    return StringSimilarity.compareTwoStrings(na, nb);
  }

  /// 计算两行的匹配评分（越低越好）
  ///
  /// w1: Item 距离权重
  /// w2: 描述相似度权重
  /// w3: Unit 不匹配惩罚权重
  /// w4: Qty 不匹配惩罚权重
  static double computeMatchScore({
    required RowFeatureVector targetVec,
    required RowFeatureVector draftVec,
    required String targetDesc,
    required String draftDesc,
    required String targetUnit,
    required String? draftUnit,
    required double targetQty,
    required double? draftQty,
    double w1 = 0.45,
    double w2 = 0.40,
    double w3 = 0.10,
    double w4 = 0.05,
  }) {
    // 1. Item 向量距离（欧氏距离）
    final itemDist = _euclideanDistance(
      targetVec.features.sublist(0, 2),
      draftVec.features.sublist(0, 2),
    );

    // 2. 描述文本相似度
    final descSim = descriptionSimilarity(targetDesc, draftDesc);
    final descPenalty = 1.0 - descSim; // 相似度越高惩罚越小

    // 3. Unit 惩罚
    double unitPenalty = 0.0;
    final tu = _normalizeText(targetUnit);
    final du = _normalizeText(draftUnit ?? '');
    if (tu.isNotEmpty && du.isNotEmpty && tu != du) {
      unitPenalty = 1.0;
    }

    // 4. Qty 惩罚（比例差）
    double qtyPenalty = 0.0;
    if (targetQty > 0 && (draftQty ?? 0) > 0) {
      final diff = (targetQty - (draftQty ?? 0)).abs();
      final maxQty = targetQty > (draftQty ?? 0) ? targetQty : (draftQty ?? 0);
      qtyPenalty = (diff / maxQty).clamp(0.0, 1.0);
    }

    final score =
        w1 * itemDist + w2 * descPenalty + w3 * unitPenalty + w4 * qtyPenalty;
    return score;
  }

  /// 欧氏距离
  static double _euclideanDistance(List<double> a, List<double> b) {
    final n = a.length;
    double sum = 0.0;
    for (var i = 0; i < n; i++) {
      final d = a[i] - b[i];
      sum += d * d;
    }
    return sum == 0 ? 0.0 : sum.sqrt();
  }
}

/// sqrt 的简单扩展
extension _SqrtExt on double {
  double sqrt() => this <= 0 ? 0.0 : MathHelper.sqrt(this);
}

/// 简单数学助手（避免引入 dart:math 时命名冲突）
class MathHelper {
  static double sqrt(double x) => x >= 0 ? x.toDouble()._sqrtNewton() : 0.0;
}

extension _Newton on double {
  double _sqrtNewton({int iterations = 8}) {
    var x = this;
    if (x <= 0) return 0.0;
    var r = x;
    for (var i = 0; i < iterations; i++) {
      r = 0.5 * (r + x / r);
    }
    return r;
  }
}

/// 文本归一化
String _normalizeText(String text) {
  return text
      .toLowerCase()
      .replaceAll(RegExp(r'\s+'), '')
      .replaceAll(RegExp(r'[()（）]'), '');
}
说明：

为了不依赖 dart:math，这里用一个简单的牛顿迭代算 sqrt（你也可以直接 import 'dart:math' as math; 然后用 math.sqrt）。

描述相似度用了 string_similarity 提供的 Levenshtein-based 评分。

三、数学版 ExcelMatchService：excel_match_service_math.dart
dart
Copy code
// lib/services/excel_match_service_math.dart

import 'dart:io';

import 'package:excel/excel.dart';
import 'package:flutter/foundation.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;

import '../models/excel_models.dart';
import 'vector_math_utils.dart';

/// 匹配服务返回结果
class MatchServiceResult {
  final bool success;
  final String message;
  final int matchedCount;
  final int totalCount;
  final List<String> logs;

  MatchServiceResult({
    required this.success,
    required this.message,
    required this.matchedCount,
    required this.totalCount,
    required this.logs,
  });
}

/// 目标文件列结构
class TargetColumns {
  final int itemCol;
  final int descriptionCol;
  final int unitCol;
  final int qtyCol;
  final int unitRateCol;
  final int amountCol;
  const TargetColumns({
    required this.itemCol,
    required this.descriptionCol,
    required this.unitCol,
    required this.qtyCol,
    required this.unitRateCol,
    required this.amountCol,
  });
}

/// 草稿文件列结构
class DraftColumns {
  final int itemCol;
  final int descriptionCol;
  final int unitCol;
  final int qtyCol;
  final int rateCol;
  final int amountCol;
  const DraftColumns({
    required this.itemCol,
    required this.descriptionCol,
    required this.unitCol,
    required this.qtyCol,
    required this.rateCol,
    required this.amountCol,
  });
}

/// 数学版 Excel 匹配服务
class ExcelMatchServiceMath {
  /// 主入口
  Future<MatchServiceResult> matchExcelFiles({
    required File draftFile,
    required File targetFile,
    required String outputPath,
  }) async {
    final logs = <String>[];
    int matchedCount = 0;
    int totalCount = 0;

    try {
      logs.add('👉 开始匹配流程 (Math Version)...');
      logs.add('草稿文件: ${draftFile.path}');
      logs.add('目标文件: ${targetFile.path}');

      // 1. 读取目标文件（Syncfusion）
      logs.add('加载目标文件...');
      final targetBytes = await targetFile.readAsBytes();
      final targetWorkbook = xlsio.Workbook.open(targetBytes);
      if (targetWorkbook.worksheets.count == 0) {
        targetWorkbook.dispose();
        return MatchServiceResult(
          success: false,
          message: '目标文件中没有工作表',
          matchedCount: 0,
          totalCount: 0,
          logs: logs,
        );
      }
      final targetSheet = targetWorkbook.worksheets[0];
      logs.add('目标工作表: ${targetSheet.name}');

      final targetColumns = _identifyTargetColumns(targetSheet);
      if (targetColumns == null) {
        targetWorkbook.dispose();
        return MatchServiceResult(
          success: false,
          message:
              '无法识别目标文件的列结构，请确保表头包含 Item / Description / Unit / Qty / Unit Rate / Amount',
          matchedCount: 0,
          totalCount: 0,
          logs: logs,
        );
      }

      logs.add(
          '目标列识别完成: item=${_colToLetter(targetColumns.itemCol)}, desc=${_colToLetter(targetColumns.descriptionCol)}, rate=${_colToLetter(targetColumns.unitRateCol)}, amount=${_colToLetter(targetColumns.amountCol)}');

      // 2. 扫描目标文件中的黄色单元格行
      logs.add('扫描目标表中的黄色单元格行...');
      final targetRows = _scanYellowRows(targetSheet, targetColumns);
      totalCount = targetRows.length;
      logs.add('共发现 $totalCount 行需要匹配填充。');

      if (targetRows.isEmpty) {
        targetWorkbook.dispose();
        return MatchServiceResult(
          success: false,
          message: '目标文件中未发现黄色单元格行，可能无需匹配。',
          matchedCount: 0,
          totalCount: 0,
          logs: logs,
        );
      }

      // 3. 构建草稿数据集
      logs.add('解析草稿文件...');
      final draftBuild = await _buildDraftMap(draftFile);
      final draftMap = draftBuild['map'] as Map<String, DraftRow>;
      final draftColumns = draftBuild['columns'] as DraftColumns?;
      if (draftColumns == null || draftMap.isEmpty) {
        targetWorkbook.dispose();
        return MatchServiceResult(
          success: false,
          message: '无法识别草稿文件的列结构或数据为空，请检查草稿文件。',
          matchedCount: 0,
          totalCount: totalCount,
          logs: logs,
        );
      }
      logs.add('草稿数据解析完成，共 ${draftMap.length} 条有效记录。');

      // 预构建 Draft 向量缓存（避免重复计算）
      final draftVectorCache = <DraftRow, RowFeatureVector>{};
      for (final d in draftMap.values) {
        draftVectorCache[d] = MatchVectorUtils.buildVector(
          item: d.item,
          unit: d.unit,
          qty: d.qty,
        );
      }

      // 4. 对每个目标行执行数学匹配
      logs.add('开始执行数学匹配...');
      final matchResults = <MatchResult>[];

      for (final t in targetRows) {
        final tVec = MatchVectorUtils.buildVector(
          item: t.item,
          unit: t.unit,
          qty: t.qty,
        );

        final matchResult = _matchTargetRowMath(
          target: t,
          targetVector: tVec,
          draftMap: draftMap,
          draftVectorCache: draftVectorCache,
        );

        matchResults.add(matchResult);

        if (matchResult.matched) {
          matchedCount++;
          logs.add(
              '✅ [${matchResult.matchType}] ${t.item} | ${t.description} -> rate=${matchResult.rate}, amount=${matchResult.amount}');
        } else {
          logs.add('⚠️ 未匹配: ${t.item} | ${t.description}');
        }
      }

      // 5. 写回目标文件并清除黄色背景
      logs.add('写回匹配值并清除背景...');
      _writeMatchedValues(targetSheet, matchResults);
      final outputBytes = targetWorkbook.saveAsStream();
      targetWorkbook.dispose();

      // 6. 保存文件
      final outFile = File(outputPath);
      if (!await outFile.parent.exists()) {
        await outFile.parent.create(recursive: true);
      }
      await outFile.writeAsBytes(outputBytes);
      logs.add('输出文件: ${outFile.path}');

      return MatchServiceResult(
        success: true,
        message: '匹配完成：成功 $matchedCount / $totalCount',
        matchedCount: matchedCount,
        totalCount: totalCount,
        logs: logs,
      );
    } catch (e, st) {
      debugPrint('[ExcelMatchServiceMath] error: $e');
      debugPrint('[ExcelMatchServiceMath] stack: $st');
      logs.add('❌ 发生异常: $e');
      return MatchServiceResult(
        success: false,
        message: '匹配失败: $e',
        matchedCount: matchedCount,
        totalCount: totalCount,
        logs: logs,
      );
    }
  }

  // =============================
  //  目标文件解析相关
  // =============================

  TargetColumns? _identifyTargetColumns(xlsio.Worksheet sheet) {
    int? itemCol, descCol, unitCol, qtyCol, rateCol, amountCol;

    // 第一行作为表头（1-based）
    for (var col = 1; col <= 50; col++) {
      final cell = sheet.getRangeByIndex(5, col); // 你的表头在第 5 行，可按需调整
      final text = (cell.displayText ?? '').toLowerCase();
      final norm = _normalizeText(text);

      if (itemCol == null &&
          _matchesHeader(norm, ['item', 'itemno', 'no'])) {
        itemCol = col - 1;
      } else if (descCol == null &&
          _matchesHeader(norm, ['description', 'descofwork'])) {
        descCol = col - 1;
      } else if (unitCol == null && _matchesHeader(norm, ['unit', 'u'])) {
        unitCol = col - 1;
      } else if (qtyCol == null &&
          _matchesHeader(norm, ['qty', 'quantity', 'a'])) {
        qtyCol = col - 1;
      } else if (rateCol == null &&
          _matchesHeader(norm, ['unitrate', 'rate', 'b'])) {
        rateCol = col - 1;
      } else if (amountCol == null &&
          _matchesHeader(norm, ['amount', 'total', 'c'])) {
        amountCol = col - 1;
      }
    }

    if (itemCol == null || descCol == null || rateCol == null || amountCol == null) {
      return null;
    }

    return TargetColumns(
      itemCol: itemCol,
      descriptionCol: descCol,
      unitCol: unitCol ?? -1,
      qtyCol: qtyCol ?? -1,
      unitRateCol: rateCol,
      amountCol: amountCol,
    );
  }

  List<TargetRow> _scanYellowRows(
    xlsio.Worksheet sheet,
    TargetColumns cols,
  ) {
    final rows = <TargetRow>[];

    // 目标文件数据从第 7 行开始（你的示例中），可按需调整
    for (var r = 7; r <= 1000; r++) {
      final oneCell = sheet.getRangeByIndex(r, cols.itemCol + 1);
      if ((oneCell.displayText ?? '').isEmpty && r > 20) {
        // 认为到尾部可以结束
        break;
      }

      final rateCell = sheet.getRangeByIndex(r, cols.unitRateCol + 1);
      final amountCell = sheet.getRangeByIndex(r, cols.amountCol + 1);

      final isYellow = _isYellowCell(rateCell) || _isYellowCell(amountCell);
      if (!isYellow) continue;

      final item = sheet
          .getRangeByIndex(r, cols.itemCol + 1)
          .displayText
          .toString()
          .trim();
      if (item.isEmpty) continue;

      final desc = sheet
          .getRangeByIndex(r, cols.descriptionCol + 1)
          .displayText
          .toString()
          .trim();

      String unit = '';
      if (cols.unitCol >= 0) {
        unit = sheet
            .getRangeByIndex(r, cols.unitCol + 1)
            .displayText
            .toString()
            .trim();
      }

      double qty = 0;
      if (cols.qtyCol >= 0) {
        final qStr = sheet
            .getRangeByIndex(r, cols.qtyCol + 1)
            .displayText
            .toString()
            .trim();
        qty = double.tryParse(qStr) ?? 0.0;
      }

      rows.add(
        TargetRow(
          item: item,
          description: desc,
          unit: unit,
          qty: qty,
          rowIndex: r - 1,
          rateColumn: cols.unitRateCol,
          amountColumn: cols.amountCol,
        ),
      );
    }

    return rows;
  }

  bool _isYellowCell(xlsio.Range cell) {
    try {
      final color = cell.cellStyle.backColor;
      if (color.isEmpty) return false;
      final c = color.toUpperCase().replaceAll('#', '');
      final hex = c.length > 6 ? c.substring(c.length - 6) : c;
      // 简单判断：R/G 都高，B 低
      final r = int.parse(hex.substring(0, 2), radix: 16);
      final g = int.parse(hex.substring(2, 4), radix: 16);
      final b = int.parse(hex.substring(4, 6), radix: 16);
      return r > 200 && g > 200 && b < 150;
    } catch (_) {
      return false;
    }
  }

  // =============================
  //  草稿文件解析相关（excel 包）
  // =============================

  Future<Map<String, dynamic>> _buildDraftMap(File draftFile) async {
    final map = <String, DraftRow>{};
    final bytes = await draftFile.readAsBytes();
    Excel excel;
    try {
      excel = Excel.decodeBytes(bytes);
    } catch (e) {
      debugPrint('[ExcelMatchServiceMath] excel decode error: $e');
      return {'map': map, 'columns': null};
    }

    DraftColumns? draftCols;

    for (final name in excel.tables.keys) {
      final sheet = excel.tables[name]!;
      if (sheet.rows.isEmpty) continue;

      draftCols ??= _identifyDraftColumns(sheet);
      if (draftCols == null) continue;

      for (final row in sheet.rows.skip(1)) {
        if (row.isEmpty) continue;
        final rowIndex = row.first?.rowIndex ?? 0;
        if (rowIndex < 1) continue;

        String item = _cellStr(sheet, rowIndex, draftCols.itemCol);
        if (item.isEmpty) continue;

        String desc = _cellStr(sheet, rowIndex, draftCols.descriptionCol);

        String? unit;
        if (draftCols.unitCol >= 0) {
          unit = _cellStr(sheet, rowIndex, draftCols.unitCol);
        }

        double? qty;
        if (draftCols.qtyCol >= 0) {
          final qStr = _cellStr(sheet, rowIndex, draftCols.qtyCol);
          qty = double.tryParse(qStr);
        }

        double? rate;
        if (draftCols.rateCol >= 0) {
          final rStr = _cellStr(sheet, rowIndex, draftCols.rateCol);
          rate = double.tryParse(rStr);
        }

        double? amount;
        if (draftCols.amountCol >= 0) {
          final aStr = _cellStr(sheet, rowIndex, draftCols.amountCol);
          amount = double.tryParse(aStr);
        }

        final key = _normalizeText('$item|$desc');
        final rowData = DraftRow(
          item: item,
          description: desc,
          unit: unit,
          qty: qty,
          rate: rate,
          amount: amount,
        );

        if (!map.containsKey(key)) {
          map[key] = rowData;
        } else {
          // 合并信息：优先有 rate/amount 的记录
          final old = map[key]!;
          map[key] = DraftRow(
            item: old.item,
            description: old.description,
            unit: old.unit ?? unit,
            qty: old.qty ?? qty,
            rate: old.rate ?? rate,
            amount: old.amount ?? amount,
          );
        }
      }
    }

    return {'map': map, 'columns': draftCols};
  }

  DraftColumns? _identifyDraftColumns(Sheet sheet) {
    final header = sheet.rows.first;
    int? itemCol, descCol, unitCol, qtyCol, rateCol, amountCol;

    for (final cell in header) {
      if (cell == null) continue;
      final idx = cell.columnIndex;
      final text = _normalizeText(cell.value.toString());

      if (itemCol == null &&
          _matchesHeader(text, ['item', 'itemno', 'no'])) {
        itemCol = idx;
      } else if (descCol == null &&
          _matchesHeader(text, ['description', 'descofwork'])) {
        descCol = idx;
      } else if (unitCol == null &&
          _matchesHeader(text, ['unit', 'u'])) {
        unitCol = idx;
      } else if (qtyCol == null &&
          _matchesHeader(text, ['qty', 'quantity'])) {
        qtyCol = idx;
      } else if (rateCol == null &&
          _matchesHeader(text, ['rate', 'unitrate', 'b'])) {
        rateCol = idx;
      } else if (amountCol == null &&
          _matchesHeader(text, ['amount', 'total', 'c'])) {
        amountCol = idx;
      }
    }

    if (itemCol == null || descCol == null || rateCol == null || amountCol == null) {
      return null;
    }

    return DraftColumns(
      itemCol: itemCol,
      descriptionCol: descCol,
      unitCol: unitCol ?? -1,
      qtyCol: qtyCol ?? -1,
      rateCol: rateCol,
      amountCol: amountCol,
    );
  }

  String _cellStr(Sheet sheet, int rowIndex, int colIndex) {
    try {
      final cell = sheet
          .cell(CellIndex.indexByColumnRow(columnIndex: colIndex, rowIndex: rowIndex));
      final v = cell.value;
      return v?.toString().trim() ?? '';
    } catch (_) {
      return '';
    }
  }

  // =============================
  //  数学匹配核心
  // =============================

  MatchResult _matchTargetRowMath({
    required TargetRow target,
    required RowFeatureVector targetVector,
    required Map<String, DraftRow> draftMap,
    required Map<DraftRow, RowFeatureVector> draftVectorCache,
  }) {
    final key = _normalizeText('${target.item}|${target.description}');
    // 0. 强匹配：完全 key 命中
    if (draftMap.containsKey(key)) {
      final d = draftMap[key]!;
      return MatchResult(
        target: target,
        draft: d,
        matched: true,
        rate: d.rate,
        amount: d.amount,
        matchType: 'strong',
      );
    }

    // 1. 遍历草稿行，计算 Score，寻找最小值
    DraftRow? bestDraft;
    double bestScore = double.infinity;
    double bestDescSim = 0.0;

    for (final d in draftMap.values) {
      final dVec = draftVectorCache[d]!;

      final score = MatchVectorUtils.computeMatchScore(
        targetVec: targetVector,
        draftVec: dVec,
        targetDesc: target.description,
        draftDesc: d.description,
        targetUnit: target.unit,
        draftUnit: d.unit,
        targetQty: target.qty,
        draftQty: d.qty,
      );

      final descSim =
          MatchVectorUtils.descriptionSimilarity(target.description, d.description);

      if (score < bestScore) {
        bestScore = score;
        bestDraft = d;
        bestDescSim = descSim;
      }
    }

    if (bestDraft == null) {
      return MatchResult(target: target, matched: false);
    }

    // 2. 根据 score & descSim 决定匹配类型
    String type;
    bool ok = false;

    if (bestDescSim >= 0.9 && bestScore <= 0.3) {
      type = 'strong';
      ok = true;
    } else if (bestDescSim >= 0.8 && bestScore <= 0.45) {
      type = 'medium';
      ok = true;
    } else if (bestDescSim >= 0.7 && bestScore <= 0.6) {
      type = 'weak';
      ok = true;
    } else {
      type = 'none';
      ok = false;
    }

    if (!ok) {
      return MatchResult(target: target, draft: bestDraft, matched: false);
    }

    return MatchResult(
      target: target,
      draft: bestDraft,
      matched: true,
      rate: bestDraft.rate,
      amount: bestDraft.amount,
      matchType: type,
    );
  }

  // =============================
  //  写回目标文件
  // =============================

  void _writeMatchedValues(
    xlsio.Worksheet sheet,
    List<MatchResult> results,
  ) {
    for (final r in results) {
      if (!r.matched) continue;
      final rowIndex = r.target.rowIndex + 1;

      if (r.rate != null) {
        final rateCell =
            sheet.getRangeByIndex(rowIndex, r.target.rateColumn + 1);
        rateCell.setNumber(r.rate!);
        rateCell.cellStyle.backColor = '#FFFFFF';
      }

      if (r.amount != null) {
        final amountCell =
            sheet.getRangeByIndex(rowIndex, r.target.amountColumn + 1);
        amountCell.setNumber(r.amount!);
        amountCell.cellStyle.backColor = '#FFFFFF';
      }
    }
  }

  // =============================
  //  工具函数
  // =============================

  bool _matchesHeader(String header, List<String> patterns) {
    for (final p in patterns) {
      if (header.contains(_normalizeText(p))) return true;
    }
    return false;
  }

  String _normalizeText(String text) {
    return text
        .toLowerCase()
        .replaceAll(RegExp(r'\s+'), '')
        .replaceAll(RegExp(r'[()（）]'), '');
  }

  String _colToLetter(int col) => String.fromCharCode(65 + col);
}
四、使用方式示例
dart
Copy code
final service = ExcelMatchServiceMath();
final result = await service.matchExcelFiles(
  draftFile: File('/path/to/draft.xlsx'),
  targetFile: File('/path/to/target.xlsx'),
  outputPath: '/path/to/output/target_filled.xlsx',
);

if (result.success) {
  print(result.message);
} else {
  print('匹配失败: ${result.message}');
}

for (final log in result.logs) {
  debugPrint(log);
}
五、总结
上面的实现把 行数据转为数值向量，通过 欧氏距离 + 文本相似度 + 单位/数量惩罚 构造了一个数学意义上的匹配评分函数。

结合工程规则（强匹配/中匹配/弱匹配阈值），可以在真实工程 BQ / Tender 的 Excel 中做到自动、鲁棒的“草稿 → 目标”价目匹配。

你可以在此基础上继续调权值、阈值或替换为更复杂的文本嵌入（例如接入 NLP 向量库）来进一步提升精度。