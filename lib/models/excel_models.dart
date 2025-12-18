import 'package:equatable/equatable.dart';

/// 单元格模型：存储行、列、值以及一些简单样式信息
class ExcelCell extends Equatable {
  final int row; // 0-based
  final int column; // 0-based
  final String? value;

  // 简化的样式信息（后续可扩展）
  final bool isBold;
  final int? backgroundColorArgb;

  const ExcelCell({
    required this.row,
    required this.column,
    this.value,
    this.isBold = false,
    this.backgroundColorArgb,
  });

  ExcelCell copyWith({
    int? row,
    int? column,
    String? value,
    bool? isBold,
    int? backgroundColorArgb,
  }) {
    return ExcelCell(
      row: row ?? this.row,
      column: column ?? this.column,
      value: value ?? this.value,
      isBold: isBold ?? this.isBold,
      backgroundColorArgb: backgroundColorArgb ?? this.backgroundColorArgb,
    );
  }

  @override
  List<Object?> get props => [row, column, value, isBold, backgroundColorArgb];
}

/// 合并单元格范围（如 A1:C3）
class MergeRange extends Equatable {
  final int startRow;
  final int startColumn;
  final int endRow;
  final int endColumn;

  const MergeRange({
    required this.startRow,
    required this.startColumn,
    required this.endRow,
    required this.endColumn,
  });

  @override
  List<Object?> get props => [startRow, startColumn, endRow, endColumn];
}

/// 一整行
class ExcelRow extends Equatable {
  final int index; // 0-based
  final List<ExcelCell> cells;

  const ExcelRow({
    required this.index,
    required this.cells,
  });

  @override
  List<Object?> get props => [index, cells];
}

/// 一个工作表的数据模型
class ExcelSheetModel extends Equatable {
  final String name;
  final List<ExcelRow> rows;
  final List<MergeRange> mergeRanges;

  const ExcelSheetModel({
    required this.name,
    required this.rows,
    this.mergeRanges = const [],
  });

  int get rowCount => rows.length;

  int get columnCount => rows.isEmpty
      ? 0
      : rows.map((r) => r.cells.length).reduce((a, b) => a > b ? a : b);

  ExcelSheetModel copyWith({
    String? name,
    List<ExcelRow>? rows,
    List<MergeRange>? mergeRanges,
  }) {
    return ExcelSheetModel(
      name: name ?? this.name,
      rows: rows ?? this.rows,
      mergeRanges: mergeRanges ?? this.mergeRanges,
    );
  }

  @override
  List<Object?> get props => [name, rows, mergeRanges];
}

/// 解锁历史记录项：记录每一次解锁操作的关键信息
class UnlockHistoryItem extends Equatable {
  final DateTime timestamp;
  final String originalPath;
  final String outputPath;
  final bool isBatch;
  final bool success;
  final String? errorMessage;

  const UnlockHistoryItem({
    required this.timestamp,
    required this.originalPath,
    required this.outputPath,
    required this.isBatch,
    required this.success,
    this.errorMessage,
  });

  @override
  List<Object?> get props =>
      [timestamp, originalPath, outputPath, isBatch, success, errorMessage];
}

/// 目标文件行结构（需要匹配的行）
/// 根据 ss 文档：从黄色单元格所在行提取字段
class TargetRow extends Equatable {
  final String item;
  final String description;
  final String unit;
  final double qty;
  final int rowIndex; // 0-based（存储时）
  final int rateColumn; // Rate 列的黄色单元格位置（0-based: 0=A, 1=B, ..., 4=E, 5=F）
  final int amountColumn; // Amount 列的黄色单元格位置（0-based）

  const TargetRow({
    required this.item,
    required this.description,
    required this.unit,
    required this.qty,
    required this.rowIndex,
    required this.rateColumn,
    required this.amountColumn,
  });

  @override
  List<Object?> get props =>
      [item, description, unit, qty, rowIndex, rateColumn, amountColumn];
}

/// 草稿文件行结构（源数据）
class DraftRow extends Equatable {
  final String item;
  final String description;
  final String? unit;
  final double? qty;
  final double? rate; // Unit Rate
  final double? amount; // Amount

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

  @override
  List<Object?> get props => [item, description, unit, qty, rate, amount];
}

/// 匹配结果
/// 根据 ss 文档：匹配成功后填写 rate 和 amount，并清除黄色背景
class MatchResult extends Equatable {
  final TargetRow target;
  final DraftRow? draft;
  final bool matched;
  final double? qty; // Qty 值（如果草稿文件有值且目标文件为空，则匹配）
  final double? rate; // Unit Rate 值（匹配到的值）
  final double? amount; // Amount 值（匹配到的值）
  final String? matchType; // 'strong', 'medium', 'weak', or null
  // 实际写入的值（可能与匹配值不同，如使用公式）
  final double? writtenQty; // 实际写入的 Qty 值
  final double? writtenRate; // 实际写入的 Rate 值
  final double? writtenAmount; // 实际写入的 Amount 值
  final String? amountFormula; // Amount 列的公式（如果有）
  // Total行相关字段
  final bool isTotalRow; // 标识是否为Total行
  final double? draftTotalAmount; // 草稿文件中的Total值（用于对比）
  final String? totalFormula; // Total行的SUM公式
  final double? calculatedTotal; // 通过公式计算的总值（用于对比）

  const MatchResult({
    required this.target,
    this.draft,
    required this.matched,
    this.qty,
    this.rate,
    this.amount,
    this.matchType,
    this.writtenQty,
    this.writtenRate,
    this.writtenAmount,
    this.amountFormula,
    this.isTotalRow = false,
    this.draftTotalAmount,
    this.totalFormula,
    this.calculatedTotal,
  });

  MatchResult copyWith({
    TargetRow? target,
    DraftRow? draft,
    bool? matched,
    double? qty,
    double? rate,
    double? amount,
    String? matchType,
    double? writtenQty,
    double? writtenRate,
    double? writtenAmount,
    String? amountFormula,
    bool? isTotalRow,
    double? draftTotalAmount,
    String? totalFormula,
    double? calculatedTotal,
  }) {
    return MatchResult(
      target: target ?? this.target,
      draft: draft ?? this.draft,
      matched: matched ?? this.matched,
      qty: qty ?? this.qty,
      rate: rate ?? this.rate,
      amount: amount ?? this.amount,
      matchType: matchType ?? this.matchType,
      writtenQty: writtenQty ?? this.writtenQty,
      writtenRate: writtenRate ?? this.writtenRate,
      writtenAmount: writtenAmount ?? this.writtenAmount,
      amountFormula: amountFormula ?? this.amountFormula,
      isTotalRow: isTotalRow ?? this.isTotalRow,
      draftTotalAmount: draftTotalAmount ?? this.draftTotalAmount,
      totalFormula: totalFormula ?? this.totalFormula,
      calculatedTotal: calculatedTotal ?? this.calculatedTotal,
    );
  }

  @override
  List<Object?> get props => [
        target,
        draft,
        matched,
        qty,
        rate,
        amount,
        matchType,
        writtenQty,
        writtenRate,
        writtenAmount,
        amountFormula,
        isTotalRow,
        draftTotalAmount,
        totalFormula,
        calculatedTotal,
      ];
}

/// 目标文件列结构（自动识别）
class TargetColumns extends Equatable {
  final int itemCol; // Item 列索引（0-based）
  final int descriptionCol; // Description 列索引
  final int unitCol; // Unit 列索引
  final int qtyCol; // Qty 列索引
  final int unitRateCol; // Unit Rate 列索引
  final int amountCol; // Amount 列索引

  const TargetColumns({
    required this.itemCol,
    required this.descriptionCol,
    required this.unitCol,
    required this.qtyCol,
    required this.unitRateCol,
    required this.amountCol,
  });

  @override
  List<Object?> get props =>
      [itemCol, descriptionCol, unitCol, qtyCol, unitRateCol, amountCol];
}

/// 草稿文件列结构（自动识别）
class DraftColumns extends Equatable {
  final int itemCol;
  final int descriptionCol;
  final int unitCol;
  final int qtyCol;
  final int rateCol; // Unit Rate 列索引
  final int amountCol; // Amount 列索引

  const DraftColumns({
    required this.itemCol,
    required this.descriptionCol,
    required this.unitCol,
    required this.qtyCol,
    required this.rateCol,
    required this.amountCol,
  });

  @override
  List<Object?> get props =>
      [itemCol, descriptionCol, unitCol, qtyCol, rateCol, amountCol];
}

/// 匹配服务结果
class MatchServiceResult extends Equatable {
  final bool success;
  final String message;
  final int matchedCount;
  final int totalCount;
  final List<String> logs;

  const MatchServiceResult({
    required this.success,
    required this.message,
    required this.matchedCount,
    required this.totalCount,
    required this.logs,
  });

  @override
  List<Object?> get props => [success, message, matchedCount, totalCount, logs];
}
