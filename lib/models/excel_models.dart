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
