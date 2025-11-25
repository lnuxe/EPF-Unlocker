import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';
import 'package:flutter/foundation.dart';
import 'package:xml/xml.dart';

import '../models/excel_models.dart';

/// 将编辑后的数据写回到原始解锁版 Excel（尽量保留原有样式和多 sheet 结构）
///
/// 当前实现：
/// - 仅更新第一个工作表对应的 sheetX.xml 中已有的单元格值
/// - 其他 sheet 和样式 XML 全部保持不变
Future<void> exportWithPreservedFormat({
  required File unlockedSource,
  required ExcelSheetModel sheetModel,
  required String savePath,
}) async {
  debugPrint(
      '[xml_preserve_export] 开始从解锁文件导出，源文件=${unlockedSource.path}, 目标=$savePath');

  final bytes = await unlockedSource.readAsBytes();
  final archive = ZipDecoder().decodeBytes(bytes);

  // 简单策略：认为当前编辑的是第一个工作表 → 选择名称最小的 sheetX.xml
  ArchiveFile? targetSheetFile;
  for (final file in archive.files) {
    if (file.isFile &&
        file.name.contains('xl/worksheets/') &&
        file.name.endsWith('.xml')) {
      if (targetSheetFile == null ||
          file.name.compareTo(targetSheetFile.name) < 0) {
        targetSheetFile = file;
      }
    }
  }

  if (targetSheetFile == null) {
    throw Exception('未找到任何工作表 XML（xl/worksheets/*.xml），无法导出');
  }

  debugPrint(
      '[xml_preserve_export] 目标工作表 XML 文件=${targetSheetFile.name}, 开始更新单元格值...');

  final raw = targetSheetFile.content;
  final contentBytes = raw is List<int> ? raw : <int>[];
  final xmlStr = utf8.decode(contentBytes);
  final doc = XmlDocument.parse(xmlStr);

  final sheetDataElements = doc.findAllElements('sheetData');
  if (sheetDataElements.isEmpty) {
    throw Exception('工作表 XML 中缺少 <sheetData> 节点，无法导出');
  }
  final sheetData = sheetDataElements.first;

  // 先把所有 <c> 元素收集起来，加速查找
  final List<XmlElement> cellElements =
      sheetData.findAllElements('c').toList(growable: false);

  XmlElement? findCellElement(String ref) {
    for (final c in cellElements) {
      final rAttr = c.getAttribute('r');
      if (rAttr == ref) return c;
    }
    return null;
  }

  for (final row in sheetModel.rows) {
    for (final cell in row.cells) {
      final value = cell.value;
      if (value == null) continue;

      final ref = _cellRef(cell.row, cell.column);
      final cElement = findCellElement(ref);
      if (cElement == null) {
        // 目前只更新已存在的单元格，避免破坏复杂结构
        continue;
      }

      // 查找或创建 <v> 节点
      XmlElement? vElement;
      for (final child in cElement.children) {
        if (child is XmlElement && child.name.local == 'v') {
          vElement = child;
          break;
        }
      }

      if (vElement != null) {
        vElement.children
          ..clear()
          ..add(XmlText(value));
      } else {
        cElement.children
            .add(XmlElement(XmlName('v'), const [], [XmlText(value)]));
      }
    }
  }

  final updatedXml = doc.toXmlString();
  final updatedBytes = utf8.encode(updatedXml);

  // 重新构建归档，替换目标 sheet 文件，其他文件保持不变
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

  final outFile = File(savePath);
  await outFile.create(recursive: true);
  await outFile.writeAsBytes(outBytes, flush: true);

  debugPrint('[xml_preserve_export] 导出完成，保留原始格式与多 sheet 结构，输出文件=$savePath');
}

String _cellRef(int rowIndexZeroBased, int columnIndexZeroBased) {
  final colLetters = _columnIndexToLetters(columnIndexZeroBased);
  final rowNumber = rowIndexZeroBased + 1;
  return '$colLetters$rowNumber';
}

String _columnIndexToLetters(int indexZeroBased) {
  var n = indexZeroBased + 1; // 列从 1 开始
  final buffer = StringBuffer();
  while (n > 0) {
    n -= 1;
    buffer.writeCharCode('A'.codeUnitAt(0) + (n % 26));
    n ~/= 26;
  }
  final str = buffer.toString();
  return String.fromCharCodes(str.codeUnits.reversed);
}
