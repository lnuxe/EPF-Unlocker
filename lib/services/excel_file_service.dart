import 'dart:convert';
import 'dart:io';
import 'package:archive/archive.dart';
import 'package:file_picker/file_picker.dart';
import 'package:file_selector/file_selector.dart';
import 'package:flutter/foundation.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:xml/xml.dart';
import 'pdf_unlock_service.dart';

/// .xls 文件不支持异常
/// 这个异常会被 UI 层捕获并显示弹窗提示
class XlsFormatNotSupportedException implements Exception {
  final String message;
  XlsFormatNotSupportedException(this.message);
  @override
  String toString() => message;
}

/// 负责文件选择、解压 .xlsx/PDF、删除保护、重打包为解锁版文件
class ExcelFileService {
  final _pdfUnlockService = PdfUnlockService();

  /// 让用户选择一个本地的文件（支持 .xlsx、.pdf）
  Future<File?> pickFile() async {
    // 在 Android 上先请求权限
    if (Platform.isAndroid) {
      try {
        if (await Permission.storage.isGranted) {
          // 已有权限
        } else {
          final status = await Permission.storage.request();
          if (!status.isGranted) {
            debugPrint('[ExcelFileService] 用户拒绝了存储权限');
            return null;
          }
        }
      } catch (e) {
        debugPrint('[ExcelFileService] 权限请求失败: $e');
        // 继续执行，某些情况下 file_picker 可能仍能工作
      }
    }
    debugPrint(
      '[ExcelFileService] pickFile() called. '
      'platform=${Platform.operatingSystem} ${Platform.version}',
    );

    try {
      // macOS: 使用 file_selector（系统原生对话框）
      if (Platform.isMacOS) {
        debugPrint(
            '[ExcelFileService] macOS detected, using file_selector.openFile(...)');
        const typeGroup = XTypeGroup(
          label: 'Excel/PDF',
          extensions: <String>['xlsx', 'pdf'],
        );
        final XFile? xfile = await openFile(
          acceptedTypeGroups: const [typeGroup],
        );
        debugPrint(
          '[ExcelFileService] file_selector.openFile returned. '
          'selectedPath=${xfile?.path ?? 'null (user cancelled or dialog failed)'}',
        );
        if (xfile == null) {
          return null;
        }
        return File(xfile.path);
      }

      // 其他平台：使用 file_picker
      debugPrint(
          '[ExcelFileService] non-macOS, using FilePicker.platform.pickFiles(...)');
      final result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx', 'pdf'],
        allowMultiple: false,
        withData: false,
      );
      debugPrint(
        '[ExcelFileService] FilePicker.pickFiles returned. '
        'selectedPath=${result?.files.single.path ?? 'null (user cancelled or dialog failed)'}',
      );

      if (result == null ||
          result.files.isEmpty ||
          result.files.single.path == null) {
        return null;
      }
      return File(result.files.single.path!);
    } catch (e, st) {
      debugPrint('[ExcelFileService] openFile ERROR: $e');
      debugPrint('[ExcelFileService] stack: $st');
      rethrow;
    }
  }

  /// 向后兼容：保留 pickExcelFile() 方法，内部调用 pickFile()
  @Deprecated('Use pickFile() instead')
  Future<File?> pickExcelFile() async {
    return pickFile();
  }

  /// 解锁工作表保护：删除所有 sheet*.xml 中的 <sheetProtection> 标签
  ///
  /// 返回一个新的、临时的 .xlsx 文件路径（原文件不修改）
  Future<File> unlockExcelSheet(File original) async {
    // 检查文件是否已经被解锁过
    final fileName = original.path.split('/').last.toLowerCase();
    if (fileName.contains('_unlocked')) {
      debugPrint('[ExcelFileService] 警告：文件名称中包含 "_unlocked"，该文件可能已经被解锁过');
    }

    final bytes = await original.readAsBytes();

    // 检查文件大小，如果太小可能是损坏的文件
    if (bytes.isEmpty) {
      throw Exception('文件为空或已损坏：${original.path}');
    }

    // 尝试解码 ZIP 归档
    Archive archive;
    try {
      archive = ZipDecoder().decodeBytes(bytes);
    } catch (e, st) {
      debugPrint('[ExcelFileService] ZIP 解码失败: $e');
      debugPrint('[ExcelFileService] 堆栈: $st');

      // 检查是否是常见的 ZIP 错误
      if (e.toString().contains('End of Central Directory')) {
        throw Exception('文件已损坏或格式不正确：${original.path}\n'
            '该文件可能已经被多次处理导致结构损坏。\n'
            '建议使用原始未解锁的文件，或者使用 Microsoft Excel 修复该文件后再试。');
      } else if (e.toString().contains('Invalid ZIP')) {
        throw Exception('无效的 ZIP 文件格式：${original.path}\n'
            '该文件可能不是有效的 .xlsx 文件，或者文件已损坏。');
      } else {
        throw Exception('无法读取文件：${original.path}\n'
            '错误信息：${e.toString()}\n'
            '文件可能已损坏，请检查文件完整性。');
      }
    }

    // Archive.files 在当前版本中是不可修改的列表，因此这里构建一个新的 Archive
    final updatedArchive = Archive();

    for (final archivedFile in archive.files) {
      if (archivedFile.isFile &&
          archivedFile.name.contains('xl/worksheets/') &&
          archivedFile.name.endsWith('.xml')) {
        // 触发解压，获取工作表 XML 内容
        final raw = archivedFile.content;
        final contentBytes = raw is List<int> ? raw : <int>[];
        final xmlStr = utf8.decode(contentBytes);

        final doc = XmlDocument.parse(xmlStr);

        final toRemove = doc.findAllElements('sheetProtection').toList();
        for (final element in toRemove) {
          element.parent?.children.remove(element);
        }

        final updatedXml = doc.toXmlString();
        final updatedBytes = utf8.encode(updatedXml);

        updatedArchive.addFile(
          ArchiveFile(archivedFile.name, updatedBytes.length, updatedBytes),
        );
      } else {
        // 其他文件/目录原样放入新的 Archive
        updatedArchive.addFile(archivedFile);
      }
    }

    final encoder = ZipEncoder();
    final unlockedBytes = encoder.encode(updatedArchive) ?? <int>[];

    final tempDir = Directory.systemTemp;
    final unlockedPath =
        '${tempDir.path}/${DateTime.now().millisecondsSinceEpoch}_unlocked.xlsx';
    final unlockedFile = File(unlockedPath);
    await unlockedFile.writeAsBytes(unlockedBytes, flush: true);
    return unlockedFile;
  }

  Future<File> unlockPdfFile(File original) async {
    debugPrint('========================================');
    debugPrint(
        '[ExcelFileService] unlockPdfFile() called for: ${original.path}');
    debugPrint('[ExcelFileService] 使用改进的 PDF 解锁服务...');

    try {
      // 使用新的 PdfUnlockService，它使用更可靠的方法
      final result = await _pdfUnlockService.unlockPdf(original);
      debugPrint('[ExcelFileService] PDF 解锁成功！');
      debugPrint('========================================');
      return result;
    } catch (e, st) {
      debugPrint('[ExcelFileService]  PDF 解锁失败: $e');
      debugPrint('[ExcelFileService] 堆栈: $st');
      debugPrint('========================================');
      rethrow;
    }
  }

  /// 统一解锁接口：根据文件扩展名自动选择对应的解锁方法
  /// 支持的文件格式：.xlsx, .pdf
  Future<File> unlockFile(File original) async {
    final extension = original.path.toLowerCase().split('.').last;
    debugPrint(
        '[ExcelFileService] unlockFile() called for extension: $extension');

    switch (extension) {
      case 'xlsx':
        return unlockExcelSheet(original);
      case 'xls':
        // 检测是否是实际的 .xlsx 格式但扩展名为 .xls
        final bytes = await original.readAsBytes();
        if (bytes.length >= 4) {
          final header = bytes.take(4).toList();
          if (header[0] == 0x50 && header[1] == 0x4B) {
            // 这是 ZIP 文件（.xlsx 格式），但扩展名是 .xls
            debugPrint(
                '[ExcelFileService] 检测到文件实际上是 .xlsx 格式，自动使用 .xlsx 解锁方法...');
            return unlockExcelSheet(original);
          }
        }
        // 是真正的 .xls 文件，不支持
        throw XlsFormatNotSupportedException('不支持 .xls 格式文件：${original.path}\n'
            '建议：使用 Microsoft Excel 打开该文件并另存为标准的 .xlsx 格式。');
      case 'pdf':
        return unlockPdfFile(original);
      default:
        throw UnsupportedError('不支持的文件格式: .$extension（仅支持 .xlsx, .pdf）');
    }
  }
}
