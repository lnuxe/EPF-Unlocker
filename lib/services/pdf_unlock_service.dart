import 'dart:async';
import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:syncfusion_flutter_pdf/pdf.dart';

class PdfUnlockService {
  Future<File> unlockPdf(File original) async {
    debugPrint('========================================');
    debugPrint('[PdfUnlockService] unlockPdf() called for: ${original.path}');
    debugPrint('[PdfUnlockService] 使用 Syncfusion PDF 库解锁...');

    try {
      // 读取原始 PDF 文件
      final bytes = await original.readAsBytes();
      debugPrint('[PdfUnlockService] 已读取文件，大小: ${bytes.length} 字节');

      // 加载 PDF 文档
      PdfDocument? document;
      try {
        // 尝试直接加载（可能有权限限制）
        document = PdfDocument(inputBytes: bytes);
        debugPrint('[PdfUnlockService] PDF 文档加载成功');
      } catch (e) {
        debugPrint('[PdfUnlockService] 直接加载失败，尝试使用空密码: $e');

        // 如果失败，尝试使用空密码
        try {
          document = PdfDocument(inputBytes: bytes, password: '');
          debugPrint('[PdfUnlockService] 使用空密码加载成功');
        } catch (e2) {
          debugPrint('[PdfUnlockService] 空密码也失败，文档可能有用户密码: $e2');
          throw Exception('PDF 文件可能有密码保护，无法直接解锁');
        }
      }

      // 检查文档是否加密
      final security = document.security;
      if (security.userPassword.isNotEmpty ||
          security.ownerPassword.isNotEmpty) {
        debugPrint('[PdfUnlockService] ⚠️ 检测到密码保护');
      }

      // 移除所有安全限制
      debugPrint('[PdfUnlockService] 移除安全限制...');
      security.userPassword = '';
      security.ownerPassword = '';

      // 设置所有权限为允许
      security.permissions.clear();
      security.permissions.addAll([
        PdfPermissionsFlags.print,
        PdfPermissionsFlags.editContent,
        PdfPermissionsFlags.copyContent,
        PdfPermissionsFlags.editAnnotations,
        PdfPermissionsFlags.fillFields,
        PdfPermissionsFlags.accessibilityCopyContent,
        PdfPermissionsFlags.assembleDocument,
        PdfPermissionsFlags.fullQualityPrint,
      ]);

      debugPrint('[PdfUnlockService] 权限设置完成，生成无保护的 PDF...');

      // 保存为新的 PDF（无密码保护）
      final List<int> unlockedBytes = await document.save();
      document.dispose();

      debugPrint(
          '[PdfUnlockService] 新 PDF 生成成功，大小: ${unlockedBytes.length} 字节');

      // 写入临时文件
      final tempDir = Directory.systemTemp;
      final unlockedPath =
          '${tempDir.path}/${DateTime.now().millisecondsSinceEpoch}_unlocked.pdf';
      final unlockedFile = File(unlockedPath);
      await unlockedFile.writeAsBytes(unlockedBytes, flush: true);

      debugPrint('[PdfUnlockService] ✅ PDF 解锁成功！');
      debugPrint('[PdfUnlockService] 输出文件: $unlockedPath');
      debugPrint('========================================');

      return unlockedFile;
    } catch (e, st) {
      debugPrint('[PdfUnlockService] ❌ PDF 解锁失败: $e');
      debugPrint('[PdfUnlockService] 堆栈: $st');
      debugPrint('========================================');

      // 提供友好的错误信息
      if (e.toString().contains('password') ||
          e.toString().contains('encrypted')) {
        throw Exception('PDF 文件有密码保护，无法自动解锁。\n'
            '请先使用正确的密码打开文件，或使用其他工具移除密码。');
      }

      throw Exception('PDF 解锁失败: $e');
    }
  }
}
