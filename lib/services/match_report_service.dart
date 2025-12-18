import 'dart:io';
import 'dart:ui' as ui;
import 'package:flutter/foundation.dart';
import 'package:syncfusion_flutter_pdf/pdf.dart';
import '../models/excel_models.dart';

/// 匹配报告生成服务：生成PDF格式的匹配结果比对报告
/// 采用SOP（Standard Operating Procedure）原则，提供标准化的报告格式
class MatchReportService {
  /// 生成匹配结果比对报告PDF
  ///
  /// [outputPath] 输出PDF文件路径
  /// [matchResults] 匹配结果列表
  /// [logs] 匹配日志列表
  /// [targetSheetName] 目标工作表名称
  /// [draftSheetName] 草稿工作表名称
  Future<File> generateMatchReport({
    required String outputPath,
    required List<MatchResult> matchResults,
    required List<String> logs,
    required String targetSheetName,
    String? draftSheetName,
  }) async {
    try {
      // 创建PDF文档
      final PdfDocument document = PdfDocument();

      // 添加A4横向页面
      // A4横向：297mm x 210mm = 1122.52点 x 794.0点（宽度1122.52点，高度794.0点）
      // 在syncfusion_flutter_pdf中，PdfPage.size是只读的，默认是A4纵向(595x842)
      // 我们需要通过调整列宽来适应页面，但页面方向无法直接修改
      // 实际上，我们可以使用更大的列宽来适应横向布局
      PdfPage currentPage = document.pages.add();
      PdfGraphics currentGraphics = currentPage.graphics;

      // 从outputPath提取文件名（不含路径和扩展名）
      final outputFile = File(outputPath);
      final fileName = outputFile.uri.pathSegments.last
          .replaceAll('_匹配报告.pdf', '')
          .replaceAll('.pdf', '');

      // 设置字体
      final PdfFont titleFont = PdfStandardFont(PdfFontFamily.helvetica, 18,
          style: PdfFontStyle.bold);
      final PdfFont headerFont = PdfStandardFont(PdfFontFamily.helvetica, 12,
          style: PdfFontStyle.bold);
      final PdfFont normalFont = PdfStandardFont(PdfFontFamily.helvetica, 10);
      final PdfFont smallFont = PdfStandardFont(PdfFontFamily.helvetica, 8);

      // 页面边距
      const double margin = 40;
      double yPosition = margin;

      // 1. 报告标题（使用文件名）
      final PdfStringFormat titleFormat =
          PdfStringFormat(alignment: PdfTextAlignment.center);
      currentGraphics.drawString(
        fileName,
        titleFont,
        format: titleFormat,
        bounds: ui.Rect.fromLTWH(0, yPosition, currentPage.size.width, 30),
      );
      yPosition += 40;

      // 2. 报告信息
      final String reportDate = DateTime.now().toString().substring(0, 19);
      currentGraphics.drawString(
        '生成时间: $reportDate',
        smallFont,
        bounds: ui.Rect.fromLTWH(
            margin, yPosition, currentPage.size.width - 2 * margin, 15),
      );
      yPosition += 20;

      currentGraphics.drawString(
        '目标工作表: $targetSheetName',
        normalFont,
        bounds: ui.Rect.fromLTWH(
            margin, yPosition, currentPage.size.width - 2 * margin, 15),
      );
      yPosition += 15;

      if (draftSheetName != null) {
        currentGraphics.drawString(
          '草稿工作表: $draftSheetName',
          normalFont,
          bounds: ui.Rect.fromLTWH(
              margin, yPosition, currentPage.size.width - 2 * margin, 15),
        );
        yPosition += 15;
      }

      // 统计信息
      final int matchedCount = matchResults.where((r) => r.matched).length;
      final int totalCount = matchResults.length;
      final double matchRate =
          totalCount > 0 ? (matchedCount / totalCount * 100) : 0;

      currentGraphics.drawString(
        '匹配统计: $matchedCount / $totalCount (${matchRate.toStringAsFixed(1)}%)',
        headerFont,
        bounds: ui.Rect.fromLTWH(
            margin, yPosition, currentPage.size.width - 2 * margin, 20),
      );
      yPosition += 30;

      // 3. Total行对比（显示在页面顶部，换行显示）
      final totalRows =
          matchResults.where((r) => r.isTotalRow && r.matched).toList();
      if (totalRows.isNotEmpty) {
        // 标题：Total行对比
        final PdfBrush titleBrush = PdfSolidBrush(PdfColor(70, 130, 180));
        currentGraphics.drawRectangle(
          brush: titleBrush,
          bounds: ui.Rect.fromLTWH(
              margin, yPosition, currentPage.size.width - 2 * margin, 30),
        );
        final PdfBrush titleTextBrush = PdfSolidBrush(PdfColor(255, 255, 255));
        currentGraphics.drawString(
          'Total行对比',
          headerFont,
          brush: titleTextBrush,
          format: PdfStringFormat(alignment: PdfTextAlignment.center),
          bounds: ui.Rect.fromLTWH(
              margin, yPosition + 7, currentPage.size.width - 2 * margin, 20),
        );
        yPosition += 35;

        // 绘制Total行对比表格（换行显示：第一行Item+描述，第二行值对比）
        final tableWidth = currentPage.size.width - 2 * margin;
        const double rowHeight = 50; // 每行高度（包含两行内容）

        // 数据行（每个Total行显示为两行）
        for (int i = 0; i < totalRows.length; i++) {
          final result = totalRows[i];
          final rowBrush = i % 2 == 0
              ? PdfSolidBrush(PdfColor(240, 240, 240))
              : PdfSolidBrush(PdfColor(255, 255, 255));

          // 绘制行背景
          currentGraphics.drawRectangle(
            brush: rowBrush,
            bounds: ui.Rect.fromLTWH(margin, yPosition, tableWidth, rowHeight),
          );

          final PdfBrush textBrush = PdfSolidBrush(PdfColor(0, 0, 0));
          final PdfStringFormat leftFormat =
              PdfStringFormat(alignment: PdfTextAlignment.left);
          final PdfStringFormat centerFormat =
              PdfStringFormat(alignment: PdfTextAlignment.center);

          // 第一行：序号 + Item + Description
          double xPos = margin;
          // 序号
          currentGraphics.drawString('${i + 1}', normalFont,
              brush: textBrush,
              format: centerFormat,
              bounds: ui.Rect.fromLTWH(xPos, yPosition + 5, 50, 15));
          xPos += 60;

          // Item和Description（分开显示，更清晰）
          final String itemText =
              result.target.item.isNotEmpty ? result.target.item : 'N/A';
          final String descText = result.target.description.isNotEmpty
              ? result.target.description
              : 'N/A';

          // Item
          currentGraphics.drawString('Item: $itemText', normalFont,
              brush: textBrush,
              format: leftFormat,
              bounds: ui.Rect.fromLTWH(xPos, yPosition + 5,
                  (tableWidth - xPos + margin) / 2 - 10, 15));

          // Description
          final double descXPos = xPos + (tableWidth - xPos + margin) / 2;
          currentGraphics.drawString('描述: $descText', normalFont,
              brush: textBrush,
              format: leftFormat,
              bounds: ui.Rect.fromLTWH(descXPos, yPosition + 5,
                  (tableWidth - xPos + margin) / 2 - 10, 15));

          // 第二行：草稿值和计算值对比
          xPos = margin + 60;
          final String draftLabel = '草稿值 (HK\$):';
          final String draftValue = result.draftTotalAmount != null
              ? result.draftTotalAmount!.toStringAsFixed(2)
              : 'N/A';
          currentGraphics.drawString(draftLabel, normalFont,
              brush: textBrush,
              format: leftFormat,
              bounds: ui.Rect.fromLTWH(xPos, yPosition + 25, 100, 15));
          xPos += 110;
          currentGraphics.drawString(draftValue, normalFont,
              brush: textBrush,
              format: leftFormat,
              bounds: ui.Rect.fromLTWH(xPos, yPosition + 25, 120, 15));

          xPos += 140;
          final String calcLabel = '计算值 (HK\$):';
          final String calcValue = result.calculatedTotal != null
              ? result.calculatedTotal!.toStringAsFixed(2)
              : 'N/A';
          currentGraphics.drawString(calcLabel, normalFont,
              brush: textBrush,
              format: leftFormat,
              bounds: ui.Rect.fromLTWH(xPos, yPosition + 25, 100, 15));
          xPos += 110;
          // 计算值（如果与草稿值不一致，显示红色 - 这是写入的值）
          final PdfBrush valueBrush = (result.draftTotalAmount != null &&
                  result.calculatedTotal != null &&
                  (result.draftTotalAmount! - result.calculatedTotal!).abs() >
                      0.01)
              ? PdfSolidBrush(PdfColor(255, 0, 0)) // 不一致时显示红色（这是写入的值）
              : textBrush;
          currentGraphics.drawString(calcValue, normalFont,
              brush: valueBrush,
              format: leftFormat,
              bounds: ui.Rect.fromLTWH(xPos, yPosition + 25, 120, 15));

          yPosition += rowHeight + 5; // 行间距
        }

        yPosition += 20;
      }

      // 4. 匹配失败行（显示在Total行对比之后，使用相同格式）
      final unmatchedRows = matchResults.where((r) => !r.matched).toList();
      if (unmatchedRows.isNotEmpty) {
        // 标题：匹配失败行
        final PdfBrush titleBrush = PdfSolidBrush(PdfColor(200, 50, 50));
        currentGraphics.drawRectangle(
          brush: titleBrush,
          bounds: ui.Rect.fromLTWH(
              margin, yPosition, currentPage.size.width - 2 * margin, 30),
        );
        final PdfBrush titleTextBrush = PdfSolidBrush(PdfColor(255, 255, 255));
        currentGraphics.drawString(
          '匹配失败行',
          headerFont,
          brush: titleTextBrush,
          format: PdfStringFormat(alignment: PdfTextAlignment.center),
          bounds: ui.Rect.fromLTWH(
              margin, yPosition + 7, currentPage.size.width - 2 * margin, 20),
        );
        yPosition += 35;

        // 绘制匹配失败行表格（使用与Total行相同的格式，显示Item和描述）
        final tableWidth = currentPage.size.width - 2 * margin;
        const double rowHeight = 50; // 每行高度（与Total行一致，包含两行内容）

        // 数据行（每个匹配失败行显示为两行：第一行Item+描述，第二行状态说明）
        for (int i = 0; i < unmatchedRows.length; i++) {
          final result = unmatchedRows[i];
          final rowBrush = i % 2 == 0
              ? PdfSolidBrush(PdfColor(255, 240, 240)) // 浅红色背景
              : PdfSolidBrush(PdfColor(255, 250, 250));

          // 绘制行背景
          currentGraphics.drawRectangle(
            brush: rowBrush,
            bounds: ui.Rect.fromLTWH(margin, yPosition, tableWidth, rowHeight),
          );

          final PdfBrush textBrush = PdfSolidBrush(PdfColor(0, 0, 0));
          final PdfStringFormat leftFormat =
              PdfStringFormat(alignment: PdfTextAlignment.left);
          final PdfStringFormat centerFormat =
              PdfStringFormat(alignment: PdfTextAlignment.center);

          // 第一行：序号 + Item + Description
          double xPos = margin;
          // 序号
          currentGraphics.drawString('${i + 1}', normalFont,
              brush: textBrush,
              format: centerFormat,
              bounds: ui.Rect.fromLTWH(xPos, yPosition + 5, 50, 15));
          xPos += 60;

          // Item和Description（分开显示，更清晰）
          final String itemText =
              result.target.item.isNotEmpty ? result.target.item : 'N/A';
          final String descText = result.target.description.isNotEmpty
              ? result.target.description
              : 'N/A';

          // Item
          currentGraphics.drawString('Item: $itemText', normalFont,
              brush: textBrush,
              format: leftFormat,
              bounds: ui.Rect.fromLTWH(xPos, yPosition + 5,
                  (tableWidth - xPos + margin) / 2 - 10, 15));

          // Description
          final double descXPos = xPos + (tableWidth - xPos + margin) / 2;
          currentGraphics.drawString('描述: $descText', normalFont,
              brush: textBrush,
              format: leftFormat,
              bounds: ui.Rect.fromLTWH(descXPos, yPosition + 5,
                  (tableWidth - xPos + margin) / 2 - 10, 15));

          // 第二行：状态说明
          xPos = margin + 60;
          const String statusText = '状态: 草稿文件中不存在匹配值';
          currentGraphics.drawString(statusText, normalFont,
              brush: textBrush,
              format: leftFormat,
              bounds: ui.Rect.fromLTWH(
                  xPos, yPosition + 25, tableWidth - xPos + margin - 10, 15));

          yPosition += rowHeight + 5; // 行间距
        }

        yPosition += 20;
      }

      // 保存PDF
      final List<int> bytes = await document.save();
      document.dispose();

      final File pdfFile = File(outputPath);
      await pdfFile.writeAsBytes(bytes, flush: true);

      debugPrint('[MatchReportService] PDF报告已生成: $outputPath');
      return pdfFile;
    } catch (e, st) {
      debugPrint('[MatchReportService] 生成PDF报告失败: $e');
      debugPrint('[MatchReportService] 堆栈: $st');
      rethrow;
    }
  }
}
